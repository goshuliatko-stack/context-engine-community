"""Persistent OAuth storage for SimpleOAuthProvider.

The upstream SimpleOAuthProvider keeps registered DCR clients and issued
access tokens in memory only, so every Railway redeploy wipes them and all
connectors (claude.ai / Cowork / Claude Code) have to re-authenticate.

PersistentOAuthProvider write-throughs clients and access tokens to the
Context Engine SQLite DB (which lives on the persistent Railway volume) and
reloads them on startup, so existing connections survive redeploys.

Short-lived state (authorization codes, login state_mapping) is intentionally
NOT persisted — it only matters within a single ~5-minute login flow on one
container; after a redeploy mid-login the user simply retries.
"""

import logging
import time

from mcp_oauth.server.auth_provider.simple_auth_provider import SimpleOAuthProvider
from mcp.server.auth.provider import AccessToken
from mcp.shared.auth import OAuthClientInformationFull

from context_engine.db import get_db

logger = logging.getLogger(__name__)

_SCHEMA = """
CREATE TABLE IF NOT EXISTS oauth_clients (
    client_id  TEXT PRIMARY KEY,
    data       TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS oauth_tokens (
    token      TEXT PRIMARY KEY,
    client_id  TEXT,
    data       TEXT NOT NULL,
    expires_at INTEGER,
    created_at TEXT DEFAULT (datetime('now'))
);
"""


class PersistentOAuthProvider(SimpleOAuthProvider):
    """SimpleOAuthProvider that persists clients + tokens to the SQLite DB."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._init_store()
        self._load_from_db()

    def _init_store(self):
        try:
            with get_db() as db:
                db.executescript(_SCHEMA)
        except Exception as e:  # pragma: no cover - best effort
            logger.warning(f"OAuth store init failed: {e}")

    def _load_from_db(self):
        try:
            now = time.time()
            with get_db() as db:
                for row in db.execute("SELECT data FROM oauth_clients"):
                    try:
                        c = OAuthClientInformationFull.model_validate_json(row["data"])
                        self.clients[c.client_id] = c
                    except Exception as e:
                        logger.warning(f"OAuth store: skip client row: {e}")
                for row in db.execute(
                    "SELECT token, data, expires_at FROM oauth_tokens"
                ):
                    if row["expires_at"] and row["expires_at"] < now:
                        continue
                    try:
                        t = AccessToken.model_validate_json(row["data"])
                        self.tokens[t.token] = t
                    except Exception as e:
                        logger.warning(f"OAuth store: skip token row: {e}")
            logger.info(
                f"OAuth store loaded: {len(self.clients)} clients, "
                f"{len(self.tokens)} tokens"
            )
        except Exception as e:  # pragma: no cover - best effort
            logger.warning(f"OAuth store load failed: {e}")

    async def register_client(self, client_info: OAuthClientInformationFull):
        await super().register_client(client_info)
        try:
            with get_db() as db:
                db.execute(
                    "INSERT OR REPLACE INTO oauth_clients (client_id, data) "
                    "VALUES (?, ?)",
                    (client_info.client_id, client_info.model_dump_json()),
                )
        except Exception as e:
            logger.warning(f"OAuth store: persist client failed: {e}")

    async def exchange_authorization_code(self, client, authorization_code):
        result = await super().exchange_authorization_code(
            client, authorization_code
        )
        try:
            token_obj = self.tokens.get(result.access_token)
            if token_obj is not None:
                with get_db() as db:
                    db.execute(
                        "INSERT OR REPLACE INTO oauth_tokens "
                        "(token, client_id, data, expires_at) VALUES (?, ?, ?, ?)",
                        (
                            token_obj.token,
                            token_obj.client_id,
                            token_obj.model_dump_json(),
                            token_obj.expires_at,
                        ),
                    )
        except Exception as e:
            logger.warning(f"OAuth store: persist token failed: {e}")
        return result

    async def revoke_token(self, token: str, token_type_hint: str | None = None):
        await super().revoke_token(token, token_type_hint)
        try:
            with get_db() as db:
                db.execute("DELETE FROM oauth_tokens WHERE token = ?", (token,))
        except Exception as e:
            logger.warning(f"OAuth store: revoke persist failed: {e}")

    async def load_access_token(self, token: str) -> AccessToken | None:
        access = await super().load_access_token(token)
        if access is None:
            # Expired/unknown — drop any stale DB row so it doesn't reload.
            try:
                with get_db() as db:
                    db.execute("DELETE FROM oauth_tokens WHERE token = ?", (token,))
            except Exception:
                pass
        return access
