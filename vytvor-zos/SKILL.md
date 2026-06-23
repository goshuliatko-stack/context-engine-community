---
name: vytvor-zos
description: "Vytvor Zmluvu o sprostredkovaní (ZoS) na predaj nehnuteľnosti ako hotový .docx vo Worde. Použi keď používateľ povie: priprav/vytvor/sprav mi zmluvu o sprostredkovaní, ZoS, sprostredkovateľskú zmluvu, zmluvu na byt/dom/pozemok, 'klient mi posiela peniaze na byt XY'. Skill nazbiera dáta (Zmluvy RAG, Vzorové zmluvy na Google Drive, LV výpisy, Context Engine), vyberie správny vzor podľa toku platby, vyplní známe údaje, neznáme vyznačí ČERVENO ako [DOPLNIŤ], zachová písmo/štýl vzoru, vygeneruje .docx cez python-docx, pošle ho do chatu a uloží do cieľovej Drive zložky."
---

# Vytvor ZoS — Zmluva o sprostredkovaní (behavioral guide)

Tento skill vyrobí kompletnú **Zmluvu o sprostredkovaní** na predaj nehnuteľnosti ako .docx súbor, štýlovo podľa vzoru, so známymi údajmi vyplnenými a neznámymi vyznačenými červeno. Sprostredkovateľom je štandardne **Venoc s. r. o.** (Georgi Conev).

> Filozofia: **nikdy nehádaj právny text** — vždy ho čerpaj zo vzoru. **Nikdy nevymýšľaj osobné údaje** — čo nevieš, vyznač červeno `[DOPLNIŤ]`. Radšej viac červených polí ako jeden tichý nesprávny údaj.

---

## 0. ČO POTREBUJEŠ OD POUŽÍVATEĽA (zozbieraj na začiatku)

Ak niečo z toho chýba a nedá sa dohľadať, **opýtaj sa** alebo vyznač červeno:

| Údaj | Príklad | Kde hľadať ak nepovie |
|------|---------|------------------------|
| Predmet (nehnuteľnosť) | „byt Schmidtovcov", „2i M. Bella" | LV výpisy, Context Engine projekt |
| Cena | 149 990 € | musí povedať používateľ |
| Provízia | 5,23 % | musí povedať používateľ |
| Doba platnosti | 150 dní od podpisu | musí povedať používateľ |
| Tok platby | „klient posiela peniaze mne" → vyber správny vzor | určuje výber vzoru |
| Vlastník/-ci (Záujemca) | Tomáš Schmidt + Natália Schmidtová | LV, Context Engine |
| Cieľová zložka | `G:\Môj disk\Goso Nehnuteľnosti\1 - 2i M.Bella - Schmidt` | Google Drive search |

**Sprostredkovateľ je štandardne Venoc s. r. o.** — ak používateľ povie inak, použi jeho zadanie.

---

## 1. ZBER DÁT — poradie zdrojov

### 1.1 Právny text + údaje sprostredkovateľa → vzor zmluvy

Postupuj v tomto poradí (každý ďalší je fallback pri zlyhaní predošlého):

```
1. Zmluvy RAG (MCP Zmluvy → search_knowledge_base / get_full_document)
   ├── Funguje → nájdi vzor "Zmluva o sprostredkovaní"
   └── TIMEOUT / Railway down (časté!) → fallback na Google Drive
2. Google Drive "Vzorové zmluvy"
   └── search_files: title contains 'Vzorové zmluvy' AND mimeType folder
       → "Zmluvy o sprostredkovaní"
       → vyber podzložku/súbor podľa toku platby (pozri 1.2)
```

⚠️ **Railway backend pre Zmluvy RAG často timeoutuje** (60 s). Nebojuj s tým donekonečna — po 1–2 pokusoch prejdi na Google Drive „Vzorové zmluvy". Ak používateľ povie že vzory sú na konkrétnej Drive ceste, choď rovno tam.

### 1.2 Výber správneho vzoru podľa toku platby

Vzorov je viac, líšia sa **kto a kedy drží peniaze**:
- **„klient posiela peniaze mne (sprostredkovateľovi)"** → vzor *„...kde po rezervácii mne klient posiela peniaze"* — tu Záujemca prijme platbu od kupujúceho a do 3 dní ju prepošle sprostredkovateľovi.
- iný tok → vyber zodpovedajúci vzor.

Hierarchia na Drive (príklad reálnej cesty):
```
Vzorové zmluvy
 └─ Zmluvy o sprostredkovaní
     └─ Zmluva o sprostredkovaní kde po rezervácií mne klient posiela peniaze
         └─ <konkrétny .docx vzor>
```

### 1.3 ⚠️ Ako čítať obsah vzoru — KRITICKÉ

- **POUŽI `mcp__Google_Drive__read_file_content`** (textová reprezentácia) — spoľahlivé.
- **NEPOUŽÍVAJ `download_file_content`** (base64) na manuálne prepisovanie — base64 .docx má desaťtisíce znakov a ručný prepis sa **ticho odsekne/poškodí**. Toto zlyhalo v minulosti.
- Z vzoru si vytiahni: celý právny text (všetky články), údaje sprostredkovateľa (Venoc s. r. o.), prílohu s marketingovými službami.

### 1.4 Údaje o osobách a nehnuteľnosti

- **Osoby (Záujemcovia):** Context Engine (`ctx_person`, `ctx_context`) + LV výpisy. Pozor na aktuálnosť — napr. priezvisko po svadbe (LV môže mať rodné priezvisko). Vždy over u používateľa pri rozpore.
- **Nehnuteľnosť:** z LV výpisov (číslo bytu, súpisné č., LV č., k. ú., parcela, výmera, podiely, titul nadobudnutia, ťarchy/záložné práva).
- LV výpisy bývajú nahrané priamo ako prílohy alebo na Drive. Disclaimer „Výpis je nepoužiteľný na právne úkony" → preto aktuálny stav ťarchy/úveru ku dňu podpisu vyznač červeno `[DOPLNIŤ]`.

### 1.5 Stabilné údaje sprostredkovateľa (Venoc s. r. o.)

Ak ich nevieš dohľadať vo vzore, použi tieto (over ak je pochybnosť):
```
Obchodné meno: Venoc s. r. o.
OR: Okresný súd Žilina, oddiel Sro, vl. č. 79531/L
IČO: 54 555 680
Sídlo: Gabajova 2593/22, 010 01 Žilina
Č. účtu: SK56 8330 0000 0026 0217 8024
Konateľ: Georgi Conev
Mail: conevreality@gmail.com
Tel.: +421 948 880 069
Web: www.georgiconev.sk
```

---

## 2. ZOSTAVENIE .DOCX — python-docx

**Negeneruj zmluvu „od oka" v texte.** Postav .docx programovo cez `python-docx`, aby si mal plnú kontrolu nad písmom, veľkosťou, kurzívou, tučným a **červenou farbou pre `[DOPLNIŤ]`**.

```bash
pip install python-docx   # ak nie je nainštalované
```

V tomto adresári je pripravený **`build_zos_template.py`** — skopíruj ho do `/tmp/contract_work/` (alebo scratchpadu), uprav dáta a spusti. Generuje celú zmluvu s pomocnými funkciami `add_p / heading / body` a značkami:
- `R("text")` → červený run (placeholder `[DOPLNIŤ]`)
- `N("text")` → normálny čierny run

### 2.1 Štýl podľa vzoru
- Písmo: **Times New Roman 12 pt** (alebo presné písmo vzoru, ak sa líši).
- Nadpis zmluvy tučný, centrovaný; podnadpis (citácia § 774 Obč. zákonníka) kurzíva, centrovaná.
- Nadpisy článkov tučné, centrované; telo zarovnané do bloku (justify).
- Okraje cca 2–2,5 cm.

### 2.2 Štruktúra (6 článkov + podpisy + príloha)
1. **Čl. I — Zmluvné strany** (Záujemca/-ovia + Sprostredkovateľ)
2. **Čl. II — Úvodné ustanovenia** (popis nehnuteľnosti z LV, titul nadobudnutia, ťarchy)
3. **Čl. III — Predmet zmluvy** (záväzky, exkluzivita, 10 % zmluvná pokuta, „Ponuka" / marketing)
4. **Čl. IV — Platobné podmienky** (cena slovom, provízia %, aukcia +30 %, tok platby podľa vzoru)
5. **Čl. V — Doba trvania** (na dobu určitú X dní od uzatvorenia)
6. **Čl. VI — Spoločné a záverečné ustanovenia**
7. **Podpisový blok** (V Žiline, dňa `[DOPLNIŤ DÁTUM]` červeno; podpisové čiary pre každého Záujemcu a Sprostredkovateľa)
8. **Príloha** — „Príklad činností vyvíjaných pre Záujemcu/-ov" (marketingové služby zo vzoru)

### 2.3 Cena slovom
Cenu vždy uveď číslom aj slovom, napr.:
`149 990,- EUR (slovom: jednostoštyridsaťdeväťtisícdeväťstodeväťdesiat EUR)`

---

## 3. JEDEN vs. VIAC VLASTNÍKOV — gramatika

Vzory bývajú písané pre **jedného** Záujemcu (jednotné číslo). Pri **viacerých spoluvlastníkoch** preklop celý text do množného čísla so správnym skloňovaním:

| pád | jednotné | množné |
|-----|----------|--------|
| 1. | Záujemca | Záujemcovia |
| 2. | Záujemcu | Záujemcov |
| 3. | Záujemcovi | Záujemcom |
| 7. | Záujemcom | Záujemcami |

- Pri každom spoluvlastníkovi uveď **spoluvlastnícky podiel** (napr. 1/2).
- Zaveď: *„(Záujemca 1 a Záujemca 2 spolu ďalej aj ako „Záujemcovia")"*.
- Počet vyhotovení = počet strán (napr. 2 Záujemcovia + Sprostredkovateľ → **3 vyhotovenia**).
- Každý Záujemca má vlastnú podpisovú čiaru.
- **Rodné priezvisko** uveď samostatným poľom, ak sa líši od súčasného (napr. po svadbe).

---

## 4. ČERVENÉ POLIA — čo treba doplniť

Vyznač červeno (`R(...)` → `RGBColor(0xFF,0x00,0x00)`) všetko, čo s istotou nevieš. Typicky:
- Rodné číslo každého Záujemcu
- Rodinný stav (ak nie je potvrdený)
- Číslo účtu / IBAN každého Záujemcu
- Aktuálny stav ťarchy / výška zostatku úveru ku dňu podpisu
- Dátum podpisu zmluvy

Údaj, ktorý používateľ **explicitne potvrdil**, NEdávaj červeno (napr. „pani je vydatá" → rodinný stav čierno).

---

## 5. DORUČENIE VÝSLEDKU

1. **Vždy pošli .docx do chatu** cez `SendUserFile` — okamžitý náhľad pre používateľa.
2. **Skús uložiť do cieľovej Drive zložky:**
   - Nájdi zložku: `search_files` (napr. parent „Goso Nehnuteľnosti" → titul „1 - 2i M.Bella - Schmidt").
   - Upload: `mcp__Google_Drive__create_file` s `parentId`, `title`, `contentMimeType = application/vnd.google-apps.document` alebo originál docx mime, `disableConversionToGoogleType` podľa potreby, `base64Content`.
   - ⚠️ Ak by upload vyžadoval **ručné prepisovanie base64** cez nástroj naslepo, **nerob to** (riziko poškodenia). Radšej daj používateľovi súbor z chatu nech ho presunie sám, alebo použi spoľahlivý upload mechanizmus. Nikdy nenahraj potenciálne poškodený súbor do cieľovej zložky.
3. **Zhrň používateľovi**:
   - čo je vyplnené automaticky,
   - **zoznam všetkých červených `[DOPLNIŤ]` polí**, ktoré musí doplniť ručne.

---

## 6. RÝCHLY CHECKLIST

- [ ] Mám cenu, províziu %, dobu platnosti, tok platby?
- [ ] Vzor načítaný cez `read_file_content` (nie base64 prepis)?
- [ ] Sprostredkovateľ = Venoc s. r. o. (alebo podľa zadania)?
- [ ] Údaje osôb z Context Engine + LV, priezviská aktuálne?
- [ ] Nehnuteľnosť kompletne z LV (byt, pozemok, podiely, titul, ťarchy)?
- [ ] Jeden/viac vlastníkov → gramatika v správnom čísle?
- [ ] Neznáme polia ČERVENO `[DOPLNIŤ]`?
- [ ] Cena číslom aj slovom?
- [ ] Doba určitá X dní od uzatvorenia (Čl. V)?
- [ ] Počet vyhotovení = počet strán?
- [ ] .docx poslaný cez SendUserFile + pokus o upload do cieľovej zložky?
- [ ] Zhrnutie + zoznam červených polí používateľovi?
