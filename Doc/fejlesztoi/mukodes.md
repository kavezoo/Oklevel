- **Típus:** fejlesztői
- **Alkalmazás:** Oklevel
- **Verzió:** 1.31
- **Dátum:** 2026-09-14
- **Állapot:** érvényes
- **Felelős:** KaveZoo

# Oklevél nyomtatás – működés

Harbour MiniGUI 3.5 asztali program. Egy Excel `.xlsx` listából egyetlen PDF-et készít: a kijelölt adatsorokból A4-es oldalakat (körlevél). Programverzió: `OKLEVEL_VERZIO` (`Main.Prg`), az ablak címsorában `v` előtaggal jelenik meg.

## Forrásfájlok

| Fájl | Szerep |
|------|--------|
| `Main.Prg` | belépési pont, INI, xlsx olvasás, grid kijelölés, PDF írás, beállításmezők |
| `Main.Fmg` | főablak Tabokkal (HMG-IDE UNICODE form) |
| `oklevel.hbp` | hbmk2 projekt |
| `oklevel.ini` | űrlapmezők és nyomtatási elrendezés (nem a forrás része) |

## Ablak és Tabok

Ablakméret: 1280×880. A `Tab_Main` négy fület tartalmaz; a PDF gomb és a progress bar a Tab alatt marad.

| Fül | Tartalom |
|-----|----------|
| Alapadatok | szerkeszthető oklevélszövegek (mintanézet) |
| Nyomtatandók | import fájl + `Grid_Lista`: Excel sorok; 1. oszlop kijelölés (`x` / üres) |
| Beállítások – elrendezés | általános + cím/klub/név/vonal/helyezés |
| Beállítások – idő / aláírás | idő, dátum, aláírás vonalak és szövegek |

Tab-oldali vezérlők elérése: `Main.Tab_Main( n ).…`. Az Alapadatok mezői az 1. oldalon; `Label_XlsxFajl` / `Button_FajlValaszt` a 2. oldalon (`AlapCtrlPage`).

## Nyomtatandók grid

Betöltés: `Main_GridBetolt()` (init, fájlválasztás, Lista frissítése). Sorok alapból kijelöltek (`x`).

| Művelet | Viselkedés |
|---------|------------|
| Dupla kattintás | 1. oszlop `x` ↔ `""` |
| Mind kijelöl | minden sor 1. oszlopa `x` |
| Kijelölés törlése | minden sor 1. oszlopa üres |

PDF: csak azok a sorok, ahol az 1. oszlop `x` (`Grid_KijeloltRekordok`).

## Űrlapmezők (Alapadatok)

| Forrás | Mező | PDF szerep |
|--------|------|------------|
| `Text_VersenyCim` | verseny címe | 1. sor |
| `Névelő` + `Klub` + `Text_KlubToldalek` | klubsor | 2. sor |
| `Név` | versenyző | 3. sor, nagybetűs |
| `Helyezés` + `Text_HelyezesSzoveg` | helyezés | 4. sor |
| `Időeredmény` | idő | 5. sor: `1:32:35` / `3:42` |
| `Label_IdoFelirat` | „időeredménnyel” | 6. sor |
| `Text_Datum` | hely, dátum | 7. sor |
| `Text_AlairasBal` / `Text_AlairasJobb` | aláíró szervezetek | alsó két hasáb |

### `[Form]`

`ImportFajl`, `VersenyCim`, `KlubToldalek`, `HelyezesSzoveg`, `Datum`, `AlairasBal`, `AlairasJobb`. Betöltés: `ON INIT`. Mentés: `ON RELEASE`, fájlválasztás, beállításmentés és PDF előtt.

### `[Nyomtatas]`

A PDF pozíciói (mm), betűméretei (pont) és félkövér kapcsolói. A Beállítások füleken szerkeszthetők (`-` / `+` gombok a számos mezőknél). Változtatás azonnal a memóriában (`BeallitasAlkalmaz`); a **Beállítások mentése** gomb az `oklevel.ini`-be ír.

Első mező: `Szoveg_Kezdo_Ymm` (felső kezdő margó, alap 152).

| Kulcs | Jelentés | Alap |
|-------|----------|------|
| `Betutipus` | betűtípus neve | Arial |
| `Oldal_Bal_Mm` / `Oldal_Jobb_Mm` | középre igazítás szélei | 18 / 192 |
| `Szoveg_Kezdo_Ymm` | szövegblokk teteje | 152 |
| `Cim_*` … `Datum_*` | szövegsorok Y / magasság / méret / félkövér | lásd `NyomtatasDefaults` |
| `NevAlattiVonal_*` / `Helyezes_*` | vonal és helyezés | |
| `Alairas*` | aláírás vonalak és szövegek | |

A `*_Ymm` értékek a `Szoveg_Kezdo_Ymm`-hez viszonyított eltolások (kivéve ahol abszolút). Logikai (`*_Felbe`): checkbox / ini `1`/`0`.

## Excel

Az első munkalap (`xl/worksheets/sheet1.xml`) olvasható Excel nélkül. Kötelező oszlopok: `Helyezés`, `Név`, `Névelő`, `Klub`, `Időeredmény`. Adat a 2. sortól.

## PDF

HMG HPDF, A4 álló. Szöveg: vektoros `HPDFPRINT` + `hupdf()` (`HB_UTF8TOSTR` → HUWIN), encoding `CP1250`. Egy dokumentum, kijelölt rekordonként egy oldal. Kimenet: `oklevel.pdf` (ütközésnél `oklevel (2).pdf` …). Mentés után megnyílik az alapértelmezett nézőben.

## Fordítás

HMG 3.5: `c:\hmg.3.5\build.bat /n oklevel.hbp`
