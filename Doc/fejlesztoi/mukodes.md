- **Típus:** fejlesztői
- **Alkalmazás:** Oklevel
- **Verzió:** 1.20
- **Dátum:** 2026-09-14
- **Állapot:** érvényes
- **Felelős:** KaveZoo

# Oklevél nyomtatás – működés

Harbour MiniGUI 3.5 asztali program. Egy Excel `.xlsx` listából egyetlen PDF-et készít: minden adatsor egy A4-es oldal (körlevél). Programverzió: `OKLEVEL_VERZIO` (`Main.Prg`), az ablak címsorában `v` előtaggal jelenik meg.

## Forrásfájlok

| Fájl | Szerep |
|------|--------|
| `Main.Prg` | belépési pont, INI, xlsx olvasás, PDF írás |
| `Main.Fmg` | főablak (HMG-IDE UNICODE form) |
| `oklevel.hbp` | hbmk2 projekt |
| `oklevel.ini` | űrlapmezők és nyomtatási elrendezés (nem a forrás része) |

## Űrlapmezők

A formon szerkeszthető szövegek a PDF minden oldalán megjelennek. Az Excelből érkező mezők a rekord szerint változnak.

| Forrás | Mező | PDF szerep |
|--------|------|------------|
| `Text_VersenyCim` | verseny címe | 1. sor |
| `Névelő` + `Klub` + `Text_KlubToldalek` | klubsor | 2. sor |
| `Név` | versenyző | 3. sor, nagybetűs |
| `Helyezés` + `Text_HelyezesSzoveg` | helyezés | 4. sor, egy betűméret, középre |
| `Időeredmény` | idő | 5. sor: `1:32:35` / `3:42` (óra vezető nulla nélkül) |
| `Label_IdoFelirat` | „időeredménnyel” | 6. sor |
| `Text_Datum` | hely, dátum | 7. sor |
| `Text_AlairasBal` / `Text_AlairasJobb` | aláíró szervezetek | alsó két hasáb |
| `Label_XlsxFajl` | importált xlsx útvonal | — |
| `Button_FajlValaszt` / `Button_PdfLetrehozas` | fájlválasztás / PDF | — |

### `[Form]`

`ImportFajl`, `VersenyCim`, `KlubToldalek`, `HelyezesSzoveg`, `Datum`, `AlairasBal`, `AlairasJobb`. Betöltés: `ON INIT` (régi `XlsxFile` / `Text1`–`Text6` kulcsok is olvashatók). Mentés: `ON RELEASE`, fájlválasztás és PDF készítés előtt.

### `[Nyomtatas]`

A PDF pozíciói (mm), betűméretei (pont) és félkövér kapcsolói. Programindításkor és PDF készítés előtt beolvasódik. Mentéskor a program kiírja az aktuális (vagy alap) értékeket, így kézzel szerkeszthető.

| Kulcs | Jelentés | Alap |
|-------|----------|------|
| `Betutipus` | betűtípus neve | Arial |
| `Oldal_Bal_Mm` / `Oldal_Jobb_Mm` | középre igazítás szélei (középpont = átlag) | 18 / 192 |
| `Szoveg_Kezdo_Ymm` | szövegblokk teteje (grafika alatt) | 152 |
| `Cim_*` | 1. sor, félkövér | 0 / 10 / 16 / 1 |
| `Klub_*` | 2. sor, félkövér, közel az 1. sorhoz | 8 / 9 / 15 / 1 |
| `Nev_*` | 3. sor, félkövér | 17 / 12 / 26 / 1 |
| `NevAlattiVonal_*` | vízszintes vonal: betűalj + `Offset_Mm` | 1 mm |
| `Helyezes_VonalAlatti_Mm` | helyezés a vonal alatt, félkövér | 3 mm / 17 / 1 |
| `Ido_*` / `IdoFelirat_*` / `Datum_*` | időblokk (idő 27 pt félkövér) | 41 / 55 / 63 |
| `Alairas_VonalAlatti_Mm`, `Alairas_Betumeret` | aláírás szöveg a vonal alatt, félkövér | 1 / 12 / 1 |

A `*_Ymm` értékek a `Szoveg_Kezdo_Ymm`-hez viszonyított eltolások. Logikai (`*_Felbe`): `1` / `0`.

## Excel

Az első munkalap (`xl/worksheets/sheet1.xml`) olvasható Excel nélkül. Kötelező oszlopok: `Helyezés`, `Név`, `Névelő`, `Klub`, `Időeredmény`. Adat a 2. sortól.

## PDF

HMG HPDF, A4 álló. Szöveg: vektoros `HPDFPRINT` + `hupdf()` (`HB_UTF8TOSTR` → HUWIN; `REQUEST HB_CODEPAGE_HUWIN`), encoding `CP1250` (ISO8859-2 helyett, hogy a Ž/Š stb. helyes legyen). Egy dokumentum, rekordonként egy oldal. Kimenet: `oklevel.pdf` (ütközésnél `oklevel (2).pdf` …). Mentés után megnyílik az alapértelmezett nézőben.

## Fordítás

HMG 3.5: `c:\hmg.3.5\build.bat /n oklevel.hbp`
