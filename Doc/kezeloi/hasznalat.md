- **Típus:** kezelői
- **Alkalmazás:** Oklevel
- **Verzió:** 1.31
- **Dátum:** 2026-09-14
- **Állapot:** érvényes
- **Felelős:** KaveZoo

# Oklevél nyomtatás

A program Excel-listából készít okleveleket egy PDF fájlba. Minden kijelölt versenyző külön oldalon jelenik meg. A szöveg az A4-es lap alsó felére kerül; a felső fél a grafikai rész, oda a program nem nyomtat.

## Előkészítés

Az Excel első sorában a mezőnevek legyenek:

Helyezés | Név | Névelő | Klub | Időeredmény

A Névelő oszlop a klubnév előtti szócska (például „a” vagy „az”); Excel-képlet is lehet, a program a kiszámolt értéket használja. A 2. sortól az adatok következnek.

Az időeredmény az oklevélen számként jelenik meg: `01:23:12` → `1:23:12`; `00:03:42` → `3:42`. Óra nélkül nincs vezető nulla a perceknél.

## Ablak fülei

1. **Alapadatok** – az oklevélen megjelenő közös szövegek (verseny címe, klubtoldalék, helyezésszöveg, dátum, aláírók).
2. **Nyomtatandók** – az importált fájl útvonala és az Excel sorai táblázatban. Az első oszlop a kijelölés: `x` = nyomtatandó. Dupla kattintással az `x` be- és kikapcsolható. A **Mind kijelöl** és **Kijelölés törlése** gombok az összes sorra hatnak.
3. **Beállítások – elrendezés** / **Beállítások – idő / aláírás** – a PDF pozíciói, betűméretei és félkövér kapcsolói. A számos mezőknél a `-` / `+` gombok léptetnek; kézzel is beírható az érték. A változtatás azonnal érvényes a következő nyomtatásra. A **Beállítások mentése** az `oklevel.ini` fájlba írja őket.

A közös szövegek a program bezárásakor is eltárolódnak.

## Oklevelek készítése

1. A Nyomtatandók fülön a `...` gombbal ki kell választani az `.xlsx` fájlt.
2. Ugyanott ellenőrizni (és szükség szerint módosítani) a kijelöléseket. Indításkor minden sor ki van jelölve.
3. A szövegeket az Alapadatok fülön a mintának megfelelően kell kitölteni.
4. A `PDF létrehozása` gomb elindítja a nyomtatást. Csak a kijelölt (`x`) sorok kerülnek a PDF-be. A kész fájl az Excel fájllal azonos mappába kerül, `oklevel.pdf` néven.
5. Ha ilyen nevű PDF már létezik, a program nem írja felül: a név végére számlálót tesz, például `oklevel (2).pdf`.
6. A mentés helye és a fájlnév a végén üzenetablakban megjelenik. A PDF az alapértelmezett nézőben (pl. Adobe Reader) megnyílik.
