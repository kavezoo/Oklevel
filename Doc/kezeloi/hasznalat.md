- **Típus:** kezelői
- **Alkalmazás:** Oklevel
- **Verzió:** 1.10
- **Dátum:** 2026-09-14
- **Állapot:** érvényes
- **Felelős:** KaveZoo

# Oklevél nyomtatás

A program Excel-listából készít okleveleket egy PDF fájlba. Minden versenyző külön oldalon jelenik meg. A szöveg az A4-es lap alsó felére kerül; a felső fél a grafikai rész, oda a program nem nyomtat.

## Előkészítés

Az Excel első sorában a mezőnevek legyenek:

Helyezés | Név | Névelő | Klub | Időeredmény

A Névelő oszlop a klubnév előtti szócska (például „a” vagy „az”); Excel-képlet is lehet, a program a kiszámolt értéket használja. A 2. sortól az adatok következnek.

Az időeredmény az oklevélen számként jelenik meg: `01:23:12` → `1:23:12`; `00:03:42` → `3:42`. Óra nélkül nincs vezető nulla a perceknél.

A képernyőn megadható szövegek (verseny címe, „csapatának tagjaként”, „helyezést ért el”, dátum, a két aláíró szervezet) a program bezárásakor eltárolódnak, és a következő indításkor visszatöltődnek.

A betűméretek, sorpozíciók és vonalak az `oklevel.ini` fájl `[Nyomtatas]` szekciójában módosíthatók. A program mentéskor kiírja a kulcsokat; a változás a következő PDF készítéskor érvényesül.

## Oklevelek készítése

1. A `...` gombbal ki kell választani az `.xlsx` fájlt. Az útvonal a felső mezőben megjelenik.
2. A szövegeket a mintának megfelelően kell kitölteni.
3. A `PDF létrehozása` gomb elindítja a nyomtatást. A kész fájl az Excel fájllal azonos mappába kerül, `oklevel.pdf` néven.
4. Ha ilyen nevű PDF már létezik, a program nem írja felül: a név végére számlálót tesz, például `oklevel (2).pdf`.
5. A mentés helye és a fájlnév a végén üzenetablakban megjelenik. A PDF az alapértelmezett nézőben (pl. Adobe Reader) megnyílik.
