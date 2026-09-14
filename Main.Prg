#include <hmg.ch>

REQUEST HB_CODEPAGE_HUWIN

#define OKLEVEL_VERZIO "1.20"

STATIC s_hNyomtatas := NIL

FUNCTION Main()

   LOAD WINDOW Main
   Main.Center
   Main.Activate

RETURN NIL

*------------------------------------------------------------------------------
FUNCTION OklevelVerzio()

RETURN OKLEVEL_VERZIO

*------------------------------------------------------------------------------
FUNCTION OklevelAblakCim()

RETURN "Országos Lassúsági Hagyományőrző Piapszívóverseny oklevél nyomtató program" + ;
   " • v" + OklevelVerzio() + " • zsfoto@gmail.com"

*------------------------------------------------------------------------------
PROCEDURE Main_OnInit()

   LOCAL cXlsx

   Main.Title := OklevelAblakCim()
   Main_IniLoad()
   Main.ProgressBar_Pdf.Value := 0

   cXlsx := AllTrim( Main.Label_XlsxFajl.Value )
   IF ! ImportFajlErvenyes( cXlsx )
      Main.Label_XlsxFajl.Value := ""
      Main_SelectFile()
   ENDIF

RETURN

*------------------------------------------------------------------------------
FUNCTION ImportFajlErvenyes( cXlsx )

   LOCAL cExt

   cXlsx := AllTrim( cXlsx )
   IF Empty( cXlsx )
      RETURN .F.
   ENDIF
   IF ! hb_FileExists( cXlsx )
      RETURN .F.
   ENDIF
   cExt := Lower( hb_FNameExt( cXlsx ) )
   IF ! ( cExt == ".xlsx" .OR. cExt == "xlsx" )
      RETURN .F.
   ENDIF

RETURN .T.

*------------------------------------------------------------------------------
PROCEDURE Main_OnRelease()

   Main_IniSave()

RETURN

*------------------------------------------------------------------------------
FUNCTION Main_IniFile()

RETURN GetStartUpFolder() + "\oklevel.ini"

*------------------------------------------------------------------------------
PROCEDURE Main_IniLoad()

   LOCAL h := IniReadFile( Main_IniFile() )
   LOCAL c

   NyomtatasBetolt( h )

   IF Empty( h )
      RETURN
   ENDIF

   c := IniFirst( h, { "ImportFajl", "XlsxFile" } )
   IF ! Empty( c )
      Main.Label_XlsxFajl.Value := c
   ENDIF
   IniSetIfFirst( h, "Text_VersenyCim", { "VersenyCim", "Text1" } )
   IniSetIfFirst( h, "Text_KlubToldalek", { "KlubToldalek", "Text2" } )
   IniSetIfFirst( h, "Text_HelyezesSzoveg", { "HelyezesSzoveg", "Text3" } )
   IniSetIfFirst( h, "Text_Datum", { "Datum", "Text4" } )
   IniSetIfFirst( h, "Text_AlairasBal", { "AlairasBal", "Text5" } )
   IniSetIfFirst( h, "Text_AlairasJobb", { "AlairasJobb", "Text6" } )

RETURN

*------------------------------------------------------------------------------
FUNCTION IniFirst( h, aKeys )

   LOCAL cKey
   LOCAL cVal

   FOR EACH cKey IN aKeys
      cVal := AllTrim( hb_HGetDef( h, cKey, "" ) )
      IF ! Empty( cVal )
         RETURN cVal
      ENDIF
   NEXT

RETURN ""

*------------------------------------------------------------------------------
PROCEDURE IniSetIf( cKey, h, cCtrl )

   LOCAL cVal := hb_HGetDef( h, cKey, "" )

   IF ! Empty( cVal )
      SetProperty( "Main", cCtrl, "Value", cVal )
   ENDIF

RETURN

*------------------------------------------------------------------------------
PROCEDURE IniSetIfFirst( h, cCtrl, aKeys )

   LOCAL cVal := IniFirst( h, aKeys )

   IF ! Empty( cVal )
      SetProperty( "Main", cCtrl, "Value", cVal )
   ENDIF

RETURN

*------------------------------------------------------------------------------
PROCEDURE Main_IniSave()

   LOCAL c := "[Form]" + Chr( 13 ) + Chr( 10 )
   LOCAL aKeys
   LOCAL cKey
   LOCAL xVal
   LOCAL cVal
   LOCAL hNyom

   c += "ImportFajl=" + Main.Label_XlsxFajl.Value + Chr( 13 ) + Chr( 10 )
   c += "VersenyCim=" + Main.Text_VersenyCim.Value + Chr( 13 ) + Chr( 10 )
   c += "KlubToldalek=" + Main.Text_KlubToldalek.Value + Chr( 13 ) + Chr( 10 )
   c += "HelyezesSzoveg=" + Main.Text_HelyezesSzoveg.Value + Chr( 13 ) + Chr( 10 )
   c += "Datum=" + Main.Text_Datum.Value + Chr( 13 ) + Chr( 10 )
   c += "AlairasBal=" + Main.Text_AlairasBal.Value + Chr( 13 ) + Chr( 10 )
   c += "AlairasJobb=" + Main.Text_AlairasJobb.Value + Chr( 13 ) + Chr( 10 )
   c += Chr( 13 ) + Chr( 10 )
   c += "; PDF nyomtatási elrendezés (mm / pont). Hiányzó kulcsnál a program alapértéket használ." + Chr( 13 ) + Chr( 10 )
   c += "[Nyomtatas]" + Chr( 13 ) + Chr( 10 )

   hNyom := Nyomtatas()
   aKeys := hb_HKeys( hNyom )
   ASort( aKeys )
   FOR EACH cKey IN aKeys
      xVal := hNyom[ cKey ]
      DO CASE
      CASE ValType( xVal ) == "N"
         IF xVal == Int( xVal )
            cVal := hb_NToS( Int( xVal ) )
         ELSE
            cVal := AllTrim( Str( xVal, 12, 3 ) )
         ENDIF
      CASE ValType( xVal ) == "L"
         cVal := iif( xVal, "1", "0" )
      OTHERWISE
         cVal := hb_CStr( xVal )
      ENDCASE
      c += cKey + "=" + cVal + Chr( 13 ) + Chr( 10 )
   NEXT

   FileWriteUtf8( Main_IniFile(), c )

RETURN

*------------------------------------------------------------------------------
FUNCTION NyomtatasDefaults()

   LOCAL h := { => }

   // Általános
   hb_HSet( h, "Betutipus", "Arial" )
   hb_HSet( h, "Oldal_Bal_Mm", 18 )
   hb_HSet( h, "Oldal_Jobb_Mm", 192 )
   hb_HSet( h, "Szoveg_Kezdo_Ymm", 152 )

   // 1. sor – verseny címe
   hb_HSet( h, "Cim_Ymm", 0 )
   hb_HSet( h, "Cim_Magassag_Mm", 10 )
   hb_HSet( h, "Cim_Betumeret", 16 )
   hb_HSet( h, "Cim_Felbe", .T. )

   // 2. sor – klub (közel az elsőhöz)
   hb_HSet( h, "Klub_Ymm", 8 )
   hb_HSet( h, "Klub_Magassag_Mm", 9 )
   hb_HSet( h, "Klub_Betumeret", 15 )
   hb_HSet( h, "Klub_Felbe", .T. )

   // 3. sor – név
   hb_HSet( h, "Nev_Ymm", 17 )
   hb_HSet( h, "Nev_Magassag_Mm", 12 )
   hb_HSet( h, "Nev_Betumeret", 26 )
   hb_HSet( h, "Nev_Felbe", .T. )

   // Vonal: betűalj + Offset_Mm; helyezés: vonal + Helyezes_VonalAlatti_Mm
   hb_HSet( h, "NevAlattiVonal_Offset_Mm", 1 )
   hb_HSet( h, "NevAlattiVonal_Bal_Mm", 30 )
   hb_HSet( h, "NevAlattiVonal_Jobb_Mm", 180 )
   hb_HSet( h, "NevAlattiVonal_Vastagsag_Mm", 0.35 )
   hb_HSet( h, "Helyezes_VonalAlatti_Mm", 3 )
   hb_HSet( h, "Helyezes_Betumeret", 17 )
   hb_HSet( h, "Helyezes_Felbe", .T. )

   // 5. sor – időeredmény
   hb_HSet( h, "Ido_Ymm", 41 )
   hb_HSet( h, "Ido_Magassag_Mm", 12 )
   hb_HSet( h, "Ido_Betumeret", 27 )
   hb_HSet( h, "Ido_Felbe", .T. )

   // 6. sor – „időeredménnyel”
   hb_HSet( h, "IdoFelirat_Ymm", 55 )
   hb_HSet( h, "IdoFelirat_Magassag_Mm", 9 )
   hb_HSet( h, "IdoFelirat_Betumeret", 14 )
   hb_HSet( h, "IdoFelirat_Felbe", .F. )

   // 7. sor – dátum / hely
   hb_HSet( h, "Datum_Ymm", 63 )
   hb_HSet( h, "Datum_Magassag_Mm", 9 )
   hb_HSet( h, "Datum_Betumeret", 15 )
   hb_HSet( h, "Datum_Felbe", .T. )

   // Aláírás vonalak
   hb_HSet( h, "AlairasVonal_Ymm", 123 )
   hb_HSet( h, "AlairasVonal_Vastagsag_Mm", 0.35 )
   hb_HSet( h, "AlairasBal_Vonal_Bal_Mm", 20 )
   hb_HSet( h, "AlairasBal_Vonal_Jobb_Mm", 92 )
   hb_HSet( h, "AlairasJobb_Vonal_Bal_Mm", 118 )
   hb_HSet( h, "AlairasJobb_Vonal_Jobb_Mm", 190 )

   // Aláíró szervezetek (közvetlenül a vonal alatt)
   hb_HSet( h, "Alairas_VonalAlatti_Mm", 1 )
   hb_HSet( h, "Alairas_Betumeret", 12 )
   hb_HSet( h, "Alairas_Felbe", .T. )
   hb_HSet( h, "AlairasBal_Bal_Mm", 20 )
   hb_HSet( h, "AlairasBal_Jobb_Mm", 92 )
   hb_HSet( h, "AlairasJobb_Bal_Mm", 118 )
   hb_HSet( h, "AlairasJobb_Jobb_Mm", 190 )

RETURN h

*------------------------------------------------------------------------------
FUNCTION Nyomtatas()

   IF s_hNyomtatas == NIL
      s_hNyomtatas := NyomtatasDefaults()
   ENDIF

RETURN s_hNyomtatas

*------------------------------------------------------------------------------
PROCEDURE NyomtatasBetolt( hIni )

   LOCAL hDef := NyomtatasDefaults()
   LOCAL h := hb_HClone( hDef )
   LOCAL cKey
   LOCAL xDef
   LOCAL cVal

   IF ValType( hIni ) == "H"
      FOR EACH cKey IN hb_HKeys( hDef )
         IF hb_HHasKey( hIni, cKey )
            cVal := AllTrim( hb_CStr( hIni[ cKey ] ) )
            xDef := hDef[ cKey ]
            DO CASE
            CASE ValType( xDef ) == "N"
               h[ cKey ] := Val( StrTran( cVal, ",", "." ) )
            CASE ValType( xDef ) == "L"
               h[ cKey ] := ( Upper( Left( cVal, 1 ) ) $ "1TIJY" )
            OTHERWISE
               h[ cKey ] := cVal
            ENDCASE
         ENDIF
      NEXT
      // Régi vegyes-méret kulcsok → új egységes kulcsok
      IF ! hb_HHasKey( hIni, "Helyezes_Betumeret" ) .AND. hb_HHasKey( hIni, "Helyezes_Szam_Betumeret" )
         h[ "Helyezes_Betumeret" ] := Val( StrTran( AllTrim( hb_CStr( hIni[ "Helyezes_Szam_Betumeret" ] ) ), ",", "." ) )
      ENDIF
      IF ! hb_HHasKey( hIni, "Helyezes_Felbe" ) .AND. hb_HHasKey( hIni, "Helyezes_Szam_Felbe" )
         h[ "Helyezes_Felbe" ] := ( Upper( Left( AllTrim( hb_CStr( hIni[ "Helyezes_Szam_Felbe" ] ) ), 1 ) ) $ "1TIJY" )
      ENDIF
      IF ! hb_HHasKey( hIni, "Ido_Betumeret" ) .AND. hb_HHasKey( hIni, "Ido_Szam_Betumeret" )
         h[ "Ido_Betumeret" ] := Val( StrTran( AllTrim( hb_CStr( hIni[ "Ido_Szam_Betumeret" ] ) ), ",", "." ) )
      ENDIF
      IF ! hb_HHasKey( hIni, "Ido_Felbe" ) .AND. hb_HHasKey( hIni, "Ido_Szam_Felbe" )
         h[ "Ido_Felbe" ] := ( Upper( Left( AllTrim( hb_CStr( hIni[ "Ido_Szam_Felbe" ] ) ), 1 ) ) $ "1TIJY" )
      ENDIF
   ENDIF

   s_hNyomtatas := h

RETURN

*------------------------------------------------------------------------------
FUNCTION NyomN( cKey )

RETURN Nyomtatas()[ cKey ]

*------------------------------------------------------------------------------
FUNCTION NyomC( cKey )

RETURN hb_CStr( Nyomtatas()[ cKey ] )

*------------------------------------------------------------------------------
FUNCTION NyomL( cKey )

RETURN Nyomtatas()[ cKey ]

*------------------------------------------------------------------------------
FUNCTION IniReadFile( cFile )

   LOCAL h := { => }
   LOCAL cTxt
   LOCAL cLine
   LOCAL nEq
   LOCAL cKey
   LOCAL cVal

   IF ! hb_FileExists( cFile )
      RETURN h
   ENDIF

   cTxt := FileReadText( cFile )
   FOR EACH cLine IN hb_ATokens( cTxt, .T. )
      cLine := AllTrim( cLine )
      IF Empty( cLine ) .OR. Left( cLine, 1 ) $ ";#"
         LOOP
      ENDIF
      IF Left( cLine, 1 ) == "["
         LOOP
      ENDIF
      nEq := At( "=", cLine )
      IF nEq < 2
         LOOP
      ENDIF
      cKey := AllTrim( Left( cLine, nEq - 1 ) )
      cVal := SubStr( cLine, nEq + 1 )
      hb_HSet( h, cKey, cVal )
   NEXT

RETURN h

*------------------------------------------------------------------------------
FUNCTION FileReadText( cFile )

   LOCAL cRaw := hb_MemoRead( cFile )
   LOCAL nLen
   LOCAL i
   LOCAL n
   LOCAL cOut := ""

   nLen := hb_BLen( cRaw )
   IF nLen >= 2 .AND. hb_BPeek( cRaw, 1 ) == 255 .AND. hb_BPeek( cRaw, 2 ) == 254
      i := 3
      DO WHILE i + 1 <= nLen
         n := hb_BPeek( cRaw, i ) + hb_BPeek( cRaw, i + 1 ) * 256
         IF n > 0
            cOut += hb_UChar( n )
         ENDIF
         i += 2
      ENDDO
      RETURN cOut
   ENDIF

   IF nLen >= 3 .AND. hb_BPeek( cRaw, 1 ) == 239 .AND. hb_BPeek( cRaw, 2 ) == 187 .AND. hb_BPeek( cRaw, 3 ) == 191
      cRaw := hb_BSubStr( cRaw, 4 )
   ENDIF

RETURN hb_UTF8ToStr( cRaw )

*------------------------------------------------------------------------------
FUNCTION FileWriteUtf8( cFile, cText )

   LOCAL nH
   LOCAL cUtf := hb_StrToUTF8( cText )
   LOCAL nOk

   nH := FCreate( cFile )
   IF nH < 0
      RETURN .F.
   ENDIF
   nOk := FWrite( nH, cUtf, hb_BLen( cUtf ) )
   FClose( nH )

RETURN nOk == hb_BLen( cUtf )

*------------------------------------------------------------------------------
PROCEDURE Main_SelectFile()

   LOCAL cStart := ""
   LOCAL cDir := ""
   LOCAL cFile
   LOCAL cPrev := AllTrim( Main.Label_XlsxFajl.Value )

   IF ! Empty( cPrev )
      hb_FNameSplit( cPrev, @cDir )
      IF ! Empty( cDir )
         cStart := cDir
      ENDIF
   ENDIF

   cFile := GetFile( { { "Excel munkafüzet (*.xlsx)", "*.xlsx" } }, ;
      "Importálandó Excel fájl", cStart, .F., .T. )

   IF Empty( cFile )
      RETURN
   ENDIF

   Main.Label_XlsxFajl.Value := cFile
   Main_IniSave()

RETURN

*------------------------------------------------------------------------------
PROCEDURE Main_CreatePdf()

   LOCAL cXlsx := AllTrim( Main.Label_XlsxFajl.Value )
   LOCAL aRecs
   LOCAL cPdf
   LOCAL cDir := "", cName := "", cExt := ""
   LOCAL lOk := .F.
   LOCAL nI
   LOCAL oRec

   Main_IniSave()
   NyomtatasBetolt( IniReadFile( Main_IniFile() ) )

   IF ! ImportFajlErvenyes( cXlsx )
      Main_SelectFile()
      cXlsx := AllTrim( Main.Label_XlsxFajl.Value )
      IF ! ImportFajlErvenyes( cXlsx )
         RETURN
      ENDIF
   ENDIF

   aRecs := XlsxReadOklevel( cXlsx )
   IF aRecs == NIL
      RETURN
   ENDIF

   IF Len( aRecs ) == 0
      MsgStop( "A fájlban nincs adat a 2. sortól.", "Oklevél" )
      RETURN
   ENDIF

   hb_FNameSplit( cXlsx, @cDir, @cName, @cExt )
   cPdf := UniquePdfName( cDir, "oklevel" )
   hb_FNameSplit( cPdf, @cDir, @cName, @cExt )

   Main.Button_PdfLetrehozas.Enabled := .F.
   Main.Button_FajlValaszt.Enabled := .F.
   Main.ProgressBar_Pdf.RangeMin := 0
   Main.ProgressBar_Pdf.RangeMax := Len( aRecs )
   Main.ProgressBar_Pdf.Value := 0
   DO EVENTS

   SELECT HPDFDOC cPdf TO lOk PAPERSIZE HPDF_PAPER_A4
   IF ! lOk
      Main.Button_PdfLetrehozas.Enabled := .T.
      Main.Button_FajlValaszt.Enabled := .T.
      MsgStop( "A PDF fájl nem hozható létre:" + CRLF + cPdf, "Oklevél" )
      RETURN
   ENDIF

   SET HPDFDOC COMPRESS ALL
   SET HPDFDOC ENCODING TO "CP1250"
   SET HPDFINFO AUTHOR  TO "Oklevél nyomtatás"
   SET HPDFINFO CREATOR TO "Oklevel"
   SET HPDFINFO TITLE   TO AllTrim( Main.Text_VersenyCim.Value )
   SET HPDFINFO DATECREATED TO Date() TIME Time()

   START HPDFDOC
      FOR nI := 1 TO Len( aRecs )
         oRec := aRecs[ nI ]
         START HPDFPAGE
            OklevelPrintPage( oRec )
         END HPDFPAGE
         Main.ProgressBar_Pdf.Value := nI
         DO EVENTS
      NEXT
   END HPDFDOC

   Main.ProgressBar_Pdf.Value := Len( aRecs )
   Main.Button_PdfLetrehozas.Enabled := .T.
   Main.Button_FajlValaszt.Enabled := .T.
   DO EVENTS

   EXECUTE FILE cPdf

   MsgInfo( "A PDF fájl elkészült." + CRLF + CRLF + ;
      "Név: " + cName + cExt + CRLF + ;
      "Mappa: " + cDir + CRLF + CRLF + ;
      "Teljes útvonal:" + CRLF + cPdf + CRLF + CRLF + ;
      "Oldalak száma: " + hb_NToS( Len( aRecs ) ), "Oklevél" )

RETURN

*------------------------------------------------------------------------------
FUNCTION UniquePdfName( cDir, cBase )

   LOCAL cFile
   LOCAL n := 2

   IF Empty( cDir )
      cDir := GetStartUpFolder() + "\"
   ELSEIF ! ( Right( cDir, 1 ) $ "\/" )
      cDir += "\"
   ENDIF

   IF Empty( cBase )
      cBase := "oklevel"
   ENDIF

   cFile := cDir + cBase + ".pdf"
   DO WHILE hb_FileExists( cFile )
      cFile := cDir + cBase + " (" + hb_NToS( n ) + ").pdf"
      n++
   ENDDO

RETURN cFile

*------------------------------------------------------------------------------
PROCEDURE OklevelPrintPage( hRec )

   LOCAL nTop
   LOCAL cCim := AllTrim( Main.Text_VersenyCim.Value )
   LOCAL cKlubSor
   LOCAL cNev := HMG_UPPER( AllTrim( hb_HGetDef( hRec, "NEV", "" ) ) )
   LOCAL cHely := FormatHelyezes( hb_HGetDef( hRec, "HELYEZES", "" ) )
   LOCAL cIdo := AllTrim( hb_HGetDef( hRec, "IDO", "" ) )
   LOCAL cIdoSzoveg := AllTrim( Main.Label_IdoFelirat.Value )
   LOCAL cDatum := AllTrim( Main.Text_Datum.Value )
   LOCAL cBal := AllTrim( Main.Text_AlairasBal.Value )
   LOCAL cJobb := AllTrim( Main.Text_AlairasJobb.Value )
   LOCAL cHelySor
   LOCAL nY
   LOCAL nNevPrintY

   cKlubSor := JoinParts( { ;
      AllTrim( hb_HGetDef( hRec, "NEVELO", "" ) ), ;
      AllTrim( hb_HGetDef( hRec, "KLUB", "" ) ), ;
      AllTrim( Main.Text_KlubToldalek.Value ) } )

   cHelySor := JoinParts( { cHely, AllTrim( Main.Text_HelyezesSzoveg.Value ) } )

   nTop := NyomN( "Szoveg_Kezdo_Ymm" )

   PdfCenter( nTop + NyomN( "Cim_Ymm" ), NyomN( "Cim_Magassag_Mm" ), cCim, ;
      NyomN( "Cim_Betumeret" ), NyomL( "Cim_Felbe" ) )
   PdfCenter( nTop + NyomN( "Klub_Ymm" ), NyomN( "Klub_Magassag_Mm" ), cKlubSor, ;
      NyomN( "Klub_Betumeret" ), NyomL( "Klub_Felbe" ) )
   PdfCenter( nTop + NyomN( "Nev_Ymm" ), NyomN( "Nev_Magassag_Mm" ), cNev, ;
      NyomN( "Nev_Betumeret" ), NyomL( "Nev_Felbe" ) )

   // Név nyomtatási Y (PdfCenter-rel azonos), vonal = betűalj + 1 mm
   nNevPrintY := nTop + NyomN( "Nev_Ymm" ) + NyomN( "Nev_Magassag_Mm" ) * 0.70
   nY := nNevPrintY + PtToMm( NyomN( "Nev_Betumeret" ) ) + NyomN( "NevAlattiVonal_Offset_Mm" )
   @ nY, NyomN( "NevAlattiVonal_Bal_Mm" ) HPDFPRINT LINE TO ;
      nY, NyomN( "NevAlattiVonal_Jobb_Mm" ) ;
      PENWIDTH NyomN( "NevAlattiVonal_Vastagsag_Mm" )

   PdfCenter( nY + NyomN( "Helyezes_VonalAlatti_Mm" ), 0, cHelySor, ;
      NyomN( "Helyezes_Betumeret" ), NyomL( "Helyezes_Felbe" ) )
   PdfCenter( nTop + NyomN( "Ido_Ymm" ), NyomN( "Ido_Magassag_Mm" ), cIdo, ;
      NyomN( "Ido_Betumeret" ), NyomL( "Ido_Felbe" ) )
   PdfCenter( nTop + NyomN( "IdoFelirat_Ymm" ), NyomN( "IdoFelirat_Magassag_Mm" ), cIdoSzoveg, ;
      NyomN( "IdoFelirat_Betumeret" ), NyomL( "IdoFelirat_Felbe" ) )
   PdfCenter( nTop + NyomN( "Datum_Ymm" ), NyomN( "Datum_Magassag_Mm" ), cDatum, ;
      NyomN( "Datum_Betumeret" ), NyomL( "Datum_Felbe" ) )

   nY := nTop + NyomN( "AlairasVonal_Ymm" )
   @ nY, NyomN( "AlairasBal_Vonal_Bal_Mm" ) HPDFPRINT LINE TO ;
      nY, NyomN( "AlairasBal_Vonal_Jobb_Mm" ) PENWIDTH NyomN( "AlairasVonal_Vastagsag_Mm" )
   @ nY, NyomN( "AlairasJobb_Vonal_Bal_Mm" ) HPDFPRINT LINE TO ;
      nY, NyomN( "AlairasJobb_Vonal_Jobb_Mm" ) PENWIDTH NyomN( "AlairasVonal_Vastagsag_Mm" )

   PdfBox( nY + NyomN( "Alairas_VonalAlatti_Mm" ), NyomN( "AlairasBal_Bal_Mm" ), ;
      0, NyomN( "AlairasBal_Jobb_Mm" ), cBal, ;
      NyomN( "Alairas_Betumeret" ), NyomL( "Alairas_Felbe" ) )
   PdfBox( nY + NyomN( "Alairas_VonalAlatti_Mm" ), NyomN( "AlairasJobb_Bal_Mm" ), ;
      0, NyomN( "AlairasJobb_Jobb_Mm" ), cJobb, ;
      NyomN( "Alairas_Betumeret" ), NyomL( "Alairas_Felbe" ) )

RETURN

*------------------------------------------------------------------------------
FUNCTION PtToMm( nPt )

RETURN nPt * 25.4 / 72

*------------------------------------------------------------------------------
PROCEDURE PdfCenter( nRow, nH, cText, nSize, lBold )

   LOCAL cFont := NyomC( "Betutipus" )
   LOCAL nMid := ( NyomN( "Oldal_Bal_Mm" ) + NyomN( "Oldal_Jobb_Mm" ) ) / 2
   LOCAL nY := nRow + nH * 0.70
   LOCAL cOut

   cText := AllTrim( cText )
   IF Empty( cText )
      RETURN
   ENDIF
   cOut := hupdf( cText )

   IF lBold
      @ nY, nMid HPDFPRINT cOut FONT cFont SIZE nSize BOLD CENTER
   ELSE
      @ nY, nMid HPDFPRINT cOut FONT cFont SIZE nSize CENTER
   ENDIF

RETURN

*------------------------------------------------------------------------------
/* UTF-8 → HUWIN (CP1250); a PDF encoding is „CP1250” (ne ISO8859-2 — Ž/Š eltér) */
FUNCTION hupdf( cText )

   IF ValType( cText ) != "C"
      RETURN ""
   ENDIF

RETURN HB_UTF8TOSTR( cText, "HUWIN" )

*------------------------------------------------------------------------------
PROCEDURE PdfBox( nRow, nCol, nH, nToCol, cText, nSize, lBold )

   LOCAL cFont := NyomC( "Betutipus" )
   LOCAL cOut
   LOCAL nMid := ( nCol + nToCol ) / 2
   LOCAL nY := nRow + nH * 0.70

   cText := AllTrim( cText )
   IF Empty( cText )
      RETURN
   ENDIF
   cOut := hupdf( cText )

   IF lBold
      @ nY, nMid HPDFPRINT cOut FONT cFont SIZE nSize BOLD CENTER
   ELSE
      @ nY, nMid HPDFPRINT cOut FONT cFont SIZE nSize CENTER
   ENDIF

RETURN

*------------------------------------------------------------------------------
FUNCTION JoinParts( aPart )

   LOCAL c := ""
   LOCAL x

   FOR EACH x IN aPart
      IF ! Empty( AllTrim( x ) )
         IF ! Empty( c )
            c += " "
         ENDIF
         c += AllTrim( x )
      ENDIF
   NEXT

RETURN c

*------------------------------------------------------------------------------
FUNCTION FormatHelyezes( cVal )

   LOCAL nVal

   cVal := AllTrim( cVal )
   IF Empty( cVal )
      RETURN ""
   ENDIF

   nVal := Val( StrTran( cVal, ",", "." ) )
   IF At( ":", cVal ) == 0 .AND. nVal != 0 .AND. nVal == Int( nVal )
      cVal := hb_NToS( Int( nVal ) )
   ENDIF

   IF Right( cVal, 1 ) != "."
      cVal += "."
   ENDIF

RETURN cVal

*------------------------------------------------------------------------------
FUNCTION FormatIdo( cVal )

   LOCAL nVal
   LOCAL nSec
   LOCAL nH := 0
   LOCAL nM := 0
   LOCAL nS := 0
   LOCAL aTok

   cVal := AllTrim( cVal )
   IF Empty( cVal )
      RETURN ""
   ENDIF

   aTok := hb_ATokens( cVal, ":" )
   IF Len( aTok ) >= 3
      nH := Int( Val( aTok[ 1 ] ) )
      nM := Int( Val( aTok[ 2 ] ) )
      nS := Int( Val( aTok[ 3 ] ) )
   ELSEIF Len( aTok ) == 2
      nM := Int( Val( aTok[ 1 ] ) )
      nS := Int( Val( aTok[ 2 ] ) )
   ELSE
      nVal := Val( StrTran( cVal, ",", "." ) )
      IF nVal > 0 .AND. nVal < 10
         nSec := Round( nVal * 86400, 0 )
         nH := Int( nSec / 3600 )
         nM := Int( ( nSec % 3600 ) / 60 )
         nS := Int( nSec % 60 )
      ELSE
         RETURN cVal
      ENDIF
   ENDIF

RETURN IdoFormat( nH, nM, nS )

*------------------------------------------------------------------------------
/* 1:32:35 ha van óra; 3:42 ha nincs (perc vezető nulla nélkül) */
FUNCTION IdoFormat( nH, nM, nS )

   IF nH > 0
      RETURN hb_NToS( nH ) + ":" + StrZero( nM, 2 ) + ":" + StrZero( nS, 2 )
   ENDIF

RETURN hb_NToS( nM ) + ":" + StrZero( nS, 2 )

*------------------------------------------------------------------------------
FUNCTION XlsxReadOklevel( cFile )

   LOCAL cSheet
   LOCAL cSst
   LOCAL aSS
   LOCAL cData
   LOCAL aRows
   LOCAL hCol
   LOCAL aRecs := {}
   LOCAL nHeader
   LOCAL nI
   LOCAL hRow
   LOCAL hRec
   LOCAL cNev
   LOCAL aHianyzik := {}

   cSheet := ZipReadUtf8( cFile, "xl/worksheets/sheet1.xml" )
   IF Empty( cSheet )
      MsgStop( "Az Excel munkalap (sheet1) nem olvasható.", "Oklevél" )
      RETURN NIL
   ENDIF

   cSst := ZipReadUtf8( cFile, "xl/sharedStrings.xml" )
   aSS := ParseSharedStrings( cSst )

   cData := XmlFirstInner( cSheet, "sheetData" )
   IF Empty( cData )
      cData := cSheet
   ENDIF

   aRows := XmlElements( cData, "row" )
   IF Len( aRows ) < 2
      MsgStop( "A munkalapon nincs elég sor (kell a fejléc és legalább egy adatsor).", "Oklevél" )
      RETURN NIL
   ENDIF

   nHeader := XlsxHeaderIndex( aRows )
   hCol := XlsxMapColumns( aRows[ nHeader ], aSS )

   IF ! hb_HHasKey( hCol, NormHeader( "Helyezés" ) )
      AAdd( aHianyzik, "Helyezés" )
   ENDIF
   IF ! hb_HHasKey( hCol, NormHeader( "Név" ) )
      AAdd( aHianyzik, "Név" )
   ENDIF
   IF ! hb_HHasKey( hCol, NormHeader( "Névelő" ) )
      AAdd( aHianyzik, "Névelő" )
   ENDIF
   IF ! hb_HHasKey( hCol, NormHeader( "Klub" ) )
      AAdd( aHianyzik, "Klub" )
   ENDIF
   IF ! hb_HHasKey( hCol, NormHeader( "Időeredmény" ) )
      AAdd( aHianyzik, "Időeredmény" )
   ENDIF

   IF Len( aHianyzik ) > 0
      hb_HSet( hCol, NormHeader( "Helyezés" ), "A" )
      hb_HSet( hCol, NormHeader( "Név" ), "B" )
      hb_HSet( hCol, NormHeader( "Névelő" ), "C" )
      hb_HSet( hCol, NormHeader( "Klub" ), "D" )
      hb_HSet( hCol, NormHeader( "Időeredmény" ), "E" )
   ENDIF

   FOR nI := nHeader + 1 TO Len( aRows )
      hRow := XlsxRowValues( aRows[ nI ], aSS )
      cNev := AllTrim( CellByField( hRow, hCol, "Név" ) )
      IF Empty( cNev )
         LOOP
      ENDIF
      hRec := { => }
      hb_HSet( hRec, "HELYEZES", CellByField( hRow, hCol, "Helyezés" ) )
      hb_HSet( hRec, "NEV", cNev )
      hb_HSet( hRec, "NEVELO", CellByField( hRow, hCol, "Névelő" ) )
      hb_HSet( hRec, "KLUB", CellByField( hRow, hCol, "Klub" ) )
      hb_HSet( hRec, "IDO", FormatIdo( CellByField( hRow, hCol, "Időeredmény" ) ) )
      AAdd( aRecs, hRec )
   NEXT

RETURN aRecs

*------------------------------------------------------------------------------
FUNCTION XlsxHeaderIndex( aRows )

   LOCAL nI
   LOCAL cR

   FOR nI := 1 TO Len( aRows )
      cR := XmlAttr( aRows[ nI, 1 ], "r" )
      IF cR == "1"
         RETURN nI
      ENDIF
   NEXT

RETURN 1

*------------------------------------------------------------------------------
FUNCTION XlsxMapColumns( aRow, aSS )

   LOCAL hCol := { => }
   LOCAL hRow := XlsxRowValues( aRow, aSS )
   LOCAL cCol
   LOCAL cTitle

   FOR EACH cCol IN hb_HKeys( hRow )
      cTitle := NormHeader( hRow[ cCol ] )
      IF ! Empty( cTitle )
         hb_HSet( hCol, cTitle, cCol )
      ENDIF
   NEXT

RETURN hCol

*------------------------------------------------------------------------------
FUNCTION XlsxRowValues( aRow, aSS )

   LOCAL h := { => }
   LOCAL aCells
   LOCAL aC
   LOCAL cRef
   LOCAL cCol
   LOCAL cType
   LOCAL cVal

   aCells := XmlElements( aRow[ 2 ], "c" )
   FOR EACH aC IN aCells
      cRef := XmlAttr( aC[ 1 ], "r" )
      cCol := CellColLetter( cRef )
      IF Empty( cCol )
         LOOP
      ENDIF
      cType := Lower( XmlAttr( aC[ 1 ], "t" ) )
      cVal := XlsxCellText( aC[ 1 ], aC[ 2 ], aSS, cType )
      hb_HSet( h, cCol, cVal )
   NEXT

RETURN h

*------------------------------------------------------------------------------
FUNCTION XlsxCellText( cOpen, cInner, aSS, cType )

   LOCAL cV := ""

   IF Empty( cOpen )
      cOpen := ""
   ENDIF

   IF cType == "s"
      cV := AllTrim( XmlFirstInner( cInner, "v" ) )
      RETURN SharedStr( aSS, Val( cV ) )
   ENDIF

   IF cType == "inlinestr"
      RETURN XmlConcatT( cInner )
   ENDIF

   IF cType == "str"
      RETURN AllTrim( XmlFirstInner( cInner, "v" ) )
   ENDIF

RETURN AllTrim( XmlFirstInner( cInner, "v" ) )

*------------------------------------------------------------------------------
FUNCTION CellByField( hRow, hCol, cField )

   LOCAL cCol := hb_HGetDef( hCol, NormHeader( cField ), "" )

   IF Empty( cCol )
      RETURN ""
   ENDIF

RETURN hb_HGetDef( hRow, cCol, "" )

*------------------------------------------------------------------------------
FUNCTION SharedStr( aSS, nIndex )

   LOCAL nPos := Int( nIndex ) + 1

   IF nPos < 1 .OR. nPos > Len( aSS )
      RETURN ""
   ENDIF

RETURN aSS[ nPos ]

*------------------------------------------------------------------------------
FUNCTION ParseSharedStrings( cXml )

   LOCAL a := {}
   LOCAL aSi
   LOCAL aOne

   IF Empty( cXml )
      RETURN a
   ENDIF

   aSi := XmlElements( cXml, "si" )
   FOR EACH aOne IN aSi
      AAdd( a, XmlConcatT( aOne[ 2 ] ) )
   NEXT

RETURN a

*------------------------------------------------------------------------------
FUNCTION NormHeader( cText )

RETURN AllTrim( HMG_UPPER( AllTrim( cText ) ) )

*------------------------------------------------------------------------------
FUNCTION CellColLetter( cRef )

   LOCAL cCol := ""
   LOCAL i
   LOCAL ch

   cRef := AllTrim( cRef )
   FOR i := 1 TO Len( cRef )
      ch := SubStr( cRef, i, 1 )
      IF ch >= "0" .AND. ch <= "9"
         EXIT
      ENDIF
      cCol += ch
   NEXT

RETURN Upper( cCol )

*------------------------------------------------------------------------------
FUNCTION ZipReadUtf8( cZip, cEntry )

   LOCAL hUnzip
   LOCAL nErr
   LOCAL cFile := ""
   LOCAL cNorm
   LOCAL cWant
   LOCAL cTmp
   LOCAL cData := ""
   LOCAL cDir

   cWant := Lower( StrTran( cEntry, "\", "/" ) )
   hUnzip := hb_unzipOpen( cZip )
   IF Empty( hUnzip )
      RETURN ""
   ENDIF

   cDir := GetTempFolder()
   IF ! Empty( cDir ) .AND. ! ( Right( cDir, 1 ) $ "\/" )
      cDir += "\"
   ENDIF
   cTmp := cDir + "oklevel_unzip.tmp"

   nErr := hb_unzipFileFirst( hUnzip )
   DO WHILE nErr == 0
      hb_unzipFileInfo( hUnzip, @cFile )
      cNorm := Lower( StrTran( cFile, "\", "/" ) )
      IF cNorm == cWant .OR. Right( cNorm, Len( cWant ) + 1 ) == "/" + cWant
         FErase( cTmp )
         hb_unzipExtractCurrentFile( hUnzip, cTmp )
         IF hb_FileExists( cTmp )
            cData := hb_MemoRead( cTmp )
            FErase( cTmp )
         ENDIF
         EXIT
      ENDIF
      nErr := hb_unzipFileNext( hUnzip )
   ENDDO
   hb_unzipClose( hUnzip )

   IF Empty( cData )
      RETURN ""
   ENDIF

   cData := hb_UTF8ToStr( cData )
   IF Asc( Left( cData, 1 ) ) == 65279
      cData := SubStr( cData, 2 )
   ENDIF

RETURN cData

*------------------------------------------------------------------------------
FUNCTION XmlElements( cXml, cTag )

   LOCAL a := {}
   LOCAL nPos := 1
   LOCAL nOpen
   LOCAL nEndOpen
   LOCAL nClose
   LOCAL cNext
   LOCAL cOpen := "<" + cTag
   LOCAL cClose := "</" + cTag + ">"
   LOCAL nTagLen := Len( cOpen )
   LOCAL cOpenTag
   LOCAL cInner

   DO WHILE .T.
      nOpen := hb_At( cOpen, cXml, nPos )
      IF nOpen == 0
         EXIT
      ENDIF
      cNext := SubStr( cXml, nOpen + nTagLen, 1 )
      IF ! Empty( cNext ) .AND. ! ( cNext $ " >/" )
         nPos := nOpen + 1
         LOOP
      ENDIF
      nEndOpen := hb_At( ">", cXml, nOpen )
      IF nEndOpen == 0
         EXIT
      ENDIF
      cOpenTag := SubStr( cXml, nOpen, nEndOpen - nOpen + 1 )
      IF SubStr( cXml, nEndOpen - 1, 1 ) == "/"
         AAdd( a, { cOpenTag, "" } )
         nPos := nEndOpen + 1
         LOOP
      ENDIF
      nClose := hb_At( cClose, cXml, nEndOpen )
      IF nClose == 0
         EXIT
      ENDIF
      cInner := SubStr( cXml, nEndOpen + 1, nClose - nEndOpen - 1 )
      AAdd( a, { cOpenTag, cInner } )
      nPos := nClose + Len( cClose )
   ENDDO

RETURN a

*------------------------------------------------------------------------------
FUNCTION XmlFirstInner( cXml, cTag )

   LOCAL a := XmlElements( cXml, cTag )

   IF Len( a ) == 0
      RETURN ""
   ENDIF

RETURN XmlDecode( a[ 1, 2 ] )

*------------------------------------------------------------------------------
FUNCTION XmlConcatT( cXml )

   LOCAL c := ""
   LOCAL aT
   LOCAL aOne

   aT := XmlElements( cXml, "t" )
   FOR EACH aOne IN aT
      c += XmlDecode( aOne[ 2 ] )
   NEXT

RETURN c

*------------------------------------------------------------------------------
FUNCTION XmlAttr( cTag, cName )

   LOCAL nPos
   LOCAL nEnd
   LOCAL cPat
   LOCAL cQ

   cPat := cName + '="'
   nPos := At( cPat, cTag )
   IF nPos == 0
      cPat := cName + "='"
      nPos := At( cPat, cTag )
      IF nPos == 0
         RETURN ""
      ENDIF
   ENDIF
   cQ := Right( cPat, 1 )
   nPos += Len( cPat )
   nEnd := hb_At( cQ, cTag, nPos )
   IF nEnd == 0
      RETURN ""
   ENDIF

RETURN SubStr( cTag, nPos, nEnd - nPos )

*------------------------------------------------------------------------------
FUNCTION XmlDecode( cText )

   cText := StrTran( cText, "&lt;", "<" )
   cText := StrTran( cText, "&gt;", ">" )
   cText := StrTran( cText, "&quot;", '"' )
   cText := StrTran( cText, "&apos;", "'" )
   cText := StrTran( cText, "&#39;", "'" )
   cText := StrTran( cText, "&amp;", "&" )

RETURN cText
