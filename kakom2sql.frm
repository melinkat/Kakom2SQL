VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Kakom Import"
   ClientHeight    =   540
   ClientLeft      =   -120
   ClientTop       =   720
   ClientWidth     =   9960
   Icon            =   "kakom2sql.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   9960
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   600
   End
   Begin VB.TextBox cStatus 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "keine Statusmeldung"
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- KakomSQL Files in Access-Datenbank exportieren ----------------
' Version für MS SQL_server
' $Header: C:\\RCS\\C\\SRC\\HATSCHER\\SQLServer\\Kakom2SQL\\kakom2sql.frm,v 1.6 2003-11-20 14:36:22+01 klaus dead $
' $Log: kakom2sql.frm,v $
' Revision 1.6  2010-06-22 14:36:22+01  klaus
' Parameter eingeführt wegen ASCON und KAKOM Kassen  Aufruf: "kakomsql ascon"  .sal wird nach ascon verschoben
'             wegen Namensgleichheit der .SAL
' hourly.sal und hourplu.sal wird bei Einlesen die Artikelmenge /1000 dividiert wegen halber Mengen
'  Satzbeispiel hourly.sal: 0,0000009024,000000,08062010,1227,0730,0759,+0000002500,+0000000191
'
' Revision 1.5  2003-09-30 16:01:59+02  klaus
' exe in KakomSQL umbenannt wegen 8Zeichen-Problem
' conAnsi erweitert, da neues Steuerzeichen herausgefiltert werden musste
'
' Revision 1.4  2003-08-25 17:34:40+02  klaus
' Kontrolle auf 5 Stellen wieder herausgenommen, da in KAKOM korrigiert
'
' Revision 1.3  2003-07-14 16:49:35+02  klaus
' Länge des TransactionsCodes wird überprüft, da in neuer Version von .sal
' in der Windows-Version 5 Stellen. Hier nur 4 Stellen wegen DOS-Version
'
' Revision 1.2  2003-07-04 14:29:53+02  klaus
' Mit neuer Version der SAL-Files ist die Filialenummer 10stellig
' Es wird grundsaetzlich auf 4 Stelle von rechts abgeschnitten
' Daraus Stelle 1+2 fuer Filialenummer und Stelle 3+4 fuer KasseNummer
'
'
' Revision 1.0  2003-07-04 13:10:03+02  klaus
' Initial revision
'
'
'
' offen: Kontrolle, ob Kasse fehlt
' 011112 Kassierername auskommentiert - es wird nur noch die Nummer
'        in TabKakomKassierer und TabImpKassierer3x1 benutzt
' 001224 Warengruppenimport o h n e Update -> Warum?
' 990224 conANSI eingebaut -> Umwandlung Zeichen Dos -> Win
' Filiale und kasse wird aus Kasse gebildet: platz 1+2 -> 90+Filiale/ 3+4->00+Kasse
' 990312 Files haben Namen ohne datum
' 990514 INI-File eingeführt Groß/Kleinschreibung beachten
'    DatenKakom  .sal Files
'    Datenbank   .mdb
'    SichPfad    ordner in dem die Tagesordner der .sal'S abgelegt werden
' 990610 nicht übernommene datensätze werden in gesondertes File abgelegt
' 990628 Level4 begonnen einzuarbeiten Tabelle und Script müssen noch bearbeitet
'    werden
' 991008 Personennamen werden übernommen aus sqlabfr.mdb und in Kassiererbericht
'    eingetragen
' 991220 Bedienerimport KassiererName und Personalnummer werden übernommen und
'    in Kassiererbericht eingetragen
' 000105 Tabelle Artikelmengen werden alle Daten zu einem Artikel in die
'    gleiche Zeile eingetragen - Index auf datum,Filiale,kasse, artikelnummer
' 001213 Artikelfrequenbericht Frequenzbericht: wenn 99:99  dann 23:59 eintragen OK
'
'
'--------------------------------------------------------------------
 
 
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim cmdP As New ADODB.Command
Dim rsP As New ADODB.Recordset

Dim cDatum As String
'für Datumkontrolle
Dim cDatumSQL As String
Dim nDoppelt As Integer
Dim nDoppeltGes As Integer
Dim cDat As String
Dim cText As String
Dim cSichPfad As String
Dim cDatenpfad As String
Dim cFile As String
Dim cFileTmp As String
Dim suchDatum As String
Dim cKasse As String            ' Für Parameter Kasse Vectron, kakom zum Speichen in Unterordner

Private Sub Form_Load()
  frmMain.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
  
  cKasse = Command()

  If Len(cKasse) Then
      cKasse = "_" & cKasse    ' Befehlszeilenargument:  cKasse
  Else
     cKasse = ""      ' Es wurden keine Befehlszeilenargumente angegeben"
  End If

End Sub
Private Sub Timer1_Timer()
  StatusAnzeige ("beginn")
  Timer1.Enabled = False         'Programm wird über Timer gestartet damit
  cDatumSQL = "X"                'Form angezeigt wird
  nDoppelt = 0
  nDoppeltGes = 0
  cFileTmp = ""
  Protokoll
  StatusAnzeige ("Protokoll angelegt")
  TagProtokoll
  StatusAnzeige ("Tagesprotokoll angelegt")
  Logbuch ("  ")
  Logbuch ("  ")
  Logbuch ("Beginn " & App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision & " am " & Date & "-" & Time & " ----")
  Logbuch ("  Abholen der Daten von " & cKasse)
  Print #2, "    Kakomsql.ini einlesen..."
  cDatenpfad = GetProfile("DatenKakom", "./")
  Print #2, "    Daten werden geholt von " & cDatenpfad
  MacheFlagDateiAus                 'Fertig.txt wird gelöscht, fals vorhanden
  MacheFlagDatei                    'Fertig.txt wird erzeugt und bleibt bis
                                    ' erfolgreiches Programmende bestehen
  suchDatum = aktuellesDatum()
  StatusAnzeige ("Datenbanken werden geöffnet")
'  dbOeffnen
  odbcOeffnen
  PLImport
  OrderImport
  RetourImport
  InventurImport
  WarengruppenImport
  FrequenzBerichtImport
  ArtikelFrequenzBerichtImport
  TransaktionenImport
  Level4Import
  KassiererImport
  BedienerImport
  If nDoppeltGes > 0 Then
    Logbuch ("!!! - " & Str(nDoppeltGes) & " Datensätze waren bereits vorhanden -> nicht übernommen-> siehe Datei: NichtUebernommen" & Format(Date, "yyyymmdd") & ".LOG !")
  End If
  Logbuch ("-------- Ende des Programms ----------")
  Logbuch ("  ")
  Logbuch ("  ")
  MacheFlagDateiAus                 ' Fertig.txt wird gelöscht -> alles OK
 End
End Sub

Function GetProfile(cKey As String, cDefault As String)
cF = CurDir() & ".ini"
cF = App.Path & "\kakomsql.ini"     '!!! alte Kakomsql.ini benutzen
GetProfile = cDefault
Open cF For Input As #9
  Do While Not EOF(9)
    Line Input #9, cline
    aa = Split(cline, "=")
    If Len(cline) > 1 Then
      If UCase(aa(0)) = UCase(cKey) Then
        GetProfile = aa(1)
        Exit Do
      End If
    End If
  Loop
  Close #9
End Function

Sub Protokoll()
  Open App.Path & "\protokol.log" For Append As #2
  Open App.Path & "\nichtUebernommen" & Format(Date, "yyyymmdd") & ".log" For Append As #8
  Print #8, " "
  Print #8, "----am " & Date & " weil bereits vorhanden nicht übernommen --------------------"
End Sub
Sub TagProtokoll()
   cF = App.Path & "\TagProt_" & cKasse & ".log"
   If Dir(cF) <> "" Then
     Kill (cF)
   End If
   Open cF For Append As #4
End Sub
Sub MacheFlagDatei()
  Open App.Path & "\fertig.txt" For Append As #3
End Sub
Sub MacheFlagDateiAus()
   cF = App.Path & "\fertig.txt"
   If Dir(cF) <> "" Then
     Close #3
     Kill App.Path & "\fertig.txt"
   End If
End Sub

Function odbcOeffnen()
' cDB = GetProfile("DSN", "")
'  Print #2, "    Zieldatenbank= SQL Server DSN = " & cDB         ' & cDatenbank
'  cOpen = ("DSN=" & cDB & "; uid=sa; pwd=;")
'  con.Open cOpen
  con.ConnectionString = "Provider=SQLOLEDB; Data Source=" & GetProfile("DataSource", "localhost") & ";" _
                         & "Initial Catalog=" & GetProfile("InitialCatalog", "business") & ";" _
                         & "user id=" & GetProfile("UserID", "sa") & ";" _
                         & "Connect Timeout=" & GetProfile("ConnectTimeout", "30")
 'con.Properties("Password").Value = GetProfile("Passwort")
 con.Properties("Integrated Security").Value = "SSPI"

  Print #2, "    Zieldatenbank: " & con.ConnectionString
  con.Open
  Print #2, "    Zieldatenbank: geöffnet"
End Function
'----( Daten einlesen )----------------------------------------------
Sub PLImport()
  cFile = cDatenpfad & "\plu.sal"
  cStatus = cFile
  cStatus.Refresh
  cFlag = "X"
  If Dir(cFile) <> "" Then
    Set cmd.ActiveConnection = con
    cmd.CommandText = "Select * from TabKakomArtikelUmsatz" _
                   & " WHERE tag = '" & str2date(suchDatum) & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenDynamic, adLockOptimistic
    Open cFile For Input As #1   ' Datei öffnen.
    Do While Not EOF(1)        ' Schleife bis Dateiende.
      Line Input #1, cText    ' Zeile in Variable einlesen.
       '0,2401,0000,19011999,0313,00000000001005,+0000004000,+0000001061,"Weizenbr.kl.                    "
       'flag,Fil,kasse,datum,hhmm,artNr,verkMenge,UBetrag,ArtBez
      aDaten = Split(cText, ",")
      aDaten(2) = Right(aDaten(2), 4)                           ' Filiale hat in neuer Version 6 Stellen
      On Error Resume Next
      If aDaten(3) = suchDatum Then
        such = "  tag = '" & str2date(aDaten(3)) & "'" _
           & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
           & " and kasse = '" & Right(aDaten(2), 2) & "'" _
           & " and artikelnummer = '" & Right(aDaten(5), 5) & "'"
        With rs
        .Filter = such
        If .RecordCount < 1 Then
          .AddNew
          !Filiale = "90" & Left(aDaten(2), 2)
          !Kasse = Right(aDaten(2), 2)
          !datum = aDaten(3) 'cDatum
          !Zeit = aDaten(4)   'cZeit
          !Artikelnummer = Right(aDaten(5), 5)
          !VerkMenge = CDbl(aDaten(6) / 1000)
          !Umsatz = CDbl(aDaten(7) / 100)
          !Artikelbezeichnung = Left(conANSI(aDaten(8)), 32)
          !Tag = str2date(aDaten(3))
          !Artikel = Right(aDaten(5), 5) & " " & Left(conANSI(aDaten(8)), 16)
          StatusAnzeige (cFile & ":Satz neu     " & cText)     ' anzeige des Fortschritts
        Else
          cVerk = !VerkMenge
          cUmsatz = !Umsatz
          !VerkMenge = CDbl(aDaten(6) / 1000)
          !Umsatz = CDbl(aDaten(7) / 100)
          If cVerk <> 0 And cVerk <> !VerkMenge Or cUmsatz <> 0 And cUmsatz <> !Umsatz Then
            StatusAnzeige (cFile & ":Satz update  " & cText)     ' anzeige des Fortschritts
            Logbuch ("!!! Datensatz geändert: VerkMenge alt:" & cVerk & " Umsatz:" & cUmsatz & " Neuer Satz: " & cText)
          Else
            StatusAnzeige (cFile & ":Satz ergänzt " & cText)     ' anzeige des Fortschritts
          End If
        End If
        .Update
        .Filter = ""
        End With
        Fehlertest
       ' DatumTest (adaten(3))
        If cFlag = "X" Then              ' begin bearb.. hier ausnahmsweise
          cFlag = "N"                    ' zeigen damit Datum in Protokoll zuerst angezeigt wird
          Logbuch ("    Beginn: " & cFile)
        End If
     Else
         Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
     End If
   Loop
   Close #1   ' Datei schließen.
   rs.Close
   cND = ""
   If nDoppelt > 0 Then
     cND = nDoppelt & " Datensätze nicht übernommen"
   End If
   Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
   nDoppelt = 0
   StatusAnzeige ("plu.sal sichern")
   FileCopy cFile, cSichPfad & "\plu.sal"
   Kill cFile
 Else
   Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
 End If
End Sub

' Bestellungen einlesen
Sub OrderImport()
cFileName = "\order.sal"
cFile = cDatenpfad & cFileName
cStatus = cFile
If Dir(cFile) <> "" Then
  Set cmd.ActiveConnection = con
  cmd.CommandText = "Select * from TabKakomArtikelUmsatz" _
                 & " WHERE tag = '" & str2date(suchDatum) & "'"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenDynamic, adLockOptimistic
  Open cFile For Input As #1   ' Datei öffnen.
  Logbuch ("    Beginn: " & cFile)
  Do While Not EOF(1)        ' Schleife bis Dateiende.
    Line Input #1, cText     ' Zeile in Variable einlesen.
    ' flag,Fil,kasse,datum,hhmm,artNr,bestellMenge,ArtikelBezeichnung
    aDaten = Split(cText, ",")
    aDaten(2) = Right(aDaten(2), 4)
    On Error Resume Next
    If aDaten(3) = suchDatum Then
      such = "tag = '" & str2date(aDaten(3)) & "'" _
                    & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
                    & " and kasse = '" & Right(aDaten(2), 2) & "'" _
                    & " and artikelnummer = '" & Right(aDaten(5), 5) & "'"
      With rs
      .Filter = such
      If .RecordCount = 0 Then
        .AddNew
        !Filiale = "90" & Left(aDaten(2), 2)
        !Kasse = Right(aDaten(2), 2)
        !datum = aDaten(3) 'cDatum
        !Zeit = aDaten(4)   'cZeit
        !Artikelnummer = Right(aDaten(5), 5)
        !Bestellmenge = CDbl(aDaten(6) / 1000)
        !Artikelbezeichnung = Left(conANSI(aDaten(7)), 32)
        !Tag = str2date(aDaten(3))
        !Artikel = Right(aDaten(5), 5) & " " & Left(conANSI(aDaten(7)), 16)
        StatusAnzeige (cFile & ":Satz neu     " & cText)
      Else
        cBestell = !Bestellmenge
        !Bestellmenge = CDbl(aDaten(6) / 1000)
        If cBestell <> 0 And cBestell <> !Bestellmenge Then
          StatusAnzeige (cFile & ":Satz update  " & cText)
          Logbuch ("!!! Datensatz geaendert: Bestellmenge alt=" & cBestell & " Neuer Satz: " & cText)
        Else
          StatusAnzeige (cFile & ":Satz ergänzt " & cText)
        End If
      End If
      .Update
      .Filter = ""
      End With
      Fehlertest
    Else
        Logbuch ("!!! - falsches Datum: " & conANSI(cText))
    End If
  Loop
  rs.Close
  Close #1   ' Datei schließen.
  cND = ""
  If nDoppelt > 0 Then
    cND = nDoppelt & " Datensätze nicht übernommen"
  End If
  Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
 nDoppelt = 0
 StatusAnzeige ("order.sal sichern")
 FileCopy cFile, cSichPfad & "\order.sal"
 Kill cFile
Else
  Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
End If
End Sub

' Retouren einlesen
Sub RetourImport()
cFile = cDatenpfad & "\return.sal"
cStatus = cFile
cStatus.Refresh
If Dir(cFile) <> "" Then
  Set cmd.ActiveConnection = con
  cmd.CommandText = "Select * from TabKakomArtikelUmsatz" _
                    & " WHERE tag = '" & str2date(suchDatum) & "'"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenDynamic, adLockOptimistic
  Open cFile For Input As #1   ' Datei öffnen.
  Logbuch ("    Beginn: " & cFile)
  Do While Not EOF(1)        ' Schleife bis Dateiende.
    Line Input #1, cText    ' Zeile in Variable einlesen.
    '0,9021,0000,19011999,0313,00000000000200,+0000001000,"* BESTELLUNG *                  "
    'flag,Fil,kasse,datum,hhmm,artNr,Retour,ArtikelBezeichnung
    aDaten = Split(cText, ",")
    aDaten(2) = Right(aDaten(2), 4)
    If InStr(1, aDaten(6), "P", vbTextCompare) > 0 Or InStr(1, aDaten(6), "N", vbTextCompare) > 0 Then
      aDaten(6) = Replace(aDaten(6), "P", "0")
      aDaten(6) = Replace(aDaten(6), "N", "0")
      Logbuch (cText & ", !!!Fehler: P oder N im Datensatz - mit 0 ersetzt")
    End If
    
    On Error Resume Next
    If aDaten(3) = suchDatum Then
      such = "tag = '" & str2date(aDaten(3)) & "'" _
        & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
        & " and kasse = '" & Right(aDaten(2), 2) & "'" _
        & " and artikelnummer = '" & Right(aDaten(5), 5) & "'"
      With rs
      .Filter = such
      If .RecordCount = 0 Then
        .AddNew
        !Filiale = "90" & Left(aDaten(2), 2)
        !Kasse = Right(aDaten(2), 2)
        !datum = aDaten(3) 'cDatum
        !Zeit = aDaten(4)   'cZeit
        !Artikelnummer = Right(aDaten(5), 5)
        !Retour = CDbl(aDaten(6) / 1000)
        !Artikelbezeichnung = Left(conANSI(aDaten(7)), 32)
        !Tag = str2date(aDaten(3))
        !Artikel = Right(aDaten(5), 5) & " " & Left(conANSI(aDaten(7)), 16)
        StatusAnzeige (cFile & ":Satz neu     " & cText)     ' anzeige des Fortschritts
      Else
        cRetour = !Retour
        !Retour = CDbl(aDaten(6) / 1000)
        If cRetour <> 0 And cRetour <> !Retour Then
          StatusAnzeige (cFile & ":Satz update  " & cText)     ' anzeige des Fortschritts
          Logbuch ("!!! Datensatz geändert: Retour alt=" & cRetour & " Neuer Satz: " & cText)
        Else
          StatusAnzeige (cFile & ":Satz ergänzt " & cText)     ' anzeige des Fortschritts
        End If
      End If
      .Update
      .Filter = ""
      End With
      Fehlertest
    Else
        Logbuch ("!!! - falsches Datum: " & conANSI(cText))
    End If
  Loop
  rs.Close
  Close #1   ' Datei schließen.
  cND = ""
  If nDoppelt > 0 Then
    cND = nDoppelt & " Datensätze nicht übernommen"
  End If
  Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
  nDoppelt = 0
  StatusAnzeige ("return.sal sichern")
  FileCopy cFile, cSichPfad & "\return.sal"
  Kill cFile
Else
  Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
End If
End Sub

' Inventur einlesen
Sub InventurImport()
  cFile = cDatenpfad & "\inventur.sal"
  cTxt = ""
  cStatus = cFile
  cStatus.Refresh
  If Dir(cFile) <> "" Then
    Set cmd.ActiveConnection = con
    cmd.CommandText = "Select * from TabKakomArtikelUmsatz" _
                  & " WHERE tag = '" & str2date(suchDatum) & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenDynamic, adLockOptimistic
    Open cFile For Input As #1   ' Datei öffnen.
    Logbuch ("    Beginn: " & cFile)
    Do While Not EOF(1)        ' Schleife bis Dateiende.
      Line Input #1, cText    ' Zeile in Variable einlesen.
      '0,9021,0000,19011999,0313,00000000000200,+0000001000,"* BESTELLUNG *                  "
      'flag,Fil,kasse,datum,hhmm,artNr,InventurMenge,ArtikelBezeichnung
      aDaten = Split(cText, ",")
      aDaten(2) = Right(aDaten(2), 4)
      If aDaten(3) = suchDatum Then
        If InStr(1, aDaten(6), "P", vbTextCompare) > 0 Or InStr(1, aDaten(6), "N", vbTextCompare) > 0 Then
          aDaten(6) = Replace(aDaten(6), "P", "0")
          aDaten(6) = Replace(aDaten(6), "N", "0")
          Logbuch (cText & ", !!!Fehler: P oder N im Datensatz - mit 0 ersetzt")
        End If
        On Error Resume Next
        With rs
        such = "tag = '" & str2date(aDaten(3)) & "'" _
            & " and filiale = '90" & Left(aDaten(2), 2) & "'" _
            & " and kasse = '" & Right(aDaten(2), 2) & "'" _
            & " and artikelnummer = '" & Right(aDaten(5), 5) & "'"
        .Filter = such
        If .RecordCount = 0 Then
          .AddNew
          !Filiale = "90" & Left(aDaten(2), 2)
          !Kasse = Right(aDaten(2), 2)
          !datum = aDaten(3) 'cDatum
          !Zeit = aDaten(4)   'cZeit
          !Inventurmenge = CDbl(aDaten(6) / 1000)
          !Artikelnummer = Right(aDaten(5), 5)
          !Artikelbezeichnung = Left(conANSI(aDaten(7)), 32)
          !Tag = str2date(aDaten(3))
          !Artikel = Right(aDaten(5), 5) & " " & Left(conANSI(aDaten(7)), 16)
          StatusAnzeige (cFile & ":Satz neu     " & cText)     ' anzeige des Fortschritts
        Else
          cInventur = !Inventurmenge
          !Inventurmenge = CDbl(aDaten(6) / 1000)
          If cInventur <> 0 And cInventur <> !Inventurmenge Then
            StatusAnzeige (cFile & ":Satz update  " & cText)     ' anzeige des Fortschritts
            Logbuch ("!!! Datensatz geändert: Inventurmenge alt=" & cInventur & " Neuer Satz: " & cText)
          Else
            StatusAnzeige (cFile & ":Satz ergänzt " & cText)     ' anzeige des Fortschritts
          End If
        End If
        .Update
        .Filter = ""
        End With
        Fehlertest
      Else
        Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
      End If
    Loop
    rs.Close
    Close #1   ' Datei schließen.
    cND = ""
    If nDoppelt > 0 Then
      cND = nDoppelt & " Datensätze nicht übernommen"
    End If
    Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
    nDoppelt = 0
    StatusAnzeige ("inventur.sal sichern")
    FileCopy cFile, cSichPfad & "\inventur.sal"
    Kill cFile
  Else
    Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
  End If
End Sub

' Warengruppen einlesen
Sub WarengruppenImport()
  cFile = cDatenpfad & "\dept.sal"
  cStatus = cFile
  cStatus.Refresh
  If Dir(cFile) <> "" Then
    Set cmd.ActiveConnection = con
    cmd.CommandText = "Select * from TabKakomWarengruppen" _
                   & " WHERE tag = '" & str2date(suchDatum) & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenDynamic, adLockOptimistic
    Open cFile For Input As #1   ' Datei öffnen.
    Logbuch ("    Beginn: " & cFile)
    Do While Not EOF(1)        ' Schleife bis Dateiende.
      Line Input #1, cText     ' Zeile in Variable einlesen.
      '0,0000,2101,17012000,0314,0001,+0000066500,+0000032798,"ýBACKWARE
      aDaten = Split(cText, ",")
      aDaten(2) = Right(aDaten(2), 4)
      If aDaten(3) = suchDatum Then
        On Error Resume Next
        such = "tag = '" & str2date(aDaten(3)) & "'" _
           & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
           & " and kasse = '" & Right(aDaten(2), 2) & "'" _
           & " and Warengruppe = '" & Right(aDaten(5), 2) & "'"
        With rs
        .Filter = such
        If .RecordCount = 0 Then
          .AddNew
          !Filiale = "90" & Left(aDaten(2), 2)
          !Kasse = Right(aDaten(2), 2)
          !datum = aDaten(3)  'Datum
          !Zeit = aDaten(4)   'Zeit
          !Warengruppe = Right(aDaten(5), 2)
          !Verkaufsmenge = CDbl(aDaten(6) / 1000)
          !Umsatz = CDbl(aDaten(7) / 100)
          !WarengruppeBez = conANSI(aDaten(8))
          !WG = Right(aDaten(5), 2) & " " & Left(conANSI(aDaten(8)), 13)
          !Tag = str2date(aDaten(3))
          StatusAnzeige (cFile & ":Satz neu     " & cText)     ' anzeige des Fortschritts
        Else
          cVerkaufsmenge = !Verkaufsmenge
          cUmsatz = !Umsatz
          !Verkaufsmenge = CDbl(aDaten(6) / 1000)
          !Umsatz = CDbl(aDaten(7) / 100)
          If cVerkaufsmenge <> 0 And cVerkaufsmenge <> !Verkaufsmenge Then
            StatusAnzeige (cFile & ":Satz update  " & cText)     ' anzeige des Fortschritts
            Logbuch ("!!! Datensatz geändert: Umsatz alt= " & cUmsatz & " Verkaufsmenge alt=" & cVerkaufsmenge & " Neuer Satz: " & cText)
          Else
            StatusAnzeige (cFile & ":Satz ergänzt " & cText)     ' anzeige des Fortschritts
          End If
        End If
        .Update
        .Filter = ""
        End With
        Fehlertest
      Else
        Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
      End If
    Loop
    rs.Close
    Close #1   ' Datei schließen.
    cND = ""
    If nDoppelt > 0 Then
      cND = nDoppelt & " Datensätze nicht übernommen"
    End If
   Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
   nDoppelt = 0
   StatusAnzeige ("dept.sal sichern")
   FileCopy cFile, cSichPfad & "\dept.sal"
   Kill cFile
 Else
   Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
 End If
End Sub

' Frequenzbericht einlesen
Sub FrequenzBerichtImport()
Dim rstFrequenz As Recordset
cFileName = "\hourly.sal"
cFile = cDatenpfad & cFileName
cStatus = cFile
cStatus.Refresh
If Dir(cFile) <> "" Then
  Set cmd.ActiveConnection = con
  Set rs = New ADODB.Recordset
  cmd.CommandText = "Select * from TabKakomFrequenz" _
                      & " WHERE tag = '" & str2date(suchDatum) & "'"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenDynamic, adLockOptimistic
  Open cFile For Input As #1   ' Datei öffnen.
  Logbuch ("    Beginn: " & cFile)

  Do While Not EOF(1)        ' Schleife bis Dateiende.
    Line Input #1, cText     ' Zeile in Variable einlesen.
    'flag,,Fil,kasse,datum,hhmm,vonhhmm,bishhmm,Verkaufsmenge,Umsatz
    aDaten = Split(cText, ",")
    aDaten(2) = Right(aDaten(2), 4)
    On Error Resume Next
    If aDaten(3) = suchDatum Then
      such = " tag = '" & str2date(aDaten(3)) & "'" _
         & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
         & " AND kasse = '" & Right(aDaten(2), 2) & "'" _
         & " AND vonZeit = '" & testZeit(aDaten(5)) & "'"
      With rs
      .Filter = such
      If .RecordCount = 0 Then
        .AddNew
        !Filiale = "90" & Left(aDaten(2), 2)
        !Kasse = Right(aDaten(2), 2)
        !datum = aDaten(3)  'Datum
        !Zeit = aDaten(4)   'Zeit
        !vonZeit = testZeit(aDaten(5))
        !bisZeit = testZeit(aDaten(6))
        !Verkaufsmenge = CDbl(aDaten(7) / 1000)
        !Umsatzbetrag = CDbl(aDaten(8) / 100)
        !Tag = str2date(aDaten(3))
        StatusAnzeige (cFile & ":Satz neu     " & cText)     ' anzeige des Fortschritts
      Else
        cVerkaufsmenge = !Verkaufsmenge
        cUmsatzbetrag = !Umsatzbetrag
        !Verkaufsmenge = CDbl(aDaten(7) / 1000)
        !Umsatzbetrag = CDbl(aDaten(8) / 100)
        If cVerkaufsmenge <> 0 And cVerkaufsmenge <> !Verkaufsmenge Then
          StatusAnzeige (cFile & ":Satz update  " & cText)     ' anzeige des Fortschritts
          Logbuch ("!!! Datensatz geändert: Umsatz alt= " & cUmsatzbetrag & " Verkaufsmenge alt=" & cVerkaufsmenge & " Neuer Satz: " & cText)
        Else
          StatusAnzeige (cFile & ":Satz ergänzt " & cText)     ' anzeige des Fortschritts
        End If
      End If
     .Update
     .Filter = ""
     End With
     Fehlertest
   Else
     Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
   End If
 Loop
 rs.Close
 Close #1   ' Datei schließen.
 cND = ""
 If nDoppelt > 0 Then
   cND = nDoppelt & " Datensätze nicht übernommen"
 End If
 Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
 nDoppelt = 0
 StatusAnzeige (cFileName & " sichern")
 FileCopy cFile, cSichPfad & cFileName
 Kill cFile
Else
  Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
End If
End Sub

' Artikel-Frequenzbericht einlesen
Sub ArtikelFrequenzBerichtImport()
Dim rstFrequenz As Recordset
cFileName = "\hourlplu.sal"
cFile = cDatenpfad & cFileName
cStatus = cFile
cStatus.Refresh
If Dir(cFile) <> "" Then
  Set cmd.ActiveConnection = con
  cmd.CommandText = "Select * from TabKakomArtikelFrequenz " _
                & " WHERE tag = '" & str2date(suchDatum) & "'"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenDynamic, adLockOptimistic
  Open cFile For Input As #1   ' Datei öffnen.
  Logbuch ("    Beginn: " & cFile)

  Do While Not EOF(1)        ' Schleife bis Dateiende.
    Line Input #1, cText     ' Zeile in Variable einlesen.
    'flag,Fil, kasse,datum,hhmm,vonhhmm,bishhmm,Verkaufsmenge,Umsatz
    '0,0000,2901,09021999,1934,00000000000200,1700,1729,0000N15000,"* BESTELLUNG *
    aDaten = Split(cText, ",")
    ' P oder N aus Datensatz entfernen, da er sonst nicht übernommen wird
    aDaten(2) = Right(aDaten(2), 4)
    If InStr(1, aDaten(8), "P", vbTextCompare) > 0 Or InStr(1, aDaten(8), "N", vbTextCompare) > 0 Then
      aDaten(8) = Replace(aDaten(8), "P", "0")
      aDaten(8) = Replace(aDaten(8), "N", "0")
      Logbuch (cText & ", !!!Fehler: P oder N im Datensatz - mit 0 ersetzt")
    End If
  
    On Error Resume Next
    If aDaten(3) = suchDatum Then
      such = "tag = '" & str2date(aDaten(3)) & "'" _
         & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
         & " AND kasse = '" & Right(aDaten(2), 2) & "'" _
         & " AND Artikelnummer = '" & Right(aDaten(5), 5) & "'" _
         & " AND vonZeit = '" & testZeit(aDaten(6)) & "'"
      With rs
      .Filter = such
      If .RecordCount = 0 Then
        .AddNew
        !Tag = str2date(aDaten(3))
        !Filiale = "90" & Left(aDaten(2), 2)
        !Kasse = Right(aDaten(2), 2)
        !Artikelnummer = Right(aDaten(5), 5)
        !Artikelbezeichnung = conANSI(aDaten(9))
        !vonZeit = testZeit(aDaten(6))
        !bisZeit = testZeit(aDaten(7))
        !Verkaufsmenge = CDbl(aDaten(8) / 1000)             ' es sollen auch halbe Stücke berechnet werden 0,5
        !datum = aDaten(3)  'Datum
        !Zeit = aDaten(4)   'Zeit
        !Artikel = Right(aDaten(5), 5) & " " & Left(conANSI(aDaten(9)), 16)
        StatusAnzeige (cFile & ":Satz neu    " & cText)     ' anzeige des Fortschritts
      Else
        StatusAnzeige (cFile & ":Satz update " & cText)     ' anzeige des Fortschritts
        cVerkaufsmenge = !Verkaufsmenge
        !Verkaufsmenge = CDbl(aDaten(8) / 1000)
        If cVerkaufsmenge <> 0 And cVerkaufsmenge <> !Verkaufsmenge Then
          Logbuch ("!!! Datensatz geändert: Verkaufsmenge alt=" & cVerkaufsmenge & " Neuer Satz: " & cText)
        End If
      End If
      .Update
      .Filter = ""
      End With
      Fehlertest
    Else
      Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
    End If
 Loop
 Close #1   ' Datei schließen.
 rs.Close
 cND = ""
 If nDoppelt > 0 Then
   cND = nDoppelt & " Datensätze nicht übernommen"
 End If
 Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
 nDoppelt = 0
 StatusAnzeige (cFileName & " sichern")
 FileCopy cFile, cSichPfad & cFileName
 Kill cFile
Else
 Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
End If
End Sub

' Transaktionen einlesen
Sub TransaktionenImport()
Dim rstDB As Recordset
cFileName = "\transact.sal"
cFile = cDatenpfad & cFileName
cStatus = cFile
cStatus.Refresh
If Dir(cFile) <> "" Then
  Set cmd.ActiveConnection = con
  cmd.CommandText = "Select * from TabKakomTransaktion" _
                   & " WHERE tag = '" & str2date(suchDatum) & "'"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenDynamic, adLockOptimistic
  Open cFile For Input As #1   ' Datei öffnen.
  Logbuch ("    Beginn: " & cFile)

  Do While Not EOF(1)        ' Schleife bis Dateiende.
    Line Input #1, cText     ' Zeile in Variable einlesen.
    'flag,Fil,kasse,datum,hhmm,vonhhmm,bishhmm,Verkaufsmenge,Umsatz
    aDaten = Split(cText, ",")
    aDaten(2) = Right(aDaten(2), 4)
    On Error Resume Next

    If aDaten(3) = suchDatum Then
      such = "tag = '" & str2date(aDaten(3)) & "'" _
                   & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
                   & " and kasse = '" & Right(aDaten(2), 2) & "'" _
                   & " and TransaktCode = '" & aDaten(5) & "'"
      
      
      With rs
      .Filter = such
      If .RecordCount = 0 Then
        .AddNew
        !Filiale = "90" & Left(aDaten(2), 2)
        !Kasse = Right(aDaten(2), 2)
        !datum = aDaten(3)  'Datum
        !Zeit = aDaten(4)   'Zeit
        !transaktCode = aDaten(5)
        !Verkaufsmenge = CDbl(aDaten(6))
        !Umsatzbetrag = CDbl(aDaten(7) / 100)
        !TransaktBez = Left(conANSI(aDaten(8)), 32)
        !Tag = str2date(aDaten(3))
        !Transaktion = aDaten(5) & " " & Left(conANSI(aDaten(8)), 22)
         StatusAnzeige (cFile & ":Satz neu   " & cText)     ' anzeige des Fortschritts
      Else
        StatusAnzeige (cFile & ":Satz update " & cText)     ' anzeige des Fortschritts
        cVerkaufsmenge = !Verkaufsmenge
        cUmsatzbetrag = !Umsatzbetrag
        !Verkaufsmenge = CDbl(aDaten(6))
        !Umsatzbetrag = CDbl(aDaten(7) / 100)
        If cVerkaufsmenge <> 0 And cVerkaufsmenge <> !Verkaufsmenge Then
          Logbuch ("!!! Datensatz geändert:  Verkaufsmenge alt=" & cVerkaufsmenge & " Neuer Satz: " & cText)
        End If
      End If
      .Update
      .Filter = ""
      End With
      Fehlertest
    Else
      Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
    End If
 Loop
 rs.Close
 Close #1   ' Datei schließen.
 cND = ""
 If nDoppelt > 0 Then
   cND = nDoppelt & " Datensätze nicht übernommen"
 End If
 Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
 nDoppelt = 0
 StatusAnzeige (cFileName & " sichern")
 FileCopy cFile, cSichPfad & cFileName
 Kill cFile
Else
  Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
End If
End Sub

' KassiererBerichte einlesen
' Kassiereberichte stehen in cashier.sal und clerk.sal
' Die Namen der Kassierer werden aus der Tabelle PERSONAL geholt

Sub KassiererImport()
Dim rsPersonal As Recordset
cFileName = "\cashier.sal"
cFile = cDatenpfad & cFileName
cStatus = cFile
cStatus.Refresh
If Dir(cFile) <> "" Then
  Set cmd.ActiveConnection = con
  cmd.CommandText = "Select * from TabKakomKassierer" _
                    & " WHERE tag = '" & str2date(suchDatum) & "'"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenDynamic, adLockOptimistic
  
  Set rsP = New ADODB.Recordset             '
  rsP.CursorLocation = adUseClient
  rsP.Open "Select * from TabHelpImpKassierName3x1", con, adOpenDynamic, adLockOptimistic
  Open cFile For Input As #1   ' Datei öffnen.
  Logbuch ("    Beginn: " & cFile)

  Do While Not EOF(1)        ' Schleife bis Dateiende.
    Line Input #1, cText     ' Zeile in Variable einlesen.
    aDaten = Split(cText, ",")
    aDaten(2) = Right(aDaten(2), 4)
    
    cSuchPersonal = "filiale = '90" & Left(aDaten(2), 2) & "' and kasseNr='" & Right(aDaten(2), 2) _
                 & "' and KassiererNr ='" & Right(aDaten(5), 2) & "'"

    With rsP         ' Datensatzgruppe auffüllen.
      .Filter = cSuchPersonal
      If .RecordCount < 1 Then
        Logbuch ("!!! Fehler: " & cSuchPersonal & " in Datei PERSONAL nicht gefunden: Personalnummer auf '0' gesetzt")
        Logbuch ("           Original-Datensatz: " & cText & "Datei: " & cFile)
        'cName = "??"
        cPersnummer = "0"
      Else
        'cName = !PersonalName
        cPersnummer = !PersonalNr
      End If
      .Filter = ""
    End With
        
    On Error Resume Next
    If aDaten(3) = suchDatum Then
      such = "tag = '" & str2date(aDaten(3)) & "'" _
                    & " AND Filiale = '90" & Left(aDaten(2), 2) & "'" _
                    & " and kasse = '" & Right(aDaten(2), 2) & "'" _
                    & " AND kassierer = '" & Right(aDaten(5), 2) & "'" _
                    & " AND TransaktionCode = '" & aDaten(6) & "'"

      With rs
      .Filter = such
      If .RecordCount < 1 Then
        .AddNew
        !Tag = str2date(aDaten(3))
        !Filiale = "90" & Left(aDaten(2), 2)
        !Kasse = Right(aDaten(2), 2)
        !kassierer = Right(aDaten(5), 2)
        !TransaktionCode = aDaten(6)
        !datum = aDaten(3)  'Datum
        !Zeit = aDaten(4)   'Zeit
        !Verkaufsmenge = CDbl(aDaten(7))
        !Umsatzbetrag = CDbl(aDaten(8) / 100)
        !TransaktBez = Left(conANSI(aDaten(9)), 32)
        '!KassiererName = cName                     ' in Datenbank nicht mehr vorhanden
        !Transaktion = aDaten(6) & " " & Left(conANSI(aDaten(9)), 22)
        !PersonalNr = cPersnummer
         StatusAnzeige (cFile & ":Satz neu    " & cText)     ' anzeige des Fortschritts
       Else
         StatusAnzeige (cFile & ":Satz update " & cText)     ' anzeige des Fortschritts
         cVerkaufsmenge = !Verkaufsmenge
         !Verkaufsmenge = CDbl(aDaten(7))
         !Umsatzbetrag = CDbl(aDaten(8) / 100)
         !TransaktBez = Left(conANSI(aDaten(9)), 32)
         !Transaktion = aDaten(6) & " " & Left(conANSI(aDaten(9)), 22)
         If cVerkaufsmenge <> 0 And cVerkaufsmenge <> !Verkaufsmenge Then
           Logbuch ("!!! Datensatz geändert: Verkaufsmenge alt =" & cVerkaufsmenge & " neu=" & !Verkafsmenge & " Neuer Satz: " & cText)
         End If
         If !PersonalNr = "0" Then
           !PersonalNr = cPersnummer
           '!KassiererName = cName
           Logbuch ("!!! Datensatz geändert: Personalnummer  von 0 auf " & cPersnummer)
         End If
       End If
       .Update
       .Filter = ""
       End With
       Fehlertest
     Else
       Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
     End If
  Loop
  rs.Close
  Close #1   ' Datei schließen.
  rsP.Close
  cND = ""
  If nDoppelt > 0 Then
    cND = nDoppelt & " Datensätze nicht übernommen-> " & rstDB.Name
  End If
  Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
  nDoppelt = 0
  StatusAnzeige (cFileName & " sichern")
  FileCopy cFile, cSichPfad & cFileName
  Kill cFile
Else
  Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
End If
End Sub

' BedienerBerichte einlesen
Sub BedienerImport()
cFileName = "\clerk.sal"
cFile = cDatenpfad & cFileName
cStatus = cFile
cStatus.Refresh
If Dir(cFile) <> "" Then
  Set cmd.ActiveConnection = con
  cmd.CommandText = "Select * from TabKakomKassierer" _
                    & " WHERE tag = '" & str2date(suchDatum) & "'"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenDynamic, adLockOptimistic
  rsP.CursorLocation = adUseClient
  rsP.Open "Select * from TabHelpImpKassierName3x1", con, adOpenDynamic, adLockOptimistic

  Open cFile For Input As #1   ' Datei öffnen.
  Logbuch ("    Beginn: " & cFile)

  Do While Not EOF(1)        ' Schleife bis Dateiende.
    Line Input #1, cText     ' Zeile in Variable einlesen.
    aDaten = Split(cText, ",")
    aDaten(2) = Right(aDaten(2), 4)
   
    cSuchPersonal = "filiale = '90" & Left(aDaten(2), 2) & "' and kasseNr='" & Right(aDaten(2), 2) _
                 & "' and KassiererNr ='" & Right(aDaten(5), 2) & "'"
    With rsP            ' Datensatzgruppe auffüllen.
      .Filter = cSuchPersonal
      If .RecordCount < 1 Then
        Logbuch ("!!! Fehler: " & cSuchPersonal & " in Datei TabHelpImpKassierName3x1 nicht gefunden: Name auf '??' gesetzt")
        Logbuch ("           Orginal-Datensatz: " & cText & "Datei: " & cFile)
        'cName = "??"
        cPersnummer = "0"
      Else
        'cName = !PersonalName
        cPersnummer = !PersonalNr
      End If
      .Filter = ""
    End With
    On Error Resume Next
    If aDaten(3) = suchDatum Then
      such = "tag = '" & str2date(aDaten(3)) & "'" _
        & " AND filiale = '90" & Left(aDaten(2), 2) & "'" _
        & " and kasse = '" & Right(aDaten(2), 2) & "'" _
        & " AND kassierer = '" & Right(aDaten(5), 2) & "'" _
        & " AND TransaktionCode = '" & aDaten(6) & "'"
      With rs
      .Filter = such
      If .RecordCount < 1 Then
        .AddNew
        !Tag = str2date(aDaten(3))
        !Filiale = "90" & Left(aDaten(2), 2)
        !Kasse = Right(aDaten(2), 2)
        !datum = aDaten(3)  'Datum
        !Zeit = aDaten(4)   'Zeit
        !kassierer = Right(aDaten(5), 2)
        !TransaktionCode = aDaten(6)
        !Verkaufsmenge = CDbl(aDaten(7))
        !Umsatzbetrag = CDbl(aDaten(8) / 100)
        !TransaktBez = Left(conANSI(aDaten(9)), 32)
        !Transaktion = aDaten(6) & " " & Left(conANSI(aDaten(9)), 22)
        '!KassiererName = cName
        !PersonalNr = cPersnummer
        StatusAnzeige (cFile & ":Satz neu    " & cText)     ' anzeige des Fortschritts
      Else
        StatusAnzeige (cFile & ":Satz update " & cText)     ' anzeige des Fortschritts
        cVerkaufsmenge = !Verkaufsmenge
        !Verkaufsmenge = CDbl(aDaten(7))
        !Umsatzbetrag = CDbl(aDaten(8) / 100)
        !TransaktBez = Left(conANSI(aDaten(9)), 32)
        !Transaktion = aDaten(6) & " " & Left(conANSI(aDaten(9)), 22)
        If cVerkaufsmenge <> 0 And cVerkaufsmenge <> !Verkaufsmenge Then
          Logbuch ("!!! Datensatz geändert: Verkaufsmenge alt =" & cVerkaufsmenge & " neu=" & !Verkafsmenge & " Neuer Satz: " & cText)
        End If
        If !PersonalNr = "0" Then
          !PersonalNr = cPersnummer
          '!KassiererName = cName
          Logbuch ("!!! Datensatz geändert: Nummer von 0 auf " & cPersnummer)
        End If
      End If
      .Update
      .Filter = ""
      End With
      Fehlertest
    Else
      Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
    End If
  Loop
  Close #1   ' Datei schließen.
  rsP.Close
  rs.Close
  cND = ""
  If nDoppelt > 0 Then
    cND = nDoppelt & " Datensätze nicht übernommen-> " & rstDB.Name
  End If
  Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
  nDoppelt = 0
  StatusAnzeige (cFileName & " sichern")
  FileCopy cFile, cSichPfad & cFileName
  Kill cFile
Else
  Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
End If
End Sub
Sub Level4Import()
  cFileName = "\level4.sal"
  cFile = cDatenpfad & cFileName
  cStatus = cFile
  cStatus.Refresh
  cFlag = "X"
  If Dir(cFile) <> "" Then
    Set cmd.ActiveConnection = con
    cmd.CommandText = "Select * from TabKakomArtikelUmsatz" _
                    & " WHERE tag = '" & str2date(suchDatum) & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenDynamic, adLockOptimistic
  
    Open cFile For Input As #1   ' Datei öffnen.
    Do While Not EOF(1)        ' Schleife bis Dateiende.
      Line Input #1, cText    ' Zeile in Variable einlesen.
'0,0000,2101,17012000,0314,00000000001444,+0000001000,"Bierkruste ofenf                "
'flag,Fil,kasse,datum,hhmm,artNr,Guschrift,ArtBez
      aDaten = Split(cText, ",")
      aDaten(2) = Right(aDaten(2), 4)
      On Error Resume Next
      If aDaten(3) = suchDatum Then
        such = "tag = '" & str2date(aDaten(3)) & "'" _
           & " AND datum = '" & aDaten(3) & "'" _
           & " and filiale = '90" & Left(aDaten(2), 2) & "'" _
           & " and kasse = '" & Right(aDaten(2), 2) & "'" _
           & " and artikelnummer = '" & Right(aDaten(5), 5) & "'"
        With rs
        .Filter = such
        If .RecordCount < 1 Then
          .AddNew
          !Filiale = "90" & Left(aDaten(2), 2)
          !Kasse = Right(aDaten(2), 2)
          !datum = aDaten(3) 'cDatum
          !Zeit = aDaten(4)   'cZeit
          !Artikelnummer = Right(aDaten(5), 5)
          !Gutschrift = CDbl(aDaten(6) / 1000)
          !Artikelbezeichnung = Left(conANSI(aDaten(7)), 32)
          !Tag = str2date(aDaten(3))
          !Artikel = Right(aDaten(5), 5) & " " & Left(conANSI(aDaten(7)), 16)
          StatusAnzeige (cFile & ":Satz neu    " & cText)     ' anzeige des Fortschritts
        Else
          StatusAnzeige (cFile & ":Satz update " & cText)     ' anzeige des Fortschritts
          cGutschrift = !Gutschrift
          !Gutschrift = CDbl(aDaten(6) / 1000)
          If cGutschrift <> 0 And cGutschrift <> !Gutschrift Then
            Logbuch ("!!! Datensatz geändert: Gutschrift alt=" & cGutschrift & " Neuer Satz: " & cText)
          End If
        End If
        .Update
        .Filter = ""
        End With
        Fehlertest
        If cFlag = "X" Then              ' begin bearb.. hier ausnahmsweise
          cFlag = "N"                    ' zeigen damit Datum in Protokoll zuerst angezeigt wird
          Logbuch ("    Beginn: " & cFile)
        End If
      Else
        Logbuch ("!!! - falsches Datum (nicht übernommen): " & conANSI(cText))
      End If
    Loop
    Close #1
    rs.Close
    cND = ""
    If nDoppelt > 0 Then
      cND = nDoppelt & " Datensätze nicht übernommen"
    End If
    Logbuch ("    Ende  : " & pad(cFile, 25) & "  " & cND)
    nDoppelt = 0
    StatusAnzeige (cFileName & " sichern")
    FileCopy cFile, cSichPfad & "\level4.sal"
    Kill cFile
  Else
    Logbuch ("!!! Fehler: " & pad(cFile, 25) & " nicht gefunden")
  End If
End Sub
' Umwandlung von DOS - Zeichen nach Win-Zeichen
Function conANSI(cTxt)
   cTextx = Replace(cTxt, Chr(34), "")
   cTextx = Replace(cTextx, Chr(132), "ä")
   cTextx = Replace(cTextx, Chr(148), "ö")
   cTextx = Replace(cTextx, Chr(129), "ü")
   cTextx = Replace(cTextx, Chr(225), "ß")
   cTextx = Replace(cTextx, Chr(142), "Ä")
   cTextx = Replace(cTextx, Chr(153), "Ö")
   cTextx = Replace(cTextx, Chr(154), "Ü")
   cTextx = Replace(cTextx, Chr(250), "")
   conANSI = Replace(cTextx, Chr(253), "")   ' ² Steuerzeichen nach leer
End Function

Private Sub Form_Unload(Cancel As Integer)
'  Print #2, "Ende Kakom import "
'  Close #2
'  Close #3
End Sub

'----( Datum ins protokoll und den Sicherungspfad einstellen )----
Sub DatumTest(cD As String)
  If cDatumSQL = "X" Then                     ' für den Ersteintrag
    cDatumSQL = cD
   ' Logbuch (" ")
   ' Logbuch (" ")
   ' Logbuch (" ")
    Logbuch ("    --- gefundenes Datum in " & cFile & " ist: " & cD & " -----------")
    csich = GetProfile("SichPfad", App.Path) & "\" & cKasse      ' Unterordner aus Parameter
    If Dir(csich, vbDirectory) = "" Then
      MkDir (csich)
    End If
    cSichPfad = csich & "\" & Right(cD, 4) & Mid$(cD, 3, 2) & Left(cD, 2)
    If Dir(cSichPfad, vbDirectory) = "" Then
      Logbuch ("    --- sichern nach " & cSichPfad & " -----------")
      MkDir (cSichPfad)
    End If
  End If
  If cD <> "" Then
    If cD <> cDatumSQL Then
      Logbuch ("!!! - falsches Datum: " & conANSI(cText))
    End If
  End If
End Sub
  
Sub Fehlertest()
  aa = Err.Number
  aa = Err.Description
  If Err.Number <> 0 And Err.Number <> 3021 Then
     If Err.Number = 3022 Then                    ' Datensätze bereits vorhanden
       If cFileTmp <> cFile Then
         Print #8, " "
         Print #8, " "
         Print #8, cFile
         cFileTmp = cFile
       End If
       Print #8, cText
       nDoppelt = nDoppelt + 1
       nDoppeltGes = nDoppeltGes + 1
     Else
       Logbuch ("!!!Fehler: " & Err.Description & Err.Number)
       Logbuch ("    -> " & cText)
     End If
     Err.Clear
  End If
End Sub

Function Logbuch(cText As String)
     Print #2, cText
     Print #4, cText        ' Tagesprotokoll
End Function

Function str2date(cD)
  str2date = CDate(Left(cD, 2) & "." & Mid(cD, 3, 2) & "." & Right(cD, 4))
End Function
Function str2datum(cD)
  str2datum = "#" & Left(cD, 2) & "/" & Mid(cD, 3, 2) & "/" & Right(cD, 4) & "#"
End Function
  
Function StatusAnzeige(cText As String)
  cStatus = cText
  cStatus.Refresh
End Function

Function fFileKopie(quellFile As String, zielfile As String)
  If Dir(sichfile & "\" & zielfile) = zielfile Then
    zielfile = zielfile & "(1)"
  End If
  FileCopy quellFile, cSichPfad & "\" & zielfile
End Function

Function aktuellesDatum()
' es wird im datenpfad irgendeine *.sal-Datei gesucht.
' deren Datum ist dann Grundlage für die weitere Arbeit des Programms

cFile = cDatenpfad & "\*.sal"
cFile = Dir(cFile)

If cFile <> "" Then
  Open cDatenpfad & "\" & cFile For Input As #1   ' Datei öffnen.
  'Do While Not EOF(1)        ' Schleife bis Dateiende.
  For n = 0 To 1
    Line Input #1, cText    ' Zeile in Variable einlesen.
    aDaten = Split(cText, ",")
    ' nachfolgenden Datensatz testen, ob gleiches Datum. Wenn nicht Fehlermeldung
    ' aber weiterarbeiten
    If n = 1 And aktuellesDatum <> aDaten(3) Then
      Logbuch ("!!! Fehler: 2 verschiedene Datum in " & cFile & " gefunden!")
      Logbuch ("            Datum auf " & aDaten(3) & " gesetzt!")
    End If
    aktuellesDatum = aDaten(3)
  Next n
  Close #1
  DatumTest (aktuellesDatum)        ' Standard Datumtest, wie bei allen Files
 Else
   ' fehler
   Logbuch ("!!! Fehler: " & cFile & " nicht gefunden : kein Datum -> Abbruch")
   MacheFlagDateiAus                 ' Fertig.txt wird gelöscht -> alles OK
   End
   Unload Me
 End If
End Function
Function pad(cS As String, nL As Integer)
  If Len(cS) < nL Then
    pad = cS & Space(nL - Len(cS))
  Else
    pad = cS
  End If
End Function

Function testZeit(cZeit)
' immer in das Format 00:00:00 wandeln
' falls in der Zeit nur 9999 stehen (bei Kassenausfall), dann 23:59:00 schreiben
' Grund: einfaches Konvertieren ins Datum/Zeit-Format
  If cZeit = "9999" Then
    testZeit = "23:59:00"
  Else
    testZeit = Left(cZeit, 2) & ":" & Right(cZeit, 2) & ":00"
  End If
End Function

