Attribute VB_Name = "Prog2"
Option Explicit
Public Grundweg As String
Public EnPicture As Boolean
Public Type Testhash
Hashvalue As Long
Hashvalue1 As Long
Hashvalue2 As Long
Name As String
Name1 As String
Name2 As String
End Type

Public Sub SucheIDNames()
Dim Testname As String
Dim Menge As Long
Dim MengeTTL As Long
Dim MengeNames As Long
Dim i As Long
Dim z As Long
Dim Testhash As Long
Dim Testhashes() As Testhash
Dim cntopen As Long
Dim cntpfad As String
Dim Inhalt() As Byte
Dim Inhaltstring As String
Dim Test As Boolean
Dim IDgerade As String
Dim Standgerade As Long
Dim Speicherpfad As String
Dim existiert As Boolean
Dim sas As Long
Dim TName As String

MengeTTL = UBound(TTLBNames)
MengeNames = UBound(IDNames)
If NumCnt = 0 Then
cntpfad = Grundweg & "\" & ProjektName & ".cnt"
If FileExists(cntpfad) Then
NumCnt = 1
ReDim Namecnt(1)
Namecnt(1) = ProjektName & ".cnt"
End If
End If
If NumCnt > 0 Then
cntopen = FreeFile
For i = 1 To NumCnt
Standgerade = 1
existiert = False
cntpfad = Grundweg & "\" & Namecnt(i)
If InStr(Namecnt(i), "\") Then
sas = InStrRev(Namecnt(i), "\")
TName = Mid(Namecnt(i), sas + 1)
Speicherpfad = Pfad & "\" & TName
Else
Speicherpfad = Pfad & "\" & Namecnt(i)
End If
If FileExists(cntpfad) Then
existiert = True
Else
sas = InStrRev(Namecnt(i), "\")
TName = Mid(Namecnt(i), sas + 1)
cntpfad = Grundweg & "\" & TName
If FileExists(cntpfad) Then
existiert = True
End If
End If
If existiert = True Then
Open cntpfad For Binary As cntopen
ReDim Inhalt(LOF(cntopen))
Get cntopen, , Inhalt
Inhaltstring = StrConv(Inhalt, vbUnicode)
Do While Test = False
Test = Splitcnt(Inhaltstring, Standgerade, IDgerade)
If IDgerade <> "" Then
Testhash = Hashing(IDgerade)
For z = 0 To MengeNames
If Namegefunden(z) = False Then
If Contexts(z).Hashvalue = Testhash Then
IDNames(z) = IDgerade
Namegefunden(z) = True
End If
End If
Next z
End If
Loop
Close cntopen
Open Speicherpfad For Binary As cntopen
Put cntopen, , Inhalt
Close cntopen
End If
Next i
End If

ReDim Testhashes(MengeTTL)
Form1.Label1.Caption = "Testing IDNames"
DoEvents
For i = 1 To MengeTTL
Testname = Replace(TTLBNames(i).Topictitle, " ", "_")
If Testname <> "" Then
Testhashes(i).Hashvalue = Hashing(TTLBNames(i).Topictitle)
Testhashes(i).Name = TTLBNames(i).Topictitle
Testhashes(i).Hashvalue1 = Hashing(Testname)
Testhashes(i).Name1 = Testname
End If
Next i

For i = 0 To MengeNames
If Namegefunden(i) = False Then
For z = 1 To MengeTTL
If Contexts(i).Hashvalue = Testhashes(z).Hashvalue Then
IDNames(i) = Testhashes(z).Name
Namegefunden(i) = True
End If
Next z
End If
Next i
End Sub
Public Function FilenameFromPath(ByVal sPfad As String) As String
Dim Standpunkt As Long
Dim i As Long
Do While i <> -1
i = InStr(i + 1, sPfad, "\")
If i = 0 Then
i = -1 'Ende
Else
Standpunkt = i
End If
Loop
FilenameFromPath = Mid(sPfad, Standpunkt + 1)
End Function

Public Function FindStringEnd(Dateistand As Long) As Long
Dim Hilfsbyte As Byte
Dim Hilfsstand As Long
Dim Hilfe As Long
Hilfsstand = Dateistand
Hilfsbyte = 1
Do While Hilfsbyte <> 0
Get Dateinummer, Hilfsstand, Hilfsbyte
Hilfsstand = Hilfsstand + 1
Loop
Hilfe = Hilfsstand - Dateistand
FindStringEnd = Hilfe
End Function

Public Function FileExists(ByVal Filename As String) As Boolean
Dim i As Integer
Err.Clear
On Error Resume Next
i = GetAttr(Filename)
  If Err.Number = 0 Then
    If Not (i And vbDirectory) Then FileExists = True
  End If
On Error GoTo 0
End Function


Private Function Splitcnt(Inhaltstring As String, Standgerade As Long, IDgerade As String)
Dim EndeZeile As Long
Dim Übergabe As Long

If Standgerade = Len(Inhaltstring) Then
Splitcnt = True
Standgerade = 0
IDgerade = ""
Exit Function
End If
If Mid(Inhaltstring, Standgerade, 1) = ":" Then
IDgerade = ""
Standgerade = InStr(Standgerade + 1, Inhaltstring, vbCrLf)
If Standgerade = 0 Then
Splitcnt = True 'Ende
Else
Standgerade = Standgerade + 2
Splitcnt = False
End If
Else
EndeZeile = InStr(Standgerade + 1, Inhaltstring, vbCrLf)
Übergabe = EndeZeile + 2
If EndeZeile = 0 Then
EndeZeile = Len(Inhaltstring)
Übergabe = EndeZeile
End If
IDgerade = Mid(Inhaltstring, Standgerade, EndeZeile - Standgerade)
If InStr(IDgerade, "=") Then
IDgerade = Mid(IDgerade, InStr(IDgerade, "=") + 1)
If InStr(IDgerade, "@") Then
IDgerade = Mid(IDgerade, 1, InStr(IDgerade, "@") - 1)
End If
If InStr(IDgerade, ">") Then
IDgerade = Mid(IDgerade, 1, InStr(IDgerade, ">") - 1)
End If
End If
Standgerade = Übergabe
End If
End Function

Public Sub RepairMacroString(NameString As String)
Dim TestString As String
Dim testlong As Long
Dim colorrgb As RGBColor
Dim Ende As Boolean
Dim Stand As Long
Dim Hilfsstring As String
Dim Zwischenstring1 As String
Dim Zwischenstring2 As String
Dim Zwischenstring3 As String
Dim woende As Long

Stand = 1
'SPC SetupColorMakro reparieren
If Left(NameString, 3) = "SPC" Or Left(NameString, 13) = "SetPopupColor" Then
If InStr(NameString, ",") = False Then 'RGB einsetzen
If Left(NameString, 3) = "SPC" Then
TestString = Mid(NameString, 5, Len(NameString) - 5)
If InStr(TestString, ")") Then
TestString = Left(TestString, InStr(TestString, ")") - 1)
End If
If IsNumeric(TestString) Then
testlong = CLng(TestString)
LongToRGB testlong, colorrgb
NameString = "SPC(" & colorrgb.Red & "," & colorrgb.Green & "," & colorrgb.Blue & ")"
End If
End If
If Left(NameString, 13) = "SetPopupColor" Then
TestString = Mid(NameString, 5, Len(NameString) - 5)
If IsNumeric(TestString) Then
testlong = CLng(TestString)
LongToRGB testlong, colorrgb
NameString = "SetPopupColor(" & colorrgb.Red & "," & colorrgb.Green & "," & colorrgb.Blue & ")"
End If
End If
End If
Else
If InStr(NameString, "SPC(") Then
Do While Ende = False
Stand = InStr(Stand, NameString, "SPC(")
If Stand = 0 Then
Ende = True
Else
Zwischenstring1 = Left(NameString, Stand - 1)
woende = InStr(Stand, NameString, ")")
Hilfsstring = Mid(NameString, Stand, woende - Stand)
Zwischenstring3 = Mid(NameString, woende)
If InStr(Hilfsstring, ",") Then
Else 'Long
Hilfsstring = Mid(Hilfsstring, 5)
If IsNumeric(Hilfsstring) Then
testlong = CLng(Hilfsstring)
LongToRGB testlong, colorrgb
Zwischenstring2 = "SPC(" & colorrgb.Red & "," & colorrgb.Green & "," & colorrgb.Blue
NameString = Zwischenstring1 & Zwischenstring2 & Zwischenstring3
End If
End If
Stand = Stand + 1
End If
Loop
End If

If InStr(NameString, "SetPopupColor(") Then
Do While Ende = False
Stand = InStr(Stand, NameString, "SetPopupColor(")
If Stand = 0 Then
Ende = True
Else
Zwischenstring1 = Left(NameString, Stand - 1)
woende = InStr(Stand, NameString, ")")
Hilfsstring = Mid(NameString, Stand, woende - Stand)
Zwischenstring3 = Mid(NameString, woende)
If InStr(Hilfsstring, ",") Then
Else 'Long
Hilfsstring = Mid(Hilfsstring, 15)
If IsNumeric(Hilfsstring) Then
testlong = CLng(Hilfsstring)
LongToRGB testlong, colorrgb
Zwischenstring2 = "SetPopupColor(" & colorrgb.Red & "," & colorrgb.Green & "," & colorrgb.Blue
NameString = Zwischenstring1 & Zwischenstring2 & Zwischenstring3
End If
End If
Stand = Stand + 1
End If
Loop
End If
End If
If Left(NameString, 3) = "PW(" Then
NameString = "PositionWindow(" & Mid(NameString, 4)
End If

If Left(NameString, 3) = "EP(" Then
NameString = "EF(" & Mid(NameString, 4)
End If
If Left(NameString, 12) = "ExecProgram(" Then
NameString = "EF(" & Mid(NameString, 4)
End If
NameString = Replace(NameString, "!EP(", "!EF(")
NameString = Replace(NameString, "!ExecProgram(", "!EF(")
NameString = Replace(NameString, Chr(96) & "EP(", Chr(96) & "EF(")
NameString = Replace(NameString, Chr(96) & "ExecProgram(", Chr(96) & "EF(")
NameString = Replace(NameString, Chr(34) & "EP(", Chr(34) & "EF(")
NameString = Replace(NameString, Chr(34) & "ExecProgram(", Chr(34) & "EF(")
If InStr(NameString, Chr(34)) Then
'Fehlerhafte Stringeingaben (z. B. New's)
TestAnf NameString
End If
End Sub

Public Sub TestMapNumber(Number As Double, Numberarray() As Double, NumberGefunden As Boolean)
Dim i As Long

For i = 0 To UBound(Numberarray)
If Number = Numberarray(i) Then
NumberGefunden = True
Number = Number + 1
TestMapNumber Number, Numberarray, NumberGefunden
Else
'OK
End If
Next i
End Sub

Public Sub TeileMacro(Macroword As String)
Dim woTeiler As Long
Dim HmAnfang As Long
Dim HmEnde As Long
Dim Teil1 As String
Dim Stand As Long
Dim Fertigstring As String
Dim Test As Long
Dim Teststand As Long
Dim woanfang As Long
Dim woende As Long
Dim Zwischenanfang As Long
Dim Zwischenende As Long
Dim Ende As Boolean

Stand = 1
Do While Ende = False
woTeiler = InStr(Stand, Macroword, ":")
If woTeiler = 0 Then
woTeiler = Len(Macroword) + 1
End If
Teil1 = Mid(Macroword, Stand, woTeiler - Stand)
HmAnfang = 0
HmEnde = 0
ZähleA_E Teil1, HmAnfang, HmEnde
If HmAnfang + Zwischenanfang = HmEnde + Zwischenende Then
Fertigstring = Fertigstring & Teil1 & vbCrLf
Zwischenanfang = 0
Zwischenende = 0
Else
Zwischenanfang = Zwischenanfang + HmAnfang
Zwischenende = Zwischenende + HmEnde
Fertigstring = Fertigstring & Teil1
End If
Stand = woTeiler + 1
If Stand > Len(Macroword) Then
Ende = True
Exit Do
End If
Loop
Macroword = Fertigstring
End Sub
Private Sub ZähleA_E(Macroteil As String, HmAnfang As Long, HmEnde As Long)
Dim Teststand As Long
Dim Test As Long
Dim woanfang As Long
Dim woende As Long

Teststand = 1
Test = 0
Do While Test = 0 'Anfang
woanfang = InStr(Teststand, Macroteil, "(")
If woanfang <> 0 Then
HmAnfang = HmAnfang + 1
Teststand = woanfang + 1
Else
Test = 1
Exit Do
End If
Loop
Test = 0
Do While Test = 0 'Ende
woende = InStr(Teststand, Macroteil, ")")
If woende <> 0 Then
HmEnde = HmEnde + 1
Teststand = woende + 1
Else
Test = 1
Exit Do
End If
Loop

End Sub

Private Sub TestAnf(MString As String)
Dim Anfang1 As Long
Dim Ende1 As Long
Dim TestString As String
Dim Stand As Long
Dim Fertigstring As String
Dim Ende As Boolean

Stand = 1
Do While Ende = False
Anfang1 = InStr(Stand, MString, Chr(34))
If Anfang1 <> 0 Then
Fertigstring = Fertigstring & Mid(MString, Stand, Anfang1 - Stand)
Ende1 = InStr(Anfang1 + 1, MString, Chr(34))
TestString = Mid(MString, Anfang1 + 1, Ende1 - Anfang1 - 1)
If InStr(TestString, "(") = False And InStr(TestString, ")") = False Then
TestString = Replace(TestString, "'", " ")
End If
Fertigstring = Fertigstring & Chr(34) & TestString & Chr(34)
Stand = Ende1 + 1
Else
Fertigstring = Fertigstring & Mid(MString, Stand)
Ende = True
Exit Do
End If
Loop
MString = Fertigstring
End Sub

Public Sub DBBtoDIB(bbits() As Byte, Bildbreiteright As Long, Bildhöhe As Long)
Dim Dibbreite As Long
Dim i As Long
Dim Übergabe() As Byte
Dim Bildbreite As Long
Dim BitZeile As Long

BitZeile = (UBound(bbits) + 1) \ Bildhöhe
If BitZeile Mod 4 <> 0 Then
   Bildbreite = BitZeile
    Dibbreite = Bildbreite * 3
    Dibbreite = Dibbreite \ 4
    Dibbreite = Dibbreite + 1
    Dibbreite = Dibbreite * 4
    Dibbreite = Dibbreite \ 2
    ReDim Übergabe(Dibbreite * Bildhöhe - 1)
  For i = 0 To Bildhöhe - 1
  CopyMemory Übergabe(i * Dibbreite), bbits(i * Bildbreite), Bildbreite
  Next i
  End If
ReDim bbits(UBound(Übergabe))
CopyMemory bbits(0), Übergabe(0), UBound(Übergabe) + 1
End Sub

