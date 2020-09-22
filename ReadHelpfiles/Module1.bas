Attribute VB_Name = "Module1"
Option Explicit
Public MoreRTF As Boolean
Public TabellenAlign As Integer
Public Scrollregion As Long
Public NonScrollregion As Long
Public ZuletztAbsatz As Boolean
Public lastcol As Long
Private IDString As String
Private jumprec As String
Public Mapstring As String
Public Deffont As Long
Public scaling As Long
Public rounderr As Long
Public fertigeIDs() As Boolean
Public strAlias As String
Public NOH As Long
Public MitteOffset As Long
Public ScrollorNoscroll As Boolean
Public anzviolas As Long
Public hasviolas As Boolean
Public Type VIOLA
Offsets As Long
Numbers As Long
End Type
Public Violafiles() As VIOLA
Public windownames() As String
Public anzwindows As Long
Private Jumpeben As String
Public Type BMTYPE
Type As Integer
Name As String
Transparent As Integer
End Type
Public anzbitmaps As Long
Public TopZähler As Long
Public Bitmaps() As BMTYPE
Public gefundeneID As Long
Public Topicanzahl As Long
Public IDNames() As String
Public Contexts() As CONTEXTLEAF
Public Optionsstring As String
Public Windowsstring As String
Public AktCharacter As Long
Public BrowseBlocknumber As Long
Public Browsenanzahl As Long
Public HasKKeywords As Boolean
Public HasAKeywords As Boolean
Public Makrostringfertig As String
Public Reserveoffset As Long
Public BrowseFalse As Boolean
Public NextGoodBrowseOffset As Long
Private Browsesequencenumber As Long
Public Browsen() As BrowseDescr
Public iscompress As Boolean
Public TTLBTreeOffsets() As Long
Public KWBTreeString()  As String
Public AWBTreeString()  As String
Public WoistKKeyword() As Long
Public WoistAKeyword() As Long
Public KWData() As Long
Public AWData() As Long
Private FontInUse As Long
Public Dateinummer As Long
Public Fonttable As String
Public colortablestring As String
Public font_descriptor() As FONTS
Public Bold As Boolean
Public Italic As Boolean
Public StrikeOut As Boolean
Public Underline As Boolean
Public DoubleUnderline As Boolean
Public SmallCaps As Boolean

Public Function Klebe(Bytefeld() As Byte, Standpunkttbhdr As Long, FileUsedspace As Long, alteTopicblocknr As Long, iscompress As Boolean) As Long
Dim testtbrhdr As TOPICBLOCKHEADER
Dim alteGröße As Long
Dim aktBlock As Long
Dim letzterBlock As Long
Dim Decgr As Long
Dim testbar() As Byte
Get Dateinummer, Standpunkttbhdr + 4096, testtbrhdr
aktBlock = alteTopicblocknr + 1
letzterBlock = testtbrhdr.LastTopicHeader \ 4096
If aktBlock < letzterBlock Then
Decgr = 4095
Else
Decgr = FileUsedspace - ((alteTopicblocknr * 4096) - 1)
End If
ReDim testbar(Decgr - 12)
Get Dateinummer, , testbar
If iscompress = True Then Decompress UBound(testbar) + 1, testbar
alteGröße = UBound(Bytefeld)
ReDim Preserve Bytefeld(alteGröße + UBound(testbar) + 1)
CopyMemory Bytefeld(alteGröße + 1), testbar(0), UBound(testbar) + 1
End Function
Public Sub DeleteDirectory(ByVal dir_name As String)
Dim file_name As String
Dim files As Collection
Dim i As Integer

    ' Get a list of files it contains.
    Set files = New Collection
    file_name = Dir$(dir_name & "\*.*", vbReadOnly + vbHidden + vbSystem + vbDirectory)
    Do While Len(file_name) > 0
        If (file_name <> "..") And (file_name <> ".") Then
            files.Add dir_name & "\" & file_name
        End If
        file_name = Dir$()
    Loop

    ' Delete the files.
    For i = 1 To files.count
        file_name = files(i)
        ' See if it is a directory.
        If GetAttr(file_name) And vbDirectory Then
            ' It is a directory. Delete it.
            DeleteDirectory file_name
        Else
            SetAttr file_name, vbNormal
            Kill file_name
        End If
    Next i

    ' The directory is now empty. Delete it.
    RmDir dir_name
End Sub


Public Function MakeRTFKopf() As String
MakeRTFKopf = "{\rtf1\ansi\deff" & Deffont & vbCrLf & Fonttable & vbCrLf & colortablestring & vbCrLf & "{\stylesheet{\fs20 \snext0 Normal;}" & vbCrLf & "}" & vbCrLf & "\pard\plain" & vbCrLf
TabellenAlign = 0
Aufmachen = Aufmachen + 1
End Function

Public Function TestID(Number As Integer, Standpunkt As Long, Dataarray() As Byte, Dateistring As String, Topictyp As Byte) As Integer
Dim Test As Integer
Dim testoffset As Long
Dim Testint As Integer
Dim tabstop As Integer
Dim tabtype As Integer
Dim WertInt As Integer
Dim Hilfsint As Integer
Dim wobinich As Long
Dim Länge As Integer
Dim i As Long
Dim HilfsbyteArray() As Byte
Dim BorderStruct As BORDERINFO
Dim NumberofTabstops As Integer

wobinich = Standpunkt
Hilfsint = Number
IDString = ""
Test = Hilfsint And 1
If Test = 1 Then 'Unknown
ReDim HilfsbyteArray(3)
CopyMemory HilfsbyteArray(0), Dataarray(wobinich), 4
wobinich = wobinich + Länge
End If

Test = Hilfsint And 2
If Test = 2 Then 'SpacingAbove
ReDim HilfsbyteArray(1)
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
WertInt = ReadCompSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
wobinich = wobinich + Länge
IDString = IDString & "\sb" & WertInt * scaling - rounderr
End If

Test = Hilfsint And 4
If Test = 4 Then 'SpacingBelow
ReDim HilfsbyteArray(1)
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
WertInt = ReadCompSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
IDString = IDString & "\sa" & WertInt * scaling - rounderr
wobinich = wobinich + Länge
End If

Test = Hilfsint And 8
If Test = 8 Then 'SpacingLines
ReDim HilfsbyteArray(1)
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
WertInt = ReadCompSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
IDString = IDString & "\sl" & WertInt * scaling - rounderr
wobinich = wobinich + Länge
End If

Test = Hilfsint And 16
If Test = 16 Then 'LeftIndent
ReDim HilfsbyteArray(1)
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
WertInt = ReadCompSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
IDString = IDString & "\li" & WertInt * scaling - rounderr
wobinich = wobinich + Länge
End If

Test = Hilfsint And 32 'RightIndent
If Test = 32 Then
ReDim HilfsbyteArray(1)
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
WertInt = ReadCompSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
IDString = IDString & "\ri" & WertInt * scaling - rounderr
wobinich = wobinich + Länge
End If

Test = Hilfsint And 64
If Test = 64 Then 'FirstlineIndent
ReDim HilfsbyteArray(1)
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
WertInt = ReadCompSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
IDString = IDString & "\fi" & WertInt * scaling - rounderr
wobinich = wobinich + Länge
End If

Test = Hilfsint And 128
If Test = 128 Then 'Unused
End If

Test = Hilfsint And 256
If Test = 256 Then 'Borderinfo
CopyMemory ByVal VarPtr(BorderStruct), Dataarray(wobinich), 3
Testint = BorderStruct.Borderparameters And 1
If Testint = 1 Then
IDString = IDString & "\box"
End If
Testint = BorderStruct.Borderparameters And 2
If Testint = 2 Then
IDString = IDString & "\brdrt"
End If
Testint = BorderStruct.Borderparameters And 4
If Testint = 4 Then
IDString = IDString & "\brdrl"
End If
Testint = BorderStruct.Borderparameters And 8
If Testint = 8 Then
IDString = IDString & "\brdrb"
End If
Testint = BorderStruct.Borderparameters And 16
If Testint = 16 Then
IDString = IDString & "\brdrr"
End If
Testint = BorderStruct.Borderparameters And 32
If Testint = 32 Then
IDString = IDString & "\brdrth"
Else
IDString = IDString & "\brdrs"
End If
Testint = BorderStruct.Borderparameters And 64
If Testint = 64 Then
IDString = IDString & "\brdrdb"
End If
If BorderStruct.BorderWidth <> 0 Then
End If
wobinich = wobinich + 3
End If

Test = Hilfsint And 512
If Test = 512 Then 'Tabinfo
ReDim HilfsbyteArray(1)
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
NumberofTabstops = scanint(HilfsbyteArray, Länge)
wobinich = wobinich + Länge
For i = 1 To NumberofTabstops
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
tabstop = ReadCompUnSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
wobinich = wobinich + Länge
Testint = tabstop And &H4000
If Testint = &H4000 Then
HilfsbyteArray(0) = Dataarray(wobinich)
HilfsbyteArray(1) = Dataarray(wobinich + 1)
tabtype = ReadCompUnSignShort(HilfsbyteArray(0), HilfsbyteArray(1), Länge)
Select Case tabtype
Case 1
IDString = IDString & "\tqr"
Case 2
IDString = IDString & "\tqc"
End Select
wobinich = wobinich + Länge
End If
IDString = IDString & "\tx" & ((tabstop And &H3FFF) * scaling - rounderr)
Next i
End If
Test = Hilfsint And 1024
If Test = 1024 Then
IDString = IDString & "\qr"
TabellenAlign = 1
End If

Test = Hilfsint And 2048
If Test = 2048 Then
IDString = IDString & "\qc"
TabellenAlign = 2
End If

Test = Hilfsint And 4096
If Test = 4096 Then
IDString = IDString & "\keep"
If ScrollorNoscroll = True Then
IDString = IDString & "\keepn"
End If
End If

TestID = wobinich - Standpunkt
If IDString <> "" Then
Select Case Topictyp
Case &H23
Dateistring = Dateistring & IDString
Case Else
If Right(Dateistring, 5) = "\pard" Then
Dateistring = Dateistring & IDString
Else
Dateistring = Dateistring & "\pard" & IDString
End If
ScrollorNoscroll = False
testoffset = AktCharacter + (Blocknummer * 32768)
If testoffset = NonScrollregion And ScrollorNoscroll = False Then
Dateistring = Dateistring & "\keepn "
ScrollorNoscroll = True
End If
End Select
End If
End Function

Public Function writeFormatcommand(Topictyp As Byte, Datenfeld() As Byte, Standpunkt As Long, Dateistring As String, Textvorhanden As Boolean) As Long
Dim Koltype As COLUMNSTRUCT
Dim TestString As String
Dim transstring As String
Dim ErsteZahl As Long
Dim ZweiteZahl As Long
Dim N1 As String
Dim N2 As String
Dim Optstr As String
Dim AlignString As String
Dim Teststr As String
Dim Namensstring As String
Dim NameWindow As String
Dim NameFile As String
Dim Hilfsbyte As Byte
Dim Hilfsint1 As Integer
Dim Bytefeld() As Byte
Dim Hilfslong As Long
Dim gr As Integer
Dim NumberofHS As Long
Dim Formatstring As String
Dim Typebyte As Byte
Dim Größe As Long
Dim Länge As Integer
Dim Hilfsint As Integer
Dim testoffset As Long
Dim i As Long
Dim Testbyte As Byte
Dim Makrostring As String
Dim Attributebyte As Byte
Dim Bytearray() As Byte
Dim Hilfsstring As String

Select Case Topictyp
Case 2
Case Else
Hilfsbyte = Datenfeld(Standpunkt)
Select Case Hilfsbyte
Case &H20
Standpunkt = Standpunkt + 5
Case &H21
Standpunkt = Standpunkt + 3
Case &H80 'Font
If Underline Then
Formatstring = Formatstring & "\ul "
End If
If DoubleUnderline Then
Formatstring = Formatstring & "\uldb "
End If
If ScrollorNoscroll Then
Formatstring = Formatstring & "\keepn"
End If
If Bold Then
Formatstring = Formatstring & "\b"
End If
If Italic Then
Formatstring = Formatstring & "\i"
End If
If StrikeOut Then
Formatstring = Formatstring & "\strike"
End If
If SmallCaps Then
Formatstring = Formatstring & "\scaps"
End If

CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt + 1), 2
If FontInUse <> font_descriptor(Hilfsint).ColorArraynumber Then 'Fontfarbe
Select Case font_descriptor(Hilfsint).ColorArraynumber
Case 0
Formatstring = Formatstring & "\cf" & font_descriptor(Hilfsint).ColorArraynumber
FontInUse = 0
Case Else
Formatstring = Formatstring & "\cf" & font_descriptor(Hilfsint).ColorArraynumber
FontInUse = font_descriptor(Hilfsint).ColorArraynumber
End Select
End If
Attributebyte = font_descriptor(Hilfsint).Attributes
If font_descriptor(Hilfsint).Fontsize <> 0 Then
Formatstring = Formatstring & "\f" & font_descriptor(Hilfsint).FontName & "\fs" & font_descriptor(Hilfsint).Fontsize
Else
Formatstring = Formatstring & "\f" & font_descriptor(Hilfsint).FontName
End If
'evtl bei Attributebyte = 0 "\plain"
Testbyte = Attributebyte And FONT_BOLD
If Testbyte = FONT_BOLD Then
If Bold = False Then
Formatstring = Formatstring & "\b"
Bold = True
End If
Else
If Bold = True Then
Formatstring = Formatstring & "\b0"
Bold = False
End If
End If
Testbyte = Attributebyte And FONT_ITAL
If Testbyte = FONT_ITAL Then
If Italic = False Then
Formatstring = Formatstring & "\i"
Italic = True
End If
Else
If Italic = True Then
Formatstring = Formatstring & "\i0"
Italic = False
End If
End If
Testbyte = Attributebyte And FONT_STRK
If Testbyte = FONT_STRK Then
If StrikeOut = False Then
Formatstring = Formatstring & "\strike"
StrikeOut = True
End If
Else
If StrikeOut = True Then
Formatstring = Formatstring & "\strike0"
StrikeOut = False
End If
End If
Testbyte = Attributebyte And FONT_DBUN
If Testbyte = FONT_DBUN Then 'doppelt unterstrichen ??
If DoubleUnderline = False Then
Formatstring = Formatstring & "\uldb"
DoubleUnderline = True
End If
Else
If DoubleUnderline = True And Jumpeben = "" Then
Formatstring = Formatstring & "\uldb0"
DoubleUnderline = False
End If
End If
Testbyte = Attributebyte And FONT_UNDR
If Testbyte = FONT_UNDR Then
    If Underline = False Then
        If Jumpeben = "" Then
        Formatstring = Formatstring & "\ul"
        Else
        If Left(Jumpeben, 12) = "\uldb0 {\v %" Then
        Jumpeben = "\uldb0 {\v *" & Mid(Jumpeben, 13)
        End If
        End If
        Underline = True
        Else
        Teststr = Left(Jumpeben, 10)
            If Teststr = "\ul0 {\v %" Then
            Jumpeben = "\ul0 {\v *" & Mid(Jumpeben, 11)
            ElseIf Teststr = "\ul0 {\v *" Then
            Jumpeben = "\ul0 {\v %" & Mid(Jumpeben, 11)
            End If
        End If
Else
            Teststr = Left(Jumpeben, 10)
            If Teststr = "\ul0 {\v %" Then
            Jumpeben = "\ul0 {\v *" & Mid(Jumpeben, 11)
            ElseIf Teststr = "\ul0 {\v *" Then
            Jumpeben = "\ul0 {\v %" & Mid(Jumpeben, 11)
            End If
    If Underline = True Then
            If Jumpeben = "" Then
        Formatstring = Formatstring & "\ul0"
        End If
        Underline = False
    End If
End If

Testbyte = Attributebyte And FONT_SMCP
If Testbyte = FONT_SMCP Then
If SmallCaps = False Then
Formatstring = Formatstring & "\scaps"
SmallCaps = True
End If
Else
If SmallCaps = True Then
Formatstring = Formatstring & "\scaps0"
SmallCaps = False
End If
End If
Dateistring = Dateistring & Formatstring & " "
Standpunkt = Standpunkt + 3
Case &H81
IDString = ""
Dateistring = Dateistring & vbCrLf & "\line "
Standpunkt = Standpunkt + 1
Case &H82 'end of Paragraph

Select Case Topictyp
Case &H20
Dateistring = Dateistring & vbCrLf & "\par "
ZuletztAbsatz = True
Case &H23
CopyMemory ByVal VarPtr(Hilfsint1), Datenfeld(Standpunkt + 2), 2
If Datenfeld(Standpunkt + 1) <> &HFF Then
Dateistring = Dateistring & vbCrLf & "\par\intbl "
ZuletztAbsatz = True
ElseIf Hilfsint1 = -1 Then
IDString = ""
Dateistring = Dateistring & "\cell\intbl\row\pard"
ElseIf Hilfsint1 = lastcol Then
Dateistring = Dateistring & vbCrLf & "\par\pard "
Else
Dateistring = Dateistring & "\cell\pard "
End If
End Select
AktCharacter = AktCharacter + 1
Standpunkt = Standpunkt + 1
Case &H83 'TAB
Standpunkt = Standpunkt + 1
Dateistring = Dateistring & "\tab "
Case &H86, &H87, &H88  'ewl or bml or ...
Select Case Hilfsbyte
Case &H86
AlignString = "c"
Case &H87
AlignString = "l"
Case &H88
AlignString = "r"
End Select
Standpunkt = Standpunkt + 1
Typebyte = Datenfeld(Standpunkt)
Standpunkt = Standpunkt + 1
ReDim Bytearray(3)
CopyMemory Bytearray(0), Datenfeld(Standpunkt), 4
Größe = scanlong(Bytearray, Länge)
Standpunkt = Standpunkt + Länge
If Typebyte = &H22 Then
ReDim Bytefeld(1)
CopyMemory Bytefeld(0), Datenfeld(Standpunkt), 2
NumberofHS = scanword(Bytefeld, gr)
NOH = NOH + NumberofHS
Standpunkt = Standpunkt + gr
AktCharacter = AktCharacter + NumberofHS
End If
If Typebyte = &H22 Or Typebyte = &H3 Then
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt), 2
Standpunkt = Standpunkt + 2
Select Case Hilfsint
Case 0
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt), 2
ZuletztAbsatz = False
transstring = ""
For i = 0 To anzbitmaps - 1
If Bitmaps(i).Name = "|bm" & CStr(Hilfsint) Then
If Bitmaps(i).Transparent <> 0 Then
transstring = "t"
End If
Select Case Bitmaps(i).Type
Case 0 'bmp
Dateistring = Dateistring & "\{bm" & AlignString & transstring & " bm" & CStr(Hilfsint) & ".bmp\}"
Case 1 'wmf
Dateistring = Dateistring & "\{bm" & AlignString & transstring & " bm" & CStr(Hilfsint) & ".wmf\}"
Case 2 'shg
Dateistring = Dateistring & "\{bm" & AlignString & transstring & " bm" & CStr(Hilfsint) & ".shg\}"
Case 3
Dateistring = Dateistring & "\{bm" & AlignString & transstring & " bm" & CStr(Hilfsint) & ".mrb\}"
End Select
End If
Next i
Standpunkt = Standpunkt + 2
Case 1
Standpunkt = Standpunkt + Größe - 2
End Select
End If
If Typebyte = 5 Then
ReDim Bytearray(Größe - 6 - 1)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 6), Größe - 6
Hilfsstring = Bytearray
Standpunkt = Standpunkt + Größe
Hilfsstring = StrConv(Hilfsstring, vbUnicode)
If InStr(Hilfsstring, Chr(0)) Then
Hilfsstring = Left(Hilfsstring, InStr(Hilfsstring, Chr(0)) - 1)
End If
If Left(Hilfsstring, 1) = "!" Then 'Button
Dateistring = Dateistring & "\{button " & Mid(Hilfsstring, 2) & "\}"
ElseIf Left(Hilfsstring, 1) = "*" Then
Teststr = Mid(Hilfsstring, 2, 5)
TestString = Mid(Hilfsstring, 8, 1)
If IsNumeric(Teststr) And IsNumeric(TestString) Then
ErsteZahl = CLng(Teststr)
ZweiteZahl = CLng(TestString)
If ZweiteZahl = 3 Then
Optstr = "REPEAT"
Else
Optstr = "PLAY"
End If
Hilfslong = ErsteZahl And 2
If Hilfslong = 2 Then
Optstr = Optstr & " NOPLAYBAR"
End If
Hilfslong = ErsteZahl And 8
If Hilfslong = 8 Then
Optstr = Optstr & " NOMENU"
End If
TestString = Mid(Hilfsstring, 10)
Teststr = LCase(TestString)
If TestExternal(TestString, N1, N2) = "ext" Then
Optstr = Optstr & " EXTERNAL"
Hilfsstring = TestString
End If

Select Case AlignString
Case "c"
Dateistring = Dateistring & "\{mci" & " " & Optstr & ", " & N2 & "\}"
Case "r"
Dateistring = Dateistring & "\{mci_right" & " " & Optstr & ", " & N2 & "\}"
Case "l"
Dateistring = Dateistring & "\{mci_left" & " " & Optstr & ", " & N2 & "\}"
End Select
End If
Else
Dateistring = Dateistring & "\{ew" & AlignString & " " & Hilfsstring & "\}"

End If
End If
Case &H89 'end of Hotspot
If Jumpeben <> "" Then
Dateistring = Dateistring & Jumpeben
Else 'Sprünge noch nicht unterstüzt
If DoubleUnderline = True Then
Dateistring = Dateistring & "\uldb0"
End If
If Underline = True Then
Dateistring = Dateistring & "\ul0"
End If
End If
DoubleUnderline = False
Underline = False
Jumpeben = ""
Standpunkt = Standpunkt + 1
Case &H8B 'non-break-space
Dateistring = Dateistring & "\~"
Standpunkt = Standpunkt + 1
Case &H8C 'non-break-hypen
Dateistring = Dateistring & "\-"
Standpunkt = Standpunkt + 1
Case &HC8 'macro
If Underline = True Then 'nur doubleunderline
Dateistring = Dateistring & "\ul0"
Underline = False
End If
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt + 1), 2
Größe = Hilfsint
ReDim Bytearray(Größe - 1)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 3), Größe
Makrostring = Bytearray
Makrostring = StrConv(Makrostring, vbUnicode)
If InStr(Makrostring, Chr(0)) Then
Makrostring = Left(Makrostring, InStr(Makrostring, Chr(0)) - 1)
End If
RepairMacroString Makrostring 'einer
TestString = Makrostring
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
Jumpeben = "\uldb0 {\v !" & TestString & "}"
jumprec = "\uldb "

Standpunkt = Standpunkt + 3 + Größe
Case &HCC 'macro without font change
If Underline = True Then
Dateistring = Dateistring & "\ul0"
Underline = False
End If
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt + 1), 2
Größe = Hilfsint
ReDim Bytearray(Größe - 1)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 3), Größe
Makrostring = Bytearray
Makrostring = StrConv(Makrostring, vbUnicode)
If InStr(Makrostring, Chr(0)) Then
Makrostring = Left(Makrostring, InStr(Makrostring, Chr(0)) - 1)
End If
RepairMacroString Makrostring
TestString = Makrostring
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
Jumpeben = "\uldb0 {\v %!" & TestString & "}"
jumprec = "\uldb "
Standpunkt = Standpunkt + 3 + Größe

Case &HE0 'popupjump
If DoubleUnderline = True Then 'nur zur Sicherheit
Dateistring = Dateistring & "\uldb0"
DoubleUnderline = False
End If
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 1), 4
For i = 0 To UBound(Contexts)
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
If TestString <> "" Then
Dateistring = Dateistring & "\ul "
Underline = True
'eben
Jumpeben = "\ul0 {\v " & TestString & "}"
End If
jumprec = "\ul"
Standpunkt = Standpunkt + 5

Case &HE1 'topicjump
If Underline = True Then
Dateistring = Dateistring & "\ul0"
Underline = False
End If
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 1), 4
For i = 0 To UBound(Contexts)
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
If TestString <> "" Then
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
Jumpeben = "\uldb0 {\v " & TestString & "}"
End If
Standpunkt = Standpunkt + 5
jumprec = "\uldb "

Case &HE2 'normaler popupjump
If DoubleUnderline = True Then
Dateistring = Dateistring & "\uldb0"
DoubleUnderline = False
End If
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 1), 4
For i = 0 To UBound(Contexts)
TestString = ""
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
If TestString <> "" Then 'zur Sicherheit
Dateistring = Dateistring & "\ul "
Underline = True
Jumpeben = "\ul0 {\v " & TestString & "}"
End If
jumprec = "\ul"
Standpunkt = Standpunkt + 5

Case &HE3 'normaler Topic Jump
If Underline = True Then
Dateistring = Dateistring & "\ul0"
Underline = False
End If
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 1), 4
For i = 0 To UBound(Contexts)
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
If TestString <> "" Then
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
Jumpeben = "\uldb0 {\v " & TestString & "}"
End If
Standpunkt = Standpunkt + 5
jumprec = "\uldb "

Case &HE6 'popupjump ohne fontchange
If DoubleUnderline = True Then
Dateistring = Dateistring & "\uldb0"
DoubleUnderline = False
End If
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 1), 4
For i = 0 To UBound(Contexts)
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
If TestString <> "" Then
Dateistring = Dateistring & "\ul "
Underline = True
Jumpeben = "\ul0 {\v " & "%" & TestString & "}"
End If
Standpunkt = Standpunkt + 5
jumprec = "\ul "

Case &HE7 'topicjump ohne fontchange
If Underline = True Then
Dateistring = Dateistring & "\ul0"
Underline = False
End If
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 1), 4
For i = 0 To UBound(Contexts)
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
If TestString <> "" Then
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
Jumpeben = "\uldb0 {\v " & "%" & TestString & "}"
End If
Standpunkt = Standpunkt + 5
jumprec = "\uldb "

Case &HEA 'Popup jump into external file with fontchange
If DoubleUnderline = True Then
Dateistring = Dateistring & "\uldb0"
DoubleUnderline = False
End If
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt + 1), 2
Größe = Hilfsint
If Größe > 0 Then
CopyMemory ByVal VarPtr(Hilfsbyte), Datenfeld(Standpunkt + 3), 1
If UBound(Datenfeld) >= Standpunkt + 8 Then 'Nur zur Sicherheit
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 4), 4
End If
Select Case Hilfsbyte
Case 1 'Offset + Windownumber
'bei Popupjump kein window
Case 4, 0
Namensstring = unhash(Hilfslong) 'externer Name
ReDim Bytearray(Größe - 6)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 8), Größe - 6
Hilfsstring = Bytearray
Hilfsstring = StrConv(Hilfsstring, vbUnicode)
If InStr(Hilfsstring, Chr(0)) Then
Länge = InStr(Hilfsstring, Chr(0))
Hilfsstring = Left(Hilfsstring, Länge - 1)
End If
Dateistring = Dateistring & "\ul "
Underline = True
TestString = Namensstring & "@" & Hilfsstring
Jumpeben = "\ul0 {\v " & TestString & "}"
jumprec = "\ul"
Case 6 'externt with windowchange
'bei popupjump kein window
End Select
End If
Standpunkt = Standpunkt + 3 + Größe

Case &HEB 'Topic into external file /sec window
If Underline = True Then
Dateistring = Dateistring & "\ul0"
Underline = False
End If
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt + 1), 2
Größe = Hilfsint
CopyMemory ByVal VarPtr(Hilfsbyte), Datenfeld(Standpunkt + 3), 1
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 4), 4
Select Case Hilfsbyte
Case 1
CopyMemory ByVal VarPtr(Testbyte), Datenfeld(Standpunkt + 8), 1
TestString = ""
For i = 0 To UBound(Contexts)
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
If TestString = "" Then
For i = 0 To UBound(Contexts)
If Contexts(i).Topicoffset = testoffset Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
End If
If TestString <> "" Then
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
If Testbyte > UBound(windownames) Or Testbyte < 0 Then
Jumpeben = "\uldb0 {\v " & TestString & "}"
Else
Jumpeben = "\uldb0 {\v " & TestString & ">" & windownames(Testbyte) & "}"
End If
End If
Case 4, 0
If UBound(Datenfeld) >= Standpunkt + 8 Then
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 4), 4
Namensstring = unhash(Hilfslong)
ReDim Bytearray(Größe - 6)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 8), Größe - 6
Hilfsstring = Bytearray
Hilfsstring = StrConv(Hilfsstring, vbUnicode)
If InStr(Hilfsstring, Chr(0)) Then
Länge = InStr(Hilfsstring, Chr(0))
Hilfsstring = Left(Hilfsstring, Länge - 1)
End If
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
TestString = Namensstring & "@" & Hilfsstring
Jumpeben = "\uldb0 {\v " & TestString & "}"
jumprec = "\uldb"
End If
Case 6
Namensstring = unhash(Hilfslong)
ReDim Bytearray(Größe - 5)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 8), Größe - 5
Hilfsstring = Bytearray
Hilfsstring = StrConv(Hilfsstring, vbUnicode)
Länge = InStr(Hilfsstring, Chr(0))
NameFile = Left(Hilfsstring, Länge - 1)
NameWindow = Mid(Hilfsstring, Länge + 1, Len(Hilfsstring) - Länge - 2)
TestString = Namensstring & ">" & NameWindow & "@" & NameFile
DoubleUnderline = True
Jumpeben = "\uldb0 {\v " & TestString & "}"
jumprec = "\uldb "
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
End Select
Standpunkt = Standpunkt + 3 + Größe

Case &HEE 'popupjump in external file without fontchange
If DoubleUnderline = True Then
Dateistring = Dateistring & "\uldb0"
DoubleUnderline = False
End If
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt + 1), 2
Größe = Hilfsint
CopyMemory ByVal VarPtr(Hilfsbyte), Datenfeld(Standpunkt + 2), 1
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 3), 4
Select Case Hilfsbyte
Case 1
'bei popupjump kein window
Case 4, 0
Namensstring = unhash(Hilfslong)
ReDim Bytearray(Größe - 6)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 8), Größe - 6
Hilfsstring = Bytearray
Hilfsstring = StrConv(Hilfsstring, vbUnicode)
If InStr(Hilfsstring, Chr(0)) Then
Länge = InStr(Hilfsstring, Chr(0))
Hilfsstring = Left(Hilfsstring, Länge - 1)
End If
Dateistring = Dateistring & "\ul "
Underline = True
TestString = Namensstring & "@" & Hilfsstring
Jumpeben = "\ul0 {\v " & "%" & TestString & "}"
jumprec = "\ul"
Case 6
'bei popupjump kein window
End Select
jumprec = "\ul "
Standpunkt = Standpunkt + 3 + Größe

Case &HEF 'topicjump in external file/ sec Window
If Underline = True Then
Dateistring = Dateistring & "\ul0"
Underline = False
End If
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt + 1), 2
Größe = Hilfsint
CopyMemory ByVal VarPtr(Hilfsbyte), Datenfeld(Standpunkt + 3), 1
CopyMemory ByVal VarPtr(Hilfslong), Datenfeld(Standpunkt + 4), 4

Select Case Hilfsbyte
Case 1
CopyMemory ByVal VarPtr(Testbyte), Datenfeld(Standpunkt + 8), 1
TestString = ""
For i = 0 To UBound(Contexts)
If Contexts(i).Hashvalue = Hilfslong Then
TestString = IDNames(i)
Exit For
End If
Next i
If TestString = "" Then
For i = 0 To UBound(Contexts)
If Contexts(i).Topicoffset = testoffset Then
TestString = IDNames(i)
testoffset = Contexts(i).Topicoffset
Exit For
End If
Next i
End If
If TestString <> "" Then
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
If UBound(windownames) >= Testbyte Then
Jumpeben = "\uldb0 {\v " & "%" & TestString & ">" & windownames(Testbyte) & "}"
Else
Jumpeben = "\uldb0 {\v " & "%" & TestString & "}"
End If
End If
Case 4, 0
ReDim Bytearray(Größe - 5)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 8), Größe - 5
Hilfsstring = Bytearray
Hilfsstring = StrConv(Hilfsstring, vbUnicode)
Länge = InStr(Hilfsstring, Chr(0))
NameFile = Left(Hilfsstring, Länge - 1)
TestString = unhash(Hilfslong)
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
Jumpeben = "\uldb0 {\v " & "%" & TestString & "@" & NameFile & "}"
Case 6
ReDim Bytearray(Größe - 5)
CopyMemory Bytearray(0), Datenfeld(Standpunkt + 8), Größe - 5
Hilfsstring = Bytearray
Hilfsstring = StrConv(Hilfsstring, vbUnicode)
Länge = InStr(Hilfsstring, Chr(0))
NameWindow = Left(Hilfsstring, Länge - 1)
NameFile = Mid(Hilfsstring, Länge + 1)
NameFile = Left(NameFile, Len(NameFile) - 2)
TestString = unhash(Hilfslong)
Dateistring = Dateistring & "\uldb "
DoubleUnderline = True
Jumpeben = "\uldb0 {\v " & "%" & TestString & ">" & NameWindow & "@" & NameFile & "}"
End Select
Standpunkt = Standpunkt + 3 + Größe

Case &HFF 'Endbyte Vorsicht bei Typ &H23
If Topictyp = &H23 Then
If Datenfeld(Standpunkt + 1) <> &HFF Then 'Endbyte
'Dateistring = Dateistring & "\cell\pard \pard\intbl" & vbCrLf
Dateistring = Dateistring & "\pard\intbl" & vbCrLf
IDString = ""
Standpunkt = Standpunkt + 1
CopyMemory ByVal VarPtr(Koltype), Datenfeld(Standpunkt), 5
lastcol = Koltype.Columnn
Standpunkt = Standpunkt + 5
Standpunkt = Standpunkt + 4 ' 'Unknown + Biased char überspringen
CopyMemory ByVal VarPtr(Hilfsint), Datenfeld(Standpunkt), 2 'id
Standpunkt = Standpunkt + 2
If Hilfsint <> 0 Then
Standpunkt = Standpunkt + TestID(Hilfsint, Standpunkt, Datenfeld, Dateistring, Topictyp)
End If
Else
'Dateistring = Dateistring & "\cell\intbl\row" & vbCrLf & "\pard"
Standpunkt = Standpunkt + 1
End If
Else
Dateistring = Dateistring & "\pard"
IDString = "\pard"

Standpunkt = Standpunkt + 1
End If
Case &H58 'TAB
Standpunkt = Standpunkt + 1
Dateistring = Dateistring & "\tab "
Case Else
Debug.Print "Fehler" & Hex(Hilfsbyte)
End Select
End Select
End Function

Public Function FindKeywordsForTopic(Buchstabe As String, TopOffset As Long, LenTopTitle As Long, Optional HasPlain As Boolean = True) As String
Dim i As Long
Dim z As Long
Dim anzahl As Long
Dim gefunden As Boolean
Dim Zwischentest() As String
Dim Zwischenliste As String
gefunden = False
anzahl = 0
Select Case Buchstabe
Case "K"
For i = 0 To UBound(KWData)
If KWData(i) >= TopOffset And KWData(i) < TopOffset + LenTopTitle Then
If anzahl = 0 Then
Zwischenliste = Zwischenliste & KWBTreeString(WoistKKeyword(i)) & ";"
ReDim Zwischentest(anzahl)
Zwischentest(0) = KWBTreeString(WoistKKeyword(i))
anzahl = 1
Else
For z = 0 To anzahl - 1
If Zwischentest(z) = KWBTreeString(WoistKKeyword(i)) Then
gefunden = True
Exit For
End If
Next z
Select Case gefunden
Case False
Zwischenliste = Zwischenliste & KWBTreeString(WoistKKeyword(i)) & ";"
ReDim Preserve Zwischentest(anzahl)
Zwischentest(0) = KWBTreeString(WoistKKeyword(i))
anzahl = anzahl + 1
End Select
gefunden = False
End If
End If
Next i

If Zwischenliste = "" Then 'Falls Fehler in Offset?
If Reserveoffset > 0 Then
For i = 0 To UBound(KWData)
If KWData(i) = Reserveoffset Then
Zwischenliste = Zwischenliste & KWBTreeString(WoistKKeyword(i)) & ";"

End If
Next i
End If
End If

If Zwischenliste <> "" Then 'Nur falls Keywords vorhanden
If HasPlain = True Then
FindKeywordsForTopic = "\pard {\up K}{\footnote\pard\plain{\up K} " & Zwischenliste & "}"
Else
FindKeywordsForTopic = "{\up K}{\footnote\pard\plain{\up K} " & Zwischenliste & "}"
End If
End If

Case "A"
For i = 0 To UBound(AWData)
If AWData(i) >= TopOffset And AWData(i) < TopOffset + LenTopTitle Then
If anzahl = 0 Then
Zwischenliste = Zwischenliste & AWBTreeString(WoistAKeyword(i)) & ";"
ReDim Zwischentest(anzahl)
Zwischentest(0) = AWBTreeString(WoistAKeyword(i))
anzahl = 1
Else
For z = 0 To anzahl - 1
If Zwischentest(z) = AWBTreeString(WoistAKeyword(i)) Then
gefunden = True
Exit For
End If
Next z
Select Case gefunden
Case False
Zwischenliste = Zwischenliste & AWBTreeString(WoistAKeyword(i)) & ";"
ReDim Preserve Zwischentest(anzahl)
Zwischentest(0) = AWBTreeString(WoistAKeyword(i))
anzahl = anzahl + 1
End Select
gefunden = False
End If
End If
Next i

If Zwischenliste = "" Then 'Falls Fehler in Offset?
If Reserveoffset > 0 Then
For i = 0 To UBound(AWData)
If AWData(i) = Reserveoffset Then
Zwischenliste = Zwischenliste & AWBTreeString(WoistAKeyword(i)) & ";"

End If
Next i
End If
End If

If Zwischenliste <> "" Then 'Nur falls Keywords vorhanden
If HasPlain = True Then
FindKeywordsForTopic = "\pard {\up A}{\footnote\pard\plain{\up A} " & Zwischenliste & "}"
Else
FindKeywordsForTopic = "{\up A}{\footnote\pard\plain{\up A} " & Zwischenliste & "}"
End If
End If
End Select
End Function
Public Sub RepariereBrowseArray(Dateinummer As Long)
Dim i As Long
Dim z As Long
Dim Test As Boolean
Dim Browsestring As String
Dim NextOffset As Long
Dim gefunden As Boolean
On Error GoTo Fehler

BrowseBlocknumber = 0
Browsesequencenumber = 1
Test = False

For i = 0 To UBound(Browsen)
Test = False
If Browsen(i).BrowseBackOffset = -1 Then
BrowseBlocknumber = BrowseBlocknumber + 1
Browsestring = "B" & Format(BrowseBlocknumber, "0000") & ":" & Format(Browsesequencenumber, "0000")
Browsesequencenumber = 2
If AktDateiname = Browsen(i).Filename Then
Put Dateinummer, Browsen(i).FileStandpunkt, Browsestring
Else
Close Dateinummer
Open Pfad & "\" & Browsen(i).Filename For Binary As Dateinummer
AktDateiname = Browsen(i).Filename
Put Dateinummer, Browsen(i).FileStandpunkt, Browsestring
End If
NextOffset = Browsen(i).BrowseForOffset
Do While Test = False
For z = 0 To UBound(Browsen)
If Browsen(z).ThisOffset = NextOffset Or Browsen(z).Reserveoffset = NextOffset Then
gefunden = True
If AktDateiname = Browsen(z).Filename Then
'OK
Else
Close Dateinummer
Open Pfad & "\" & Browsen(z).Filename For Binary As Dateinummer
AktDateiname = Browsen(z).Filename
End If
Browsestring = "B" & Format(BrowseBlocknumber, "0000") & ":" & Format(Browsesequencenumber, "0000")
Browsesequencenumber = Browsesequencenumber + 1
Put Dateinummer, Browsen(z).FileStandpunkt, Browsestring
NextOffset = Browsen(z).BrowseForOffset
If NextOffset = -1 Then
Test = True
End If
Exit For
End If
Next z
If gefunden = False Then
MsgBox "nicht gefunden"
Test = True
Exit For
End If
Loop
End If
Next i
Exit Sub
Fehler:
Exit Sub
End Sub

Public Sub CreateBrowseSequence(AktTopOffset As Long, NextBrowseOffset As Long, BackBrowseOffset As Long, TopicNummer As Long, Dateistring As String, BisherigeDateigröße As Long)
Dim Browsestring As String
If NextBrowseOffset <> -1 Or BackBrowseOffset <> -1 Then
Select Case BackBrowseOffset
Case -1 'Beginn of Browse Sequence
BrowseBlocknumber = BrowseBlocknumber + 1
Browsestring = "B" & Format(BrowseBlocknumber, "0000") & ":0001"
Browsesequencenumber = 2
NextGoodBrowseOffset = NextBrowseOffset
Case Else
If AktTopOffset = NextGoodBrowseOffset Or Reserveoffset = NextGoodBrowseOffset Then  'IO
If NextGoodBrowseOffset <> 0 Then
AktTopOffset = NextGoodBrowseOffset 'Falls Reserveoffset richtig
End If
Browsestring = "B" & Format(BrowseBlocknumber, "0000") & ":" & Format(Browsesequencenumber, "0000")
Browsesequencenumber = Browsesequencenumber + 1
NextGoodBrowseOffset = NextBrowseOffset
Else 'verkehrte Reihenfolge
BrowseFalse = True
Browsestring = "B0000:0000"
End If
End Select
ReDim Preserve Browsen(Browsenanzahl)
Browsen(Browsenanzahl).BrowseBackOffset = BackBrowseOffset
Browsen(Browsenanzahl).BrowseForOffset = NextBrowseOffset
Browsen(Browsenanzahl).ThisOffset = AktTopOffset
Browsen(Browsenanzahl).Reserveoffset = Reserveoffset
Browsen(Browsenanzahl).Topicnum = TopicNummer
Browsen(Browsenanzahl).FileStandpunkt = BisherigeDateigröße + Len(Dateistring) + 36
Browsen(Browsenanzahl).Filename = AktDateiname
Browsenanzahl = Browsenanzahl + 1
Browsestring = "{\up +}{\footnote\pard\plain{\up +}" & Browsestring & "}" & vbCrLf
Dateistring = Dateistring & Browsestring
Makrostringfertig = "BrowseButtons()" & vbCrLf
End If
End Sub

Public Sub Createhpjfile()
Dim Nummer As Long
Dim i As Long
Dim Filetext As String

If Contentsnumber <> -1 Then
Optionsstring = Optionsstring & vbCrLf & "CONTENTS=" & IDNames(Contentsnumber)
End If
If RTFGröße < 1000000 Then 'ab ca 1 MB komprimieren (vorher evtl Fehler - Bug in Helpcompiler-Workshop)
Filetext = "[OPTIONS]" & vbCrLf
Else
Filetext = "[OPTIONS]" & vbCrLf & "COMPRESS=12 Hall Zeck" & vbCrLf
End If
If Optionsstring <> "" Then
Filetext = Filetext & Optionsstring & vbCrLf
End If
Filetext = Filetext & "REPORT = Yes" & vbCrLf
Filetext = Filetext & "HLP=" & ProjektName & ".hlp" & vbCrLf & vbCrLf
If PetraAnzahl > 0 And MoreRTF = True Then
Filetext = Filetext & "[FILES]" & vbCrLf
For i = 1 To PetraAnzahl
Filetext = Filetext & PetraFile(i).RTFName & vbCrLf
Next i
Filetext = Filetext & vbCrLf
Else
Filetext = Filetext & "[FILES]" & vbCrLf & PetraFile(0).RTFName & vbCrLf & vbCrLf
End If
If Windowsstring <> "" Then
Filetext = Filetext & Windowsstring
End If
If Len(Aliasstring) > 9 Then
Filetext = Filetext & Aliasstring & vbCrLf
End If
If Mapstring <> "" Then
Filetext = Filetext & Mapstring
End If
If Configstring <> "" Then
Filetext = Filetext & Configstring
End If
If Makrostringfertig <> "" Then 'zur Sicherheit (Browsebutton)
If InStr(MacrostringAll, Makrostringfertig) Then
Else
MacrostringAll = MacrostringAll & Makrostringfertig
End If
End If
Filetext = Filetext & "[CONFIG]" & vbCrLf & MacrostringAll
If AnzBaggage > 0 Then
Baggagestring = ""
For i = 0 To AnzBaggage - 1
If BaggageInFile(i) = 0 Then
Baggagestring = Baggagestring & Baggagefiles(i) & vbCrLf
End If
Next i
Filetext = Filetext & "[BAGGAGE]" & vbCrLf & Baggagestring
End If

Nummer = FreeFile
Open Pfad & "\" & ProjektName & ".hpj" For Binary As Nummer
Put Nummer, , Filetext
Close Nummer
End Sub


Public Sub SortContexts()
Dim erledigt() As Boolean
Dim Übergabe1() As Long
Dim Übergabe2() As Long
Dim Übergabe3() As String
Dim Testar() As Long
Dim i As Long
Dim z As Long
Dim Größe As Long
Größe = UBound(Contexts)
ReDim Übergabe1(Größe)
ReDim Übergabe2(Größe)
ReDim erledigt(Größe)
ReDim Übergabe3(Größe)
ReDim Testar(Größe)
For i = 0 To Größe
Testar(i) = Contexts(i).Topicoffset
Next i
For i = 0 To Größe

Next i
QuickSort Testar

For i = 0 To Größe

For z = 0 To Größe
If Testar(i) = Contexts(z).Topicoffset And erledigt(z) = False Then
Übergabe1(i) = Contexts(z).Topicoffset
Übergabe2(i) = Contexts(z).Hashvalue
Übergabe3(i) = IDNames(z)
erledigt(z) = True
Exit For
End If
Next z
Next i
For i = 0 To Größe
Contexts(i).Topicoffset = Übergabe1(i)
Contexts(i).Hashvalue = Übergabe2(i)
IDNames(i) = Übergabe3(i)

Next i
End Sub

Public Sub Findviolas(Aktuellestopicoffset As Long, Fertigstring As String)
Dim i As Long
Dim Testat As Boolean
Testat = False
For i = 0 To anzviolas - 1
If Violafiles(i).Offsets = Aktuellestopicoffset Then 'Or Violafiles(i).Offsets = Reserveoffset Then
If Violafiles(i).Numbers > -1 And Violafiles(i).Numbers < UBound(windownames) Then
'eben
Fertigstring = Fertigstring & "{\up >}{\footnote\pard\plain{\up >} " & windownames(Violafiles(i).Numbers) & "}" & vbCrLf
End If
Testat = True
Exit For
End If
Next i
If Testat = False Then 'falls fehler in logik
For i = 0 To anzviolas - 1
If Violafiles(i).Offsets = Reserveoffset And Reserveoffset <> 0 Then
'eben
Fertigstring = Fertigstring & "{\up >}{\footnote\pard\plain{\up >} " & windownames(Violafiles(i).Numbers) & "}" & vbCrLf
Testat = True
Exit For
End If
Next i
End If

End Sub

Public Sub BerechneTabellenbreite(BreitenArray() As Long, tabletype As Byte, trleft As Double)
Dim Gesamtbreite As Long
Dim i As Long
Dim AnzahlSpalten As Long
Dim Rest As Long
Dim AnzZuKlein As Long
Dim KleinsteBreite As Long
Dim Zugabe As Long
Dim Breitenzähler As Long
Dim Breitenstand As Long

Breitenzähler = 0
KleinsteBreite = 500
AnzZuKlein = 0
Gesamtbreite = 0
AnzahlSpalten = UBound(BreitenArray) + 1
Breitenstand = trleft
Select Case tabletype
Case 0, 2
If trleft = -1 Then
For i = 0 To AnzahlSpalten - 1
If Breitenstand > BreitenArray(i) Then
BreitenArray(i) = Breitenstand + 50
End If
If BreitenArray(i) - Breitenstand < KleinsteBreite Then
AnzZuKlein = AnzZuKlein + 1
End If
Breitenstand = BreitenArray(i)
Next i
Gesamtbreite = BreitenArray(AnzahlSpalten - 1)
If 9000 > Gesamtbreite Then
Rest = 9000 - Gesamtbreite
If AnzZuKlein > 0 Then
Zugabe = Rest \ AnzZuKlein
End If
Select Case TabellenAlign
Case 0, 1 'Left + Right
For i = 0 To AnzahlSpalten - 1
If BreitenArray(i) < KleinsteBreite Then
BreitenArray(i) = BreitenArray(i) + Zugabe + Breitenzähler
End If
Breitenzähler = Breitenzähler + BreitenArray(i)
Next i
Case 2 'Center
For i = 0 To AnzahlSpalten - 1
If BreitenArray(i) < KleinsteBreite Then
BreitenArray(i) = BreitenArray(i) + Zugabe + Breitenzähler
End If
Breitenzähler = Breitenzähler + BreitenArray(i)
Next i
End Select
End If
End If
End Select
End Sub

Public Sub TestStringHash(TestString As String)
Dim Hashvalue As Long
Dim gefunden As Boolean
Dim z As Long
Hashvalue = Hashing(TestString)
gefunden = False
For z = 0 To UBound(Contexts)
If Hashvalue = Contexts(z).Hashvalue Then
IDNames(z) = TestString
Namegefunden(z) = True
gefunden = True
Exit For
End If
Next z

End Sub

Private Function TestExternal(StringToTest As String, HelpFilename As String, Medianame As String) As String
'ext int
Dim Uberstring As String
Dim TestString As String
Dim THname As String
Dim TMname As String
Dim i As Long
Medianame = ""
HelpFilename = ""
If InStr(StringToTest, "+") = 0 Then 'kein +
Medianame = StringToTest
TestExternal = "ext"
Exit Function
End If

If AnzBaggage = 0 Then 'keine Datei
Medianame = StringToTest
TestExternal = "ext"
Exit Function
Else
For i = 0 To AnzBaggage - 1
Uberstring = Mid(StringToTest, Len(StringToTest) - Len(Baggagefiles(i)) + 1)
If LCase(Baggagefiles(i)) = LCase(Uberstring) Then
If Mid(StringToTest, Len(StringToTest) - Len(Uberstring), 1) = "+" Then
BaggageInFile(i) = 1
Medianame = Uberstring
HelpFilename = Left(StringToTest, Len(StringToTest) - Len(Medianame) - 1)
TestExternal = "int"
Exit Function
End If
End If
Next i
End If
TestString = LCase(StringToTest)
If InStr(TestString, ".hlp+") Then
THname = Mid(StringToTest, 1, InStr(TestString, ".hlp+") + 3)
TMname = Mid(StringToTest, InStr(TestString, ".hlp+") + 5)
For i = 0 To AnzBaggage - 1
If LCase(TMname) = LCase(Baggagefiles(i)) Then
BaggageInFile(i) = 1
HelpFilename = THname
Medianame = TMname
TestExternal = "int"
Exit Function
End If
Next i
Else
Medianame = StringToTest
TestExternal = "ext"
Exit Function
End If

Medianame = StringToTest
TestExternal = "ext"

End Function
