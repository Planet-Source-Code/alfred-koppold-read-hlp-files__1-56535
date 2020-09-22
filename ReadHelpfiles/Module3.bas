Attribute VB_Name = "Module3"
Option Explicit
Public Aliasstring As String
Private unhashar As Boolean
Public Testarray() As Variant
Private Const OFFSET_4 As Double = 4294967296#
Public MappingDa() As String
Private Pictureoffsets() As Long

Private Sub CreateHashArray(HashArray As Variant)
HashArray = Array(&H0, &HD1, &HD2, &HD3, &HD4, &HD5, &HD6, &HD7, &HD8, &HD9, &HDA, &HDB, &HDC, &HDD, &HDE, &HDF, &HE0, &HE1, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9, &HEA, &HEB, &HEC, &HED, &HEE, &HEF, -16, &HB, -14, &HF3, -12, &HF5, -10, 16, -8, -7, &HFA, &HFB, -4, -3, &HC, -1, &HA, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, &H20, &H21, &H22, &H23, &H24, &H25, &H26, &H27, &H28, &H29, &H2A, &HB, &HC, &HD, &HE, &HD, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, &H20, &H21, &H22, &H23, &H24, &H25, &H26, &H27, &H28, &H29, &H2A, &H2B, &H2C, &H2D, &H2E, &H2F, &H50, &H51, &H52, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H5B, &H5C, &H5D, &H5E, &H5F, &H60, &H61, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H6A, &H6B, &H6C, &H6D, &H6E, &H6F, _
    &H70, &H71, &H72, &H73, &H74, &H75, &H76, -137, &H78, &H79, &H7A, &H7B, &H7C, &H7D, &H7E, &H7F, &H80, &H81, &H82, &H83, &HB, &H85, &H86, &H87, &H88, &H89, &H8A, &H8B, &H8C, &H8D, &H8E, &H8F, &H90, &H91, &H92, &H93, &H94, &H95, &H96, &H97, &H98, &H99, &H9A, &H9B, &H9C, &H9D, &H9E, &H9F, &HA0, &HA1, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HAA, &HAB, &HAC, &HAD, &HAE, -81, &HB0, &HB1, &HB2, &HB3, -76, &HB5, &HB6, &HB7, &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF, &HC0, &HC1, &HC2, &HC3, &HC4, &HC5, -58, &HC7, &HC8, &HC9, &HCA, &HCB, -52, &HCD, &HCE, &HCF)
End Sub

Public Function Hashing(strName As String) As Long
Dim hArray As Variant
Dim hash As Double
Dim i As Long
Dim Buchstabe As String
Dim BuchstabeAscii As Long
Dim Minusoffset As Double
Dim Obergrenze As Double
Dim Untergrenze As Double

Obergrenze = 2147483647#
Untergrenze = -2147483648#
Minusoffset = -4294967296#
Hashing = 0
CreateHashArray hArray
If strName = "" Then
Hashing = 1
Else
For i = 1 To Len(strName)
Buchstabe = Mid(strName, i, 1)
BuchstabeAscii = Asc(Buchstabe)
hash = (hash * 43) + hArray(BuchstabeAscii)
        If hash > OFFSET_4 Then
        Do While hash > OFFSET_4
        hash = hash - OFFSET_4
        Loop

        End If
        If hash < Minusoffset Then
        Do While hash < Minusoffset
        hash = hash + OFFSET_4
        Loop
        End If

Next i
If hash > Obergrenze And hash < OFFSET_4 Then
Hashing = UnsignedToLong(hash)
ElseIf hash < Untergrenze And hash > Minusoffset Then
hash = Abs(hash)
hash = UnsignedToLong(hash)
Hashing = 0 - hash
Else
Hashing = hash
End If
End If
End Function
Private Sub CreateUnHashArray()
Testarray = Array("Ä", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "Ä", "Ä", "Ä", "Ä", "Ä", "Ä", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "Ä")
unhashar = True
End Sub

Public Function unhash(Number As Long, Optional Offset1 As Double = 0) As String
Dim Übergabe As Double
Dim Test As Double

If unhashar = False Then CreateUnHashArray
Übergabe = LongToUnsigned(Number)
Übergabe = Übergabe + Offset1
Test = DblMod(Übergabe, 43)
If Test = 0 Then
Test = 43
End If
unhash = Testarray(Test)
Do While Übergabe <> 0
Übergabe = Übergabe - Test
If Übergabe = 0 Then Exit Do
Do While Übergabe <> 0
Übergabe = Übergabe / 43
If Übergabe < 43 Then
unhash = Testarray(Übergabe) & unhash
Übergabe = 0
Exit Do
Else
Test = DblMod(Übergabe, 43)
If Test = 0 Then
Test = 43
End If
unhash = Testarray(Test) & unhash
If Übergabe = 0 Then Exit Do
Übergabe = Übergabe - Test
End If
Loop
Loop
If InStr(unhash, "Ä") Then unhash = unhash(Number, Offset1 + OFFSET_4)
End Function

Public Function DblMod(Wert As Double, Divisor As Long) As Double
Dim Test As Double
Dim Testdouble As Double
Dim TestString As String
If Wert <= Divisor Then
DblMod = Wert
Else
Test = Wert / Divisor
TestString = CStr(Test)
If InStr(TestString, ",") Then
TestString = Left(TestString, InStr(TestString, ",") - 1)
Test = CDbl(TestString)
Else
Test = Test 'OK
End If
Testdouble = Test * Divisor
DblMod = Wert - Testdouble
End If
End Function

Public Sub MakeAlias()
Dim Gutesoffset As Long
Dim Letzter As Long
Dim LetztesOffset As Long
Dim i As Long

Aliasstring = "[ALIAS]" & vbCrLf
Letzter = 1
For i = 0 To UBound(Contexts) - 1
If Contexts(i).Topicoffset = Contexts(i + Letzter).Topicoffset Then
LetztesOffset = Contexts(i + 1).Topicoffset
Gutesoffset = i + 1
    Do While Contexts(i).Topicoffset = LetztesOffset And Letzter + i < UBound(Contexts)
    Letzter = Letzter + 1
    LetztesOffset = Contexts(i + Letzter).Topicoffset
    If LetztesOffset = Contexts(i).Topicoffset Then
    Gutesoffset = i + Letzter
    End If
    Loop
    Aliasstring = Aliasstring & IDNames(i) & "=" & IDNames(Gutesoffset) & vbCrLf
End If
Letzter = 1
Next i
End Sub

Private Sub SetBitmapIDs(HotspotArray() As Byte, Picturename As String)
Dim Pruef As Long
Dim LenString As Long
Dim Übergabedata() As Byte
Dim Zähler As Long
Dim IDStringarray() As String
Dim HsBytearray() As Byte
Dim Makroarray() As Byte
Dim HSNameArray() As String
Dim NumberOfHotspots As Integer
Dim gefunden As Boolean
Dim Länge1 As Long
Dim Hashvalue As Long
Dim SizeOfMakrodata As Long
Dim ArrayLen As Long
Dim z As Long
Dim Test() As Byte
Dim TestString As String
Dim HotspotName As String
Dim ContextorMacroName As String
Dim Stand As Long
Dim Stringdata As String
Dim Hotspot() As HOTSPOT_TYPE
Dim Länge As Long
Dim INDatei As Boolean
Dim sprung As Long
Dim Hsbgröße As Long
Dim i As Long
Dim Agröße As Long
Zähler = 0

ArrayLen = UBound(HotspotArray) + 1
CopyMemory ByVal VarPtr(NumberOfHotspots), HotspotArray(1), 2
CopyMemory ByVal VarPtr(SizeOfMakrodata), HotspotArray(3), 4
ReDim Hotspot(NumberOfHotspots - 1)
ReDim IDStringarray(NumberOfHotspots - 1)
ReDim HSNameArray(NumberOfHotspots - 1)
sprung = 0
For i = 0 To NumberOfHotspots - 1
Hotspot(i).id0 = HotspotArray(7 + sprung)
sprung = sprung + 1
Hotspot(i).id1 = HotspotArray(7 + sprung)
sprung = sprung + 1
Hotspot(i).id2 = HotspotArray(7 + sprung)
sprung = sprung + 1
CopyMemory ByVal VarPtr(Hotspot(i).x), HotspotArray(7 + sprung), 2
sprung = sprung + 2
CopyMemory ByVal VarPtr(Hotspot(i).Y), HotspotArray(7 + sprung), 2
sprung = sprung + 2
CopyMemory ByVal VarPtr(Hotspot(i).w), HotspotArray(7 + sprung), 2
sprung = sprung + 2
CopyMemory ByVal VarPtr(Hotspot(i).h), HotspotArray(7 + sprung), 2
sprung = sprung + 2
CopyMemory ByVal VarPtr(Hotspot(i).hash), HotspotArray(7 + sprung), 4
sprung = sprung + 4
Next i
Stand = 7 + (NumberOfHotspots * 15) 'hsptype

If SizeOfMakrodata > 0 Then
ReDim Makroarray(SizeOfMakrodata - 1)
CopyMemory Makroarray(0), HotspotArray(Stand), SizeOfMakrodata
Stand = Stand + SizeOfMakrodata
End If

ReDim Test(ArrayLen - Stand)
CopyMemory Test(0), HotspotArray(Stand), ArrayLen - Stand + 1
TestString = Test
TestString = StrConv(TestString, vbUnicode)
Stand = 1
For i = 0 To NumberOfHotspots - 1
Länge = InStr(Stand, TestString, Chr(0)) - Stand
HotspotName = Mid(TestString, Stand, Länge)
HSNameArray(i) = HotspotName
Stand = Stand + Länge + 1
Länge = InStr(Stand, TestString, Chr(0)) - Stand
Pruef = InStr(Stand, TestString, Chr(0))
If Pruef <= 0 Then
Länge = Len(TestString) - Stand
End If
ContextorMacroName = Mid(TestString, Stand, Länge)
IDStringarray(i) = ContextorMacroName
Select Case Hotspot(i).id0
Case &HC8, &HCC 'Makro
INDatei = False
Case &HE2, &HE3, &HE6, &HE7 'gleich
INDatei = True
Case &HEA, &HEB, &HEE, &HEF 'extern
If InStr(ContextorMacroName, "@") Then 'extern
INDatei = False
Else 'intern
INDatei = True
If InStr(ContextorMacroName, ">") Then
Länge1 = InStr(ContextorMacroName, ">")
ContextorMacroName = Left(ContextorMacroName, Länge1 - 1)
End If
End If
Case Else
INDatei = False
End Select
Select Case INDatei
Case True
Hashvalue = Hashing(ContextorMacroName)
gefunden = False
For z = 0 To UBound(Contexts)
If Hashvalue = Contexts(z).Hashvalue Then
IDNames(z) = ContextorMacroName
Namegefunden(z) = True
ReDim Preserve MappingDa(UBound(MappingDa) + 1)
MappingDa(UBound(MappingDa)) = ContextorMacroName
gefunden = True
Exit For
End If
Next z
If gefunden = True Then
Zähler = Zähler + 1
Hsbgröße = Hsbgröße + 15
ReDim Preserve HsBytearray(Hsbgröße - 1)
HsBytearray(Hsbgröße - 15) = Hotspot(i).id0
HsBytearray(Hsbgröße - 14) = Hotspot(i).id1
HsBytearray(Hsbgröße - 13) = Hotspot(i).id2
CopyMemory HsBytearray(Hsbgröße - 12), ByVal VarPtr(Hotspot(i).x), 2
CopyMemory HsBytearray(Hsbgröße - 10), ByVal VarPtr(Hotspot(i).Y), 2
CopyMemory HsBytearray(Hsbgröße - 8), ByVal VarPtr(Hotspot(i).w), 2
CopyMemory HsBytearray(Hsbgröße - 6), ByVal VarPtr(Hotspot(i).h), 2
CopyMemory HsBytearray(Hsbgröße - 4), ByVal VarPtr(Hotspot(i).hash), 4
Stringdata = Stringdata & HSNameArray(i) & Chr(0) & IDStringarray(i) & Chr(0)
End If
Case False
Zähler = Zähler + 1
Hsbgröße = Hsbgröße + 15
ReDim Preserve HsBytearray(Hsbgröße - 1)
HsBytearray(Hsbgröße - 15) = Hotspot(i).id0
HsBytearray(Hsbgröße - 14) = Hotspot(i).id1
HsBytearray(Hsbgröße - 13) = Hotspot(i).id2
CopyMemory HsBytearray(Hsbgröße - 12), ByVal VarPtr(Hotspot(i).x), 2
CopyMemory HsBytearray(Hsbgröße - 10), ByVal VarPtr(Hotspot(i).Y), 2
CopyMemory HsBytearray(Hsbgröße - 8), ByVal VarPtr(Hotspot(i).w), 2
CopyMemory HsBytearray(Hsbgröße - 6), ByVal VarPtr(Hotspot(i).h), 2
CopyMemory HsBytearray(Hsbgröße - 4), ByVal VarPtr(Hotspot(i).hash), 4
Stringdata = Stringdata & HSNameArray(i) & Chr(0) & IDStringarray(i) & Chr(0)
End Select
Stand = Stand + Länge + 1
Next i
NumberOfHotspots = Zähler
LenString = Len(Stringdata)
Agröße = 7 + (NumberOfHotspots * 15) + SizeOfMakrodata + LenString
ReDim Übergabedata(Agröße - 1)
Übergabedata(0) = 1
CopyMemory Übergabedata(1), ByVal VarPtr(NumberOfHotspots), 2
CopyMemory Übergabedata(3), ByVal VarPtr(SizeOfMakrodata), 4
CopyMemory Übergabedata(7), HsBytearray(0), NumberOfHotspots * 15
Stand = 7 + (NumberOfHotspots * 15)
If SizeOfMakrodata > 0 Then
CopyMemory Übergabedata(Stand), Makroarray(0), SizeOfMakrodata
Stand = Stand + SizeOfMakrodata
End If
Stringdata = StrConv(Stringdata, vbFromUnicode)
CopyMemory Übergabedata(Stand), ByVal StrPtr(Stringdata), LenString
ReDim HotspotArray(UBound(Übergabedata))
CopyMemory HotspotArray(0), Übergabedata(0), UBound(Übergabedata) + 1
End Sub

Public Sub SHGMRB_BM(Bildtyp As Integer, Pictypestandpunkt As Long, Picturezahl As Integer, Bildarray() As Byte, Picturename As String, Compressiontyp As Byte, BMBeginn As Long, BeginnCompressedSize As Long, Hotspotsize As Long, Hotspotoffset As Long, CompressedOffset As Long, palettestand As Long, Optional Picturenumber As Long = 0)
Dim BMHeadArray() As Byte
Dim shgmrbhead As SHGMRB_HEAD
Dim Palettengröße As Long
Dim shgmrbhead2 As SHGMRB_HEAD2
Dim shgheadteil As SHG_HEAD
Dim LenHeadbisCompSize As Long
Dim StandAnfang As Long
Dim Hilfslong As Long
Dim HotspotData() As Byte
Dim PaletteData() As Byte
Dim Haspalette As Boolean
Dim Speichernummer As Long
Dim Dateiname As String
Dim shgNormalPicOffset As Long
Dim Rechcompoffset As Long
Dim IsMRB As Boolean

shgNormalPicOffset = 8
If Picturenumber = 0 Then '1. oder einziges Bild
Select Case Picturezahl
Case 1
Bitmaps(anzbitmaps).Type = 2 'shg
IsMRB = False
Case Else
Bitmaps(anzbitmaps).Type = 3 'mrb
IsMRB = True
End Select
Get Dateinummer, BMBeginn, shgmrbhead
shgmrbhead.NumberofPictures = 1
shgmrbhead.Magic = &H506C
ReDim Pictureoffsets(Picturezahl - 1)
Pictureoffsets(Picturenumber) = 4 + (Picturezahl * 4)  'bei 0 = 9
shgmrbhead2.PictureType = Bildtyp
shgmrbhead2.PackingMethod = Compressiontyp
LenHeadbisCompSize = BeginnCompressedSize - Pictypestandpunkt - 2
ReDim BMHeadArray(LenHeadbisCompSize - 1)
StandAnfang = Pictypestandpunkt + 2
Get Dateinummer, StandAnfang, BMHeadArray

If Hotspotsize > 0 Then 'mit Hotspot-Daten
ReDim HotspotData(Hotspotsize - 1)
Get Dateinummer, Pictypestandpunkt + Hotspotoffset, HotspotData
SetBitmapIDs HotspotData, Picturename
End If
shgheadteil.CompSize = MakeCompressedUnsignedLong(UBound(Bildarray) + 1)
If palettestand < (Pictypestandpunkt + CompressedOffset) And Bildtyp <> 8 Then 'mit Palette
Haspalette = True
Palettengröße = Pictypestandpunkt + CompressedOffset - palettestand
ReDim PaletteData(Palettengröße - 1)
Get Dateinummer, palettestand, PaletteData
Else
Haspalette = False
End If
Rechcompoffset = 2 + LenHeadbisCompSize + 16 + Palettengröße
shgheadteil.CompOffset = Rechcompoffset
If Hotspotsize > 0 Then
shgheadteil.Hotspotsize = MakeCompressedUnsignedLong(UBound(HotspotData) + 1)
shgheadteil.HSPOffset = Rechcompoffset + UBound(Bildarray) + 1
Else 'ohne Hotspot
shgheadteil.Hotspotsize = MakeCompressedUnsignedLong(0)
shgheadteil.HSPOffset = 0
End If
Speichernummer = FreeFile
If IsMRB = True Then
Debug.Print 1
End If
Dateiname = Pfad & "\" & Mid(Picturename, 2)
If IsMRB = False Then
'als shg speichern
If Picturezahl > 1 Then
Open Dateiname & "(" & Picturenumber & ")" & ".shg" For Binary As Speichernummer 'Test
Else
Open Dateiname & ".shg" For Binary As Speichernummer
End If
Put Speichernummer, , shgmrbhead
Put Speichernummer, , shgNormalPicOffset  'Pictureoffsets(0)
Put Speichernummer, , shgmrbhead2
Put Speichernummer, , BMHeadArray
Put Speichernummer, , shgheadteil
If Haspalette = True Then
Put Speichernummer, , PaletteData
End If
Put Speichernummer, , Bildarray
If Hotspotsize > 0 Then
Put Speichernummer, , HotspotData
End If
Close Speichernummer
End If
If Picturezahl > 1 Then
'1.Bild als mrb speichern
shgmrbhead.Magic = &H706C
Open Dateiname & ".mrb" For Binary As Speichernummer
Put Speichernummer, , shgmrbhead
Put Speichernummer, , Pictureoffsets
Put Speichernummer, , shgmrbhead2
Put Speichernummer, , BMHeadArray
Put Speichernummer, , shgheadteil
If Haspalette = True Then
Put Speichernummer, , PaletteData
End If
Put Speichernummer, , Bildarray
If Hotspotsize > 0 Then
Put Speichernummer, , HotspotData
End If
Close Speichernummer
End If
End If

If Picturezahl > 1 And Picturenumber > 0 Then 'mrb
shgmrbhead2.PictureType = Bildtyp
shgmrbhead2.PackingMethod = Compressiontyp
LenHeadbisCompSize = BeginnCompressedSize - Pictypestandpunkt - 2
ReDim BMHeadArray(LenHeadbisCompSize - 1)
StandAnfang = Pictypestandpunkt + 2
Get Dateinummer, StandAnfang, BMHeadArray
If Hotspotsize > 0 Then 'mit Hotspot-Daten
ReDim HotspotData(Hotspotsize - 1)
Get Dateinummer, Pictypestandpunkt + Hotspotoffset, HotspotData
SetBitmapIDs HotspotData, Picturename
End If
shgheadteil.CompSize = MakeCompressedUnsignedLong(UBound(Bildarray) + 1)
If palettestand < (Pictypestandpunkt + CompressedOffset) Then 'mit Palette
Haspalette = True
Palettengröße = Pictypestandpunkt + CompressedOffset - palettestand
ReDim PaletteData(Palettengröße - 1)
Get Dateinummer, palettestand, PaletteData
Else
Haspalette = False
End If
Rechcompoffset = 2 + LenHeadbisCompSize + 16 + Palettengröße
shgheadteil.CompOffset = Rechcompoffset
If Hotspotsize > 0 Then
shgheadteil.Hotspotsize = MakeCompressedUnsignedLong(UBound(HotspotData) + 1)
shgheadteil.HSPOffset = Rechcompoffset + UBound(Bildarray) + 1
Else 'ohne Hotspot
shgheadteil.Hotspotsize = MakeCompressedUnsignedLong(0)
shgheadteil.HSPOffset = 0
End If
Speichernummer = FreeFile
'MRB als shg speichern
'Dateiname = Pfad & "\" & Mid(Picturename, 2)
'Open Dateiname & "(" & Picturenumber & ")" & ".shg" For Binary As Speichernummer
'shgmrbhead.NumberofPictures = 1
'shgmrbhead.Magic = &H506C
'Put Speichernummer, , shgmrbhead
'Put Speichernummer, , shgNormalPicOffset  'Pictureoffsets(0)
'Put Speichernummer, , shgmrbhead2
'Put Speichernummer, , BMHeadArray
'Put Speichernummer, , shgheadteil
'If Haspalette = True Then
'Put Speichernummer, , PaletteData
'End If
'Put Speichernummer, , Bildarray
'If Hotspotsize > 0 Then
'Put Speichernummer, , HotspotData
'End If
'Close Speichernummer
'als mrb speichern
Open Dateiname & ".mrb" For Binary As Speichernummer
Hilfslong = LOF(Speichernummer)
Pictureoffsets(Picturenumber) = Hilfslong
Put Speichernummer, 5 + (Picturenumber * 4), Hilfslong
Put Speichernummer, LOF(Speichernummer) + 1, shgmrbhead2
Put Speichernummer, , BMHeadArray
Put Speichernummer, , shgheadteil
If Haspalette = True Then
Put Speichernummer, , PaletteData
End If
Put Speichernummer, , Bildarray
If Hotspotsize > 0 Then
Put Speichernummer, , HotspotData
End If
Close Speichernummer
End If

End Sub

Public Sub TestMakro(StringMakro As String)
If InStr(StringMakro, "'") Or InStr(StringMakro, "`") Then
StringMakro = Replace(StringMakro, "'", "")
StringMakro = Replace(StringMakro, "`", "")
End If
End Sub
