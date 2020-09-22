Attribute VB_Name = "mainmodule"
Option Explicit
Public SavePhrase As Boolean
Public MacrostringAll As String
Public Contentsnumber As Long
Public Namecnt() As String
Public NumCnt As Long
Public Namegefunden() As Boolean
Public AktDateiname As String
Public BaggageInFile() As Byte
Public AnzBaggage As Long
Public Baggagefiles() As String
Public Baggagestring As String
Public Configstring As String
Public RTFGröße As Long
Public Macrostring As String
Private LetzterString As String
Public Aktuellestopicoffset As Long
Public Blocknummer As Long
Private CharacterCount As Long
Public Fertigstring As String
Public Hasphrase As Boolean
Private FaceName() As String
Private Fontfamilys() As String
Public PetraAnzahl As Long
Public PetraFile() As PETRA_TYPE
Private Phrimage() As Byte
Public Header As HELPHEADER
Public Filehdr As FILEHEADER
Public BTreeHdr As BTREEHEADER
Public CurrNode As BTREENODEHEADER
Public DirLeafEntry() As DIRECTORYLEAFENTRY
Private sysheader As SYSTEMHEADER
Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Type PETRA_TYPE
RTFName As String
FirstNumber As Long
End Type
Public Type TTLBTREE_TYPE
Topicoffset As Long
Topictitle As String
End Type
Public TTLBNames() As TTLBTREE_TYPE
Public AnzTTLBTREE As Long


Public Function GetbmFile(FileOffset As Long, Picturename As String) As Long
Dim f As Long
Dim bmfilehead As BITMAPFILEHEADER
Dim palettestand As Long
Dim PictureType() As Integer
Dim PicturePacking() As Integer
Dim Palettearray() As PALETTEENTRY
Dim Pictypestandpunkt() As Long
Dim CompressedOffset As Long
Dim Hotspotoffset As Long
Dim bmihead As BITMAPINFOHEADER
Dim CompressedSize As Long
Dim Hotspotsize As Long
Dim picfhead As PICTUREFILEHEADER
Dim Stand As Long
Dim Xdpi As Long
Dim Ydpi As Long
Dim Weite As Integer
Dim Hoehe As Integer
Dim Hilfsint1 As Integer
Dim Hilfsint2 As Integer
Dim Hilfsbyte1 As Byte
Dim Hilfsbyte2 As Byte
Dim Bildarray() As Byte
Dim Colors As Long
Dim Intfeld(9) As Integer
Dim StandCompressedSize As Long
Dim Picoffsets() As Long
Dim Längen As Integer
Dim i As Long
Dim z As Long
Dim Metaplacehead As METAPLACEABLEHEADER
Dim Checksum As Integer
Dim Pictype As Byte
Dim PackingMethod As Byte
Dim BMBeginn As Long
Dim DecompSize As Long
Dim IsMRB_SHG As Boolean
Dim Ende As Long

ReDim Preserve Bitmaps(anzbitmaps)
Bitmaps(anzbitmaps).Name = Picturename
Get Dateinummer, FileOffset + 1, Filehdr
Stand = Seek(Dateinummer)
BMBeginn = Stand
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Get Dateinummer, , picfhead
If picfhead.NumberofPictures <= 0 Then
MsgBox "Error in Helpfile"
End
End If
ReDim Pictureoffsets(picfhead.NumberofPictures - 1)
ReDim Picoffsets(picfhead.NumberofPictures - 1)
For i = 0 To picfhead.NumberofPictures - 1
Get Dateinummer, , Picoffsets(i)
Next i
ReDim Pictypestandpunkt(picfhead.NumberofPictures - 1)
ReDim PictureType(picfhead.NumberofPictures - 1)
ReDim PicturePacking(picfhead.NumberofPictures - 1)
For i = 0 To picfhead.NumberofPictures - 1
Get Dateinummer, Stand + Picoffsets(i), Pictype
Pictypestandpunkt(i) = Stand + Picoffsets(i)
PictureType(i) = Pictype
Get Dateinummer, , PackingMethod
PicturePacking(i) = PackingMethod
Next i

For i = 0 To picfhead.NumberofPictures - 1
IsMRB_SHG = False
Select Case PictureType(i)
Case 6
Form1.Label1.Caption = "Loading " & Mid(Picturename, 2) & ".bmp"
Stand = Pictypestandpunkt(i) + 2
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
Xdpi = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
bmihead.biXPelsPerMeter = (Xdpi * 79 + 1) \ 2
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
Ydpi = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
bmihead.biYPelsPerMeter = (Ydpi * 79 + 1) \ 2
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsbyte1
Get Dateinummer, , Hilfsbyte2
bmihead.biPlanes = ReadCompUnSignShort(Hilfsbyte1, Hilfsbyte2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsbyte1
Get Dateinummer, , Hilfsbyte2
bmihead.biBitCount = ReadCompUnSignShort(Hilfsbyte1, Hilfsbyte2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biWidth = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biHeight = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biClrUsed = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Colors = bmihead.biClrUsed
If Colors = 0 Then Colors = ShiftLeft06(1, bmihead.biBitCount)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biClrImportant = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
StandCompressedSize = Stand
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
CompressedSize = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
bmihead.biCompression = 0 'wird entkomprimert
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
Hotspotsize = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, CompressedOffset
Get Dateinummer, , Hotspotoffset
Stand = Stand + 8
palettestand = Stand
If CompressedSize < 1 Then Exit Function 'Vorläufig (MRB-Files)
ReDim Bildarray(CompressedSize - 1)
Get Dateinummer, Pictypestandpunkt(i) + CompressedOffset, Bildarray
Stand = Seek(Dateinummer)
Select Case PackingMethod

Case 0 'unkomprimiert
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then  'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 0, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
Case 1 'RunLen
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb oder shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 1, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
DecRunLen Bildarray
Case 2 'LZ77
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 2, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
Decompress UBound(Bildarray) + 1, Bildarray
Case 3 'LZ77 + RunLen
Decompress UBound(Bildarray) + 1, Bildarray
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 1, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
DecRunLen Bildarray

End Select
If IsMRB_SHG = False Or EnPicture = True Then
bmfilehead.bfType = 19778
If Colors <= 256 Then  'mit Palette
bmfilehead.bfSize = Len(bmfilehead) + Len(bmihead) + (Colors * 4) + UBound(Bildarray) + 1
bmfilehead.bfOffBits = Len(bmfilehead) + Len(bmihead) + (Colors * 4)
Else 'ohne Palette
bmfilehead.bfSize = Len(bmfilehead) + Len(bmihead) + UBound(Bildarray) + 1
bmfilehead.bfOffBits = Len(bmfilehead) + Len(bmihead)
End If
bmihead.biSizeImage = (((bmihead.biWidth * bmihead.biBitCount + 31) \ 32) * 4) * bmihead.biHeight
bmihead.biSize = Len(bmihead)
f = FreeFile
'Bei BMP lassen: Bitmaps(x).Type = 0
If i = 0 Then
If IsMRB_SHG = False Then
Open Pfad & "\" & Mid(Picturename, 2) & ".bmp" For Binary As f
Else
Open Pfad & "\" & Mid(Picturename, 2) & "(0).bmp" For Binary As f
End If
Else
Open Pfad & "\" & Mid(Picturename, 2) & "(" & i & ")" & ".bmp" For Binary As f
End If

Put f, , bmfilehead
Put f, , bmihead
Bitmaps(anzbitmaps).Transparent = bmihead.biClrImportant
If Colors <= 256 Then 'Palette
ReDim Palettearray(Colors - 1)
Get Dateinummer, palettestand, Palettearray
Put f, , Palettearray
End If
Put f, , Bildarray
Close f
End If
Case 5 'DBB
Form1.Label1.Caption = "Loading " & Mid(Picturename, 2) & ".bmp"
Stand = Pictypestandpunkt(i) + 2
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
Xdpi = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
bmihead.biXPelsPerMeter = (Xdpi * 79 + 1) \ 2
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
Ydpi = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
bmihead.biYPelsPerMeter = (Ydpi * 79 + 1) \ 2
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsbyte1
Get Dateinummer, , Hilfsbyte2
bmihead.biPlanes = ReadCompUnSignShort(Hilfsbyte1, Hilfsbyte2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsbyte1
Get Dateinummer, , Hilfsbyte2
bmihead.biBitCount = ReadCompUnSignShort(Hilfsbyte1, Hilfsbyte2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biWidth = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biHeight = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biClrUsed = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Colors = bmihead.biClrUsed
If Colors = 0 Then Colors = ShiftLeft06(1, bmihead.biBitCount)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
bmihead.biClrImportant = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
StandCompressedSize = Stand
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
CompressedSize = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
bmihead.biCompression = 0 'wird entkomprimert
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
Hotspotsize = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, CompressedOffset
Get Dateinummer, , Hotspotoffset
Stand = Stand + 8
If CompressedSize < 1 Then Exit Function 'Vorläufig (MRB-Files)
ReDim Bildarray(CompressedSize - 1)
Get Dateinummer, Pictypestandpunkt(i) + CompressedOffset, Bildarray
Stand = Seek(Dateinummer)
Select Case PackingMethod

Case 0 'unkomprimiert
DBBtoDIB Bildarray, bmihead.biWidth, bmihead.biHeight
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then  'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 0, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
Case 1 'RunLen
DecRunLen Bildarray
DBBtoDIB Bildarray, bmihead.biWidth, bmihead.biHeight
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb oder shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 0, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
Case 2 'LZ77
Decompress UBound(Bildarray) + 1, Bildarray
DBBtoDIB Bildarray, bmihead.biWidth, bmihead.biHeight
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 0, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
Case 3 'LZ77 + RunLen
Decompress UBound(Bildarray) + 1, Bildarray
DecRunLen Bildarray
DBBtoDIB Bildarray, bmihead.biWidth, bmihead.biHeight
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 0, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, palettestand, i
IsMRB_SHG = True
End If
End Select
If IsMRB_SHG = False Or EnPicture = True Then
bmfilehead.bfType = 19778
If Colors <= 256 Then  'mit Palette
bmfilehead.bfSize = Len(bmfilehead) + Len(bmihead) + (Colors * 4) + UBound(Bildarray) + 1
bmfilehead.bfOffBits = Len(bmfilehead) + Len(bmihead) + (Colors * 4)
Else 'ohne Palette
bmfilehead.bfSize = Len(bmfilehead) + Len(bmihead) + UBound(Bildarray) + 1
bmfilehead.bfOffBits = Len(bmfilehead) + Len(bmihead)
End If
bmihead.biSizeImage = (((bmihead.biWidth * bmihead.biBitCount + 31) \ 32) * 4) * bmihead.biHeight
bmihead.biSize = Len(bmihead)
f = FreeFile
'Bei BMP lassen: Bitmaps(x).Type = 0
If i = 0 Then
If IsMRB_SHG = False Then
Open Pfad & "\" & Mid(Picturename, 2) & ".bmp" For Binary As f
Else
Open Pfad & "\" & Mid(Picturename, 2) & "(0).bmp" For Binary As f
End If
Else
Open Pfad & "\" & Mid(Picturename, 2) & "(" & i & ")" & ".bmp" For Binary As f
End If

Put f, , bmfilehead
Put f, , bmihead
Bitmaps(anzbitmaps).Transparent = bmihead.biClrImportant
'Palette machen
ReDim Palettearray(Colors - 1)
Palettearray(1).peBlue = 255
Palettearray(1).peGreen = 255
Palettearray(1).peRed = 255
Put f, , Palettearray

Put f, , Bildarray
Close f
End If

Case 8
Form1.Label1.Caption = "Loading " & Mid(Picturename, 2) & ".wmf"
Stand = Pictypestandpunkt(i) + 3
Get Dateinummer, Stand, Weite 'Right
Stand = Stand + 2
Get Dateinummer, Stand, Hoehe 'Bottom
Stand = Stand + 2
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
DecompSize = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
StandCompressedSize = Stand
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
CompressedSize = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, Hilfsint1
Get Dateinummer, , Hilfsint2
Hotspotsize = ReadCompUnSignLong(Hilfsint1, Hilfsint2, Längen)
Stand = Stand + Längen
Get Dateinummer, Stand, CompressedOffset
Get Dateinummer, , Hotspotoffset
Stand = Stand + 8
ReDim Bildarray(CompressedSize - 1)
Get Dateinummer, Pictypestandpunkt(i) + CompressedOffset, Bildarray
Stand = Seek(Dateinummer)

Select Case PackingMethod
Case 0
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 0, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, 0, i
IsMRB_SHG = True
End If
Case 1
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 1, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, 0, i
IsMRB_SHG = True
End If
DecRunLen Bildarray
Case 2
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 2, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, 0, i
IsMRB_SHG = True
End If
Decompress UBound(Bildarray) + 1, Bildarray
Case 3
Decompress UBound(Bildarray) + 1, Bildarray
If picfhead.NumberofPictures > 1 Or Hotspotoffset <> 0 Then 'mrb shg
SHGMRB_BM PictureType(i), Pictypestandpunkt(i), picfhead.NumberofPictures, Bildarray, Picturename, 1, BMBeginn, StandCompressedSize, Hotspotsize, Hotspotoffset, CompressedOffset, 0, i
IsMRB_SHG = True
End If
DecRunLen Bildarray
End Select
If IsMRB_SHG = False Or EnPicture = True Then
Metaplacehead.mtKey = &H9AC6CDD7
Metaplacehead.mtRight = Weite
Metaplacehead.mtBottom = Hoehe
Metaplacehead.mtInch = 2540
CopyMemory ByVal VarPtr(Intfeld(0)), ByVal VarPtr(Metaplacehead), 20
For z = 0 To 9
Checksum = Checksum Xor Intfeld(z)
Next z
Metaplacehead.mtCheckSum = Checksum
f = FreeFile
If IsMRB_SHG = False Then
Open Pfad & "\" & Mid(Picturename, 2) & ".wmf" For Binary As f
Else
Open Pfad & "\" & Mid(Picturename, 2) & "(0).wmf" For Binary As f
End If
If Bitmaps(anzbitmaps).Type = 0 Then
Bitmaps(anzbitmaps).Type = 1
End If
Put f, , Metaplacehead
Put f, , Bildarray
Close f
End If
End Select
DoEvents
Next i
anzbitmaps = anzbitmaps + 1
End Function
Public Function GetCFFile(FileOffset As Long, Configname As String) As Long
Dim Stand As Long
Dim Beginn As Long
Dim Ende As Long
Dim Länge As Long
Dim CNum As Long
Dim Makroword As String
CNum = Mid(Configname, 4)
Get Dateinummer, FileOffset + 1, Filehdr
Stand = Seek(Dateinummer)
Beginn = Stand
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Länge = Ende - Beginn
Makroword = Space(Länge)
Get Dateinummer, Beginn, Makroword
If InStr(Makroword, ":") Then
TeileMacro Makroword 'mehrere Macros duch : geteilt
End If
RepairMacroString Makroword
Configstring = Configstring & "[CONFIG:" & CNum & "]" & vbCrLf & Makroword & vbCrLf & vbCrLf
Beginn = Beginn + Länge
End Function

Public Function GetBaggageFile(FileOffset As Long, strFilename As String) As Long
Dim Stand As Long
Dim Ende As Long
Dim FileBeginn As Long
Dim Filearray() As Byte
Dim Filenumber As Long
Dim Dateiname As String

Get Dateinummer, FileOffset + 1, Filehdr
Stand = Seek(Dateinummer)
FileBeginn = Stand
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
ReDim Filearray(Ende - FileBeginn)
Get Dateinummer, FileBeginn, Filearray
Dateiname = Pfad & "\" & strFilename
Filenumber = FreeFile
Open Dateiname For Binary As Filenumber
Put Filenumber, , Filearray
Close Filenumber
ReDim Preserve Baggagefiles(AnzBaggage)
ReDim Preserve BaggageInFile(AnzBaggage)

Baggagefiles(AnzBaggage) = strFilename
AnzBaggage = AnzBaggage + 1
End Function
Public Function GetCtxoMapFile(FileOffset As Long) As Long
Dim NummerGefunden As Boolean
Dim z As Long
Dim testlong As Long
Dim Testname As String
Dim i As Long
Dim ctxEntry() As CTXOMAPENTRY
Dim Hilfsint As Integer
Dim gefunden As Boolean
Dim Gleich As Boolean
Dim gemapt() As Byte
Dim aufname As Long
Dim MapNumber As Double
Dim Numbers() As Double
Dim First As Boolean
Dim Numm As String
Dim TName As String
Dim ÜName As String
Dim Testhash As Long

MapNumber = 2345
ReDim gemapt(UBound(IDNames))
Get Dateinummer, FileOffset + 1, Filehdr
Get Dateinummer, , Hilfsint
If Hilfsint = 0 Then Exit Function '0 Mapped ID
ReDim ctxEntry(Hilfsint - 1)
Get Dateinummer, , ctxEntry
ReDim Numbers(UBound(ctxEntry))
For i = 0 To UBound(ctxEntry)
Numbers(i) = LongToUnsigned(ctxEntry(i).MapID)
Next i
For i = 0 To UBound(Numbers)
If i < UBound(Numbers) Then
For z = i + 1 To UBound(Numbers)
If Numbers(i) = Numbers(z) Then
TestMapNumber MapNumber, Numbers, NummerGefunden 'keine doppelten Nummern
NummerGefunden = False
Numbers(z) = MapNumber
End If
Next z
End If
Next i
Mapstring = "[MAP]" & vbCrLf
For i = 0 To UBound(ctxEntry)
Gleich = False
If i > 0 Then
If ctxEntry(i).MapID = ctxEntry(i - 1).MapID Then 'Fehlerkorrektur
Gleich = True
End If
End If
gefunden = False
For z = 0 To UBound(Contexts)
If ctxEntry(i).Topicoffset = Contexts(z).Topicoffset Then
If Gleich = False Then
If InStr(IDNames(z), " ") = False Then
If First = False Then
Numm = Numbers(i)
If Right(IDNames(z), Len(Numm)) = Numm Then
Mapstring = Mapstring & IDNames(z) & " " & Numbers(i) & vbCrLf
TName = Left(IDNames(z), Len(IDNames(z)) - Len(Numm))
First = True
End If
Else
If Namegefunden(z) = False Then
ÜName = TName & CStr(Numbers(i))
Testhash = Hashing(ÜName)
If Testhash = Contexts(z).Hashvalue Then
Mapstring = Mapstring & ÜName & " " & Numbers(i) & vbCrLf
Namegefunden(z) = True
IDNames(z) = ÜName
Else
Mapstring = Mapstring & IDNames(z) & " " & Numbers(i) & vbCrLf
End If
Else
Mapstring = Mapstring & IDNames(z) & " " & Numbers(i) & vbCrLf
End If
End If
Else 'Fehler bei Leertaste
testlong = Hashing(IDNames(z))
Testname = unhash(testlong)
Mapstring = Mapstring & Testname & " " & Numbers(i) & vbCrLf
End If
gemapt(z) = 1
gefunden = True
Exit For
Else 'Fehler - gleiche ID mehrmals
gefunden = True
Exit For
End If
ElseIf ctxEntry(i).Topicoffset = -1 Then 'Fehler - ID nicht in Helpfile vorhanden !
gefunden = True
Exit For
End If
Next z
Next i

For i = 0 To UBound(IDNames)
'If gemapt(i) = 0 Then
If LCase(Left(IDNames(i), 4)) = "idh_" Then
TestMapNumber MapNumber, Numbers, NummerGefunden
If NummerGefunden = True Then
ReDim Preserve Numbers(UBound(Numbers) + 1)
Numbers(UBound(Numbers)) = MapNumber
NummerGefunden = False
End If
If InStr(IDNames(i), " ") = False Then
Mapstring = Mapstring & IDNames(i) & " " & LongToUnsigned(CLng(MapNumber)) & vbCrLf
Else 'Fehler bei Leertaste
testlong = Hashing(IDNames(i))
Testname = unhash(testlong)
Mapstring = Mapstring & Testname & " " & LongToUnsigned(CLng(MapNumber))
End If
MapNumber = MapNumber + 1
End If
'End If
Next i
Mapstring = Mapstring & vbCrLf
End Function

Public Function GetFontFile(FileOffset As Long) As Long
Dim GefundeneNummer As Long
Dim gefunden As Boolean
Dim ofont As OLDFONT
Dim fonthdr As FONTHEADER
Dim i As Long
Dim gefund As Boolean
Dim z As Long
Dim x As Long
Dim Farbenanzahl As Long
Dim Leng As Long
Dim Hilfsstring As String

gefunden = False
Get Dateinummer, FileOffset + 1, Filehdr
Get Dateinummer, , fonthdr
Seek Dateinummer, fonthdr.DescriptorsOffset + 9 + FileOffset + 1
ReDim font_descriptor(fonthdr.NumDescriptors - 1)
For i = 0 To fonthdr.NumDescriptors - 1
Get Dateinummer, , ofont
font_descriptor(i).Fontsize = ofont.HalfPoints
font_descriptor(i).FontName = ofont.FontName
font_descriptor(i).FontFamily = ofont.FontFamily
font_descriptor(i).Attributes = ofont.Attributes
CopyMemory font_descriptor(i).Fontfarbe(0), ofont.FGRGB(0), 3
Next i
gefund = False
ReDim Colortables(0)
Colortables(0).rgbtBlue = 1
Colortables(0).rgbtGreen = 1
Colortables(0).rgbtRed = 0
For i = 0 To fonthdr.NumDescriptors - 1
gefund = False
GefundeneNummer = 0
If font_descriptor(i).Fontfarbe(0) = 1 And font_descriptor(i).Fontfarbe(1) = 1 And font_descriptor(i).Fontfarbe(2) = 0 Then  'Farbwechsel ?
'Normal 1 1 0
gefund = True
Else
For z = 0 To Farbenanzahl
If font_descriptor(i).Fontfarbe(0) = Colortables(z).rgbtRed And font_descriptor(i).Fontfarbe(1) = Colortables(z).rgbtGreen And font_descriptor(i).Fontfarbe(2) = Colortables(z).rgbtBlue Then  'Farbwechsel ?
GefundeneNummer = z
gefund = True
End If
Next z
If gefund = False Then
Farbenanzahl = Farbenanzahl + 1
font_descriptor(i).ColorArraynumber = Farbenanzahl
ReDim Preserve Colortables(Farbenanzahl)
Colortables(Farbenanzahl).rgbtRed = font_descriptor(i).Fontfarbe(0)
Colortables(Farbenanzahl).rgbtGreen = font_descriptor(i).Fontfarbe(1)
Colortables(Farbenanzahl).rgbtBlue = font_descriptor(i).Fontfarbe(2)
Else
font_descriptor(i).ColorArraynumber = GefundeneNummer
End If
End If
gefund = False
Next i
Leng = (fonthdr.DescriptorsOffset - fonthdr.FacenamesOffset) / fonthdr.NumFacenames
Seek Dateinummer, fonthdr.FacenamesOffset + 9 + FileOffset + 1
ReDim FaceName(fonthdr.NumFacenames - 1)
ReDim Fontfamilys(fonthdr.NumFacenames - 1)
For i = 0 To fonthdr.NumFacenames - 1
Hilfsstring = Space(Leng)
Get Dateinummer, , Hilfsstring
Hilfsstring = Left(Hilfsstring, InStr(Hilfsstring, Chr(0)) - 1)
For x = 1 To 31
If InStr(Hilfsstring, Chr(x)) Then
Hilfsstring = Replace(Hilfsstring, Chr(x), "") 'manchmal Fehler in Helpfile
End If
Next x
FaceName(i) = Hilfsstring
Next i
'Übernehmen
For i = 0 To fonthdr.NumDescriptors - 1
Select Case font_descriptor(i).FontFamily
Case &H1
font_descriptor(i).FontFamily = "modern"
Case &H2
font_descriptor(i).FontFamily = "roman"
Case &H3
font_descriptor(i).FontFamily = "swiss"
Case &H4
font_descriptor(i).FontFamily = "script"
Case &H5
font_descriptor(i).FontFamily = "decor"
End Select
Fontfamilys(font_descriptor(i).FontName) = font_descriptor(i).FontFamily
font_descriptor(i).Fontnamefertig = FaceName(font_descriptor(i).FontName)

Next i
'aufmachen + 1
Fonttable = "{\fonttbl"
Aufmachen = Aufmachen + 1
For i = 0 To fonthdr.NumFacenames - 1 'Fonttable machen

If font_descriptor(0).Fontnamefertig = FaceName(i) And gefunden = False Then
Deffont = i
gefunden = True
End If
If Fontfamilys(i) = "" Then Fontfamilys(i) = "nil"
For z = 0 To UBound(font_descriptor)
If font_descriptor(z).Fontnamefertig = FaceName(i) Then
Fonttable = Fonttable & "{\f" & i & "\f" & Fontfamilys(i) & " " & FaceName(i) & ";}"
Exit For
End If
Next z
Next i

'For i = 0 To fonthdr.NumDescriptors - 1 'Fonttable machen
'Fonttable = Fonttable & "{\f" & i & "\f" & font_descriptor(i).FontFamily & " " & font_descriptor(i).Fontnamefertig & ";}"
'Next i

'zumachen + 1
Fonttable = Fonttable & "}"
zumachen = zumachen + 1
'aufmachen + 1
colortablestring = "{\colortbl;"
Aufmachen = Aufmachen + 1
For i = 1 To UBound(Colortables)
colortablestring = colortablestring & "\red" & Colortables(i).rgbtRed & "\green" & Colortables(i).rgbtGreen & "\blue" & Colortables(i).rgbtBlue & ";"
Next i
'zumachen + 1
colortablestring = colortablestring & "}"
zumachen = zumachen + 1
Select Case fonthdr.FacenamesOffset
Case Is >= 16, Is >= 12
scaling = 1
rounderr = 0
Case Else
scaling = 10
rounderr = 5
End Select
End Function

Public Function GetTopicID(FileOffset As Long) As Long
Dim btreehead As BTREEHEADER
Dim Keyword As String
Dim btreeNodhead As BTREENODEHEADER
Dim Ende As Long
Dim Testnummer As Long
Dim i As Long
Dim z As Long
Dim Topicoffset As Long
Dim Länge As Long
Dim Dateistand As Long
Dim headerstand As Long

Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Get Dateinummer, , btreehead

Select Case btreehead.NLevels
Case 1 'Leafpage
For i = 1 To btreehead.NLevels 'bei Leafentry 1
Get Dateinummer, , btreeNodhead
For z = 1 To btreeNodhead.NEntries
Get Dateinummer, , Topicoffset
Dateistand = Seek(Dateinummer)
Länge = FindStringEnd(Dateistand)
Keyword = Space(Länge)
Get Dateinummer, Dateistand, Keyword
If InStr(Keyword, Chr(0)) Then
Keyword = Left(Keyword, InStr(Keyword, Chr(0)) - 1)
End If
TestStringHash Keyword
Next z
Next i
Case Else 'Indexpage
headerstand = Seek(Dateinummer)
For i = 1 To btreehead.TotalPages
If i - 1 = btreehead.RootPage Then
Else
Get Dateinummer, headerstand, btreeNodhead
For z = 0 To btreeNodhead.NEntries 'Entries einlesen
Get Dateinummer, , Topicoffset
ReDim Preserve TTLBTreeOffsets(Testnummer)
TTLBTreeOffsets(Testnummer) = Topicoffset
Testnummer = Testnummer + 1
Dateistand = Seek(Dateinummer)
Länge = FindStringEnd(Dateistand)
Keyword = Space(Länge)
Get Dateinummer, Dateistand, Keyword
If InStr(Keyword, Chr(0)) Then
Keyword = Left(Keyword, InStr(Keyword, Chr(0)) - 1)
End If
TestStringHash Keyword
Next z
End If
headerstand = headerstand + btreehead.PageSize
Next i
End Select
End Function

Public Function GetViolaFile(FileOffset As Long) As Long
Dim Hilfslong As Long
Dim btreehead As BTREEHEADER
Dim btreeNodhead As BTREENODEHEADER
Dim Ende As Long
Dim i As Long
Dim z As Long
Dim Topicoffset As Long
Dim headerstand As Long
Dim Dateistand As Long

Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
hasviolas = True
Get Dateinummer, , btreehead
Select Case btreehead.NLevels
Case 1 'Leafpage
For i = 1 To btreehead.NLevels 'bei Leafentry 1
Get Dateinummer, , btreeNodhead
For z = 1 To btreeNodhead.NEntries
Get Dateinummer, , Topicoffset
Dateistand = Seek(Dateinummer)
Get Dateinummer, Dateistand, Hilfslong
ReDim Preserve Violafiles(anzviolas)
Violafiles(anzviolas).Offsets = Topicoffset
Violafiles(anzviolas).Numbers = Hilfslong
anzviolas = anzviolas + 1
Next z
Next i
Case Else 'Indexpage
headerstand = Seek(Dateinummer)
For i = 1 To btreehead.TotalPages
If i - 1 = btreehead.RootPage Then
Else
Get Dateinummer, headerstand, btreeNodhead
For z = 1 To btreeNodhead.NEntries - 1
Get Dateinummer, , Topicoffset
Dateistand = Seek(Dateinummer)
Get Dateinummer, Dateistand, Hilfslong
ReDim Preserve Violafiles(anzviolas)
Violafiles(anzviolas).Offsets = Topicoffset
Violafiles(anzviolas).Numbers = Hilfslong
anzviolas = anzviolas + 1
Next z
End If
headerstand = headerstand + btreehead.PageSize
Next i
End Select
End Function

Public Function GetPhrase(FileOffset As Long) As Long
Dim Nummer As Long
Dim Name As String
Dim Ende As Long
Dim Beginn As Long
Dim PhraseOffset() As Integer
Dim Phrhdr As PHRASEHEADER
Dim Stand As Long
Dim Länge As Long
Dim i As Long
Dim PhraseLongOffset() As Long
Hasphrase = True
Get Dateinummer, FileOffset + 1, Filehdr
Stand = Seek(Dateinummer)
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
ReDim Phrimage(Filehdr.UsedSpace - 1)
Get Dateinummer, Stand, Phrhdr
ReDim PhraseOffset(Phrhdr.NumPhrases)
ReDim PhraseLongOffset(Phrhdr.NumPhrases)
Get Dateinummer, , PhraseOffset
If Phrhdr.NumPhrases Then 'Falls Phrasen vorhanden
For i = 0 To Phrhdr.NumPhrases
PhraseLongOffset(i) = IntegerToUnsigned(PhraseOffset(i))
Next i
Stand = Seek(Dateinummer)
Länge = Ende - Stand
Dim a() As Byte
ReDim a(Länge)
Dim b As String
Get Dateinummer, , a
Decompress UBound(a) + 1, a
b = a
b = StrConv(b, vbUnicode)
ReDim PhraseArray(Phrhdr.NumPhrases - 1)
Beginn = (Phrhdr.NumPhrases + 1) * 2
For i = 0 To Phrhdr.NumPhrases - 1
PhraseArray(i) = Mid(b, PhraseLongOffset(i) - Beginn + 1, PhraseLongOffset(i + 1) - PhraseLongOffset(i))
Name = Name & PhraseArray(i) & vbCrLf
Next i
If SavePhrase Then
Nummer = FreeFile
Open Pfad & "\" & ProjektName & ".ph" For Binary As Nummer
Put Nummer, , Name
Close Nummer
End If
End If
End Function

Public Function GetPhrImage(FileOffset As Long) As Long
Dim Stand As Long
Dim Bytearray() As Byte

Get Dateinummer, FileOffset + 1, Filehdr
Stand = Seek(Dateinummer)
ReDim Phrimage(Filehdr.UsedSpace - 1)
Get Dateinummer, Stand, Phrimage
End Function

Public Function GetPhrIndex(FileOffset As Long) As Long
Dim Nummer As Long
Dim Name As String
Dim Anfangst As Long
Dim Phrimagestring As String
Dim Endst As Long
Dim anzahl As Long
Dim i As Long
Dim Ende As Long
Dim BCount As Long
Dim Value() As Byte
Dim Arrayfertig() As Long
Dim Phrinhdr As PHRINDEXHDR

Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
ReDim Value(Filehdr.UsedSpace)
Get Dateinummer, , Phrinhdr
If UBound(Phrimage) + 1 <> Phrinhdr.PhrImageSize Then
Decompress UBound(Phrimage) + 1, Phrimage
End If
Phrimagestring = Phrimage
Phrimagestring = StrConv(Phrimagestring, vbUnicode)
Get Dateinummer, , Value
BCount = Phrinhdr.BitCount And &HF
Hall Value, Phrinhdr.NEntries, BCount, Arrayfertig
anzahl = UBound(Arrayfertig) + 1
ReDim PhraseImageArray(anzahl - 1)
Anfangst = 0
If SavePhrase Then
Nummer = FreeFile
Open Pfad & "\" & ProjektName & ".ph" For Binary As Nummer
End If
For i = 0 To anzahl - 1
Endst = Arrayfertig(i)
PhraseImageArray(i) = Mid(Phrimagestring, Anfangst + 1, Endst - Anfangst)
Name = Name & PhraseImageArray(i) & vbCrLf
Anfangst = Endst
Next i
If SavePhrase Then
Put Nummer, , Name
Close Nummer
End If
End Function

Public Function GetxWMAPFile(FileOffset As Long, Buchstabe As String) As Long
Exit Function 'wird nicht benötigt
Dim Ende As Long
Dim i As Long
Dim KWMapEntry As KWMAPREC
Dim Hilfsint As Integer
Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Get Dateinummer, , Hilfsint
For i = 1 To Hilfsint
Get Dateinummer, , KWMapEntry
Next i
End Function

Public Function GetxWDataFile(FileOffset As Long, Buchstabe As String) As Long
Dim Ende As Long

Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Select Case Buchstabe
Case "K"
ReDim KWData((Filehdr.UsedSpace / 4) - 1)
Get Dateinummer, , KWData
HasKKeywords = True
Case "A"
ReDim AWData((Filehdr.UsedSpace / 4) - 1)
Get Dateinummer, , AWData
HasAKeywords = True
End Select
End Function

Public Function GetSystemFile(FileOffset As Long) As Long
Dim testlong As Long
Dim colorrgb As RGBColor
Dim Testdouble As Double
Dim CNTFilename As String
Dim RGBString As String
Dim Hilfsbytes() As Byte
Dim RGBSRString As String
Dim x As String
Dim Y As String
Dim Width As String
Dim Height As String
Dim systemrec As SYSTEMRECORD
Dim Ende As Long
Dim LCID(4) As Integer
Dim Bytearray() As Byte
Dim TestString As String
Dim a As SECWINDOW
Dim i As Long
Dim Test As Boolean
Dim NameString As String
Dim Hilfslong As Long
Dim Hilfsstring As String
Dim AdjustResolution As Boolean
Dim Farbe As Long
Dim b As String
Dim Stand As Long
Dim OnTop As Boolean
Dim AutoSize As Boolean
Dim Flags As Long
Dim AndTest As Long
Dim Maximize As String
Dim LCIDString As String

OnTop = False
AutoSize = False
AdjustResolution = True
Windowsstring = ""
Optionsstring = ""
Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Get Dateinummer, , sysheader
If sysheader.Minor <= 16 Then
'not compressed, TopicBlockSize 2k
iscompress = False
Else
Select Case sysheader.Flags
Case 0
iscompress = False
'Flags=0: not compressed, TopicBlockSize 4k
Case 4
iscompress = True
'Flags=4: LZ77 compressed, TopicBlockSize 4k
Case 8
iscompress = True
'Flags=8: LZ77 compressed, TopicBlockSize 2k
End Select
End If
Do While Test = False
Get Dateinummer, , systemrec
Select Case systemrec.RecordType
Case 1 'help file title
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
Optionsstring = Optionsstring & vbCrLf & "TITLE=" & NameString
Case 2 'copyright notice shown in AboutBox
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
If InStr(NameString, vbCrLf) Then
NameString = Left(NameString, InStr(NameString, vbCrLf))
End If

NameString = Left(NameString, (Len(NameString) - 1))
If NameString <> "" Then
Optionsstring = Optionsstring & vbCrLf & "COPYRIGHT=" & NameString
End If
Case 3 'Contents topic offset of starting topic
Get Dateinummer, , Hilfslong
For i = 0 To UBound(Contexts)
If UBound(Contexts) = 0 And IDNames(0) = "" Then Exit For
If Contexts(i).Topicoffset = Hilfslong Then
Contentsnumber = i
Exit For
End If
Next i
Case 4 'all macros executed on opening
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
RepairMacroString NameString 'einer
MacrostringAll = MacrostringAll & NameString & vbCrLf
Case 5 'Windows *.ICO file  See WIN31WH on icon file format
ReDim Hilfsbytes(systemrec.DataSize - 1)
Get Dateinummer, , Hilfsbytes 'Icon ist in Array
Case 6 'Windows defined in the HPJ-file
Get Dateinummer, , a
AndTest = a.Flags And WSYSFLAG_NAME
If AndTest = WSYSFLAG_NAME Then
b = a.Names
b = StrConv(b, vbUnicode)
If InStr(b, vbNullChar) Then
b = Left(b, InStr(b, vbNullChar) - 1)
End If
ReDim Preserve windownames(anzwindows)
windownames(anzwindows) = b
anzwindows = anzwindows + 1
Else
b = ""
End If
Hilfsstring = b & "="
AndTest = a.Flags And WSYSFLAG_CAPTION
If AndTest = WSYSFLAG_CAPTION Then
b = a.Caption
b = StrConv(b, vbUnicode)
b = Left(b, InStr(b, vbNullChar) - 1)
Else
b = ""
End If
Hilfsstring = Hilfsstring & Chr(34) & b & Chr(34) & ","
If a.x = -1 Then
x = ""
Else
x = a.x
End If
AndTest = a.Flags And WSYSFLAG_X
If AndTest <> WSYSFLAG_X Then x = ""
If a.Y = -1 Then
Y = ""
Else: Y = a.Y
End If
AndTest = a.Flags And WSYSFLAG_Y
If AndTest <> WSYSFLAG_Y Then Y = ""
If a.Height = -1 Then
Height = ""
Else
Height = a.Height
End If
AndTest = a.Flags And WSYSFLAG_HEIGHT
If AndTest <> WSYSFLAG_HEIGHT Then Height = ""
If a.Width = -1 Then
Width = ""
Else: Width = a.Width
End If
AndTest = a.Flags And WSYSFLAG_WIDTH
If AndTest <> WSYSFLAG_WIDTH Then Width = ""
If x = "" And Y = "" And Height = "" And Width = "" Then
Hilfsstring = Hilfsstring & ","
Else
Hilfsstring = Hilfsstring & "(" & x & "," & Y & "," & Width & "," & Height & "),"
End If
Testdouble = IntegerToUnsigned(a.Maximize)
AndTest = a.Flags And WSYSFLAG_MAXIMIZE
If AndTest = WSYSFLAG_MAXIMIZE Then
If a.Maximize = 0 Then
Maximize = ""
Else
Maximize = Testdouble 'a.Maximize
End If
Else
If a.Maximize = 0 Then
Maximize = ""
Else
Maximize = Testdouble 'a.Maximize
End If
End If
AndTest = a.Flags And WSYSFLAG_RGB
If AndTest = WSYSFLAG_RGB Then
Farbe = Rgb(a.Rgb(0), a.Rgb(1), a.Rgb(2))
RGBString = "(r" & Farbe & ")"
Else
RGBString = ""
End If
Hilfsstring = Hilfsstring & Maximize & "," & RGBString & ","
AndTest = a.Flags And WSYSFLAG_RGBNSR
If AndTest = WSYSFLAG_RGBNSR Then
Farbe = Rgb(a.RgbNsr(0), a.RgbNsr(1), a.RgbNsr(2))
RGBSRString = "(r" & Farbe & ")"
Else
RGBSRString = ""
End If
Hilfsstring = Hilfsstring & RGBSRString
Flags = IntegerToUnsigned(a.Flags)
AndTest = Flags And WSYSFLAG_ADJUSTRRESOLUTION
If AndTest = WSYSFLAG_ADJUSTRRESOLUTION Then AdjustResolution = False
AndTest = Flags And WSYSFLAG_AUTOSIZEHEIGHT
If AndTest = WSYSFLAG_AUTOSIZEHEIGHT Then AutoSize = True
AndTest = Flags And WSYSFLAG_TOP
If AndTest = WSYSFLAG_TOP Then OnTop = True
If AutoSize = False And OnTop = True And AdjustResolution = True Then
Hilfsstring = Hilfsstring & ",f1"
End If
If AutoSize = True And OnTop = False And AdjustResolution = True Then
Hilfsstring = Hilfsstring & ",f2"
End If
If AutoSize = True And OnTop = True And AdjustResolution = True Then
Hilfsstring = Hilfsstring & ",f3"
End If
If AutoSize = False And OnTop = False And AdjustResolution = False Then
Hilfsstring = Hilfsstring & ",f4"
End If
If AutoSize = False And OnTop = True And AdjustResolution = False Then
Hilfsstring = Hilfsstring & ",f5"
End If
If AutoSize = True And OnTop = False And AdjustResolution = False Then
Hilfsstring = Hilfsstring & ",f6"
End If
If AutoSize = True And OnTop = True And AdjustResolution = True Then
Hilfsstring = Hilfsstring & ",f7"
End If
'wenn nur Adjust gar nichts
AndTest = a.Flags And WSYSFLAG_TYPE
If AndTest = WSYSFLAG_TYPE Then
b = a.Types
b = StrConv(b, vbUnicode)
b = Left(b, InStr(b, vbNullChar) - 1)
End If
Windowsstring = Windowsstring & Hilfsstring & vbCrLf
Case 8 'Citation  the Citation printed
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
If NameString <> "" Then
Optionsstring = Optionsstring & vbCrLf & "CITATION=" & NameString
End If
Case 9 'language ID, Windows 95 (HCW 4.00)
Get Dateinummer, , LCID
LCIDString = "LCID=0x" & Hex(LCID(4)) & " 0x" & Hex(LCID(0)) & " 0x" & Hex(LCID(1))
Optionsstring = Optionsstring & vbCrLf & LCIDString
Case 10 'ContentFileName CNT file name, Windows 95 (HCW 4.00)
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
CNTFilename = NameString
NumCnt = NumCnt + 1
ReDim Preserve Namecnt(NumCnt)
Namecnt(NumCnt) = CNTFilename
Optionsstring = Optionsstring & vbCrLf & "CNT=" & CNTFilename
Case 11 'Charset charset, Windows 95 (HCW 4.00)
ReDim Bytearray(systemrec.DataSize - 1)
Get Dateinummer, , Bytearray '5 Bytes???
Case 12 'default dialog font, Windows 95 (HCW 4.00)
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
Case 13 'Group  defined GROUPs, Multimedia Help File
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
Case 14 'IndexSeparators separators, Windows 95 (HCW 4.00)
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
Case 18 'defined language, Multimedia Help Files
NameString = Space(systemrec.DataSize)
Get Dateinummer, , NameString
NameString = Left(NameString, (Len(NameString) - 1))
MsgBox NameString
Case 19 'defined DLLMAPS, Multimedia Help Files
''nicht unterstützt
ReDim Hilfsbytes(systemrec.DataSize - 1)
Get Dateinummer, , Hilfsbytes
Case Else
Test = True
End Select
Stand = Seek(Dateinummer)
If Stand >= Ende Then
Exit Do
End If
Loop
If Windowsstring <> "" Then
Windowsstring = "[WINDOWS]" & vbCrLf & Windowsstring & vbCrLf
End If
If Optionsstring <> "" Then
Optionsstring = Mid(Optionsstring, 3) ' 1.vbcrlf abschneiden
End If
End Function

Public Function GetTopicFile(FileOffset As Long) As Long
Dim FirstTop As Boolean
Dim AnfangString As Long
Dim EndeString As Long
Dim Berechnungsarray() As Long
Dim testx1 As Long
Dim trgaph As Double
Dim testberechnung As Long
Dim testberechnung1 As Long
Dim testlong As Long
Dim LenTitle As Long
Dim OK As Boolean
Dim endtest As Boolean
Dim cell As Double
Dim trleft As Double
Dim x1 As Long
Dim Textvorhanden As Boolean
Dim BeginnofFormatCode As Long
Dim wielang As Integer
Dim Hilfsarray() As Byte
Dim Topicsize As Long
Dim TopicLength As Integer
Dim NumberofColumns As Byte
Dim MinTableWidth As Integer
Dim tabletype As Byte
Dim hil As String
Dim Decgr As Long
Dim Enden As Boolean
Dim AktPetra As Long
Dim starte As Long
Dim alteBlockNr As Long
Dim wo As Long
Dim Ende As Long
Dim Dnr As Long
Dim z As Long
Dim l1 As Long
Dim bargröße As Long
Dim testoffset As Long
Dim Hilfslong As Long
Dim Hilfsint As Integer
Dim Hilfsbyte As Byte
Dim Hilfsbyte1 As Byte
Dim LAll As Long
Dim Dateigröße As Long
Dim TestString As String
Dim bar() As Byte
Dim Topictitle As String
Dim TopicBlocknumber As Long
Dim Standpunktformat As Long
Dim test1 As Boolean
Dim Phrasecompr As Boolean
Dim Standpunkt As Long
Dim Testbyte() As Byte
Dim tbhdr As TOPICBLOCKHEADER
Dim tlnk As TOPICLINK
Dim Ausgleich As Long
Dim Linkstandpunkt As Long
Dim XEnd As Boolean
Dim TopName As String
Dim Länge As Long
Dim Standpunkttbhdr As Long
Dim thdr As TOPICHEADER
Dim Kolumnen() As Integer
Dim RecordType As Byte
Dim te As Long
Dim Koltype As COLUMNSTRUCT
Dim ab As Long
Dim Formatstring As String
Dim NextTophdr As Long
Dim AnfangsID As Long

FirstTop = False
RTFGröße = 0
AktPetra = 1
ReDim Browsen(0)
NOH = 0
BrowseBlocknumber = 0
Browsenanzahl = 0
Makrostringfertig = ""
Aktuellestopicoffset = 0
Standpunkt = 0
CharacterCount = 0
BrowseFalse = False
AktCharacter = 0
Blocknummer = 0
Fertigstring = ""
Dateigröße = 0
gefundeneID = 0
Reserveoffset = 0
NextGoodBrowseOffset = 0
Aktuellestopicoffset = 0
Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Standpunkttbhdr = Seek(Dateinummer)
Get Dateinummer, , tbhdr
If Filehdr.UsedSpace > 4096 Then
Decgr = 4095
Else
Decgr = Filehdr.UsedSpace - 1
End If
ReDim bar(Decgr - 12)
Get Dateinummer, , bar
If iscompress = True Then Decompress UBound(bar) + 1, bar
Standpunkt = tbhdr.FirstTopicLink - 12
bargröße = UBound(bar)
TopicBlocknumber = Standpunkt \ &H4000
alteBlockNr = TopicBlocknumber
Standpunkt = Standpunkt Mod &H4000
XEnd = False
Do While XEnd = False
LetzterString = TestString
TestString = ""
hil = ""
If Standpunkt + 21 > UBound(bar) Then 'Typelink geht in neuen Block
Klebe bar, Standpunkttbhdr, Filehdr.UsedSpace, TopicBlocknumber, iscompress
End If
CopyMemory ByVal VarPtr(tlnk), bar(Standpunkt), 21
Linkstandpunkt = Standpunkt
If tlnk.NextBlock \ &H4000 > alteBlockNr Then 'Nächster Link in nächstem Block
If tlnk.NextBlock Mod &H4000 > 12 Then 'nächster Link nicht gleich am Anfang
Klebe bar, Standpunkttbhdr, Filehdr.UsedSpace, TopicBlocknumber, iscompress
Reserveoffset = (NextTophdr + 1) * 32768
Else
Reserveoffset = (NextTophdr + 1) * 32768
End If
End If
If tlnk.NextBlock = -1 Then
XEnd = True
End If
If tlnk.DataLen2 > tlnk.BlockSize - tlnk.DataLen1 Then
Phrasecompr = True
Else
Phrasecompr = False
End If
Reserveoffset = 0
RecordType = tlnk.RecordType
Select Case tlnk.RecordType
Case 2
Topictitle = ""
Textvorhanden = False
If Standpunkt + 21 + tlnk.BlockSize < bargröße Then
Aktuellestopicoffset = CharacterCount + (Blocknummer * 32768)
Else
Aktuellestopicoffset = (Blocknummer + 1) * 32768
Reserveoffset = CharacterCount + (Blocknummer * 32768)
End If
CopyMemory ByVal VarPtr(thdr), bar(Standpunkt + 21), 28 'erster Topic = 12 - 12 = 0
If thdr.NextTopic = -1 Then
End If
If thdr.NextTopic <> -1 Then
NextTophdr = thdr.NextTopic \ &H4000
TopZähler = thdr.Topicnum
FirstTop = False
If TopZähler = 0 Then
If PetraAnzahl = 0 Or MoreRTF = False Then
Fertigstring = ""
Fertigstring = MakeRTFKopf
Dnr = FreeFile
Open Pfad & "\" & PetraFile(0).RTFName For Binary As Dnr
AktDateiname = PetraFile(0).RTFName
Dateigröße = 0
End If
End If

If PetraAnzahl > 0 And MoreRTF = True Then
If TopZähler = 0 Then
Fertigstring = ""
Fertigstring = MakeRTFKopf
Dnr = FreeFile
Open Pfad & "\" & PetraFile(1).RTFName For Binary As Dnr
AktDateiname = PetraFile(1).RTFName
RTFGröße = 0
AktPetra = 2
Else
If AktPetra <= PetraAnzahl Then
If PetraFile(AktPetra).FirstNumber = TopZähler + 1 Then
FirstTop = True
ab = LOF(Dnr)
Dateigröße = Dateigröße + Len(Fertigstring)
Fertigstring = Fertigstring & "}"
Put Dnr, ab + 1, Fertigstring
Fertigstring = ""
Fertigstring = MakeRTFKopf
RTFGröße = RTFGröße + Len(Dnr)
Close Dnr
Dnr = FreeFile
Open Pfad & "\" & PetraFile(AktPetra).RTFName For Binary As Dnr
AktDateiname = PetraFile(AktPetra).RTFName
Dateigröße = 0
AktPetra = AktPetra + 1
End If
End If
End If
End If
Form1.Label1.Caption = "Read Topic " & TopZähler + 1
DoEvents
If TopZähler > 0 Then
If thdr.NextTopic > 0 Then
If ScrollorNoscroll = True Then
If Underline Then Formatstring = Formatstring & "\ul "
Fertigstring = Fertigstring & "\plain " & vbCrLf 'Sicherheit
TabellenAlign = 0
ScrollorNoscroll = False
End If

Select Case ZuletztAbsatz
Case True
If FirstTop = False Then
Fertigstring = Fertigstring & vbCrLf & "\plain\pard\page" & vbCrLf
End If
TabellenAlign = 0
Case False
If FirstTop = False Then
Fertigstring = Fertigstring & vbCrLf & "\plain\par\pard\page" & vbCrLf
End If
TabellenAlign = 0
End Select
DoEvents
End If
End If
If tlnk.DataLen2 > 0 Then
ReDim Testbyte(tlnk.DataLen2)
CopyMemory Testbyte(0), bar(Standpunkt + 20 + 28 + 1), tlnk.DataLen2
If Phrasecompr = True Then
TestString = Dephrase(Hasphrase, Testbyte, tlnk.DataLen2)
Else
TestString = Testbyte 'keine Komprimierung
End If
End If
TopName = ""
TestString = StrConv(TestString, vbUnicode)
LenTitle = Len(TestString)
If gefundeneID <= UBound(Contexts) Then
'Fehlerkorrektur
Do While Contexts(gefundeneID).Topicoffset < Aktuellestopicoffset
TopName = IDNames(gefundeneID)
fertigeIDs(gefundeneID) = True
gefundeneID = gefundeneID + 1 'zur Sicherheit
Loop
If (Contexts(gefundeneID).Topicoffset >= Aktuellestopicoffset And Contexts(gefundeneID).Topicoffset < Aktuellestopicoffset + LenTitle) Or ((Contexts(gefundeneID).Topicoffset >= Reserveoffset And Contexts(gefundeneID).Topicoffset < Reserveoffset + LenTitle) And Reserveoffset <> 0) Then

Hilfsbyte = 1
Do While Hilfsbyte = 1
If gefundeneID < UBound(Contexts) Then
If Contexts(gefundeneID).Topicoffset = Contexts(gefundeneID + 1).Topicoffset Then
gefundeneID = gefundeneID + 1
Else
Hilfsbyte = 0
End If
Else
Hilfsbyte = 0
End If
Loop
TopName = IDNames(gefundeneID)
fertigeIDs(gefundeneID) = True
gefundeneID = gefundeneID + 1
End If
End If

If TopName <> "" Then
'eben
Fertigstring = Fertigstring & "{\up #}{\footnote\pard\plain{\up #} " & TopName & "}" & vbCrLf
End If
If hasviolas And TopName <> "" Then
Findviolas Aktuellestopicoffset, Fertigstring
End If

Scrollregion = thdr.Scroll
NonScrollregion = thdr.NonScroll
If TestString <> "" And TestString <> Chr(0) Then
Hilfslong = InStr(TestString, Chr(0))
If Hilfslong = 0 Then
Topictitle = TestString
Else 'Hat Macros
Topictitle = Left(TestString, Hilfslong - 1)
Macrostring = Mid(TestString, Hilfslong + 1)
If InStr(Macrostring, Chr(0)) Then
testlong = InStr(Macrostring, Chr(0))
Macrostring = Left(Macrostring, testlong - 1)
End If
End If
If Topictitle <> "" Then
Fertigstring = Fertigstring & "{\up $}{\footnote\pard\plain{\up $} " & Topictitle & "}" & vbCrLf
End If
If Hilfslong > 0 Then
If Macrostring <> "" Then
RepairMacroString Macrostring 'einer
TestMakro Macrostring
Fertigstring = Fertigstring & "{\up !}{\footnote\pard\plain{\up !} " & Macrostring & "}" & vbCrLf 'Macros nicht unterstützt
End If
End If
End If
CreateBrowseSequence Aktuellestopicoffset, thdr.BrowseFor, thdr.BrowseBck, thdr.Topicnum, Fertigstring, Dateigröße
If HasKKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("K", Aktuellestopicoffset, LenTitle)
End If
If HasAKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("A", Aktuellestopicoffset, LenTitle)
End If

Else
End If
Topictitle = ""
Macrostring = ""
TestString = ""
Case &H23 'Tabelle
If Right(Fertigstring, 1) = "}" Then 'direkt vor Tabelle Fußnoten
Fertigstring = Fertigstring & vbCrLf & "\par" & vbCrLf
End If
Fertigstring = Fertigstring & vbCrLf & "\trowd"
TabellenAlign = 0
Linkstandpunkt = Standpunkt
ReDim Hilfsarray(3)
CopyMemory Hilfsarray(0), bar(Standpunkt + 21), 4
Topicsize = scanlong(Hilfsarray, wielang) 'TopicSize compressed
Standpunkt = Standpunkt + wielang + 21
Hilfsbyte = bar(Standpunkt)
Hilfsbyte1 = bar(Standpunkt + 1)
TopicLength = ReadCompUnSignShort(Hilfsbyte, Hilfsbyte1, wielang) 'TopicLength compressed
'MitteOffset = MitteOffset + TopicLength
CharacterCount = CharacterCount + TopicLength
Standpunkt = Standpunkt + wielang
NumberofColumns = bar(Standpunkt)
Standpunkt = Standpunkt + 1
tabletype = bar(Standpunkt)
Standpunkt = Standpunkt + 1
Select Case tabletype
Case 0, 2
CopyMemory ByVal VarPtr(MinTableWidth), bar(Standpunkt), 2
Fertigstring = Fertigstring & "\trqc"
l1 = MinTableWidth
Standpunkt = Standpunkt + 2
Case Else
l1 = 32767
End Select
CopyMemory ByVal VarPtr(Hilfslong), bar(Standpunkt), 4
If NumberofColumns > 1 Then
ReDim Kolumnen((NumberofColumns * 2) - 1)
CopyMemory ByVal VarPtr(Kolumnen(0)), bar(Standpunkt), 4 * NumberofColumns
Standpunkt = Standpunkt + 4 * NumberofColumns
x1 = Kolumnen(0) + Kolumnen(1) + Kolumnen(3) \ 2
trgaph = ((Kolumnen(3) * scaling - rounderr) * l1) \ 32767
Fertigstring = Fertigstring & "\trgaph" & trgaph
trleft = (((Kolumnen(1) - Kolumnen(3)) * scaling - rounderr) * l1 - 32767) \ 32767
Fertigstring = Fertigstring & "\trleft" & trleft
ReDim Berechnungsarray(NumberofColumns - 1) 'Berechnung der Spaltenbreite
testberechnung = ((x1 * scaling - rounderr) * l1) \ 32767
Berechnungsarray(0) = testberechnung
testberechnung1 = (((x1 + Kolumnen(2) + Kolumnen(3)) * scaling - rounderr) * l1) \ 32767
Berechnungsarray(1) = testberechnung1
testx1 = x1
If NumberofColumns > 2 Then
testx1 = testx1 + Kolumnen(2) + Kolumnen(3)
For z = 2 To NumberofColumns - 1
testx1 = testx1 + Kolumnen(2 * z) + Kolumnen(2 * z + 1)
Berechnungsarray(z) = ((testx1 * scaling - rounderr) * l1) \ 32767
Next z
End If
BerechneTabellenbreite Berechnungsarray, tabletype, trleft 'Spaltenbreiten berechnen
'1.Spalte
cell = ((x1 * scaling - rounderr) * l1) \ 32767
If cell <> Berechnungsarray(0) Then
cell = Berechnungsarray(0)
End If
Fertigstring = Fertigstring & " \cellx" & cell
'2.Spalte
cell = (((x1 + Kolumnen(2) + Kolumnen(3)) * scaling - rounderr) * l1) \ 32767
If cell <> Berechnungsarray(1) Then
cell = Berechnungsarray(1)
End If
'End If
Fertigstring = Fertigstring & "\cellx" & cell
x1 = x1 + Kolumnen(2) + Kolumnen(3)
For z = 2 To NumberofColumns - 1
x1 = x1 + Kolumnen(2 * z) + Kolumnen(2 * z + 1)

'3.Spalte aufwärts
cell = ((x1 * scaling - rounderr) * l1) \ 32767
If cell <> Berechnungsarray(z) Then
cell = Berechnungsarray(z)
End If
Fertigstring = Fertigstring & "\cellx" & cell
Next z
Else
ReDim Kolumnen(1)
CopyMemory ByVal VarPtr(Kolumnen(0)), bar(Standpunkt), 4 * NumberofColumns
Standpunkt = Standpunkt + 4
trleft = ((Kolumnen(1) * scaling - rounderr) * l1 - 32767) \ 32767
Fertigstring = Fertigstring & "\trleft" & trleft
cell = ((Kolumnen(0) * scaling - rounderr) * l1) \ 32767
If cell < 500 Then 'Achtung!!
cell = 9000
End If
Fertigstring = Fertigstring & "\cellx" & cell

End If

CopyMemory ByVal VarPtr(Koltype), bar(Standpunkt), 5

lastcol = Koltype.Columnn
Standpunkt = Standpunkt + 5
Standpunkt = Standpunkt + 2 'EndNullen
Standpunkt = Standpunkt + 2 ' 'Unknown + Biased char
CopyMemory ByVal VarPtr(Hilfsint), bar(Standpunkt), 2 'id
Standpunkt = Standpunkt + 2
If Hilfsint <> 0 Then
Standpunkt = Standpunkt + TestID(Hilfsint, Standpunkt, bar, Fertigstring, RecordType)
End If
Fertigstring = Fertigstring & "\pard\intbl" & vbCrLf
BeginnofFormatCode = Standpunkt
Case &H20
Linkstandpunkt = Standpunkt
ReDim Hilfsarray(3)
CopyMemory Hilfsarray(0), bar(Standpunkt + 21), 4
Topicsize = scanlong(Hilfsarray, wielang) 'TopicSize compressed
Standpunkt = Standpunkt + wielang + 21
Hilfsbyte = bar(Standpunkt)
Hilfsbyte1 = bar(Standpunkt + 1)
TopicLength = ReadCompUnSignShort(Hilfsbyte, Hilfsbyte1, wielang) 'TopicLength compressed
'MitteOffset = MitteOffset + TopicLength
CharacterCount = CharacterCount + TopicLength
Standpunkt = Standpunkt + wielang
Standpunkt = Standpunkt + 2 'EndNullen
Standpunkt = Standpunkt + 2 ' 'Unknown + Biased char
CopyMemory ByVal VarPtr(Hilfsint), bar(Standpunkt), 2 'id
Standpunkt = Standpunkt + 2
If Hilfsint <> 0 Then
Standpunkt = Standpunkt + TestID(Hilfsint, Standpunkt, bar, Fertigstring, RecordType)
End If
BeginnofFormatCode = Standpunkt
End Select
If tlnk.DataLen2 > 0 And tlnk.RecordType <> 2 Then
ReDim Testbyte(tlnk.DataLen2 - 1)
Länge = tlnk.DataLen2
CopyMemory Testbyte(0), bar(0 + tlnk.DataLen1 + Linkstandpunkt), Länge
'Phrase
If Phrasecompr Then
TestString = Dephrase(Hasphrase, Testbyte, tlnk.DataLen2)
Else
TestString = Testbyte 'keine Komprimierung
End If
testoffset = AktCharacter + (Blocknummer * 32768)
If testoffset >= NonScrollregion And testoffset < Scrollregion And ScrollorNoscroll = False Then
Fertigstring = Fertigstring & "\keepn "
ScrollorNoscroll = True
End If

If testoffset >= Scrollregion And ScrollorNoscroll = True And Scrollregion <> -1 Then
Select Case RecordType
Case &H23
Fertigstring = Fertigstring & "\plain "
TabellenAlign = 0
Case &H20
Fertigstring = Fertigstring & "\pard\plain "
TabellenAlign = 0
End Select
ScrollorNoscroll = False
End If
hil = StrConv(TestString, vbUnicode)
Enden = False
starte = 1
Standpunktformat = BeginnofFormatCode
'LAll = 0
Do While Enden = False
wo = InStr(starte, hil, Chr(0))
AnfangString = 0
If RecordType = &H23 Then
Ausgleich = 0
Else
Ausgleich = 0
End If
If wo > Len(hil) - Ausgleich Then Exit Do '-1 oder????

If wo = 0 Then
Enden = True
Exit Do
Else
If wo = starte Then
EndeString = wo + 1
AnfangString = starte
Else
AnfangString = starte
EndeString = wo
End If
If EndeString = 0 Then
EndeString = Len(hil) + 1
AnfangString = starte
Enden = True
End If
TestString = Mid(hil, AnfangString, EndeString - AnfangString)

If TestString <> "" And TestString <> Chr(0) Then

Textvorhanden = True
End If
Formatstring = ""
writeFormatcommand RecordType, bar, Standpunktformat, Formatstring, Textvorhanden
te = LAll + NOH + (Blocknummer * 32768)
If TestString <> "" And TestString <> Chr(0) Then
LAll = LAll + Len(TestString) + 1
Else
LAll = LAll + 1
End If
OK = False
If gefundeneID < UBound(Contexts) Then
Hilfsbyte = 1
Do While Hilfsbyte = 1
If gefundeneID < UBound(Contexts) Then
If Contexts(gefundeneID).Topicoffset = Contexts(gefundeneID + 1).Topicoffset Then
gefundeneID = gefundeneID + 1
Else
Hilfsbyte = 0
End If
Else
Hilfsbyte = 0
End If
Loop

If te = Contexts(gefundeneID).Topicoffset Then
TopName = IDNames(gefundeneID)
If TopName <> "" Then
Fertigstring = Fertigstring & "{\up #}{\footnote\pard\plain{\up #} " & TopName & "}" & vbCrLf
If hasviolas And TopName <> "" Then
Findviolas Aktuellestopicoffset, Fertigstring
End If
If HasKKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("K", Contexts(gefundeneID).Topicoffset, 1, False)
End If
If HasAKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("A", Contexts(gefundeneID).Topicoffset, 1, False)
End If
End If
fertigeIDs(gefundeneID) = True
gefundeneID = gefundeneID + 1 'zur Sicherheit
OK = True
End If
If gefundeneID <= UBound(Contexts) Then
If te + Len(TestString) > Contexts(gefundeneID).Topicoffset And TestString <> Chr(0) And OK = False Then
endtest = False
Do While Contexts(gefundeneID).Topicoffset < te + Len(TestString) And endtest = False
TopName = IDNames(gefundeneID)
test1 = False
Do While test1 = False
If gefundeneID < UBound(Contexts) Then
If Contexts(gefundeneID).Topicoffset = Contexts(gefundeneID + 1).Topicoffset Then
gefundeneID = gefundeneID + 1
Else
test1 = True
End If
Else
test1 = True
gefundeneID = gefundeneID - 1 'zurückstellen
endtest = True
End If
Loop
If TopName <> "" Then
Fertigstring = Fertigstring & "{\up #}{\footnote\pard\plain{\up #} " & TopName & "}" & vbCrLf
If hasviolas And TopName <> "" Then
Findviolas Aktuellestopicoffset, Fertigstring
End If
If HasKKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("K", Contexts(gefundeneID).Topicoffset, 1, False)
End If
If HasAKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("A", Contexts(gefundeneID).Topicoffset, 1, False)
End If
End If
fertigeIDs(gefundeneID) = True
gefundeneID = gefundeneID + 1 'zur Sicherheit
Loop
End If
End If
End If
If CharacterCount + (Blocknummer * 32768) > te Then
If gefundeneID <= UBound(Contexts) Then
If te = Contexts(gefundeneID).Topicoffset Then
AnfangsID = gefundeneID
If gefundeneID < UBound(Contexts) Then
If Contexts(gefundeneID).Topicoffset = Contexts(gefundeneID + 1).Topicoffset Then
fertigeIDs(gefundeneID) = True
gefundeneID = gefundeneID + 1
Do While Contexts(gefundeneID).Topicoffset = Contexts(gefundeneID + 1).Topicoffset
fertigeIDs(gefundeneID) = True
gefundeneID = gefundeneID + 1
Loop
End If
End If
'eben
If IDNames(AnfangsID) <> "" Then
Fertigstring = Fertigstring & "{\up #}{\footnote\pard\plain{\up #} " & IDNames(AnfangsID) & "}" & vbCrLf
If hasviolas And TopName <> "" Then
Findviolas Aktuellestopicoffset, Fertigstring
End If
If HasKKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("K", Contexts(gefundeneID).Topicoffset, 1, False)
End If
If HasAKeywords And Topictitle <> "" Then
Fertigstring = Fertigstring & FindKeywordsForTopic("A", Contexts(gefundeneID).Topicoffset, 1, False)
End If
End If
fertigeIDs(gefundeneID) = True
gefundeneID = gefundeneID + 1
'End If
End If
End If
End If

If TestString <> "" And TestString <> Chr(0) Then
TestString = Replace(TestString, "\", "\'5c") 'RTF-Steuerzeichen im Text ersetzen
TestString = Replace(TestString, "{", "\{\-")
TestString = Replace(TestString, "}", "\'7d")
TestString = Replace(TestString, Chr(255), "\'" & Hex(255))
TestString = Replace(TestString, Chr(149), "\bullet")
Fertigstring = Fertigstring & TestString & Formatstring
If InStr(Formatstring, "\par\") Or InStr(Formatstring, "\par ") Or Right(Formatstring, 4) = "\par" Then
ZuletztAbsatz = True
Else
ZuletztAbsatz = False
End If
Formatstring = ""
Textvorhanden = True
Else
Fertigstring = Fertigstring & Formatstring

End If
If TestString = "" Or TestString = Chr(0) Then
starte = EndeString
Else
starte = EndeString + 1 'EndNull
End If
End If
Loop
End If
If tlnk.NextBlock <> -1 Then
Standpunkt = tlnk.NextBlock - 12
TopicBlocknumber = Standpunkt \ &H4000
If TopicBlocknumber <> alteBlockNr Then
LAll = 0
NOH = 0
alteBlockNr = TopicBlocknumber
Standpunkttbhdr = Standpunkttbhdr + 4096
Get Dateinummer, Standpunkttbhdr, tbhdr
Blocknummer = Blocknummer + 1
MitteOffset = 0
CharacterCount = 0
If Filehdr.UsedSpace > 4096 * (TopicBlocknumber + 1) Then
Decgr = 4095
Else
Decgr = Filehdr.UsedSpace - (4096 * TopicBlocknumber) - 1
End If
ReDim bar(Decgr - 12)
Get Dateinummer, , bar
If iscompress = True Then Decompress UBound(bar) + 1, bar
bargröße = UBound(bar)
End If
Standpunkt = Standpunkt Mod &H4000
Else
XEnd = True
End If
If Len(Fertigstring) > 10000 Then
Dateigröße = Dateigröße + Len(Fertigstring)
ab = LOF(Dnr)
Put Dnr, ab + 1, Fertigstring
RTFGröße = LOF(Dnr)
Fertigstring = ""
End If
AktCharacter = tlnk.NextBlock
Loop
'zumachen + 1
Fertigstring = Fertigstring & "}"
zumachen = zumachen + 1
ab = LOF(Dnr)
Dateigröße = Dateigröße + Len(Fertigstring)
Put Dnr, ab + 1, Fertigstring
If BrowseFalse Then 'verkehrte Reihenfolge
RepariereBrowseArray Dnr
End If
RTFGröße = RTFGröße + LOF(Dnr)
Close Dnr
End Function

Public Function GetxWBTREEFile(FileOffset As Long, Buchstabe As String) As Long
Dim btreehead As BTREEHEADER
Dim Keyword As String
Dim count As Integer
Dim xWDataOffset As Long
Dim xWBTreeArray() As Byte
Dim btreeNodhead As BTREENODEHEADER
Dim Ende As Long
Dim i As Long
Dim z As Long
Dim Länge As Long
Dim Aufzahl As Long
Dim Dateistand As Long
Dim DataOffset As Long
Dim Standpunkt As Long
Dim KeyBytes() As Byte
Dim KeywordArray() As String
Dim bisherigeAnzahlstrings As Long
Dim wobinich As Long
Dim w As Long
Dim anzahlkeywords As Long

If Buchstabe = "K" Or Buchstabe = "A" Then 'A und K-Keywords
ReDim KeywordArray(0)
Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
ReDim xWBTreeArray(Filehdr.UsedSpace - 1)
Dateistand = Seek(Dateinummer)
Get Dateinummer, Dateistand, xWBTreeArray
CopyMemory ByVal VarPtr(btreehead), xWBTreeArray(0), 38 'head einlesen
Standpunkt = 38
Select Case Buchstabe
Case "K"
ReDim WoistKKeyword(UBound(KWData))
Case "A"
ReDim WoistAKeyword(UBound(AWData))
End Select
Select Case btreehead.NLevels
Case 1
CopyMemory ByVal VarPtr(btreeNodhead), xWBTreeArray(Standpunkt), 8 'Nodehead einlesen
Standpunkt = Standpunkt + 8
anzahlkeywords = anzahlkeywords + btreeNodhead.NEntries
Select Case Buchstabe
Case "K"
ReDim Preserve KWBTreeString(anzahlkeywords - 1)
Case "A"
ReDim Preserve AWBTreeString(anzahlkeywords - 1)
End Select
For z = 0 To btreeNodhead.NEntries - 1     'Entries einlesen
Länge = lstrlen(VarPtr(xWBTreeArray(Standpunkt)))
ReDim KeyBytes(Länge - 1)
CopyMemory KeyBytes(0), xWBTreeArray(Standpunkt), Länge
Keyword = StrConv(KeyBytes, vbUnicode)
Select Case Buchstabe
Case "K"
KWBTreeString(z) = Keyword
Case "A"
AWBTreeString(z) = Keyword
End Select
CopyMemory ByVal VarPtr(count), xWBTreeArray(Standpunkt + Länge + 1), 2
CopyMemory ByVal VarPtr(xWDataOffset), xWBTreeArray(Standpunkt + Länge + 3), 4
wobinich = xWDataOffset \ 4
For w = 1 To count
Select Case Buchstabe
Case "K"
WoistKKeyword(wobinich) = z 'Indexnummer des Strings
Case "A"
WoistAKeyword(wobinich) = z 'Indexnummer des Strings
End Select
wobinich = wobinich + 1
Next w
Standpunkt = Standpunkt + Länge + 6 + 1
Next z

Case Else
For i = 1 To btreehead.TotalPages
If i - 1 = btreehead.RootPage Then
Else
CopyMemory ByVal VarPtr(btreeNodhead), xWBTreeArray(Standpunkt), 8 'Nodehead einlesen
Aufzahl = Aufzahl + 8
anzahlkeywords = anzahlkeywords + btreeNodhead.NEntries
Select Case Buchstabe
Case "K"
ReDim Preserve KWBTreeString(anzahlkeywords - 1)
Case "A"
ReDim Preserve AWBTreeString(anzahlkeywords - 1)
End Select
For z = 0 + bisherigeAnzahlstrings To btreeNodhead.NEntries + bisherigeAnzahlstrings + -1   'Entries einlesen
Länge = lstrlen(VarPtr(xWBTreeArray(Standpunkt + Aufzahl)))
ReDim KeyBytes(Länge - 1)
CopyMemory KeyBytes(0), xWBTreeArray(Standpunkt + Aufzahl), Länge
Keyword = StrConv(KeyBytes, vbUnicode)
Select Case Buchstabe
Case "K"
KWBTreeString(z) = Keyword
Case "A"
AWBTreeString(z) = Keyword
End Select
CopyMemory ByVal VarPtr(count), xWBTreeArray(Standpunkt + Aufzahl + Länge + 1), 2
CopyMemory ByVal VarPtr(xWDataOffset), xWBTreeArray(Standpunkt + Aufzahl + Länge + 3), 4
DataOffset = xWDataOffset \ 4
For w = 1 To count
Select Case Buchstabe
Case "K"
WoistKKeyword(DataOffset) = z 'Indexnummer des Strings
Case "A"
WoistAKeyword(DataOffset) = z 'Indexnummer des Strings
End Select
DataOffset = DataOffset + 1
wobinich = wobinich + 1
Next w
Aufzahl = Aufzahl + Länge + 6 + 1
Next z
bisherigeAnzahlstrings = bisherigeAnzahlstrings + btreeNodhead.NEntries
End If
Aufzahl = 0
Standpunkt = Standpunkt + btreehead.PageSize
Next i
End Select
End If
End Function

Public Function GetTtlbtreeFile(FileOffset As Long) As Long ' $ - Keyword - Order
Dim btreehead As BTREEHEADER
Dim headerstand As Long
Dim Keyword As String
Dim btreeNodhead As BTREENODEHEADER
Dim Ende As Long
Dim Testnummer As Long
Dim i As Long
Dim z As Long
Dim Topicoffset As Long
Dim Länge As Long
Dim Dateistand As Long

Get Dateinummer, FileOffset + 1, Filehdr
Ende = FileOffset + Filehdr.UsedSpace + Len(Filehdr)
Get Dateinummer, , btreehead
Select Case btreehead.NLevels
Case 1 'Leafpage
For i = 1 To btreehead.NLevels 'bei Leafentry 1
Get Dateinummer, , btreeNodhead
For z = 1 To btreeNodhead.NEntries - 1
Get Dateinummer, , Topicoffset
Dateistand = Seek(Dateinummer) + 1 'EndNull
Länge = FindStringEnd(Dateistand)
Keyword = Space(Länge - 1)
Get Dateinummer, Dateistand, Keyword
AnzTTLBTREE = AnzTTLBTREE + 1
ReDim Preserve TTLBNames(AnzTTLBTREE)
TTLBNames(AnzTTLBTREE).Topictitle = Keyword
TTLBNames(AnzTTLBTREE).Topicoffset = Topicoffset
Next z
Next i
Case Else 'Indexpage
headerstand = Seek(Dateinummer)
For i = 1 To btreehead.TotalPages
If i - 1 = btreehead.RootPage Then
Else
Get Dateinummer, headerstand, btreeNodhead
For z = 1 To btreeNodhead.NEntries - 1
Get Dateinummer, , Topicoffset
ReDim Preserve TTLBTreeOffsets(Testnummer)
TTLBTreeOffsets(Testnummer) = Topicoffset
Testnummer = Testnummer + 1
Dateistand = Seek(Dateinummer) + 1
Länge = FindStringEnd(Dateistand)
Keyword = Space(Länge - 1)
Get Dateinummer, Dateistand, Keyword
AnzTTLBTREE = AnzTTLBTREE + 1
ReDim Preserve TTLBNames(AnzTTLBTREE)
TTLBNames(AnzTTLBTREE).Topictitle = Keyword
TTLBNames(AnzTTLBTREE).Topicoffset = Topicoffset
Next z
End If
headerstand = headerstand + btreehead.PageSize
Next i
End Select
End Function

Public Function GetContextFile(FileOffset As Long) As Long
Dim btreehead As BTREEHEADER
Dim IDNummer As Long
Dim IDNString As String
Dim btreeNodhead As BTREENODEHEADER
Dim headerstand As Long
Dim ConLeaf As CONTEXTLEAF
Dim i As Long
Dim anzahl As Long
Dim z As Long

Get Dateinummer, FileOffset + 1, Filehdr
Get Dateinummer, , btreehead
If btreehead.NLevels < 1 Then
ReDim Contexts(0)
ReDim IDNames(0)
Exit Function
End If
Select Case btreehead.NLevels
Case 1 'Leafpage
Get Dateinummer, , btreeNodhead
ReDim Contexts(btreeNodhead.NEntries - 1)
Get Dateinummer, , Contexts
Case Else 'Indexpage
headerstand = Seek(Dateinummer)
For i = 1 To btreehead.TotalPages
If i - 1 = btreehead.RootPage Then
Else
Get Dateinummer, headerstand, btreeNodhead
For z = 1 To btreeNodhead.NEntries
ReDim Preserve Contexts(anzahl)
Get Dateinummer, , ConLeaf
Contexts(anzahl).Hashvalue = ConLeaf.Hashvalue
Contexts(anzahl).Topicoffset = ConLeaf.Topicoffset
anzahl = anzahl + 1
Next z
End If
headerstand = headerstand + btreehead.PageSize
Next i
End Select
Topicanzahl = UBound(Contexts) + 1
ReDim Namegefunden(Topicanzahl - 1)
ReDim IDNames(Topicanzahl - 1)
ReDim fertigeIDs(Topicanzahl - 1)
For i = 0 To Topicanzahl - 1
IDNummer = Contexts(i).Hashvalue
IDNString = unhash(IDNummer)
IDNames(i) = IDNString
Next i
SortContexts
End Function

Public Function GetPetraFile(FileOffset As Long) As Long
If MoreRTF = False Then Exit Function 'dann unnötig
Dim btreehead As BTREEHEADER
Dim RTFName As String
Dim OldRTFName As String
Dim btreeNodhead As BTREENODEHEADER
Dim Testnummer As Long
Dim i As Long
Dim z As Long
Dim Topicoffset As Long
Dim Länge As Long
Dim Dateistand As Long
Dim headerstand As Long

Get Dateinummer, FileOffset + 1, Filehdr
Get Dateinummer, , btreehead

Select Case btreehead.NLevels
Case 1 'Leafpage
For i = 1 To btreehead.NLevels 'bei Leafentry 1
Get Dateinummer, , btreeNodhead
For z = 1 To btreeNodhead.NEntries
Get Dateinummer, , Topicoffset
Dateistand = Seek(Dateinummer)
Länge = FindStringEnd(Dateistand)
RTFName = Space(Länge)
Get Dateinummer, Dateistand, RTFName
If InStr(RTFName, Chr(0)) Then
RTFName = Left(RTFName, InStr(RTFName, Chr(0)) - 1)
RTFName = FilenameFromPath(RTFName)
If OldRTFName <> RTFName Then
If RTFName <> "" Then
PetraAnzahl = PetraAnzahl + 1
ReDim Preserve PetraFile(PetraAnzahl)
PetraFile(PetraAnzahl).RTFName = RTFName
PetraFile(PetraAnzahl).FirstNumber = Topicoffset
End If
OldRTFName = RTFName
End If
End If
Next z
Next i
Case Else 'Indexpage
headerstand = Seek(Dateinummer)
For i = 1 To btreehead.TotalPages
If i - 1 = btreehead.RootPage Then
Else
Get Dateinummer, headerstand, btreeNodhead
For z = 0 To btreeNodhead.NEntries 'Entries einlesen
Get Dateinummer, , Topicoffset
ReDim Preserve TTLBTreeOffsets(Testnummer)
TTLBTreeOffsets(Testnummer) = Topicoffset
Testnummer = Testnummer + 1
Dateistand = Seek(Dateinummer)
Länge = FindStringEnd(Dateistand)
RTFName = Space(Länge)
Get Dateinummer, Dateistand, RTFName
If InStr(RTFName, Chr(0)) Then
RTFName = Left(RTFName, InStr(RTFName, Chr(0)) - 1)
RTFName = FilenameFromPath(RTFName)
If OldRTFName <> RTFName Then
If RTFName <> "" Then
PetraAnzahl = PetraAnzahl + 1
ReDim Preserve PetraFile(PetraAnzahl)
PetraFile(PetraAnzahl).RTFName = RTFName
PetraFile(PetraAnzahl).FirstNumber = Topicoffset
End If
OldRTFName = RTFName
End If
End If
Next z
End If
headerstand = headerstand + btreehead.PageSize
Next i
End Select
End Function



