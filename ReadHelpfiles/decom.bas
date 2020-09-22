Attribute VB_Name = "Deco"
Option Explicit
Public Aufmachen As Long
Public zumachen As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Maske() As Byte
Public Pfad As String
Public ProjektName As String
Public PhraseImageArray() As String
Public PhraseArray() As String
Private wobinich As Long
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767
Public Type RGBColor
    Red As Byte
    Green As Byte
    Blue As Byte
End Type


Public Function Decompress(CompSize As Long, Buffer() As Byte) As Long
Dim Inbytes As Long
Dim OutBytes As Long
Dim Bitmap As Byte
Dim Sets() As Byte
Dim NumToRead As Integer
Dim Index As Integer
Dim Counter As Integer
Dim Length As Integer
Dim Distance As Integer
Dim CurrPos As Long
Dim CodePtr As Long
Dim Buffertemp() As Byte
ReDim Buffertemp(CompSize * 8)
Dim Zählbyte As Long
On Error Resume Next
Zählbyte = 0
Do While Inbytes < CompSize '+ 1
    CopyMemory Bitmap, Buffer(Inbytes), 1
    Zählbyte = Zählbyte + 1
    If Zählbyte > CompSize + 1 Then
    Exit Do
    End If
    NumToRead = BytesToRead(Bitmap)
        If (CompSize - Inbytes) < NumToRead Then
        NumToRead = CompSize - Inbytes + 1
        Else
        NumToRead = NumToRead
        End If
    ReDim Sets(16) 'NumToRead - 1)
    CopyMemory Sets(0), Buffer(Inbytes + 1), NumToRead
    Inbytes = Inbytes + NumToRead + 1
    Index = 0
        For Counter = 0 To 7
            If BitSet(Bitmap, Counter) Then
            If Zählbyte + 2 > CompSize Then
            Exit Do
            End If
            Length = ((Sets(Index + 1) And &HF0) \ 16) + 3
            Distance = (256 * (Sets(Index + 1) And &HF)) + Sets(Index) + 1
            Zählbyte = Zählbyte + 2
            CodePtr = CurrPos - Distance
                Do While Length
                If CurrPos > UBound(Buffertemp) Then
                ReDim Preserve Buffertemp(UBound(Buffertemp) + CompSize)
                End If
                Buffertemp(CurrPos) = Buffertemp(CodePtr)
                CurrPos = CurrPos + 1
                CodePtr = CodePtr + 1
                OutBytes = OutBytes + 1
                Length = Length - 1
                Loop
            Index = Index + 2
            Else
            If Zählbyte + 1 > CompSize Then
            Exit Do
            End If
                If CurrPos > UBound(Buffertemp) Then
                ReDim Preserve Buffertemp(UBound(Buffertemp) + CompSize)
                End If
            Buffertemp(CurrPos) = Sets(Index)
            CurrPos = CurrPos + 1
            Index = Index + 1
            OutBytes = OutBytes + 1
            Zählbyte = Zählbyte + 1
            End If
        Next Counter
    Loop
    ReDim Buffer(OutBytes - 1)
    CopyMemory Buffer(0), Buffertemp(0), OutBytes
    Decompress = OutBytes
End Function
Private Function BitSet(FBitmap As Byte, Bit As Integer) As Integer
Dim Bitmap As Byte
Bitmap = FBitmap
BitSet = 0
Select Case Bit
Case 0
Bitmap = Bitmap And 1
If Bitmap = 1 Then BitSet = 1
Case 1
Bitmap = Bitmap And 2
If Bitmap = 2 Then BitSet = 1
Case 2
Bitmap = Bitmap And 4
If Bitmap = 4 Then BitSet = 1
Case 3
Bitmap = Bitmap And 8
If Bitmap = 8 Then BitSet = 1
Case 4
Bitmap = Bitmap And 16
If Bitmap = 16 Then BitSet = 1
Case 5
Bitmap = Bitmap And 32
If Bitmap = 32 Then BitSet = 1
Case 6
Bitmap = Bitmap And 64
If Bitmap = 64 Then BitSet = 1
Case 7
Bitmap = Bitmap And 128
If Bitmap = 128 Then BitSet = 1
End Select
End Function


Private Function BytesToRead(FBitmap As Byte) As Long
Dim Bitmap As Byte
Dim TempSum As Integer
Dim Counter As Integer
Bitmap = FBitmap
TempSum = 8
For Counter = 0 To 7
TempSum = TempSum + BitSet(Bitmap, Counter)
Next Counter
BytesToRead = TempSum
End Function

Private Function GetBit(Number As Integer) As Integer
Static mask As Long
Static Value As Long
Dim maskvalue As Long
Dim Test As Long
GetBit = 0

If Number Then
mask = ShiftLeft06(mask, 1)
    If mask = 0 Then
    CopyMemory ByVal VarPtr(maskvalue), Maske(wobinich), 4
    Value = maskvalue
    If Value = 1218961 Then
    End If
    mask = 1
    wobinich = wobinich + 4
    End If
Else
mask = 0
End If
Test = Value And mask
If Test = mask Then GetBit = 1
End Function


Public Sub Hall(bytValue() As Byte, Entries As Long, BitCount As Long, PhraseOffsets() As Long)
Dim Offset As Long
Dim Stepgr As Long
Dim Phrindb As Long
Dim l As Long
Dim n As Long

ReDim Maske(UBound(bytValue))
wobinich = 0
CopyMemory Maske(0), bytValue(0), UBound(Maske) + 1
Phrindb = BitCount
ReDim PhraseOffsets(Entries - 1) '8 Stück (Entries)
Stepgr = ShiftLeft06(1, BitCount)
GetBit 0 'initialisieren
PhraseOffsets(0) = Offset '0
For l = 0 To Entries - 1 '8 Stück (Entries)
    For n = 1 To GetBit(1) Step Stepgr '1<<2 (Bits)
    Do While GetBit(1) = 1
    n = n + Stepgr
    Loop
    Next n
    If GetBit(1) Then n = n + 1
    If Phrindb > 1 Then
    If GetBit(1) Then n = n + 2
    End If
    If Phrindb > 2 Then
    If GetBit(1) Then n = n + 4
    End If
    If Phrindb > 3 Then
    If GetBit(1) Then n = n + 8
    End If
    If Phrindb > 4 Then
    If GetBit(1) Then n = n + 16
    End If
    Offset = Offset + n
    PhraseOffsets(l) = Offset
Next l

End Sub

Private Function EvenNumber(ByVal n As Long) As Boolean
  EvenNumber = Not CBool(n And 1&)
End Function

Public Function ShiftLeft06(ByVal lngValue As Long, ByVal ShiftCount As Long) As Long
  Select Case ShiftCount
  Case 0&
    ShiftLeft06 = lngValue
  Case 1&
    If lngValue And &H40000000 Then
      ShiftLeft06 = (lngValue And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If lngValue And &H20000000 Then
      ShiftLeft06 = (lngValue And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If lngValue And &H10000000 Then
      ShiftLeft06 = (lngValue And &HFFFFFFF) * &H8& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If lngValue And &H8000000 Then
      ShiftLeft06 = (lngValue And &H7FFFFFF) * &H10& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If lngValue And &H4000000 Then
      ShiftLeft06 = (lngValue And &H3FFFFFF) * &H20& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If lngValue And &H2000000 Then
      ShiftLeft06 = (lngValue And &H1FFFFFF) * &H40& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If lngValue And &H1000000 Then
      ShiftLeft06 = (lngValue And &HFFFFFF) * &H80& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &HFFFFFF) * &H80&
    End If
  Case 8&
    If lngValue And &H800000 Then
      ShiftLeft06 = (lngValue And &H7FFFFF) * &H100& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H7FFFFF) * &H100&
    End If
  Case 9&
    If lngValue And &H400000 Then
      ShiftLeft06 = (lngValue And &H3FFFFF) * &H200& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3FFFFF) * &H200&
    End If
  Case 10&
    If lngValue And &H200000 Then
      ShiftLeft06 = (lngValue And &H1FFFFF) * &H400& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1FFFFF) * &H400&
    End If
  Case 11&
    If lngValue And &H100000 Then
      ShiftLeft06 = (lngValue And &HFFFFF) * &H800& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &HFFFFF) * &H800&
    End If
  Case 12&
    If lngValue And &H80000 Then
      ShiftLeft06 = (lngValue And &H7FFFF) * &H1000& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H7FFFF) * &H1000&
    End If
  Case 13&
    If lngValue And &H40000 Then
      ShiftLeft06 = (lngValue And &H3FFFF) * &H2000& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3FFFF) * &H2000&
    End If
  Case 14&
    If lngValue And &H20000 Then
      ShiftLeft06 = (lngValue And &H1FFFF) * &H4000& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1FFFF) * &H4000&
    End If
  Case 15&
    If lngValue And &H10000 Then
      ShiftLeft06 = (lngValue And &HFFFF&) * &H8000& Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &HFFFF&) * &H8000&
    End If
  Case 16&
    If lngValue And &H8000& Then
      ShiftLeft06 = (lngValue And &H7FFF&) * &H10000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H7FFF&) * &H10000
    End If
  Case 17&
    If lngValue And &H4000& Then
      ShiftLeft06 = (lngValue And &H3FFF&) * &H20000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3FFF&) * &H20000
    End If
  Case 18&
    If lngValue And &H2000& Then
      ShiftLeft06 = (lngValue And &H1FFF&) * &H40000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1FFF&) * &H40000
    End If
  Case 19&
    If lngValue And &H1000& Then
      ShiftLeft06 = (lngValue And &HFFF&) * &H80000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &HFFF&) * &H80000
    End If
  Case 20&
    If lngValue And &H800& Then
      ShiftLeft06 = (lngValue And &H7FF&) * &H100000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H7FF&) * &H100000
    End If
  Case 21&
    If lngValue And &H400& Then
      ShiftLeft06 = (lngValue And &H3FF&) * &H200000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3FF&) * &H200000
    End If
  Case 22&
    If lngValue And &H200& Then
      ShiftLeft06 = (lngValue And &H1FF&) * &H400000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1FF&) * &H400000
    End If
  Case 23&
    If lngValue And &H100& Then
      ShiftLeft06 = (lngValue And &HFF&) * &H800000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &HFF&) * &H800000
    End If
  Case 24&
    If lngValue And &H80& Then
      ShiftLeft06 = (lngValue And &H7F&) * &H1000000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H7F&) * &H1000000
    End If
  Case 25&
    If lngValue And &H40& Then
      ShiftLeft06 = (lngValue And &H3F&) * &H2000000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3F&) * &H2000000
    End If
  Case 26&
    If lngValue And &H20& Then
      ShiftLeft06 = (lngValue And &H1F&) * &H4000000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1F&) * &H4000000
    End If
  Case 27&
    If lngValue And &H10& Then
      ShiftLeft06 = (lngValue And &HF&) * &H8000000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &HF&) * &H8000000
    End If
  Case 28&
    If lngValue And &H8& Then
      ShiftLeft06 = (lngValue And &H7&) * &H10000000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H7&) * &H10000000
    End If
  Case 29&
    If lngValue And &H4& Then
      ShiftLeft06 = (lngValue And &H3&) * &H20000000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H3&) * &H20000000
    End If
  Case 30&
    If lngValue And &H2& Then
      ShiftLeft06 = (lngValue And &H1&) * &H40000000 Or &H80000000
    Else
      ShiftLeft06 = (lngValue And &H1&) * &H40000000
    End If
  Case 31&
    If lngValue And &H1& Then
      ShiftLeft06 = &H80000000
    Else
      ShiftLeft06 = &H0&
    End If
  End Select
End Function

Public Function Dephrase(Hasphrase As Boolean, Bytearray() As Byte, DataLen2 As Long) As String
On Error GoTo Error1
Dim Endpunkt As Boolean
Dim TestString As String
Dim fertiglänge As Long
Dim i As Long
Dim testlong As Long
Dim Bytetest As Byte
Dim probe1 As Long
Dim Probe As Byte
Dim z As Long
Dim Y() As Byte
Dim str As String
Dim Hilfslong As Long
testlong = 0
fertiglänge = 0
If Hasphrase Then 'nur Phrase
Endpunkt = False
TestString = ""
TestString = ""
fertiglänge = 0
i = 0
testlong = Bytearray(i)
Bytetest = Bytearray(i)
Do While Endpunkt = False
testlong = Bytearray(i)
Bytetest = Bytearray(i)
Select Case Bytetest
Case Is > 15, 0
TestString = TestString & StrConv(Chr(Bytetest), vbFromUnicode)
i = i + 1
fertiglänge = fertiglänge + 1
Case Else
testlong = (((Bytetest * 256) - 256) + Bytearray(i + 1))
i = i + 2
If EvenNumber(testlong) Then
testlong = testlong / 2
TestString = TestString & StrConv(PhraseArray(testlong), vbFromUnicode)
fertiglänge = fertiglänge + Len(PhraseArray(testlong))
Else
testlong = (testlong - 1) / 2
TestString = TestString & StrConv(PhraseArray(testlong), vbFromUnicode)
TestString = TestString & StrConv(" ", vbFromUnicode)
fertiglänge = fertiglänge + Len(PhraseArray(testlong)) + 1
End If
End Select
If fertiglänge >= DataLen2 Then Endpunkt = True
Loop
Else 'Phraseimage + Index
Endpunkt = False
TestString = ""
fertiglänge = 0
i = 0
testlong = Bytearray(i)
Bytetest = Bytearray(i)

Do While Endpunkt = False 'Anfang Phrase
testlong = Bytearray(i)
Bytetest = Bytearray(i)
If EvenNumber(testlong) Then
TestString = TestString & StrConv(PhraseImageArray(testlong \ 2), vbFromUnicode)
fertiglänge = fertiglänge + Len(PhraseImageArray(testlong \ 2))
Else
Probe = Bytetest And 15
If Probe = 15 Then
Probe = Bytetest \ 16 + 1
For z = 1 To Probe
TestString = TestString & StrConv(Chr(0), vbFromUnicode)
Next z
fertiglänge = fertiglänge + Probe
Else
Probe = Bytetest And 7
If Probe = 7 Then
Probe = Bytetest \ 16 + 1
For z = 1 To Probe
TestString = TestString & StrConv(" ", vbFromUnicode)
Next z
fertiglänge = fertiglänge + Probe
Else
Probe = Bytetest And 3
probe1 = Probe
If Probe = 3 Then
Probe = Bytetest \ 8 + 1
If probe1 <> Bytetest Then
End If

fertiglänge = fertiglänge + Probe
ReDim Y(Probe - 1)
CopyMemory Y(0), Bytearray(i + 1), Probe
str = Y

TestString = TestString & str
i = i + Probe
Else
Probe = Bytetest And 1
If Probe = 1 Then
Hilfslong = (Bytetest * 64) + 64 + Bytearray(i + 1)
TestString = TestString & StrConv(PhraseImageArray(Hilfslong), vbFromUnicode)
fertiglänge = fertiglänge + Len(PhraseImageArray(Hilfslong))
i = i + 1
End If
End If
End If
End If
End If
i = i + 1
If fertiglänge >= DataLen2 Then Endpunkt = True
Loop
End If
Dephrase = TestString
Error1: 'verlasse Function
End Function

Public Function DecRunLen(Bytes() As Byte) As Long
Dim Ende As Boolean
Dim n As Long
Dim i As Long
Dim Größe As Long
Dim Test As Long
Dim Byteübergabe() As Byte
Dim Länge As Long
Dim Wiederholbyte As Byte
Dim Standpunkt As Long
Dim Fertigstandpunkt As Long
ReDim Byteübergabe(0)
Ende = False
Größe = UBound(Bytes)
Standpunkt = 0
Fertigstandpunkt = 0
Do While Ende = False
n = Bytes(Standpunkt)
Test = n And &H80
If Test = &H80 Then
Länge = n And &H7F
ReDim Preserve Byteübergabe(Fertigstandpunkt + Länge - 1)
If Länge > 0 Then
CopyMemory Byteübergabe(Fertigstandpunkt), Bytes(Standpunkt + 1), Länge
End If
Fertigstandpunkt = Fertigstandpunkt + Länge
Standpunkt = Standpunkt + Länge + 1
Else
ReDim Preserve Byteübergabe(Fertigstandpunkt + n)
Wiederholbyte = Bytes(Standpunkt + 1)
For i = 0 To n - 1
Byteübergabe(Fertigstandpunkt + i) = Wiederholbyte
Next i
Standpunkt = Standpunkt + 2
Fertigstandpunkt = Fertigstandpunkt + n
End If
If Standpunkt >= Größe Then Ende = True
Loop
ReDim Bytes(Fertigstandpunkt - 1)
CopyMemory Bytes(0), Byteübergabe(0), Fertigstandpunkt
End Function
Public Function ReadCompUnSignShort(ErstesByte As Byte, ZweitesByte As Byte, Länge As Integer) As Integer
Dim testlong As Long
Dim Hilfsint As Integer
Dim Testint As Integer
testlong = ErstesByte
Testint = ErstesByte
If EvenNumber(testlong) Then
ReadCompUnSignShort = Testint / 2
Länge = 1
Else
Hilfsint = ZweitesByte * 128
ReadCompUnSignShort = ((Testint - 1) / 2) + Hilfsint
Länge = 2
End If
End Function

Public Function ReadCompSignShort(ErstesByte As Byte, ZweitesByte As Byte, Länge As Integer) As Integer
Dim testlong As Long
Dim Hilfsint As Integer
Dim Testint As Integer
testlong = ErstesByte
Testint = ErstesByte
If EvenNumber(testlong) Then
ReadCompSignShort = (Testint / 2) - 64
Länge = 1
Else
Hilfsint = ZweitesByte * 128
ReadCompSignShort = ((Testint - 1) / 2) + Hilfsint - 16384
Länge = 2
End If
End Function

Public Function ReadCompUnSignLong(Hilfsint1 As Integer, Hilfsint2 As Integer, Länge As Integer) As Long
Dim testlong As Long
Dim Hilfslong As Long
Dim ErstesInt As Long
Dim ZweitesInt As Long
ErstesInt = IntegerToUnsigned(Hilfsint1)
ZweitesInt = IntegerToUnsigned(Hilfsint2)

testlong = ErstesInt
If EvenNumber(testlong) Then
ReadCompUnSignLong = testlong / 2
Länge = 2
Else
Hilfslong = ZweitesInt * 32768
ReadCompUnSignLong = ((testlong - 1) / 2) + Hilfslong
Länge = 4
End If
End Function

Public Function scanint(Testint() As Byte, Länge As Integer) As Integer
Dim Testbyte As Byte
Dim testlong As Long
Dim Testint1 As Integer
Testbyte = Testint(0)
CopyMemory ByVal VarPtr(Testint1), Testint(0), 2
testlong = Testbyte
If EvenNumber(testlong) = True Then 'gerade
scanint = (Testbyte \ 2) - &H40
Länge = 1
Else 'ungerade
scanint = (Testint1 \ 2) - &H4000
Länge = 2
End If
End Function
Public Function scanword(Testint() As Byte, Länge As Integer) As Integer
Dim Testbyte As Byte
Dim testlong As Long
Dim Testbyte1 As Byte

CopyMemory ByVal VarPtr(Testbyte), Testint(0), 1
CopyMemory ByVal VarPtr(Testbyte1), Testint(1), 1
testlong = Testbyte
If EvenNumber(testlong) = True Then 'gerade
scanword = (Testbyte \ 2)
Länge = 1
Else 'ungerade
scanword = (Testbyte1 \ 2)
Länge = 2
End If
End Function

Public Function scanlong(testlong() As Byte, Länge As Integer) As Long
Dim Testdouble1 As Double
Dim Testlong1 As Long
Dim Testlong2 As Long
Dim Testlong3 As Long
Dim Testdouble As Double
Dim Testinteger As Integer
Testlong1 = testlong(0)
If EvenNumber(Testlong1) = True Then 'gerade
CopyMemory ByVal VarPtr(Testinteger), testlong(0), 2
Testlong2 = IntegerToUnsigned(Testinteger)
scanlong = (Testlong2 \ 2) - &H4000
Länge = 2
Else 'ungerade
CopyMemory ByVal VarPtr(Testlong3), testlong(0), 4
Testdouble = LongToUnsigned(Testlong3)
Testdouble1 = (Testdouble / 2)
Testdouble1 = Testdouble1 - &H40000000
scanlong = CLng(Testdouble1)
Länge = 4
End If
End Function
Public Function IntegerToUnsigned(intValue As Integer) As Long
        If intValue < 0 Then
          IntegerToUnsigned = intValue + OFFSET_2
        Else
          IntegerToUnsigned = intValue
        End If
End Function

     Public Function UnsignedToLong(dblValue As Double) As Long
        If dblValue <= MAXINT_4 Then
          UnsignedToLong = dblValue
        Else
          UnsignedToLong = dblValue - OFFSET_4
        End If
      End Function

  Public Function LongToUnsigned(lngValue As Long) As Double

        If lngValue < 0 Then
          LongToUnsigned = lngValue + OFFSET_4
        Else
          LongToUnsigned = lngValue
        End If
End Function
Public Function MakeCompressedUnsignedLong(Number As Long) As Long
Dim rechnen As Long
Dim Übergabe As Long
Übergabe = Number

rechnen = Übergabe * 2
If EvenNumber(rechnen) Then
rechnen = rechnen + 1
End If
MakeCompressedUnsignedLong = rechnen
End Function


Public Sub LongToRGB(pColor As Long, pRGBentry As RGBColor)
       CopyMemory ByVal VarPtr(pRGBentry), ByVal VarPtr(pColor), 3
End Sub

