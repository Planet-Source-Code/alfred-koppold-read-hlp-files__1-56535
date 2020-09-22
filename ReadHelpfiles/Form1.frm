VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5496
   ClientLeft      =   1080
   ClientTop       =   1812
   ClientWidth     =   7440
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5496
   ScaleWidth      =   7440
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check3 
      Caption         =   "Save Phrase-File"
      Height          =   252
      Left            =   4080
      TabIndex        =   5
      Top             =   1200
      Width           =   3012
   End
   Begin VB.CheckBox Check2 
      Caption         =   "BMPs and WMFs from SHG and MRB"
      Height          =   492
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   3132
   End
   Begin VB.CheckBox Check1 
      Caption         =   "More rtfs"
      Height          =   492
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   2652
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   2280
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      Filter          =   "(*.hlp)|*.hlp"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Datei auswählen"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Height          =   612
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   2052
   End
   Begin VB.Label Label1 
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   6492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Dim sas As Long
Dim Dateiname As String
Dim Standpunkt As Long
Dim Länge As Long
Dim KWBTreeArray() As Byte
Dim Dateistand As Long
Dim n As Long
Dim z As Long
Dim KeyBytes() As Byte
Dim Keyword As String
Dim Antwort As Long
Dim bisherigeAnzahlstrings As Long
Dim Hilfslong As Long
Dim i As Long
Dim anzahlkeywords As Long
Dim al() As Byte

Zurücksetzen
CommonDialog1.ShowOpen
Dateiname = CommonDialog1.Filename
If Dateiname = "" Then
Command1.Enabled = True
Exit Sub
End If
sas = InStrRev(Dateiname, "\")
ProjektName = Mid(CommonDialog1.Filename, sas + 1, Len(CommonDialog1.Filename) - sas - 4)
Grundweg = Mid(CommonDialog1.Filename, 1, sas - 1)
PetraFile(0).RTFName = ProjektName & ".rtf"
Pfad = App.Path & "\" & ProjektName
If FolderExists(Pfad) Then
Antwort = MsgBox("Soll Ordner " & Pfad & " gelöscht werden?", vbQuestion Or vbYesNo, "Question")
If Antwort = vbNo Then
Command1.Enabled = True
Exit Sub
End If
DeleteDirectory Pfad
End If
MkDir Pfad
Dateinummer = FreeFile
Open Dateiname For Binary As Dateinummer
Get Dateinummer, , Header
If Header.Magic <> &H35F3F Then
MsgBox "No Windows Helpfile!"
Close Dateinummer
Command1.Enabled = True
Exit Sub
End If
Get Dateinummer, Header.DirectoryStart + 1, Filehdr
Get Dateinummer, , BTreeHdr
Dateistand = Seek(Dateinummer)

  Select Case BTreeHdr.NLevels
Case 1
    ReDim al(Filehdr.UsedSpace)
    Get Dateinummer, Dateistand + 8, al
    Dim s As String
    Dim te As Long
    s = al
    s = StrConv(s, vbUnicode)
        Dim Ende As Integer
        Dim Anfang As Integer
Anfang = 1
    For n = 0 To BTreeHdr.TotalBtreeEntries - 1
    Ende = InStr(Anfang, s, Chr(0))
    ReDim Preserve DirLeafEntry(n)
    DirLeafEntry(n).Filename = Mid(s, Anfang, Ende - Anfang)
    CopyMemory ByVal VarPtr(te), al(Ende), 4
    DirLeafEntry(n).FileOffset = te
    Anfang = Ende + 5
    Next n

Case Else
ReDim KWBTreeArray(Filehdr.UsedSpace - 1 - 38)
Get Dateinummer, Dateistand, KWBTreeArray
Standpunkt = 0
For i = 1 To BTreeHdr.TotalPages
If i - 1 = BTreeHdr.RootPage Then
Standpunkt = Standpunkt + BTreeHdr.PageSize
Else
CopyMemory ByVal VarPtr(CurrNode), KWBTreeArray(Standpunkt), 8 'Nodehead einlesen
Standpunkt = Standpunkt + 8
anzahlkeywords = anzahlkeywords + CurrNode.NEntries
ReDim Preserve DirLeafEntry(anzahlkeywords)
For z = 0 + bisherigeAnzahlstrings To CurrNode.NEntries + bisherigeAnzahlstrings + -1   'Entries einlesen
Länge = lstrlen(VarPtr(KWBTreeArray(Standpunkt)))
ReDim KeyBytes(Länge - 1)
CopyMemory KeyBytes(0), KWBTreeArray(Standpunkt), Länge
Keyword = StrConv(KeyBytes, vbUnicode)
DirLeafEntry(z).Filename = Keyword
CopyMemory ByVal VarPtr(Hilfslong), KWBTreeArray(Standpunkt + Länge + 1), 4
DirLeafEntry(z).FileOffset = Hilfslong
Standpunkt = Standpunkt + Länge + 4 + 1
Next z
bisherigeAnzahlstrings = bisherigeAnzahlstrings + CurrNode.NEntries
Standpunkt = Standpunkt + CurrNode.Unused '7
End If
Next i
End Select
    
            For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|Petra" Then
        Label1.Caption = "Read Petra File"
        DoEvents
        GetPetraFile DirLeafEntry(n).FileOffset
        End If
        Next n

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|CONTEXT" Then
        Label1.Caption = "Read Context File"
        DoEvents
        GetContextFile (DirLeafEntry(n).FileOffset)
        End If
        Next n
        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|TopicId" Then
        Label1.Caption = "Read TopicID"
        DoEvents
        GetTopicID DirLeafEntry(n).FileOffset
        End If
        Next n

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If Left(DirLeafEntry(n).Filename, 1) <> "|" Then
        Label1.Caption = "Read Baggage File"
        DoEvents
        GetBaggageFile (DirLeafEntry(n).FileOffset), DirLeafEntry(n).Filename
        End If
        Next n

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If Left(DirLeafEntry(n).Filename, 3) = "|bm" Then
        Label1.Caption = "Read Bitmaps"
        DoEvents
        GetbmFile DirLeafEntry(n).FileOffset, DirLeafEntry(n).Filename
        End If
        Next n

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If Left(DirLeafEntry(n).Filename, 3) = "|CF" Then
        Label1.Caption = "Read Macros"
        DoEvents
        GetCFFile DirLeafEntry(n).FileOffset, DirLeafEntry(n).Filename
        End If
        Next n

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|TTLBTREE" Then
        Label1.Caption = "Read Ttlbtree File"
        DoEvents
        GetTtlbtreeFile (DirLeafEntry(n).FileOffset)
        End If
        Next n

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|SYSTEM" Then
        Label1.Caption = "Read Systemfile"
        DoEvents
        GetSystemFile (DirLeafEntry(n).FileOffset)
        End If
        Next n

        SucheIDNames

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|VIOLA" Then
        Label1.Caption = "Read Viola File"
        DoEvents
        GetViolaFile (DirLeafEntry(n).FileOffset)
        End If
        Next n

        

                        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|Phrase" Or DirLeafEntry(n).Filename = "|Phrases" Then
        Label1.Caption = "Decompress Phrases"
        DoEvents
        GetPhrase (DirLeafEntry(n).FileOffset)
        End If
        Next n
                For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|PhrImage" Then
        Label1.Caption = "Decompress Phrases"
        DoEvents
        GetPhrImage (DirLeafEntry(n).FileOffset)
        End If
        Next n

        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|PhrIndex" Then
        Label1.Caption = "Decompress Phrases"
        DoEvents
        GetPhrIndex (DirLeafEntry(n).FileOffset)
        End If
        Next n



                For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|FONT" Then
        Label1.Caption = "Read Font File"
        DoEvents
        GetFontFile (DirLeafEntry(n).FileOffset)
        End If
        Next n
        
        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If Left(DirLeafEntry(n).Filename, 1) = "|" And Mid(DirLeafEntry(n).Filename, 3) = "WMAP" Then
        Label1.Caption = "Read xWMAP File"
        DoEvents
        GetxWMAPFile DirLeafEntry(n).FileOffset, Mid(DirLeafEntry(n).Filename, 2, 1)
        End If
        Next n
        
        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If Left(DirLeafEntry(n).Filename, 1) = "|" And Mid(DirLeafEntry(n).Filename, 3) = "WDATA" Then
        Label1.Caption = "Read xWDATA File"
        DoEvents
        GetxWDataFile (DirLeafEntry(n).FileOffset), Mid(DirLeafEntry(n).Filename, 2, 1)
        End If
        Next n
        
        For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If Left(DirLeafEntry(n).Filename, 1) = "|" And Mid(DirLeafEntry(n).Filename, 3) = "WBTREE" Then
        Label1.Caption = "Read xWBTREE File"
        DoEvents
        GetxWBTREEFile (DirLeafEntry(n).FileOffset), Mid(DirLeafEntry(n).Filename, 2, 1)
        End If
        Next n
        
                For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|CTXOMAP" Then
        Label1.Caption = "Read Ctxomap File"
        DoEvents
        GetCtxoMapFile (DirLeafEntry(n).FileOffset)
        End If
        Next n

        
            For n = 0 To BTreeHdr.TotalBtreeEntries - 1
        If DirLeafEntry(n).Filename = "|TOPIC" Then
        Label1.Caption = "Read Topic File"
        DoEvents
        GetTopicFile (DirLeafEntry(n).FileOffset)
        End If
        Next n
        


Close Dateinummer
MakeAlias

Createhpjfile
Command1.Enabled = True
Label1.Caption = ""
End Sub




Private Sub Zurücksetzen()
AnzTTLBTREE = 0
ReDim TTLBNames(0)
Label1.Caption = ""
MoreRTF = Check1
EnPicture = Check2
SavePhrase = Check3
RTFGröße = 0
PetraAnzahl = 0
ReDim PetraFile(0)
ReDim Bitmaps(0)
HasAKeywords = False
HasKKeywords = False
anzbitmaps = 0
anzwindows = 0
Mapstring = ""
ZuletztAbsatz = False
Macrostring = ""
MacrostringAll = ""
Makrostringfertig = ""
Configstring = ""
Baggagestring = ""
AnzBaggage = 0
ReDim MappingDa(0)
Hasphrase = False
Aufmachen = 0
ReDim fertigeIDs(0)
zumachen = 0
Aliasstring = ""
TopZähler = 0
MitteOffset = 0
Aktuellestopicoffset = 0
anzviolas = 0
ScrollorNoscroll = False
anzwindows = 0
hasviolas = False
anzbitmaps = 0
Command1.Enabled = False
ReDim Namecnt(0)
NumCnt = 0
Contentsnumber = -1
ReDim Namegefunden(0)
ReDim IDNames(0)
ReDim BaggageInFile(0)
ReDim Baggagefiles(0)

End Sub

