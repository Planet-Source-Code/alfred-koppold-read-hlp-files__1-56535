Attribute VB_Name = "Module2"
Option Explicit
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Const FAM_DECOR = &H5
Public Const FAM_MODERN = &H1
Public Const FAM_NIL = &H3
Public Const FAM_ROMAN = &H2
Public Const FAM_SCRIPT = &H4
Public Const FAM_SWISS = &H3
Public Const FAM_TECH = &H3
Public Const FONT_BOLD = &H1
Public Const FONT_DBUN = &H10
Public Const FONT_ITAL = &H2
Public Const FONT_NORM = &H0
Public Const FONT_SMCP = &H20
Public Const FONT_STRK = &H8
Public Const FONT_UNDR = &H4
Public Const TL_DISPLAY = &H20
Public Const TL_DISPLAY30 = &H1
Public Const TL_TABLE = &H23
Public Const TL_TOPICHDR = &H2
Public Const WSYSFLAG_ADJUSTRRESOLUTION = &H1000 '/* Adjust Screen Resolution
Public Const WSYSFLAG_AUTOSIZEHEIGHT = &H800 '/* Auto-Size Height */
Public Const WSYSFLAG_CAPTION = &H4 '/* Caption is valid */
Public Const WSYSFLAG_HEIGHT = &H40 '/* Height is valid */
Public Const WSYSFLAG_MAXIMIZE = &H80 '/* Maximize is valid */
Public Const WSYSFLAG_NAME = &H2 '/* Name is valid */
Public Const WSYSFLAG_RGB = &H100 '/* Rgb is valid */
Public Const WSYSFLAG_RGBNSR = &H200 '/* RgbNsr is valid */
Public Const WSYSFLAG_TOP = &H400 '/* On top was set in HPJ file */
Public Const WSYSFLAG_TYPE = &H1 '/* Type is valid */
Public Const WSYSFLAG_WIDTH = &H20 '/* Width is valid */
Public Const WSYSFLAG_X = &H8 '/* X is valid */
Public Const WSYSFLAG_Y = &H10 '/* Y is valid */

Public Type HELPHEADER '/* structure at beginning of help file */
    Magic As Long '              /* 0x00035F3F */
    DirectoryStart As Long '  /* offset of FILEHEADER of internal direcory */
    FreeChainStart As Long '     /* offset of FILEHEADER or -1L */
    EntireFileSize As Long '     /* size of entire help file in bytes */
End Type

Public Type FILEHEADER    '/* structure at FileOffset of each internal file */
    ReservedSpace As Long    '  /* reserved space in help file incl. FILEHEADER */
    UsedSpace As Long     '     /* used space in help file excl. FILEHEADER */
    FileFlags As Byte ' /* normally 4 */
End Type

Public Type BTREEHEADER '  /* structure after FILEHEADER of each Btree */
    Magic As Integer  '  /* 0x293B */
    Flags As Integer '    /* bit 0x0002 always 1, bit 0x0400 1 if direcory */
    PageSize As Integer ' /* 0x0400=1k if directory, 0x0800=2k else */
    Structure(15) As Byte ' /* string describing structure of data */
    MustBeZero As Integer '       /* 0 */
    PageSplits As Integer '        /* number of page splits Btree has suffered */
    RootPage As Integer '          /* page number of Btree root page */
    MustBeNegOne As Integer '     /* 0xFFFF */
    TotalPages As Integer '       /* number of Btree pages */
    NLevels As Integer '           /* number of levels of Btree */
    TotalBtreeEntries As Long ' /* number of entries in Btree */
End Type

Public Type BTREEINDEXHEADER '/* structure at beginning of every index-page */
    Unused As Integer '  /* unused Bytes on the End of this page */
    NEntries As Integer '          /* number of entries in this index-page */
    PreviousPage As Integer '      /* page number of previous page */
End Type

Public Type BTREENODEHEADER '/* structure at beginning of every leaf-page */
    Unused As Integer '  /* unused bytes at End of this leaf-page */
    NEntries As Integer '          /* number of entires in this leaf-page */
    PreviousPage As Integer '      /* page number of preceeding leaf-page or -1 */
    NextPage As Integer '          /* page number of next leaf-page or -1 */
End Type

Public Type SYSTEMHEADER  '/* structure at beginning of |SYSTEM file */
    Magic As Integer '    /* 0x036C */
    Minor As Integer '    /* help file format version number */
    Major As Integer '    /* 1 */
    GenDate As Long       '   /* date/time the help file was generated or 0 */
    Flags As Integer '    /* tells you how the help file is compressed */
End Type

Public Type SYSTEMRECORD ' /* internal structure */
    'FILE *File;
    'SavePos As Long
    'Remaining As Long
    RecordType As Integer ' /* type of data in record */
    DataSize As Integer '   /* size of data */
    'Data(9) As Byte
End Type

Public Type SECWINDOW     '/* structure of data following RecordType 6 */
    Flags As Integer '    /* flags (see below) */
    Types(9) As Byte ' /* type of window */
    Names(8) As Byte ' /* window name */
    Caption(50) As Byte ' /* caption for window */
    x As Integer ' /* x coordinate of window (0..1000) */
    Y As Integer '                /* y coordinate of window (0..1000) */
    Width As Integer '/* width of window (0..1000) */
    Height As Integer '/* height of window (0..1000) */
    Maximize As Integer '          /* maximize flag and window styles */
    Rgb(2) As Byte '/* color of scrollable region */
    Unknown1 As Byte
    RgbNsr(2) As Byte '/* color of non-scrollable region */
    Unknown2 As Byte
End Type

Public Type MVBWINDOW
    Flags As Integer '    /* flags (see below) */
    Types(9) As Byte ' /* type of window */
    Names(8) As Byte ' /* window name */
    Caption(50) As Byte ' /* caption for window */
    MoreFlags As Byte
    x As Integer '                 /* x coordinate of window (0..1000) */
    Y As Integer '                 /* y coordinate of window (0..1000) */
    Width As Integer '            /* width of window (0..1000) */
    Height As Integer '            /* height of window (0..1000) */
    Maximize As Integer '          /* maximize flag and window styles */
    TopRgb(2) As Byte '
    Unknown0 As Byte
    Unknown(243) As Byte
    Rgb(2) As Byte ' /* color of scrollable region */
    Unknown1 As Byte
    RgbNsr(2) As Byte ' /* color of non-scrollable region */
    Unknown2 As Byte
    X2 As Integer
    Y2 As Integer
    Width2 As Integer
    Height2 As Integer
    X3 As Integer
    Y3 As Integer
    End Type

Public Type KEYINDEX '        /* structure of data following RecordType 14 */
    btreename(9) As Byte
    mapname(9) As Byte
    dataname(9) As Byte
    title(79) As Byte
End Type


Public Type PHRINDEXHDR '  /* structure of beginning of |PhrIndex file */
    Magic As Long '              /* sometimes 0x0001 */
    NEntries As Long '                 /* number of phrases */
    CompressedSize As Long '          /* size of PhrIndex file */
    PhrImageSize As Long '            /* size of decompressed PhrImage file */
    PhrImagecompressedsize As Long '  /* size of PhrImage file */
    Always0 As Long '
    BitCount As Byte '4 Bits BitCount, 4 Bits unknown
    unknown_12 As Byte ' 8 Bits unknown
    always4A00 As Integer '    /* sometimes 0x4A01, 0x4A02 */
End Type

Public Type FONTHEADER   ' /* structure of beginning of |FONT file */
    NumFacenames As Integer '       /* number of face names */
    NumDescriptors As Integer '     /* number of font descriptors */
    FacenamesOffset As Integer '    /* offset of face name array */
    DescriptorsOffset As Integer '  /* offset of descriptors array */
    NumFormats As Integer '         /* only if FacenamesOffset >= 12 */
    FormatsOffset As Integer '      /* offset of formats array */
    NumCharmaps As Integer '        /* only if FacenamesOffset >= 16 */
    CharmapsOffset As Integer '     /* offset of charmapnames array */
End Type

Public Type FONTDESCRIPTOR '/* internal font descriptor */
    Bold As Byte
    Italic As Byte
    Underline As Byte
    StrikeOut As Byte
    DoubleUnderline As Byte
    SmallCaps As Byte
    HalfPoints As Byte
    FontFamily As Byte 'see below
    FontName As Integer
    textcolor As Byte
    backcolor As Byte
    style As Integer
    expndtw As Integer
    up As Byte
End Type

Public Type OLDFONT  '     /* non-Multimedia font descriptor
    Attributes As Byte ' /* Font Attributes See values FONT_..
    HalfPoints As Byte ' /* PointSize * 2 */
    FontFamily As Byte ' /* Font Family. See values below */
    FontName As Integer '  /* Number of font in Font List */
    FGRGB(2) As Byte '/* RGB values of foreground */
    BGRGB(2) As Byte '/* unused background RGB Values */
End Type
Public Type NEWFONT  '      /* structure located at DescriptorsOffset */
    Unknown1 As Byte
    FontName As Integer
    FGRGB(2) As Byte
    BGRGB(2) As Byte
    unknown5 As Byte
    unknown6 As Byte
    unknown7 As Byte
    unknown8 As Byte
    unknown9 As Byte
    Height As Long
    mostlyzero(11) As Byte
    Weight As Integer
    unknown10 As Byte
    unknown11 As Byte
    Italic As Byte
    Underline As Byte
    StrikeOut As Byte
    DoubleUnderline As Byte
    SmallCaps As Byte
    unknown17 As Byte
    unknown18 As Byte
    PitchAndFamily As Byte
End Type

Public Type NEWSTYLE
    StyleNum As Integer
    BasedOn As Integer
    font As NEWFONT
    Unknown(34) As Byte
    StyleName(64) As Byte
    End Type
    
Public Type DIRECTORYLEAFENTRY
Filename As String
FileOffset As Long
End Type

Public Type FONTINTERNAL
NumFacenames As Integer
NumDescriptors As Integer
FacenamesOffset As Integer
DescriptorsOffset As Integer
NumStyles As Integer
StyleOffset As Integer
NumCharMapTables As Integer
CharMapTableOffset As Integer
End Type

Public Type TOPICBLOCKHEADER
LastTopicLink As Long 'points to last topic link in previous block or -1L
FirstTopicLink As Long 'points to first topic link in this block
LastTopicHeader As Long 'points to topic link of last topic header or 0L, -1L
End Type
Public Type TOPICLINK
    BlockSize As Long 'size of this link + LinkData1 + LinkData2
    DataLen2 As Long 'length of decompressed LinkData2
    PrevBlock As Long 'Windows 3.1 (HC31): TOPICPOS of previous TOPICLINK */
    NextBlock As Long 'Windows 3.1 (HC31): TOPICPOS of next TOPICLINK
    DataLen1 As Long 'includes size of TOPICLINK
    RecordType As Byte 'See below
End Type
Public Type TOPICHEADER
BlockSize As Long 'size of topic, including internal topic links
BrowseBck  As Long 'topic offset for prev topic in browse sequence
BrowseFor As Long 'topic offset for next topic in browse sequence
Topicnum  As Long 'topic number
NonScroll As Long 'start of non-scrolling region (topic offset) or -1L
Scroll As Long 'start of scrolling region (topic offset)
NextTopic As Long 'start of next type 2 record
End Type
Public Type CTXOMAPENTRY
MapID As Long
Topicoffset As Long
End Type
Public Type KWMAPREC      'structure of |xWMAP leaf-page entries */
    FirstRec As Long           'index number of first keyword on leaf page */
    PageNum As Integer   'page number that keywords are associated with */
End Type
Public Type FONTS
    Attributes As Byte ' /* Font Attributes See values FONT_..
    Fontsize As Byte
    FontFamily As String ' /* Font Family. See values below */
    FontName As String
    Fontnamefertig As String
    Fontfarbe(2) As Byte '/* RGB values of foreground */
    ColorArraynumber As Long
End Type
Public Type TTLINDEX
Topicoffset As Long
PageNumber As Integer
End Type
Public Type CONTEXTINDEX
Hashvalue As Long
PageNumber As Integer
End Type
Public Type CONTEXTLEAF
Hashvalue As Long
Topicoffset As Long
End Type
Public Type PHRASEHEADER
NumPhrases As Integer
OneHundred As Integer
DecompressSize As Long
End Type
Public Type PICTUREFILEHEADER
Magic As Integer
NumberofPictures As Integer
End Type
Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type
Type METAPLACEABLEHEADER
  mtKey             As Long         'must be 0x9AC6CDD7
  mtHandle          As Integer      'must be 0
  mtLeft            As Integer
  mtTop             As Integer
  mtRight           As Integer
  mtBottom          As Integer
  mtInch            As Integer
  mtReserved        As Long         'must be 0
  mtCheckSum        As Integer
End Type
Public Type COLUMNSTRUCT
Columnn As Integer
Unknown As Integer
Always0 As Integer
End Type
Public Type BORDERINFO
Borderparameters As Byte
BorderWidth As Integer
End Type
Public Type Topicbeschreibung
Topicoffset As Long
TopicNummer As Long
Nummer As Long
End Type
Public Type RGBTRIPLE
        rgbtBlue As Byte
        rgbtGreen As Byte
        rgbtRed As Byte
End Type
Public Colortables() As RGBTRIPLE
Public Type BrowseDescr
Topicnum As Long
ThisOffset As Long
Reserveoffset As Long
BrowseBackOffset As Long
BrowseForOffset As Long
FileStandpunkt As Long
Filename As String
End Type
Public Type BROWSEFERTIG
Topicoffset As Long
TopicNumber As Long
Number As Long
End Type

Public Type SHG_HEAD
CompSize As Long
Hotspotsize As Long
CompOffset As Long
HSPOffset As Long
End Type
Public Type HOTSPOT_TYPE
id0 As Byte
id1 As Byte
id2 As Byte
x As Integer
Y As Integer
w As Integer
h As Integer
hash As Long
End Type
Public Type SHGMRB_HEAD
Magic As Integer
NumberofPictures As Integer
End Type
Public Type SHGMRB_HEAD2
PictureType As Byte
PackingMethod As Byte
End Type
Public Function FolderExists(ByVal Path$) As Boolean
Dim i As Integer
On Error Resume Next
i = GetAttr(Path)
If i And vbDirectory Then FolderExists = True
On Error GoTo 0
End Function

