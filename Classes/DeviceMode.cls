VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CDS_UPDATEREGISTRY     As Long = &H1
Private Const CDS_TEST               As Long = &H2
Private Const CDS_FULLSCREEN         As Long = &H4&
Private Const CDS_GLOBAL             As Long = &H8&
Private Const CDS_SET_PRIMARY        As Long = &H10&
Private Const CDS_VIDEOPARAMETERS    As Long = &H20&
Private Const CDS_NORESET            As Long = &H10000000
Private Const CDS_RESET              As Long = &H40000000
Private Const CDS_FORCE              As Long = &H80000000

Private Const CCHDEVICENAME          As Long = 32&
Private Const CCHFORMNAME            As Long = 32&

Private Const DISP_CHANGE_FAILED     As Long = -1&
Private Const DISP_CHANGE_SUCCESSFUL As Long = 0&
Private Const DISP_CHANGE_RESTART    As Long = 1&

Private Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1&
Private Const DISPLAY_DEVICE_PRIMARY_DEVICE      As Long = &H4&
Private Const DISPLAY_PRIMARY_DEVICE             As Long = &H4& 'Primary device

Private Const DISPLAY_DEVICE_MIRRORING_DRIVER    As Long = &H8&
Private Const DISPLAY_DEVICE_VGA_COMPATIBLE      As Long = &H10&
Private Const DISPLAY_DEVICE_REMOVABLE           As Long = &H20&
Private Const DISPLAY_DEVICE_MODESPRUNED         As Long = &H8000000

'Flags, die angeben, welche Werte enthalten sind, oder geändert werden sollen
Private Const DM_ORIENTATION         As Long = &H1&
'Private Const DM_GRAYSCALE           As Long = &H1&
'Private Const DM_OUT_DEFAULT         As Long = &H1&
'Private Const DM_UPDATE              As Long = &H1&
'Private Const DM_OUT_BUFFER          As Long = &H2&
'Private Const DM_COPY                As Long = &H2&
'Private Const DM_INTERLACED          As Long = &H2&
Private Const DM_PAPERSIZE           As Long = &H2&
'Private Const DM_IN_PROMPT           As Long = &H4&
Private Const DM_PAPERLENGTH         As Long = &H4&
'Private Const DM_PROMPT              As Long = &H4&
'Private Const DM_IN_BUFFER           As Long = &H8&
'Private Const DM_MODIFY              As Long = &H8&
Private Const DM_PAPERWIDTH          As Long = &H8&
Private Const DM_SCALE               As Long = &H10&
Private Const DM_POSITION            As Long = &H20&
Private Const DM_NUP                 As Long = &H40&
Private Const DM_DISPLAYORIENTATION  As Long = &H80&    ' XP only
Private Const DM_COPIES              As Long = &H100&
Private Const DM_DEFAULTSOURCE       As Long = &H200&
Private Const DM_SPECVERSION         As Long = &H320&
Private Const DM_GETDEFID            As Long = &H400&
Private Const DM_PRINTQUALITY        As Long = &H400&
Private Const DM_SETDEFID            As Long = &H401&
Private Const DM_REPOSITION          As Long = &H402&
Private Const DM_COLOR               As Long = &H800&
Private Const DM_DUPLEX              As Long = &H1000&
Private Const DM_YRESOLUTION         As Long = &H2000&
Private Const DM_TTOPTION            As Long = &H4000&
Private Const DM_COLLATE             As Long = &H8000&
Private Const DM_FORMNAME            As Long = &H10000
Private Const DM_LOGPIXELS           As Long = &H20000
Private Const DM_BITSPERPEL          As Long = &H40000  ' Farbtiefe
Private Const DM_PELSWIDTH           As Long = &H80000  ' Breite
Private Const DM_PELSHEIGHT          As Long = &H100000 ' Höhe
Private Const DM_DISPLAYFLAGS        As Long = &H200000
Private Const DM_DISPLAYFREQUENCY    As Long = &H400000 ' Bild-Wiederholfrequenz
Private Const DM_RESERVED1           As Long = &H800000
Private Const DM_RESERVED2           As Long = &H1000000
Private Const DM_ICMMETHOD           As Long = &H2000000
Private Const DM_ICMINTENT           As Long = &H4000000
Private Const DM_MEDIATYPE           As Long = &H8000000
'Private Const DM_PANNINGWIDTH        As Long = &H8000000
Private Const DM_DITHERTYPE          As Long = &H10000000
Private Const DM_PANNINGHEIGHT       As Long = &H10000000
Private Const DM_DISPLAYFIXEDOUTPUT  As Long = &H20000000 ' XP only


'for printers:
Private Const DMORIENT_PORTRAIT  As Long = 0
Private Const DMORIENT_LANDSCAPE As Long = 0
Private Const DMRES_HIGH         As Long = 0
Private Const DMRES_MEDIUM       As Long = 0
Private Const DMRES_LOW          As Long = 0
Private Const DMRES_DRAFT        As Long = 0

'for displays Display-Orientation
Private Const DMDO_DEFAULT           As Long = 0 ' The current mode's display device orientation is the natural orientation of the device, and should be used as the default.
Private Const DMDO_90                As Long = 1 ' The display device orientation is 90 degrees (measured clockwise) from that of DMDO_DEFAULT.
Private Const DMDO_180               As Long = 2 ' The display device orientation is 180 degrees (measured clockwise) from that of DMDO_DEFAULT.
Private Const DMDO_270               As Long = 3 ' The display device orientation is 270 degrees (measured clockwise) from that of DMDO_DEFAULT.


Private Const EDS_RAWMODE            As Long = 2& ' Wenn diese Einstellung festgelegt ist, gibt die Funktion unabhängig von den Überwachungsfunktionen alle Grafikmodi zurück, die vom Adaptertreiber gemeldet werden. Andernfalls werden nur Modi zurückgegeben, die mit aktuellen Monitoren kompatibel sind.
Private Const EDS_ROTATEDMODE        As Long = 0& ' Wenn festgelegt, gibt die Funktion Grafikmodi in allen Ausrichtungen zurück. Andernfalls werden nur Modi zurückgegeben, die die gleiche Ausrichtung haben wie die derzeit für die angeforderte Anzeige festgelegte.

Private Const ENUM_CURRENT_SETTINGS  As Long = -1
Private Const ENUM_REGISTRY_SETTINGS As Long = -2

Private Const MONITORINFOF_PRIMARY   As Long = &H1&

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-devmodew

'typedef struct _devicemodeW {
'  WCHAR dmDeviceName[CCHDEVICENAME];
'  WORD  dmSpecVersion;
'  WORD  dmDriverVersion;
'  WORD  dmSize;
'  WORD  dmDriverExtra;
'  DWORD dmFields;
'  union {
'    struct {
'      short dmOrientation;
'      short dmPaperSize;
'      short dmPaperLength;
'      short dmPaperWidth;
'      short dmScale;
'      short dmCopies;
'      short dmDefaultSource;
'      short dmPrintQuality;
'    } DUMMYSTRUCTNAME;
'    POINTL dmPosition;
'    struct {
'      POINTL dmPosition;
'      DWORD  dmDisplayOrientation;
'      DWORD  dmDisplayFixedOutput;
'    } DUMMYSTRUCTNAME2;
'  } DUMMYUNIONNAME;
'  short dmColor;
'  short dmDuplex;
'  short dmYResolution;
'  short dmTTOption;
'  short dmCollate;
'  WCHAR dmFormName[CCHFORMNAME];
'  WORD  dmLogPixels;
'  DWORD dmBitsPerPel;
'  DWORD dmPelsWidth;
'  DWORD dmPelsHeight;
'  union {
'    DWORD dmDisplayFlags;
'    DWORD dmNup;
'  } DUMMYUNIONNAME2;
'  DWORD dmDisplayFrequency;
'  DWORD dmICMMethod;
'  DWORD dmICMIntent;
'  DWORD dmMediaType;
'  DWORD dmDitherType;
'  DWORD dmReserved1;
'  DWORD dmReserved2;
'  DWORD dmPanningWidth;
'  DWORD dmPanningHeight;
'} DEVMODEW, *PDEVMODEW, *NPDEVMODEW, *LPDEVMODEW;

Private Type POINT
    X As Long ' 4
    Y As Long ' 4
End Type ' Sum: 8

Private Type PAPER
    dmOrientation   As Integer ' 2
    dmPaperSize     As Integer ' 2
    dmPaperLength   As Integer ' 2
    dmPaperWidth    As Integer ' 2
    dmScale         As Integer ' 2
    dmCopies        As Integer ' 2
    dmDefaultSource As Integer ' 2
    dmPrintQuality  As Integer ' 2
End Type                 ' Sum: 16

'Die EnumDisplaySettings-Funktion legt Werte für die folgenden fünf DEVMODE-Member fest:
'dmBitsPerPel
'dmPelsWidth
'dmPelsHeight
'dmDisplayFlags
'dmDisplayFrequency

Private Type DEVICEMODEW
    dmDeviceName(1 To CCHDEVICENAME * 2) As Byte '64
    'dmDeviceName As String * CCHDEVICENAME '(1 To CCHDEVICENAME * 2) As Byte '64
    dmSpecVersion        As Integer '  2
    dmDriverVersion      As Integer '  2
    dmSize               As Integer '  2
    dmDriverExtra        As Integer '  2
    dmFields             As Long    ' 4
    'Union
    '{
        'struct
        '{
        '    dmOrientation   As Integer
        '    dmPaperSize     As Integer
        '    dmPaperLength   As Integer
        '    dmPaperWidth    As Integer
        '    dmScale         As Integer
        '    dmCopies        As Integer
        '    dmDefaultSource As Integer
        '    dmPrintQuality  As Integer
        '} DUMMYSTRUCTNAME;
        'dmPosition As POINT
        'struct
        '{
    dmPosition           As POINT   '  8
    dmDisplayOrientation As Long    ' 4
    dmDisplayFixedOutput As Long    ' 4
        '} DUMMYSTRUCTNAME2;
    '} DUMMYUNIONNAME;
    
    dmColor              As Integer '  2
    dmDuplex             As Integer '  2
    dmYResolution        As Integer '  2
    dmTTOption           As Integer '  2
    dmCollate            As Integer    ' 4
    dmFormName(1 To CCHFORMNAME * 2) As Byte  ' 64
    'dmFormName As String * CCHFORMNAME
    dmLogPixels          As Integer '  2
    dmBitsPerPel         As Long 'Integer '  2 <<---- ' 8, 16, 24, 32
    'dmUnusedPadding      As Integer '  2
    dmPelsWidth          As Long    ' 4 <<---- '800, 1024, 1280, 1600, . . .
    dmPelsHeight         As Long    ' 4 <<---- '600,  768, 1024,  900, . . .
    'union {
    dmDisplayFlags       As Long    ' 4 <<----
    '    DWORD dmNup;
    '} DUMMYUNIONNAME2;
    dmDisplayFrequency   As Long    ' 4 <<---- ' 60, 100, 120, 143, 166 . . .
    dmICMMethod          As Long
    dmICMIntent          As Long
    dmMediaType          As Long
    dmDitherType         As Long
    dmReserved1          As Long
    dmReserved2          As Long
    dmPanningWidth       As Long
    dmPanningHeight      As Long
End Type '} DEVMODEW, *PDEVMODEW, *NPDEVMODEW, *LPDEVMODEW;

Private m_DevMode As DEVICEMODEW
Private m_bPapGet As Boolean
Private m_bPapSet As Boolean
Private m_Paper   As PAPER
Private m_dmExtra(0 To 512) As Byte

Private Sub Class_Initialize()
    'Debug.Print LenB(m_DevMode)
    m_DevMode.dmSize = LenB(m_DevMode) '188
    'm_DevMode.dmDriverExtra = 512 '8192
End Sub
'
'Friend Sub New_(ByVal DeviceName As String)
'    'm_DevMode.dmDeviceName = DeviceName
'    Dim maxbytes As Long: maxbytes = 64
'    Dim l As Long: l = LenB(DeviceName)
'    maxbytes = IIf(l <= 64, l, maxbytes)
'    RtlMoveMemory m_DevMode.dmDeviceName(1), ByVal StrPtr(DeviceName), maxbytes
'    'MsgBox LenB(m_DevMode.dmDeviceName)
'End Sub

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_DevMode)
End Property

Public Property Get DeviceName() As String
    DeviceName = Trim0(m_DevMode.dmDeviceName)
End Property

Public Property Get VersionSpec() As Integer
     VersionSpec = m_DevMode.dmSpecVersion
End Property

Public Property Get VersionDriver() As Integer
     VersionDriver = m_DevMode.dmDriverVersion
End Property

Public Property Get Size() As Integer
     Size = m_DevMode.dmSize
End Property

Public Property Get DriverExtraMem() As Integer
     DriverExtraMem = m_DevMode.dmSize
End Property

Public Property Get Fields() As Long
    Fields = m_DevMode.dmFields
End Property

Private Sub GetPaper()
    RtlMoveMemory m_Paper, m_DevMode.dmPosition, LenB(m_Paper)
    m_bPapGet = True
End Sub

Private Sub SetPaper()
    RtlMoveMemory m_DevMode.dmPosition, m_Paper, LenB(m_Paper)
    m_bPapSet = True
End Sub

Public Property Get PaperOrientation() As Integer
    If Not m_bPapGet Then GetPaper
    PaperOrientation = m_Paper.dmOrientation
End Property
Friend Function PaperOrientation_ToStr() As String
    If Not m_bPapGet Then GetPaper
    'PaperOrientation = m_Paper.dmOrientation
End Function
Public Property Get PaperSize() As Integer
    If Not m_bPapGet Then GetPaper
    PaperSize = m_Paper.dmPaperSize
End Property
Public Property Get PaperLength() As Integer ' 2
    If Not m_bPapGet Then GetPaper
    PaperLength = m_Paper.dmPaperLength
End Property
Public Property Get PaperWidth() As Integer  ' 2
    If Not m_bPapGet Then GetPaper
    PaperWidth = m_Paper.dmPaperWidth
End Property
Public Property Get PaperScale() As Integer       ' 2
    If Not m_bPapGet Then GetPaper
    PaperScale = m_Paper.dmScale
End Property
Public Property Get PaperCopies() As Integer      ' 2
    If Not m_bPapGet Then GetPaper
    PaperCopies = m_Paper.dmCopies
End Property
Public Property Get PrinterDefaultSource() As Integer ' 2
    If Not m_bPapGet Then GetPaper
    PrinterDefaultSource = m_Paper.dmDefaultSource
End Property
Public Property Get PrinterPrintQuality() As Integer ' 2
    If Not m_bPapGet Then GetPaper
    PrinterPrintQuality = m_Paper.dmPrintQuality
End Property


Public Property Get PositionX() As Long
    PositionX = m_DevMode.dmPosition.X
End Property

Public Property Get PositionY() As Long
    PositionY = m_DevMode.dmPosition.Y
End Property

Public Property Get Position_ToStr() As String
    Position_ToStr = Me.PositionX & "; " & Me.PositionY
End Property

Public Property Get DisplayOrientation() As Long
    DisplayOrientation = m_DevMode.dmDisplayOrientation
End Property

Public Function DisplayOrientation_ToStr() As String
    Dim s As String, o As Long: o = Me.DisplayOrientation
    Select Case o
    Case DMDO_DEFAULT: s = "Default 0°"
    Case DMDO_90:      s = "90°"
    Case DMDO_180:     s = "180°"
    Case DMDO_270:     s = "270°"
    End Select
    DisplayOrientation_ToStr = s
End Function

Public Property Get DisplayFixedOutput() As Long
    DisplayFixedOutput = m_DevMode.dmDisplayFixedOutput
End Property

Public Property Get Color() As Integer
    Color = m_DevMode.dmColor
End Property

Public Property Get Duplex() As Integer
    Duplex = m_DevMode.dmDuplex
End Property

Public Property Get ResolutionY() As Integer
    ResolutionY = m_DevMode.dmYResolution
End Property

Public Property Get TTOption() As Integer
    TTOption = m_DevMode.dmTTOption
End Property

Public Property Get Collate() As Integer
    Collate = m_DevMode.dmCollate
End Property

Public Property Get FormName() As String
    FormName = Trim0(m_DevMode.dmFormName)
End Property

Public Property Get LogPixels() As Integer
    LogPixels = m_DevMode.dmLogPixels
End Property

'Display settings
Public Property Get BitsPerPixel() As Long
    BitsPerPixel = m_DevMode.dmBitsPerPel
End Property

Public Property Get PixelsWidth() As Long
    PixelsWidth = m_DevMode.dmPelsWidth
End Property

Public Property Get PixelsHeight() As Long
    PixelsHeight = m_DevMode.dmPelsHeight
End Property

Public Property Get DisplayFlags() As Long
    DisplayFlags = m_DevMode.dmDisplayFlags
End Property

Public Property Get DisplayFrequency() As Long
    DisplayFrequency = m_DevMode.dmDisplayFrequency
End Property

'Public Property Get Orientation() As Long
'    Orientation = m_DevMode.dmOrientation
'End Property

Friend Function ShortToStr() As String
    ShortToStr = Me.PixelsWidth & " x " & Me.PixelsHeight & " " & Me.BitsPerPixel & "Bit @" & Me.DisplayFrequency & "Hz"
End Function

Friend Function ToStr() As String
    Dim f As Long: f = Me.Fields
    Dim s As String: s = "DeviceMode {" & vbCrLf
    s = s & "    DeviceName         = " & Me.DeviceName & vbCrLf
    s = s & "    Fields             = " & Me.Fields & vbCrLf
    s = s & "    Size               = " & Me.Size & vbCrLf
    s = s & "    VersionDriver      = " & Me.VersionDriver & vbCrLf
    If (f And DM_ORIENTATION) = DM_ORIENTATION Then s = s & "    Orientation        = " & Me.PaperOrientation_ToStr & vbCrLf
    If (f And DM_PAPERSIZE) = DM_PAPERSIZE Then s = s & "    Papersize          = " & Me.PaperSize & vbCrLf
    If (f And DM_PAPERLENGTH) = DM_PAPERLENGTH Then s = s & "    Paperlength        = " & Me.PaperLength & vbCrLf
    If (f And DM_PAPERWIDTH) = DM_PAPERWIDTH Then s = s & "    PaperWidth         = " & Me.PaperWidth & vbCrLf
    If (f And DM_SCALE) = DM_SCALE Then s = s & "    Scale              = " & Me.PaperScale & vbCrLf
    If (f And DM_POSITION) = DM_POSITION Then s = s & "    Position           = " & Me.Position_ToStr & vbCrLf
    'If (f And DM_NUP) = DM_NUP Then                     s = s & "    Nup                = " & Me.Nup & vbCrLf
    If (f And DM_DISPLAYORIENTATION) = DM_DISPLAYORIENTATION Then s = s & "    DisplayOrientation = " & Me.DisplayOrientation_ToStr & vbCrLf
    If (f And DM_COPIES) = DM_COPIES Then s = s & "    Copies             = " & Me.PaperCopies & vbCrLf
    If (f And DM_DEFAULTSOURCE) = DM_DEFAULTSOURCE Then s = s & "    DefaultSource      = " & Me.PrinterDefaultSource & vbCrLf
    If (f And DM_SPECVERSION) = DM_SPECVERSION Then s = s & "    SpecVersion        = " & Me.VersionSpec & vbCrLf
    'If (f And DM_GETDEFID) = DM_GETDEFID Then           s = s & "    GetDefID           = " & Me.GetDefID & vbCrLf
    If (f And DM_PRINTQUALITY) = DM_PRINTQUALITY Then s = s & "    PrintQuality       = " & Me.PrinterPrintQuality & vbCrLf
    'If (f And DM_SETDEFID) = DM_SETDEFID Then           s = s & "    SetDefID           = " & Me.SetDefID & vbCrLf
    'If (f And DM_REPOSITION) = DM_REPOSITION Then       s = s & "    Reposition         = " & Me.Reposition & vbCrLf
    If (f And DM_COLOR) = DM_COLOR Then s = s & "    Color              = " & Me.Color & vbCrLf
    If (f And DM_DUPLEX) = DM_DUPLEX Then s = s & "    Duplex             = " & Me.Duplex & vbCrLf
    If (f And DM_YRESOLUTION) = DM_YRESOLUTION Then s = s & "    Resolution-Y       = " & Me.ResolutionY & vbCrLf
    If (f And DM_TTOPTION) = DM_TTOPTION Then s = s & "    TTOption           = " & Me.TTOption & vbCrLf
    If (f And DM_COLLATE) = DM_COLLATE Then s = s & "    Collate            = " & Me.Collate & vbCrLf
    If (f And DM_FORMNAME) = DM_FORMNAME Then s = s & "    FormName           = " & Me.FormName & vbCrLf
    If (f And DM_LOGPIXELS) = DM_LOGPIXELS Then s = s & "    LogPixels          = " & Me.LogPixels & vbCrLf
    If (f And DM_BITSPERPEL) = DM_BITSPERPEL Then s = s & "    BitsPerPixel       = " & Me.BitsPerPixel & vbCrLf
    If (f And DM_PELSWIDTH) = DM_PELSWIDTH Then s = s & "    PixelsWidth        = " & Me.PixelsWidth & vbCrLf
    If (f And DM_PELSHEIGHT) = DM_PELSHEIGHT Then s = s & "    PixelsHeight       = " & Me.PixelsHeight & vbCrLf
    If (f And DM_DISPLAYFLAGS) = DM_DISPLAYFLAGS Then s = s & "    DisplayFlags       = " & Me.DisplayFlags & vbCrLf
    If (f And DM_DISPLAYFREQUENCY) = DM_DISPLAYFREQUENCY Then s = s & "    DisplayFrequency   = " & Me.DisplayFrequency & vbCrLf
    If (f And DM_DISPLAYFIXEDOUTPUT) = DM_DISPLAYFIXEDOUTPUT Then s = s & "    DisplayFixedOutput (XPOnly)" & vbCrLf
    'If (f And DM_RESERVED1)           = &H800000
    'If (f And DM_RESERVED2)           = &H1000000
    'If (f And DM_ICMMETHOD) = DM_ICMMETHOD Then         s = s & "    IcmMethod          = " & Me.IcmMethod & vbCrLf
    'If (f And DM_ICMINTENT) = DM_ICMINTENT Then         s = s & "    IcmIntent          = " & Me.IcmIntent & vbCrLf
    'If (f And DM_MEDIATYPE) = DM_MEDIATYPE Then         s = s & "    MediaType          = " & Me.MediaType & vbCrLf
    'If (f And DM_DITHERTYPE) = DM_DITHERTYPE Then       s = s & "    DitherType         = " & Me.DitherType & vbCrLf
    'If (f And DM_PANNINGHEIGHT) = DM_PANNINGHEIGHT Then s = s & "    PanningHeight      = " & Me.PanningHeight & vbCrLf
    
    
    's = s & "    FormName         = " & Me.FormName & vbCrLf
    's = s & "    VersionSpec      = " & Me.VersionSpec & vbCrLf
    's = s & "    Position X; Y    = " & Me.Position_ToStr & vbCrLf
    
    's = s & "    BitsPerPixel     = " & Me.BitsPerPixel & vbCrLf
    's = s & "    Width  [Pixels]  = " & Me.PixelsWidth & vbCrLf
    's = s & "    Height [Pixels]  = " & Me.PixelsHeight & vbCrLf
    's = s & "    Flags            = " & Me.DisplayFlags & vbCrLf
    's = s & "    Frequency [Hz]   = " & Me.DisplayFrequency & vbCrLf
    ToStr = s & "}"
End Function
