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

'Flags, die angeben, welche Werte geändert werden sollen
Private Const DM_POSITION            As Long = &H20&
Private Const DM_DISPLAYORIENTATION  As Long = &H80&    ' XP only
Private Const DM_BITSPERPEL          As Long = &H40000  ' Farbtiefe
Private Const DM_PELSWIDTH           As Long = &H80000  ' Breite
Private Const DM_PELSHEIGHT          As Long = &H100000 ' Höhe
Private Const DM_DISPLAYFLAGS        As Long = &H200000
Private Const DM_DISPLAYFREQUENCY    As Long = &H400000 ' Bild-Wiederholfrequenz
'Private Const DM_DISPLAYFIXEDOUTPUT As Long = &H20000000 ' XP only

Private Const EDS_RAWMODE            As Long = 2& ' Wenn diese Einstellung festgelegt ist, gibt die Funktion unabhängig von den Überwachungsfunktionen alle Grafikmodi zurück, die vom Adaptertreiber gemeldet werden. Andernfalls werden nur Modi zurückgegeben, die mit aktuellen Monitoren kompatibel sind.
Private Const EDS_ROTATEDMODE        As Long = 0& ' Wenn festgelegt, gibt die Funktion Grafikmodi in allen Ausrichtungen zurück. Andernfalls werden nur Modi zurückgegeben, die die gleiche Ausrichtung haben wie die derzeit für die angeforderte Anzeige festgelegte.

Private Const ENUM_CURRENT_SETTINGS  As Long = -1
Private Const ENUM_REGISTRY_SETTINGS As Long = -2

Private Const MONITORINFOF_PRIMARY   As Long = &H1&


'https://learn.microsoft.com/de-de/windows/win32/api/wingdi/ns-wingdi-devmodew

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
    dmDeviceName(1 To CCHDEVICENAME * 2) As Byte
    dmSpecVersion   As Integer
    dmDriverVersion As Integer
    dmSize          As Integer
    dmDriverExtra   As Integer
    dmFields        As Long
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
    dmPosition           As POINT
    dmDisplayOrientation As Long
    dmDisplayFixedOutput As Long
        '} DUMMYSTRUCTNAME2;
    '} DUMMYUNIONNAME;
    
    dmColor       As Integer
    dmDuplex      As Integer
    dmYResolution As Integer
    dmTTOption    As Integer
    dmCollate     As Integer
    dmFormName(1 To CCHFORMNAME * 2) As Byte
    dmLogPixels        As Integer
    dmBitsPerPel       As Long ' <<---- ' 8, 16, 24, 32
    dmPelsWidth        As Long ' <<---- '800, 1024, 1280, 1600, . . .
    dmPelsHeight       As Long ' <<---- '600,  768, 1024,  900, . . .
    'union {
    dmDisplayFlags     As Long ' <<----
    '    DWORD dmNup;
    '} DUMMYUNIONNAME2;
    dmDisplayFrequency As Long ' <<---- ' 60, 100, 120, 143, 166 . . .
    dmICMMethod        As Long
    dmICMIntent        As Long
    dmMediaType        As Long
    dmDitherType       As Long
    dmReserved1        As Long
    dmReserved2        As Long
    dmPanningWidth     As Long
    dmPanningHeight    As Long
End Type '} DEVMODEW, *PDEVMODEW, *NPDEVMODEW, *LPDEVMODEW;

Private m_DevMode As DEVICEMODEW
Private m_dmExtra(0 To 512) As Byte

Private Sub Class_Initialize()
    'm_DevMode.dmSize = LenB(m_DevMode)
    'm_DevMode.dmDriverExtra = 512 '8192
End Sub

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

Public Property Get PositionX() As Long
    PositionX = m_DevMode.dmPosition.X
End Property

Public Property Get PositionY() As Long
    PositionY = m_DevMode.dmPosition.Y
End Property

Public Property Get DisplayOrientation() As Long
    DisplayOrientation = m_DevMode.dmDisplayOrientation
End Property

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

'now for display settings
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




'Public Property Get BitsPerPixel() As Long
'    BitsPerPixel = m_DevMode.dmBitsPerPel
'End Property
'
'Public Property Get PixelsWidth() As Long
'    PixelsWidth = m_DevMode.dmPelsWidth
'End Property
'
'Public Property Get PixelsHeight() As Long
'    PixelsHeight = m_DevMode.dmPelsHeight
'End Property
'
'Public Property Get Flags() As Long
'    Flags = m_DevMode.dmDisplayFlags
'End Property
'
'Public Property Get Frequency() As Long
'    Frequency = m_DevMode.dmDisplayFrequency
'End Property

Public Property Get Orientation() As Long
    Orientation = m_DevMode.dmDisplayOrientation
End Property

Friend Function ToStr() As String
    Dim s As String: s = "DeviceMode {" & vbCrLf
    s = s & "    BitsPerPixel     = " & Me.BitsPerPixel & vbCrLf
    s = s & "    PixelsWidth      = " & Me.PixelsWidth & vbCrLf
    s = s & "    PixelsHeight     = " & Me.PixelsHeight & vbCrLf
    s = s & "    DisplayFlags     = " & Me.DisplayFlags & vbCrLf
    s = s & "    DisplayFrequency = " & Me.DisplayFrequency & vbCrLf
    ToStr = s & "}"
End Function
