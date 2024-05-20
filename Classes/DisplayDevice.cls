VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'iModeNum
Private Const ENUM_CURRENT_SETTINGS  As Long = -1& ' Rufen Sie die aktuellen Einstellungen für das Anzeigegerät ab.
Private Const ENUM_REGISTRY_SETTINGS As Long = -2& ' Rufen Sie die Einstellungen für das Anzeigegerät ab, die derzeit in der Registrierung gespeichert sind.

'dwFlags
Private Const EDS_RAWMODE     As Long = 2 ' Wenn diese Einstellung festgelegt ist, gibt die Funktion unabhängig von den Überwachungsfunktionen alle Grafikmodi zurück, die vom Adaptertreiber gemeldet werden. Andernfalls werden nur Modi zurückgegeben, die mit aktuellen Monitoren kompatibel sind.
Private Const EDS_ROTATEDMODE As Long = 4 ' Wenn festgelegt, gibt die Funktion Grafikmodi in allen Ausrichtungen zurück. Andernfalls werden nur Modi zurückgegeben, die die gleiche Ausrichtung haben wie die derzeit für die angeforderte Anzeige festgelegte.

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-display_devicew
'typedef struct _DISPLAY_DEVICEW {
'  DWORD cb;
'  WCHAR DeviceName[32];
'  WCHAR DeviceString[128];
'  DWORD StateFlags;
'  WCHAR DeviceID[128];
'  WCHAR DeviceKey[128];
'} DISPLAY_DEVICEW, *PDISPLAY_DEVICEW, *LPDISPLAY_DEVICEW;
Private Type DISPLAY_DEVICEW
    cb                         As Long
    DeviceName(1 To 32 * 2)    As Byte
    DeviceString(1 To 128 * 2) As Byte
    StateFlags                 As Long
    DeviceID(1 To 128 * 2)     As Byte
    DeviceKey(1 To 128 * 2)    As Byte
End Type '} DISPLAY_DEVICEW, *PDISPLAY_DEVICEW, *LPDISPLAY_DEVICEW;

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaysettingsw
'BOOL EnumDisplaySettingsW([in] LPCWSTR lpszDeviceName, [in] DWORD iModeNum, [out] DEVMODEW *lpDevMode);
'Private Declare Function EnumDisplaySettingsW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, ByVal lpDevMode As LongPtr) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaysettingsexw
'BOOL EnumDisplaySettingsExW([in] LPCWSTR lpszDeviceName, [in] DWORD iModeNum, [out] DEVMODEW *lpDevMode, [in] DWORD dwFlags);
Private Declare Function EnumDisplaySettingsExW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, ByVal lpDevMode As LongPtr, ByVal dwFlags As Long) As Long

Private m_Display     As DISPLAY_DEVICEW
Private m_CurSetting  As DeviceMode
Private m_RegSetting  As DeviceMode
Private m_Settings    As Collection
Private m_GraCard     As Collection
Private m_RegSettings As Collection 'Of DeviceMode

Private Sub Class_Initialize()
    m_Display.cb = LenB(m_Display)
End Sub

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_Display)
End Property

Public Property Get DeviceName() As String
    DeviceName = Trim0(m_Display.DeviceName)
End Property

Public Property Get Display() As String
    Dim dn As String: dn = m_Display.DeviceName
    Dim pos As Long: pos = InStrRev(dn, "\")
    If pos > 0 Then dn = Left(dn, pos - 1)
    Display = dn 'Trim0(m_Display.DeviceName)
End Property

Public Property Get Device() As String
    Device = Trim0(m_Display.DeviceString)
End Property

Public Property Get Flags() As Long
    Flags = m_Display.StateFlags
End Property
    
Public Property Get ID() As String
    ID = Trim0(m_Display.DeviceID)
End Property

Public Property Get Key() As String
    Key = Trim0(m_Display.DeviceKey)
End Property

Private Function GetDeviceMode(ByVal DevName As String, ByVal iMode As Long, ByVal iFlag As Long) As DeviceMode
    Dim dm As New DeviceMode
    Dim lpDevNam As LongPtr
    If Len(DevName) Then
        DevName = DevName & vbNullChar
        lpDevNam = StrPtr(DevName)
    End If
    Dim hr As Long: hr = EnumDisplaySettingsExW(lpDevNam, iMode, dm.Ptr, iFlag)
    If hr Then Set GetDeviceMode = dm
End Function

Private Function GetDeviceModes(ByVal DevName As String, ByVal iFlag As Long) As Collection
    Dim modes As New Collection
    Dim dm As DeviceMode
    Dim i As Long
    Do
        Set dm = GetDeviceMode(DevName, i, iFlag)
        If dm Is Nothing Then Exit Do
        modes.Add dm
        i = i + 1
    Loop
    Set GetDeviceModes = modes
End Function

Public Property Get Settings() As Collection
    If m_Settings Is Nothing Then
        Set m_Settings = GetDeviceModes(Me.Display, EDS_ROTATEDMODE)  ' EDS_RAWMODE) '1)
    End If
    Set Settings = m_Settings
End Property

Public Property Get SettingsGraphicsCard() As Collection 'Of DeviceMode
    If m_GraCard Is Nothing Then
        Set m_GraCard = GetDeviceModes(vbNullString, EDS_ROTATEDMODE) ' EDS_RAWMODE) '1)
    End If
    Set SettingsGraphicsCard = m_GraCard
End Property

Public Property Get SettingCurrent() As DeviceMode
    If m_CurSetting Is Nothing Then Set m_CurSetting = GetDeviceMode(Me.Display, ENUM_CURRENT_SETTINGS, EDS_ROTATEDMODE) ' EDS_RAWMODE) '1)
    Set SettingCurrent = m_CurSetting
End Property

Public Property Get SettingRegistry() As DeviceMode
    If m_RegSetting Is Nothing Then Set m_RegSetting = GetDeviceMode(Me.Display, ENUM_REGISTRY_SETTINGS, EDS_ROTATEDMODE) ' EDS_RAWMODE) '1)
    Set SettingRegistry = m_RegSetting
End Property

Friend Function ToStr() As String
    Dim s As String
    s = "Display {" & vbCrLf
    s = s & "    ID         = " & Me.ID & vbCrLf
    s = s & "    Key        = " & Me.Key & vbCrLf
    s = s & "    DeviceName = " & Me.DeviceName & vbCrLf
    s = s & "    Flags      = " & Me.Flags & vbCrLf
    s = s & "    Device     = " & Me.Device & vbCrLf
    s = s & "    cur-set    = " & Me.SettingCurrent.ShortToStr & vbCrLf
    s = s & "    reg-set    = " & Me.SettingRegistry.ShortToStr & vbCrLf
    ToStr = s & "}" '& vbCrLf
End Function
