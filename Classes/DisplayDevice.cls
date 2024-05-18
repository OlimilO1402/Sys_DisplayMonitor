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

'https://learn.microsoft.com/de-de/windows/win32/api/wingdi/ns-wingdi-display_devicew
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

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-enumdisplaysettingsw
'BOOL EnumDisplaySettingsW([in] LPCWSTR lpszDeviceName, [in] DWORD iModeNum, [out] DEVMODEW *lpDevMode);
'Private Declare Function EnumDisplaySettingsW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, ByVal lpDevMode As LongPtr) As Long
                 
'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-enumdisplaysettingsexw
'BOOL EnumDisplaySettingsExW([in] LPCWSTR lpszDeviceName, [in] DWORD iModeNum, [out] DEVMODEW *lpDevMode, [in] DWORD dwFlags);
Private Declare Function EnumDisplaySettingsExW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, ByVal lpDevMode As LongPtr, ByVal dwFlags As Long) As Long

Private m_Display     As DISPLAY_DEVICEW
Private m_CurSettings As Collection 'Of DeviceMode
Private m_RegSettings As Collection 'Of DeviceMode

Private Sub Class_Initialize()
    m_Display.cb = LenB(m_Display)
End Sub

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_Display)
End Property

Public Property Get Name() As String
    Name = Trim0(m_Display.DeviceName)
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

Private Function GetDeviceModes(ByVal DevName As String, ByVal iMode As Long, ByVal iFlag As Long) As Collection
    Dim Settings As New Collection
    Dim dm As DeviceMode
    Dim i As Long
    DevName = DevName & vbNullChar
    'DevName = "\\.\DISPLAY2" & vbNullChar
    Do
        Set dm = New DeviceMode ' MNew.DeviceMode(DevName)
        Dim hr As Long
        hr = EnumDisplaySettingsExW(StrPtr(DevName), iMode, dm.Ptr, iFlag)
        'hr = EnumDisplaySettingsExW(StrPtr(DevName), i, dm.Ptr, iFlag)
        'hr = EnumDisplaySettingsExW(0&, i, dm.Ptr, iFlag)
        If hr Then Settings.Add dm
        i = i + 1
    Loop Until hr = 0
    Set GetDeviceModes = Settings
End Function

Public Property Get SettingsCurrent() As Collection
    If m_CurSettings Is Nothing Then Set m_CurSettings = GetDeviceModes(Me.Display, ENUM_CURRENT_SETTINGS, EDS_ROTATEDMODE) ' EDS_RAWMODE) '1)
    Set SettingsCurrent = m_CurSettings
End Property

Public Property Get SettingsRegistry() As Collection
    If m_RegSettings Is Nothing Then Set m_RegSettings = GetDeviceModes(Me.Display, ENUM_REGISTRY_SETTINGS, EDS_ROTATEDMODE)
    'Set m_RegSettings = New Collection
    Set SettingsRegistry = m_RegSettings
End Property

Friend Function ToStr() As String
    Dim s As String
    s = "Display {" & vbCrLf
    s = s & "    ID     = " & Me.ID & vbCrLf
    s = s & "    Key    = " & Me.Key & vbCrLf
    s = s & "    Name   = " & Me.Name & vbCrLf
    s = s & "    Flags  = " & Me.Flags & vbCrLf
    s = s & "    Device = " & Me.Device & vbCrLf
    ToStr = s & "}"
End Function
