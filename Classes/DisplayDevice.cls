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
Private Const ENUM_CURRENT_SETTINGS  As Long = -1& ' Rufen Sie die aktuellen Einstellungen f�r das Anzeigeger�t ab.
Private Const ENUM_REGISTRY_SETTINGS As Long = -2& ' Rufen Sie die Einstellungen f�r das Anzeigeger�t ab, die derzeit in der Registrierung gespeichert sind.

'dwFlags
Private Const EDS_RAWMODE     As Long = 2 ' Wenn diese Einstellung festgelegt ist, gibt die Funktion unabh�ngig von den �berwachungsfunktionen alle Grafikmodi zur�ck, die vom Adaptertreiber gemeldet werden. Andernfalls werden nur Modi zur�ckgegeben, die mit aktuellen Monitoren kompatibel sind.
Private Const EDS_ROTATEDMODE As Long = 4 ' Wenn festgelegt, gibt die Funktion Grafikmodi in allen Ausrichtungen zur�ck. Andernfalls werden nur Modi zur�ckgegeben, die die gleiche Ausrichtung haben wie die derzeit f�r die angeforderte Anzeige festgelegte.

Private Const EDD_GET_DEVICE_INTERFACE_NAME As Long = &H1&

'StateFlags:

Private Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1&
Private Const DISPLAY_DEVICE_ACTIVE              As Long = &H1&       ' specifies whether a monitor is presented as being "on" by the respective GDI view. Windows Vista: EnumDisplayDevices will only enumerate monitors that can be presented as being "on."
Private Const DISPLAY_DEVICE_MULTI_DRIVER        As Long = &H2&
Private Const DISPLAY_DEVICE_PRIMARY_DEVICE      As Long = &H4&       ' The primary desktop is on the device. For a system with a single display card, this is always set. For a system with multiple display cards, only one device can have this set.
Private Const DISPLAY_DEVICE_MIRRORING_DRIVER    As Long = &H8&       ' Represents a pseudo device used to mirror application drawing for remoting or other purposes. An invisible pseudo monitor is associated with this device. For example, NetMeeting uses it. Note that GetSystemMetrics (SM_MONITORS) only accounts for visible display monitors.
Private Const DISPLAY_DEVICE_VGA_COMPATIBLE      As Long = &H10&      ' The device is VGA compatible.
Private Const DISPLAY_DEVICE_REMOVABLE           As Long = &H20&      ' The device is removable; it cannot be the primary display.
Private Const DISPLAY_DEVICE_MODESPRUNED         As Long = &H8000000  ' The device has more display modes than its output devices support.

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

#If VBA7 Then
    Private Declare PtrSafe Function EnumDisplaySettingsExW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, ByVal lpDevMode As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function EnumDisplayDevicesW Lib "user32" (ByVal lpDevice As LongPtr, ByVal iDevNum As Long, ByVal lpDisplayDevice As LongPtr, ByVal dwFlags As Long) As Long
#Else
    'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaysettingsw
    'BOOL EnumDisplaySettingsW([in] LPCWSTR lpszDeviceName, [in] DWORD iModeNum, [out] DEVMODEW *lpDevMode);
    'Private Declare Function EnumDisplaySettingsW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, ByVal lpDevMode As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaysettingsexw
    'BOOL EnumDisplaySettingsExW([in] LPCWSTR lpszDeviceName, [in] DWORD iModeNum, [out] DEVMODEW *lpDevMode, [in] DWORD dwFlags);
    Private Declare Function EnumDisplaySettingsExW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, ByVal lpDevMode As LongPtr, ByVal dwFlags As Long) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-enumdisplaydevicesw
    'BOOL EnumDisplayDevicesW( [in] LPCWSTR lpDevice, [in] DWORD iDevNum, [out] PDISPLAY_DEVICEW lpDisplayDevice, [in] DWORD dwFlags);
    Private Declare Function EnumDisplayDevicesW Lib "user32" (ByVal lpDevice As LongPtr, ByVal iDevNum As Long, ByVal lpDisplayDevice As LongPtr, ByVal dwFlags As Long) As Long
    
#End If
Private m_DevName     As String
Private m_Display     As DISPLAY_DEVICEW
Private m_CurSetting  As DeviceMode
Private m_RegSetting  As DeviceMode
Private m_Settings    As Collection
Private m_GraCard     As Collection
Private m_RegSettings As Collection 'Of DeviceMode

Private Sub Class_Initialize()
    m_Display.cb = LenB(m_Display)
End Sub

Friend Sub New_(Optional ByVal DevName As String = vbNullString, Optional ByRef i_inout As Long = 0)
    m_DevName = DevName
    Dim lpDevNam As LongPtr
    If Len(DevName) Then
        DevName = DevName & vbNullChar
        lpDevNam = StrPtr(DevName)
    End If
    Dim hr As Long: hr = EnumDisplayDevicesW(lpDevNam, i_inout, VarPtr(m_Display), EDD_GET_DEVICE_INTERFACE_NAME)
    If hr Then
        i_inout = i_inout + 1
    Else
        i_inout = -1
        'Fehlermeldung ausgeben
    End If
End Sub

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_Display)
End Property

'Public Property Get PropDisplay() As String
'    Dim d0 As String: d0 = Me.DeviceName
'    Dim pos As Long: pos = InStrRev(d0, "\")
'    Dim dn As String: If pos > 0 Then dn = Left(d0, pos - 1)
'    If Len(dn) < 5 Then dn = d0
'    PropDisplay = dn 'Trim0(m_Display.DeviceName)
'End Property

Public Property Get DeviceName() As String
    DeviceName = Trim0(m_Display.DeviceName)
End Property

Private Property Get DevName() As String
    DevName = m_DevName: If Len(DevName) = 0 Then DevName = Me.DeviceName
End Property

Public Property Get DeviceString() As String
    DeviceString = Trim0(m_Display.DeviceString)
End Property

Public Property Get StateFlags() As Long
    StateFlags = m_Display.StateFlags
End Property

Friend Function StateFlags_ToStr() As String
    Dim sf As Long: sf = m_Display.StateFlags
    Dim s As String
    If sf And DISPLAY_DEVICE_ACTIVE Then s = s & IIf(Len(s), " Or ", "") & "DD_ACTIVE"           ' specifies whether a monitor is presented as being "on" by the respective GDI view. Windows Vista: EnumDisplayDevices will only enumerate monitors that can be presented as being "on."
    If sf And DISPLAY_DEVICE_MULTI_DRIVER Then s = s & IIf(Len(s), " Or ", "") & "DD_MULTI_DRIVER"      '
    If sf And DISPLAY_DEVICE_PRIMARY_DEVICE Then s = s & IIf(Len(s), " Or ", "") & "DD_PRIMARY_DEVICE"   ' The primary desktop is on the device. For a system with a single display card, this is always set. For a system with multiple display cards, only one device can have this set.
    If sf And DISPLAY_DEVICE_MIRRORING_DRIVER Then s = s & IIf(Len(s), " Or ", "") & "DD_MIRRORING_DRIVER" ' Represents a pseudo device used to mirror application drawing for remoting or other purposes. An invisible pseudo monitor is associated with this device. For example, NetMeeting uses it. Note that GetSystemMetrics (SM_MONITORS) only accounts for visible display monitors.
    If sf And DISPLAY_DEVICE_VGA_COMPATIBLE Then s = s & IIf(Len(s), " Or ", "") & "DD_VGA_COMPATIBLE"   ' The device is VGA compatible.
    If sf And DISPLAY_DEVICE_REMOVABLE Then s = s & IIf(Len(s), " Or ", "") & "DD_REMOVABLE"        ' The device is removable; it cannot be the primary display.
    If sf And DISPLAY_DEVICE_MODESPRUNED Then s = s & IIf(Len(s), " Or ", "") & "DD_MODESPRUNED"      ' The device has more display modes than its output devices support.
    'case DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1&
    'Case Else:
    'End Select
    StateFlags_ToStr = s
End Function

Public Property Get DeviceID() As String
    DeviceID = Trim0(m_Display.DeviceID)
End Property

Public Property Get DeviceKey() As String
    DeviceKey = Trim0(m_Display.DeviceKey)
End Property

Friend Function GetDeviceMode(ByVal DevName As String, ByVal iMode As Long, ByVal iFlag As Long) As DeviceMode
    Dim dm As New DeviceMode
    Dim lpDevNam As LongPtr
    If Len(DevName) Then
        DevName = DevName & vbNullChar
        lpDevNam = StrPtr(DevName)
    End If
    Dim hr As Long: hr = EnumDisplaySettingsExW(lpDevNam, iMode, dm.Ptr, iFlag)
    If hr Then Set GetDeviceMode = dm
End Function

Friend Function GetDeviceModes(ByVal DevName As String, ByVal iFlag As Long) As Collection
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

Friend Function GetDisplayDevice(ByVal DevName As String, ByVal i As Long) As DisplayDevice
    Dim dd As New DisplayDevice
    Dim lpDevNam As LongPtr
    If Len(DevName) Then
        DevName = DevName & vbNullChar
        lpDevNam = StrPtr(DevName)
    End If
    Dim hr As Long: hr = EnumDisplayDevicesW(lpDevNam, i, dd.Ptr, EDD_GET_DEVICE_INTERFACE_NAME)
    If hr Then Set GetDisplayDevice = dd
End Function

Friend Function GetDisplayDevices(ByVal DevName As String) As Collection
    Dim DisplayDevices As New Collection
    Dim dd As DisplayDevice
    Dim i As Long
    Do
        Set dd = GetDisplayDevice(DevName, i)
        If dd Is Nothing Then Exit Do
        DisplayDevices.Add dd
        i = i + 1
    Loop
    Set GetDisplayDevices = DisplayDevices
End Function

Public Property Get Settings() As Collection
    If m_Settings Is Nothing Then
        Set m_Settings = GetDeviceModes(DevName, EDS_ROTATEDMODE)  ' EDS_RAWMODE) '1)
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
    If m_CurSetting Is Nothing Then Set m_CurSetting = GetDeviceMode(DevName, ENUM_CURRENT_SETTINGS, EDS_ROTATEDMODE) ' EDS_RAWMODE) '1)
    Set SettingCurrent = m_CurSetting
End Property

Public Property Get SettingRegistry() As DeviceMode
    If m_RegSetting Is Nothing Then Set m_RegSetting = GetDeviceMode(DevName, ENUM_REGISTRY_SETTINGS, EDS_ROTATEDMODE) ' EDS_RAWMODE) '1)
    Set SettingRegistry = m_RegSetting
End Property

Friend Function ToStr() As String
    Dim s As String
    s = "DISPLAY_DEVICEW {" & vbCrLf
    's = s & "    PropDisplay  = " & Me.PropDisplay & vbCrLf
    s = s & "    DeviceName   = " & Me.DeviceName & vbCrLf
    s = s & "    DeviceString = " & Me.DeviceString & vbCrLf
    s = s & "    StateFlags   = " & Me.StateFlags & " := " & Me.StateFlags_ToStr & vbCrLf
    s = s & "    DeviceID     = " & Me.DeviceID & vbCrLf
    s = s & "    DeviceKey    = " & Me.DeviceKey & vbCrLf
    If Not m_CurSetting Is Nothing Then
        s = s & "    cur-set    = " & Me.SettingCurrent.ShortToStr & vbCrLf
    End If
    If Not m_RegSetting Is Nothing Then
        s = s & "    reg-set    = " & Me.SettingRegistry.ShortToStr & vbCrLf
    End If
    ToStr = s & "}" '& vbCrLf
End Function
