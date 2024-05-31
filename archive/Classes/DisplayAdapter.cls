VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class DisplayAdapter As DisplayDevice

Private m_Display As DisplayDevice
Private m_Connexs As Collection 'Of DisplayConnector

Private Sub Class_Initialize()
    Set m_Display = MNew.DisplayDevice()
    Set m_Connexs = New Collection
    Dim dc As DisplayConnector
    Dim i As Long, nm As String: nm = vbNullString 'm_Display.DeviceName
    Do
        Set dc = New DisplayConnector: dc.New_ nm, i
        If i < 0 Then Exit Do
        m_Connexs.Add dc
    Loop
End Sub

Public Property Get DeviceString() As String
    DeviceString = m_Display.DeviceString
End Property

Public Property Get DeviceID() As String
    DeviceID = m_Display.DeviceID
End Property

Public Property Get DeviceKey() As String
    DeviceKey = m_Display.DeviceKey
End Property

'Public Property Get Device() As String
'    Device = m_Display.Device
'End Property

Public Property Get DeviceName() As String
    DeviceName = m_Display.DeviceName
End Property

Public Property Get StateFlags() As Long
    StateFlags = m_Display.StateFlags
End Property

Friend Function StateFlags_ToStr() As String
    StateFlags_ToStr = m_Display.StateFlags_ToStr
End Function

Public Property Get Connectors() As Collection 'Of DisplayConnector
    If m_Connexs Is Nothing Then Set m_Connexs = m_Display.GetDisplayDevices(Me.DeviceName)
    Set Connectors = m_Connexs
End Property

Friend Function ToStr() As String
    Dim s As String
    s = "DisplayAdapter{" & vbCrLf
    s = s & "    DeviceName   = " & Me.DeviceName & vbCrLf
    s = s & "    DeviceString = " & Me.DeviceString & vbCrLf
    s = s & "    StateFlags   = " & Me.StateFlags & " := " & Me.StateFlags_ToStr & vbCrLf
    s = s & "    DeviceID     = " & Me.DeviceID & vbCrLf
    s = s & "    DeviceKey    = " & Me.DeviceKey & vbCrLf
    's = s & "    Device     = " & Me.Device & vbCrLf
    's = s & "    Display    = " & Me.Display & vbCrLf
    'If Not m_CurSetting Is Nothing Then
    '    s = s & "    cur-set    = " & Me.SettingCurrent.ShortToStr & vbCrLf
    'End If
    'If Not m_RegSetting Is Nothing Then
    '    s = s & "    reg-set    = " & Me.SettingRegistry.ShortToStr & vbCrLf
    'End If
    ToStr = s & "}" '& vbCrLf
End Function

