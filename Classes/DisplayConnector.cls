VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class DisplayConnector As DisplayDevice

Private m_Display As DisplayDevice
Private m_Monitor As Monitor

Friend Sub New_(ByVal DevName As String, ByRef i_inout As Long)
    Set m_Display = MNew.DisplayDevice(DevName, i_inout)
    Set m_Monitor = MPC.ItemByKey(m_Display.DeviceName)
End Sub

Public Property Get DeviceName() As String
    DeviceName = m_Display.DeviceName
End Property

Public Property Get DeviceString() As String
    DeviceString = m_Display.DeviceString
End Property

Public Property Get StateFlags() As Long
    StateFlags = m_Display.StateFlags
End Property

Friend Function StateFlags_ToStr() As String
    StateFlags_ToStr = m_Display.StateFlags_ToStr
End Function

Public Property Get DeviceID() As String
    DeviceID = m_Display.DeviceID
End Property

Public Property Get DeviceKey() As String
    DeviceKey = m_Display.DeviceKey
End Property

Public Property Get Monitor() As Monitor
    Set Monitor = m_Monitor
End Property

Friend Function Monitor_ToStr() As String
    If m_Monitor Is Nothing Then Exit Function
    Monitor_ToStr = m_Monitor.ToStr
End Function

Friend Function ToStr() As String
    Dim s As String
    s = "DisplayConnector {" & vbCrLf
    s = s & "    DeviceName   = " & Me.DeviceName & vbCrLf
    s = s & "    DeviceString = " & Me.DeviceString & vbCrLf
    s = s & "    StateFlags   = " & Me.StateFlags & " := " & Me.StateFlags_ToStr & vbCrLf
    s = s & "    DeviceID     = " & Me.DeviceID & vbCrLf
    s = s & "    DeviceKey    = " & Me.DeviceKey & vbCrLf
    'If Not m_CurSetting Is Nothing Then
    '    s = s & "    cur-set    = " & Me.SettingCurrent.ShortToStr & vbCrLf
    'End If
    'If Not m_RegSetting Is Nothing Then
    '    s = s & "    reg-set    = " & Me.SettingRegistry.ShortToStr & vbCrLf
    'End If
    ToStr = s & "}" '& vbCrLf
End Function


