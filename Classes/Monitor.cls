VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MONITOR_DEFAULTTONULL    As Long = &H0& 'Gibt NULL zurück.
Private Const MONITOR_DEFAULTTOPRIMARY As Long = &H1& 'Gibt ein Handle an den primären Anzeigemonitor zurück.
Private Const MONITOR_DEFAULTTONEAREST As Long = &H2& 'Gibt ein Handle an den Anzeigemonitor zurück, das dem Punkt am nächsten ist.

Private Const MONITORINFOF_PRIMARY     As Long = &H1&

Public Enum EMonitorDefault
    ToNull = 0
    ToPrimary = 1
    ToNearest = 2
End Enum

Private Const CCHDEVICENAME            As Long = 32&

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const EDD_GET_DEVICE_INTERFACE_NAME As Long = &H1&
Private Const GUID_DEVINTERFACE_MONITOR As Long = 0

'https://learn.microsoft.com/de-de/windows/win32/api/Winuser/ns-winuser-monitorinfo
'https://learn.microsoft.com/de-de/windows/win32/api/winuser/ns-winuser-monitorinfoexw
Private Type MONITORINFO
    cbSize    As Long
    rcMonitor As RECT
    rcWork    As RECT
    dwFlags   As Long
    szDevice(1 To CCHDEVICENAME * 2) As Byte
End Type

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-monitorfrompoint
'HMONITOR MonitorFromPoint([in] POINT pt, [in] DWORD dwFlags);
Private Declare Function MonitorFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long, ByVal dwFlags As Long) As LongPtr

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-monitorfromrect
'HMONITOR MonitorFromRect([in] LPCRECT lprc, [in] DWORD dwFlags);
Private Declare Function MonitorFromRect Lib "user32" (ByRef lprc As RECT, ByVal dwFlags As Long) As LongPtr

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-monitorfromwindow
'HMONITOR MonitorFromWindow([in] HWND  hwnd, [in] DWORD dwFlags);
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As LongPtr

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-getmonitorinfow
'BOOL GetMonitorInfoW([in] HMONITOR hMonitor, [out] LPMONITORINFO lpmi);
Private Declare Function GetMonitorInfoW Lib "user32" (ByVal HMONITOR As LongPtr, ByRef lpmi_out As MONITORINFO) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-enumdisplaydevicesw
'BOOL EnumDisplayDevicesW( [in] LPCWSTR lpDevice, [in] DWORD iDevNum, [out] PDISPLAY_DEVICEW lpDisplayDevice, [in] DWORD dwFlags);
Private Declare Function EnumDisplayDevicesW Lib "user32" (ByVal lpDevice As LongPtr, ByVal iDevNum As Long, ByVal lpDisplayDevice As LongPtr, ByVal dwFlags As Long) As Long

Private m_Handle   As LongPtr
Private m_Info     As MONITORINFO
Private m_Flag     As Long
Private m_hDC      As LongPtr
Private m_Clip     As RECT
'Private m_Displays As DisplayDevices
Private m_Displays As Collection

Private Sub Class_Initialize()
    m_Info = GetMonitor(MonitorFromPoint(0, 0, MONITOR_DEFAULTTOPRIMARY)) 'default to primary
End Sub

Friend Sub New_(ByVal HMONITOR As LongPtr, ByVal hDC As LongPtr, ByVal lpClipRECT As LongPtr)
    m_hDC = hDC
    RtlMoveMemory m_Clip, ByVal lpClipRECT, LenB(m_Clip)
    m_Info = GetMonitor(HMONITOR)
End Sub

Friend Sub NewP(ByVal X As Long, ByVal Y As Long, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest)
    m_Info = GetMonitor(MonitorFromPoint(X, Y, aFlag))
End Sub

Friend Sub NewR(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest)
    Dim R As RECT
    With R: .Left = X: .Top = Y: .Right = .Left + Width: .Bottom = .Top + Height: End With
    m_Info = GetMonitor(MonitorFromRect(R, aFlag))
End Sub

Friend Sub NewW(ByVal hWnd As LongPtr, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest)
    m_Info = GetMonitor(MonitorFromWindow(hWnd, aFlag))
End Sub

Private Function GetMonitor(ByVal Handle As LongPtr) As MONITORINFO
    m_Handle = Handle
    If m_Handle <> 0 Then
        GetMonitor.cbSize = LenB(GetMonitor)
        Dim hr As Long: hr = GetMonitorInfoW(m_Handle, GetMonitor)
    End If
End Function

Public Property Get Handle() As LongPtr
    Handle = m_Handle
End Property

Public Property Get Name() As String
    Name = Trim0(m_Info.szDevice)
End Property

Public Property Get Left() As Long
    Left = m_Info.rcMonitor.Left
End Property
Public Property Get Top() As Long
    Top = m_Info.rcMonitor.Top
End Property
Public Property Get Width() As Long
    Width = m_Info.rcMonitor.Right - Left
End Property
Public Property Get Height() As Long
    Height = m_Info.rcMonitor.Bottom - Top
End Property

Public Property Get WorkLeft() As Long
    WorkLeft = m_Info.rcWork.Left
End Property
Public Property Get WorkTop() As Long
    WorkTop = m_Info.rcWork.Top
End Property
Public Property Get WorkWidth() As Long
    WorkWidth = m_Info.rcWork.Right - WorkLeft
End Property
Public Property Get WorkHeight() As Long
    WorkHeight = m_Info.rcWork.Bottom - WorkTop
End Property

Public Property Get ClipLeft() As Long
    ClipLeft = m_Clip.Left
End Property
Public Property Get ClipTop() As Long
    ClipTop = m_Clip.Top
End Property
Public Property Get ClipWidth() As Long
    ClipWidth = m_Clip.Right - WorkLeft
End Property
Public Property Get ClipHeight() As Long
    ClipHeight = m_Clip.Bottom - WorkTop
End Property

Public Property Get Displays() As Collection 'Of DisplayDevice
    If m_Displays Is Nothing Then Set m_Displays = GetDisplayDevices(Me.Name)
    Set Displays = m_Displays
End Property
'Public Property Get Display(ByVal Index As Long) As DisplayDevice
'    If m_Displays Is Nothing Then Set m_Displays = GetDisplayDevices(Me.Name)
'    If Index < m_Displays.Count Then Set Display = m_Displays.Item(Index + 1)
'End Property

Private Function GetDisplayDevices(ByVal DevName As String) As Collection
    Dim Displays As New Collection
    Dim dd As DisplayDevice
    Dim i As Long
    Do
        Set dd = New DisplayDevice
        DevName = DevName & vbNullChar
        Dim hr As Long: hr = EnumDisplayDevicesW(StrPtr(DevName), i, dd.Ptr, EDD_GET_DEVICE_INTERFACE_NAME)
        If hr Then Displays.Add dd
        i = i + 1
    Loop Until hr = 0
    Set GetDisplayDevices = Displays
End Function

Private Function RECT_ToStr(R As RECT) As String
    With R
        RECT_ToStr = "RECT{L=" & .Left & "; T=" & .Top & "; R=" & .Right & "; B=" & .Bottom & "}"
    End With
End Function

Private Function RECT_IsNull(R As RECT) As Boolean
    Dim b As Boolean
    With R
        b = .Left <> 0:   If b Then Exit Function
        b = .Top <> 0:    If b Then Exit Function
        b = .Right <> 0:  If b Then Exit Function
        b = .Bottom <> 0: If b Then Exit Function
    End With
    RECT_IsNull = True
End Function

Friend Function ToStr() As String
    Dim s As String
    s = "Monitor{" & vbCrLf
    s = s & "    Name  = " & Me.Name & vbCrLf
    s = s & "    Handle= " & m_Handle & vbCrLf
    s = s & "    W / H = " & Me.Width & " / " & Me.Height & vbCrLf
    s = s & "    Moni-" & RECT_ToStr(m_Info.rcMonitor) & vbCrLf
    s = s & "    Work-" & RECT_ToStr(m_Info.rcWork) & vbCrLf
    If m_hDC <> 0 Then
        s = s & "    hDC   = " & m_hDC & vbCrLf
    End If
    If Not RECT_IsNull(m_Clip) Then
        s = s & "    Clip-" & RECT_ToStr(m_Clip) & vbCrLf
    End If
    ToStr = s & "}" & vbCrLf
End Function
