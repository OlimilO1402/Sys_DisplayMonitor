VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "View & Change Display-Settings"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Form2"
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   0
      Width           =   855
   End
   Begin VB.ListBox List9 
      Appearance      =   0  '2D
      Height          =   1815
      ItemData        =   "FMain.frx":1782
      Left            =   12360
      List            =   "FMain.frx":1784
      TabIndex        =   12
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ListBox List8 
      Appearance      =   0  '2D
      Height          =   1815
      ItemData        =   "FMain.frx":1786
      Left            =   9960
      List            =   "FMain.frx":1788
      TabIndex        =   11
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ListBox List7 
      Appearance      =   0  '2D
      Height          =   1815
      ItemData        =   "FMain.frx":178A
      Left            =   7560
      List            =   "FMain.frx":178C
      TabIndex        =   10
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ListBox List6 
      Appearance      =   0  '2D
      Height          =   1815
      ItemData        =   "FMain.frx":178E
      Left            =   5160
      List            =   "FMain.frx":1790
      TabIndex        =   9
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   7
      Top             =   7800
      Width           =   5175
   End
   Begin VB.ListBox List4 
      Appearance      =   0  '2D
      Height          =   4620
      ItemData        =   "FMain.frx":1792
      Left            =   7560
      List            =   "FMain.frx":1794
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   7800
      Width           =   5175
   End
   Begin VB.ListBox List3 
      Appearance      =   0  '2D
      Height          =   4620
      ItemData        =   "FMain.frx":1796
      Left            =   0
      List            =   "FMain.frx":1798
      TabIndex        =   4
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   2520
      Width           =   12735
   End
   Begin VB.ListBox List5 
      Appearance      =   0  '2D
      Height          =   1050
      ItemData        =   "FMain.frx":179A
      Left            =   0
      List            =   "FMain.frx":179C
      TabIndex        =   8
      Top             =   4080
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Appearance      =   0  '2D
      Height          =   795
      ItemData        =   "FMain.frx":179E
      Left            =   0
      List            =   "FMain.frx":17A0
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox TBMonitor 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   0
      Width           =   12735
   End
   Begin VB.ListBox LBMonitors 
      Appearance      =   0  '2D
      Height          =   2070
      ItemData        =   "FMain.frx":17A2
      Left            =   0
      List            =   "FMain.frx":17A9
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label LlMonitors 
      Caption         =   "Monitors: "
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Label1"
      Height          =   375
      Left            =   12360
      TabIndex        =   16
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Label1"
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   5280
      Width           =   2415
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_GraCard As DisplayAdapter

Private m_Monitor As Monitor
Private m_Device  As DisplayDevice
Private m_Mode    As DeviceMode
Private m_Filter  As DisplayModeFilter


Private Sub Command1_Click()
    FMain2.Show
End Sub

' Also für mich werden da erstmal veschiedene Begrifflichkeiten wild durcheinander geworfen.
' Da tauchen display devices, display monitors sowie display adapters auf.
' Anyway, ich vermute mal, du machst das hier:
'
'    To obtain information on a display monitor, first call EnumDisplayDevices with lpDevice set to NULL.
'    Then call EnumDisplayDevices with lpDevice set to DISPLAY_DEVICE.DeviceName from the first call to
'    EnumDisplayDevices and with iDevNum set to zero. Then DISPLAY_DEVICE.DeviceString is the monitor name.
' Beim ersten Aufruf mit NULL kommt der funktionierende String ("\\.\DISPLAY*") zurück (display adapter, schätze ich).
' Füttert man das nochmal in EnumDisplayDevices, kommt das, was im Quote als display monitor bezeichnet wird ("\\.\DISPLAY*\Monitor*").
' Was EnumDisplaySettingsEx anscheinend erwartet, ist ein display adapter.
' Warum das jetzt aber so nicht dasteht, wird nur MS beantworten können.


Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Monitors.Init
    UpdateViewMonitors
    LBMonitors.ListIndex = 0
    'Ähm so ein Mist meine Verschachtelungstheorie ist Mist
    'wir haben
    'Monitors -> List Of Monitor -> List Of DisplayDevice -> List Of DeviceMode
    'stattdessen sollte es sein:
    'PC -> List Of Monitor -> List Of DisplayDevice
    '                      -> List Of DeviceMode
    'PC -> List Of GraCard -> List Of DisplayDevice
    '                      -> List Of DeviceMode
    
    
    ' PC -> (DisplayDevice/DisplayAdapter) GraCard-01 -> (DisplayDevice) Display-01 & MonitorInfo
    '                                                                    -> DeviceMode-01
    '                                                                    '> DeviceMode-02
    '                                                                    '> DeviceMode-..
    '                                                                    '> DeviceMode-nn
    '                                                 '> (DisplayDevice) Display-02 & MonitorInfo
    '                                                                    -> DeviceMode-01
    '                                                                    '> DeviceMode-nn
    '                                                 '> (DisplayDevice) Display-..
    '                                                                    -> DeviceMode-01
    '                                                                    '> DeviceMode-nn
    '                                                 '> (DisplayDevice) Display-nn
    '                                                                    -> DeviceMode-01
    '                                                                    '> DeviceMode-nn
    '
    ' PC -> (DisplayDevice/DisplayAdapter) GraCard-02 -> (DisplayDevice) Display-01
    '                                                 '> (DisplayDevice) Display-02
    '                                                 '> (DisplayDevice) Display-..
    '                                                 '> (DisplayDevice) Display-nn
    '
    ' PC -> (DisplayDevice/DisplayAdapter) GraCard-.. -> (DisplayDevice) Display-01
    '                                                 '> (DisplayDevice) Display-02
    '                                                 '> (DisplayDevice) Display-..
    '                                                 '> (DisplayDevice) Display-nn
    '
    ' PC -> (DisplayDevice/DisplayAdapter) GraCard-nn -> (DisplayDevice) Display-01
    '                                                 '> (DisplayDevice) Display-02
    '                                                 '> (DisplayDevice) Display-..
    '                                                 '> (DisplayDevice) Display-nn
    
    

'
'EnumDisplayDevicesW
'===================
'
'-> To query all display devices in the current session, call this function in a loop,
'   starting with iDevNum set to 0, and incrementing iDevNum until the function fails.
'
'-> To select all display devices in the desktop, use only the display devices that
'   have the DISPLAY_DEVICE_ATTACHED_TO_DESKTOP flag in the DISPLAY_DEVICE structure.
'
'-> To get information on the display adapter, call EnumDisplayDevices with lpDevice set to NULL.
'   For example, DISPLAY_DEVICE.DeviceString contains the adapter name.
'
'-> To obtain information on a display monitor, first call EnumDisplayDevices with lpDevice set to NULL.
'   Then call EnumDisplayDevices with lpDevice set to DISPLAY_DEVICE.DeviceName from the first call to
'   EnumDisplayDevices and with iDevNum set to zero. Then DISPLAY_DEVICE.DeviceString is the monitor name.
'
'To query all monitor devices associated with an adapter, call EnumDisplayDevices in a
'loop with lpDevice set to the adapter name, iDevNum set to start at 0, and iDevNum set
'to increment until the function fails.

'----------------------------------------------------------------------------------------------------------------------
'Note that DISPLAY_DEVICE.DeviceName changes with each call for monitor information, so you must save the adapter name.
'----------------------------------------------------------------------------------------------------------------------

'The function fails when there are no more monitors for the adapter.
'
'
'EnumDisplaySettingsExW
'======================
'To obtain information for all of a display device's graphics modes,
'make a series of calls to EnumDisplaySettingsEx, as follows:
'Set iModeNum to zero for the first call, and increment iModeNum by one for each subsequent call.
'Continue calling the function until the return value is zero.
    
End Sub

Sub UpdateViewMonitors()
    Dim m As Monitor, i As Long, c As Long: c = Monitors.Count
    LlMonitors.Caption = "Monitors: " & c
    LBMonitors.Clear
    For i = 0 To c - 1
        Set m = Monitors.Item(i)
        LBMonitors.AddItem m.Device
    Next
End Sub

Private Sub LBMonitors_Click()
    Dim i As Long: i = LBMonitors.ListIndex
    Dim c As Long: c = Monitors.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Monitor = Monitors.Item(i)
    UpdateViewMonitor m_Monitor
End Sub

Private Sub List2_Click()
    If m_Monitor Is Nothing Then Exit Sub
    If m_Monitor.Displays Is Nothing Then Exit Sub
    Dim i As Long: i = List2.ListIndex
    Dim c As Long: c = m_Monitor.Displays.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Device = m_Monitor.Displays.Item(i + 1)
    Set m_Filter = MNew.DisplayModeFilter(m_Device.Settings)
    UpdateViewFilter m_Filter
    UpdateViewDisplayDevice m_Device
End Sub

Private Sub List5_Click()
    If m_Monitor Is Nothing Then Exit Sub
    If m_Monitor.Adapters Is Nothing Then Exit Sub
    Dim i As Long: i = List5.ListIndex
    Dim c As Long: c = m_Monitor.Adapters.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Device = m_Monitor.Adapters.Item(i + 1)
    If Not m_Device.Settings Is Nothing Then
        If m_Device.Settings.Count Then
            Set m_Filter = MNew.DisplayModeFilter(m_Device.Settings)
        Else
            If Not m_Device.SettingsGraphicsCard Is Nothing Then
                If m_Device.SettingsGraphicsCard.Count Then
                    Set m_Filter = MNew.DisplayModeFilter(m_Device.SettingsGraphicsCard)
                End If
            End If
        End If
    ElseIf Not m_Device.SettingsGraphicsCard Is Nothing Then
        If m_Device.SettingsGraphicsCard.Count Then
            Set m_Filter = MNew.DisplayModeFilter(m_Device.SettingsGraphicsCard)
        End If
    End If
    UpdateViewFilter m_Filter
    UpdateViewDisplayDevice m_Device
End Sub

Private Sub List3_Click()
    If m_Device Is Nothing Then Exit Sub
    If m_Device.Settings Is Nothing Then Exit Sub
    Dim i As Long: i = List3.ListIndex
    Dim c As Long: c = m_Device.Settings.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Mode = m_Device.Settings.Item(i + 1)
    UpdateViewDeviceMode m_Mode
End Sub

Private Sub List4_Click()
    If m_Device Is Nothing Then Exit Sub
    If m_Device.SettingsGraphicsCard Is Nothing Then Exit Sub
    Dim i As Long: i = List4.ListIndex
    Dim c As Long: c = m_Device.SettingsGraphicsCard.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Mode = m_Device.SettingsGraphicsCard.Item(i + 1)
    UpdateViewDeviceMode2 m_Mode
End Sub

Sub UpdateViewMonitor(Monitor As Monitor)
    TBMonitor.Text = vbNullString
    Text2.Text = vbNullString
    Text3.Text = vbNullString
    List2.Clear
    List5.Clear
    If Monitor Is Nothing Then Exit Sub
    TBMonitor.Text = Monitor.ToStr
    Dim i As Long, c As Long, dd As DisplayDevice
    c = Monitor.Displays.Count
    If c = 0 Then Exit Sub
    For i = 1 To c
        Set dd = m_Monitor.Displays.Item(i)
        List2.AddItem dd.DeviceName
    Next
    c = Monitor.Adapters.Count
    If c = 0 Then Exit Sub
    For i = 1 To c
        Set dd = m_Monitor.Adapters.Item(i)
        List5.AddItem dd.DeviceName
    Next
End Sub

Sub UpdateViewDisplayDevice(Display As DisplayDevice)
    Text2.Text = vbNullString
    Text3.Text = vbNullString
    List3.Clear
    List4.Clear
    If Display Is Nothing Then Exit Sub
    Text2.Text = Display.ToStr
    If Display.Settings Is Nothing Then Exit Sub
    Dim i As Long, c As Long, dm As DeviceMode
    c = Display.Settings.Count
    If c > 0 Then
        For i = 1 To c
            Set dm = Display.Settings.Item(i)
            List3.AddItem i
        Next
    End If
    c = Display.SettingsGraphicsCard.Count
    If c > 0 Then
        For i = 1 To c
            Set dm = Display.SettingsGraphicsCard.Item(i)
            List4.AddItem i
        Next
    End If
End Sub

Sub UpdateViewDeviceMode(Mode As DeviceMode)
    Text3.Text = vbNullString
    If Mode Is Nothing Then Exit Sub
    Text3.Text = Mode.ToStr
End Sub

Sub UpdateViewDeviceMode2(Mode As DeviceMode)
    Text4.Text = vbNullString
    If Mode Is Nothing Then Exit Sub
    Text4.Text = Mode.ToStr
End Sub

Sub UpdateViewFilter(Filter As DisplayModeFilter)
    List6.Clear
    List7.Clear
    List8.Clear
    List9.Clear
    If Filter Is Nothing Then Exit Sub
    
    Dim i As Long
    If Not Filter.FilterSize Is Nothing Then
        For i = 1 To Filter.FilterSize.Count
            List6.AddItem Filter.FilterSize.Item(i)
        Next
        'List6.Sorted = True
        Label6.Caption = List6.ListCount
    End If
    If Not Filter.FilterRotation Is Nothing Then
        For i = 1 To Filter.FilterRotation.Count
            List7.AddItem Filter.FilterRotation.Item(i)
        Next
        'List7.Sorted = True
        Label7.Caption = List7.ListCount
    End If
    If Not Filter.FilterBitsPerPixel Is Nothing Then
        For i = 1 To Filter.FilterBitsPerPixel.Count
            List8.AddItem Filter.FilterBitsPerPixel.Item(i)
        Next
        'List8.Sorted = True
        Label8.Caption = List8.ListCount
    End If
    If Not Filter.FilterFrequency Is Nothing Then
        For i = 1 To Filter.FilterFrequency.Count
            List9.AddItem Filter.FilterFrequency.Item(i)
        Next
        'List9.Sorted = True
        Label9.Caption = List9.ListCount
    End If
    
End Sub
