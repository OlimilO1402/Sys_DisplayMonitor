VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "View & Change Display-Settings"
   ClientHeight    =   10455
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
   ScaleHeight     =   10455
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows-Standard
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
      Height          =   5295
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   8
      Top             =   5160
      Width           =   5175
   End
   Begin VB.ListBox List4 
      Height          =   5160
      ItemData        =   "FMain.frx":1782
      Left            =   7560
      List            =   "FMain.frx":1784
      TabIndex        =   7
      Top             =   5160
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
      Height          =   5295
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   6
      Top             =   5160
      Width           =   5175
   End
   Begin VB.ListBox List3 
      Height          =   4905
      ItemData        =   "FMain.frx":1786
      Left            =   0
      List            =   "FMain.frx":1788
      TabIndex        =   5
      Top             =   5160
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
      TabIndex        =   4
      Top             =   2520
      Width           =   12735
   End
   Begin VB.ListBox List5 
      Height          =   1335
      ItemData        =   "FMain.frx":178A
      Left            =   0
      List            =   "FMain.frx":178C
      TabIndex        =   9
      Top             =   3840
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   1335
      ItemData        =   "FMain.frx":178E
      Left            =   0
      List            =   "FMain.frx":1790
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   360
      Width           =   12735
   End
   Begin VB.ListBox List1 
      Height          =   2100
      ItemData        =   "FMain.frx":1792
      Left            =   0
      List            =   "FMain.frx":1794
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Monitors"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Monitor As Monitor
Private m_Device  As DisplayDevice
Private m_Mode    As DeviceMode

Private Sub Form_Load()
    'Ähm so ein Mist meine Verschachtelungstheorie ist Mist
    'wir haben
    'Monitors -> List Of Monitor -> List Of DisplayDevice -> List Of DeviceMode
    'stattdessen sollte es sein:
    'PC -> List Of Monitor -> List Of DisplayDevice
    '                      -> List Of DeviceMode
    'PC -> List Of GraCard -> List Of DisplayDevice
    '                      -> List Of DeviceMode
End Sub

Private Sub Command1_Click()
    Monitors.Init
    UpdateViewMonitors
End Sub

Sub UpdateViewMonitors()
    Dim i As Long, m As Monitor
    For i = 0 To Monitors.Count - 1
        Set m = Monitors.Item(i)
        List1.AddItem m.Name
    Next
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
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
    UpdateViewDisplayDevice m_Device
End Sub

Private Sub List5_Click()
    If m_Monitor Is Nothing Then Exit Sub
    If m_Monitor.Adapters Is Nothing Then Exit Sub
    Dim i As Long: i = List5.ListIndex
    Dim c As Long: c = m_Monitor.Adapters.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Device = m_Monitor.Adapters.Item(i + 1)
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
    Text1.Text = vbNullString
    Text2.Text = vbNullString
    Text3.Text = vbNullString
    List2.Clear
    List5.Clear
    If Monitor Is Nothing Then Exit Sub
    Text1.Text = Monitor.ToStr
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
