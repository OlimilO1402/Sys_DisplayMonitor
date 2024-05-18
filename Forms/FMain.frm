VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "View & Change Display-Settings"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14775
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
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows-Standard
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
      Height          =   5775
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   6
      Top             =   4680
      Width           =   12375
   End
   Begin VB.ListBox List3 
      Height          =   5670
      ItemData        =   "FMain.frx":1782
      Left            =   0
      List            =   "FMain.frx":1784
      TabIndex        =   5
      Top             =   4680
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
      Height          =   2175
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   2520
      Width           =   12375
   End
   Begin VB.ListBox List2 
      Height          =   2100
      ItemData        =   "FMain.frx":1786
      Left            =   0
      List            =   "FMain.frx":1788
      TabIndex        =   3
      Top             =   2520
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
      Width           =   12375
   End
   Begin VB.ListBox List1 
      Height          =   2100
      ItemData        =   "FMain.frx":178A
      Left            =   0
      List            =   "FMain.frx":178C
      TabIndex        =   0
      Top             =   360
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

Private Sub Command1_Click()
    Monitors.Init
    UpdateViewMonitors
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Set m_Monitor = Monitors.Item(i)
    UpdateViewMonitor m_Monitor
End Sub

Private Sub List2_Click()
    Dim i As Long: i = List2.ListIndex
    If i < 0 Then Exit Sub
    Set m_Device = m_Monitor.Displays.Item(i + 1)
    UpdateViewDisplayDevice m_Device
End Sub

Private Sub List3_Click()
    Dim i As Long: i = List3.ListIndex
    If i < 0 Then Exit Sub
    Dim c As Long: c = m_Device.SettingsCurrent.Count
    If i <= c Then
        Set m_Mode = m_Device.SettingsCurrent.Item(i + 1)
    Else
        'i = i - c
        'Set m_Mode = m_Device.SettingsRegistry.Item(i + 1)
    End If
    UpdateViewDeviceMode m_Mode
End Sub

Sub UpdateViewMonitors()
    Dim i As Long, m As Monitor
    For i = 0 To Monitors.Count - 1
        Set m = Monitors.Item(i)
        List1.AddItem m.Name
    Next
End Sub

Sub UpdateViewMonitor(Monitor As Monitor)
    If Monitor Is Nothing Then
        Text1.Text = vbNullString
        Text2.Text = vbNullString
        Text3.Text = vbNullString
        Exit Sub
    End If
    Text1.Text = Monitor.ToStr
    Dim c As Long: c = Monitor.Displays.Count
    If c Then
        List2.Clear
        Dim i As Long, dd As DisplayDevice
        For i = 1 To c
            Set dd = m_Monitor.Displays.Item(i)
            List2.AddItem dd.Name 'ToStr
        Next
    End If
End Sub

Sub UpdateViewDisplayDevice(Display As DisplayDevice)
    If Display Is Nothing Then
        Text2.Text = vbNullString
        Text3.Text = vbNullString
        Exit Sub
    End If
    Text2.Text = Display.ToStr
    Dim c As Long: c = Display.SettingsCurrent.Count
    If c Then
        List3.Clear
        Dim i As Long, dm As DeviceMode
        For i = 1 To c
            Set dm = Display.SettingsCurrent.Item(i)
            'List3.AddItem "ModeCurrent-" & i 'dm.ToStr 'Name
            List3.AddItem i
        Next
'        c = Display.SettingsRegistry.Count
'        For i = 1 To c
'            Set dm = Display.SettingsRegistry.Item(i)
'            List3.AddItem "ModeRegistry-" & i 'dm.ToStr
'        Next
    End If
End Sub

Sub UpdateViewDeviceMode(Mode As DeviceMode)
    If Mode Is Nothing Then
        Text3.Text = vbNullString
        Exit Sub
    End If
    Text3.Text = Mode.ToStr
    'Debug.Print Mode.Size
    
End Sub
