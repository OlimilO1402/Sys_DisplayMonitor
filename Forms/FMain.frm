VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   13455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15375
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "FMain"
   ScaleHeight     =   13455
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List6 
      Appearance      =   0  '2D
      Height          =   1785
      ItemData        =   "FMain.frx":0000
      Left            =   2400
      List            =   "FMain.frx":0002
      TabIndex        =   8
      Top             =   6120
      Width           =   2415
   End
   Begin VB.ListBox List7 
      Appearance      =   0  '2D
      Height          =   1785
      ItemData        =   "FMain.frx":0004
      Left            =   4800
      List            =   "FMain.frx":0006
      TabIndex        =   7
      Top             =   6120
      Width           =   2415
   End
   Begin VB.ListBox List8 
      Appearance      =   0  '2D
      Height          =   1785
      ItemData        =   "FMain.frx":0008
      Left            =   7200
      List            =   "FMain.frx":000A
      TabIndex        =   6
      Top             =   6120
      Width           =   2415
   End
   Begin VB.ListBox List9 
      Appearance      =   0  '2D
      Height          =   1785
      ItemData        =   "FMain.frx":000C
      Left            =   9600
      List            =   "FMain.frx":000E
      TabIndex        =   5
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox TBDispMonitor 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   2400
      Width           =   12975
   End
   Begin VB.TextBox TBDispConnector 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   480
      Width           =   12975
   End
   Begin VB.ListBox LBDispConnectors 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      ItemData        =   "FMain.frx":0010
      Left            =   0
      List            =   "FMain.frx":0017
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Label1"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Label1"
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label LlDispConnectors 
      AutoSize        =   -1  'True
      Caption         =   "DisplayConnectors: "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1725
   End
   Begin VB.Label LlDispAdapter 
      AutoSize        =   -1  'True
      Caption         =   "DisplayAdapter: "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1440
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_GrafxCard As DisplayAdapter
Private m_Connector As DisplayConnector
Private m_Monitor   As Monitor
Private m_Filter    As DisplayModeFilter

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_GrafxCard = MPC.GraphicsCard
    UpdateViewGraCard m_GrafxCard
End Sub

Private Sub LBDispConnectors_Click()
    If m_GrafxCard Is Nothing Then Exit Sub
    Dim i As Long: i = LBDispConnectors.ListIndex
    Dim cons As Collection: Set cons = m_GrafxCard.Connectors
    If cons Is Nothing Then Exit Sub
    Dim c As Long: c = cons.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Connector = cons.Item(i + 1)
    If m_Connector Is Nothing Then Exit Sub
    UpdateViewDispConnector m_Connector
    Set m_Monitor = m_Connector.Monitor
    If m_Monitor Is Nothing Then Exit Sub
    Set m_Filter = MNew.DisplayModeFilter(m_Monitor.Settings)
    UpdateViewFilter m_Filter
    
End Sub

Sub UpdateViewGraCard(GraCard As DisplayAdapter)
    LlDispAdapter.Caption = "DisplayAdapter: " & GraCard.DeviceString
    LlDispConnectors.Caption = "DisplayConnectors: " & GraCard.Connectors.Count
    LBDispConnectors.Clear
    Dim v, dc As DisplayConnector
    For Each v In GraCard.Connectors
        Set dc = v
        LBDispConnectors.AddItem dc.DeviceName
    Next
End Sub
    
Sub UpdateViewDispConnector(Connector As DisplayConnector)
    If Connector Is Nothing Then Exit Sub
    TBDispConnector.Text = Connector.ToStr
    If Connector.Monitor Is Nothing Then TBDispMonitor.Text = vbNullString: Exit Sub
    TBDispMonitor.Text = Connector.Monitor_ToStr
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

