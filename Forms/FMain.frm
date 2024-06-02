VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "DisplayMonitor"
   ClientHeight    =   9375
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
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   9375
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows-Standard
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
      Height          =   4875
      ItemData        =   "FMain.frx":1782
      Left            =   0
      List            =   "FMain.frx":1789
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   12480
      TabIndex        =   14
      Top             =   6360
      Width           =   2535
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   9960
      TabIndex        =   13
      Top             =   6360
      Width           =   2535
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   7440
      TabIndex        =   12
      Top             =   6360
      Width           =   2535
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   4920
      TabIndex        =   11
      Top             =   6360
      Width           =   2535
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  '2D
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox TBDispMonitor 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   2400
      Width           =   12975
   End
   Begin VB.TextBox TBDispConnector 
      Appearance      =   0  '2D
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
   Begin VB.Label Label5 
      Caption         =   "PixelsWidth "
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "PixelsHeight "
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "DisplayRotation "
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "BitsPerPixel"
      Height          =   375
      Left            =   9960
      TabIndex        =   6
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "DisplayFrequency"
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   6000
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
    Dim col As Collection: Set col = m_Monitor.Settings
    MPtr.Col_Sort col
    Set m_Filter = MNew.DisplayModeFilter(col) 'm_Monitor.Settings)
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
    Combo5.Clear
    Combo6.Clear
    Combo7.Clear
    Combo8.Clear
    Combo9.Clear
    If Filter Is Nothing Then Exit Sub
    
    Dim i As Long
    If Not Filter.FilterWMax Is Nothing Then
        Col_ToListControl Filter.FilterWMax, Combo5, True
        'For i = 1 To Filter.FilterWMax.Count
        '    Combo5.AddItem Filter.FilterWMax.Item(i)
        'Next
        'List6.Sorted = True
        Label5.Caption = Combo5.ListCount
    End If
    If Not Filter.FilterHMin Is Nothing Then
        Col_ToListControl Filter.FilterHMin, Combo6, True
        'For i = 1 To Filter.FilterHMin.Count
        '    Combo6.AddItem Filter.FilterHMin.Item(i)
        'Next
        'List6.Sorted = True
        Label6.Caption = Combo6.ListCount
    End If
    If Not Filter.FilterRotation Is Nothing Then
        Col_ToListControl Filter.FilterRotation, Combo7, True
        'For i = 1 To Filter.FilterRotation.Count
        '    Combo7.AddItem Filter.FilterRotation.Item(i)
        'Next
        'List7.Sorted = True
        Label7.Caption = Combo7.ListCount
    End If
    If Not Filter.FilterBitsPerPixel Is Nothing Then
        Col_ToListControl Filter.FilterBitsPerPixel, Combo8, True
        'For i = 1 To Filter.FilterBitsPerPixel.Count
        '    Combo8.AddItem Filter.FilterBitsPerPixel.Item(i)
        'Next
        'List8.Sorted = True
        Label8.Caption = Combo8.ListCount
    End If
    If Not Filter.FilterFrequency Is Nothing Then
        Col_ToListControl Filter.FilterFrequency, Combo9, True
        'For i = 1 To Filter.FilterFrequency.Count
        '    Combo9.AddItem Filter.FilterFrequency.Item(i)
        'Next
        'List9.Sorted = True
        Label9.Caption = Combo9.ListCount
    End If
    
End Sub

