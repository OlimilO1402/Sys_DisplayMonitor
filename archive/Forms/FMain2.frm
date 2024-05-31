VERSION 5.00
Begin VB.Form FMain2 
   Caption         =   "Form1"
   ClientHeight    =   8175
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows-Standard
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
      Height          =   4335
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   3840
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
      TabIndex        =   3
      Top             =   1920
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
      Height          =   1560
      ItemData        =   "FMain2.frx":0000
      Left            =   0
      List            =   "FMain2.frx":0007
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox TBDispAdapter 
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
      TabIndex        =   1
      Top             =   0
      Width           =   12975
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
      TabIndex        =   4
      Top             =   1980
      Width           =   1725
   End
   Begin VB.Label LlDispAdapter 
      AutoSize        =   -1  'True
      Caption         =   "DispAdapter: "
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
      Width           =   1200
   End
End
Attribute VB_Name = "FMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_GraCard   As DisplayAdapter
Private m_Connector As DisplayConnector

Private Sub Form_Load()
    Set m_GraCard = MPC.GraphicsCard
    LlDispAdapter.Caption = "DisplayAdapter: " & vbCrLf & m_GraCard.DeviceString
    'LBDispAdapter.Clear
    'LBDispAdapter.AddItem GraCard.Name
    TBDispAdapter.Text = m_GraCard.ToStr
    
    LlDispConnectors.Caption = "DispConnectors: " & m_GraCard.Connectors.Count
    LBDispConnectors.Clear
    Dim v, dc As DisplayConnector
    For Each v In m_GraCard.Connectors
        Set dc = v
        LBDispConnectors.AddItem dc.DeviceName
    Next
    
End Sub

Private Sub LBDispConnectors_Click()
    If m_GraCard Is Nothing Then Exit Sub
    Dim i As Long: i = LBDispConnectors.ListIndex
    Dim c As Long: c = m_GraCard.Connectors.Count
    If i < 0 Or c < i Then Exit Sub
    Set m_Connector = m_GraCard.Connectors.Item(i + 1)
    TBDispConnector.Text = m_Connector.ToStr
    If m_Connector.Monitor Is Nothing Then TBDispMonitor.Text = vbNullString: Exit Sub
    TBDispMonitor.Text = m_Connector.Monitor_ToStr & vbCrLf & _
                         m_Connector.Monitor.DisplayDevice.ToStr
End Sub
