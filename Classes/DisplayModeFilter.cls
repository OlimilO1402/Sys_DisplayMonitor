VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayModeFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_DeviceModes As Collection 'Of DeviceMode
Private m_DMFiltered  As Collection 'Of DeviceMode
Private m_FilterWMax  As String
Private m_FilterHMin  As String
Private m_FilterRota  As String
Private m_FilterBpPx  As String
Private m_FilterFreq  As String
Private m_AllWMax     As Collection 'Of Long, String
Private m_AllHMin     As Collection 'Of Long, String
Private m_AllRota     As Collection 'Of Long, String
Private m_AllBpPx     As Collection 'Of Long, String
Private m_AllFreq     As Collection 'Of Long, String

Private Sub Class_Initialize()
    Set m_AllWMax = New Collection ' Of Long, String
    Set m_AllHMin = New Collection ' Of Long, String
    Set m_AllRota = New Collection ' Of Long, String
    Set m_AllBpPx = New Collection ' Of Long, String
    Set m_AllFreq = New Collection ' Of Long, String
End Sub

Friend Sub New_(List As Collection)  'Of DeviceMode
    Set m_DeviceModes = List
    Dim v, dm As DeviceMode
    Dim lkey As Long, skey As String
    For Each v In m_DeviceModes
        Set dm = v
        lkey = dm.WHMax:              skey = GetKeyWMax(dm): If Not AllWMax_Contains(skey) Then m_AllWMax.Add lkey, skey
        lkey = dm.WHMin:              skey = GetKeyHMin(dm): If Not AllHMin_Contains(skey) Then m_AllHMin.Add lkey, skey
        lkey = dm.DisplayOrientation: skey = GetKeyRota(dm): If Not AllRota_Contains(skey) Then m_AllRota.Add lkey, skey
        lkey = dm.BitsPerPixel:       skey = GetKeyBpPx(dm): If Not AllBpPx_Contains(skey) Then m_AllBpPx.Add lkey, skey
        lkey = dm.DisplayFrequency:   skey = GetKeyFreq(dm): If Not AllFreq_Contains(skey) Then m_AllFreq.Add lkey, skey
    Next
End Sub

Private Function AllWMax_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_AllWMax(key)) Then: 'DoNothing
    AllWMax_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function AllHMin_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_AllHMin(key)) Then: 'DoNothing
    AllHMin_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function AllRota_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_AllRota(key)) Then: 'DoNothing
    AllRota_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function AllBpPx_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_AllBpPx(key)) Then: 'DoNothing
    AllBpPx_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function AllFreq_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_AllFreq(key)) Then: 'DoNothing
    AllFreq_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GetKeyWMax(dm As DeviceMode) As String
    GetKeyWMax = CStr(dm.WHMax)
End Function

Private Function GetKeyHMin(dm As DeviceMode) As String
    GetKeyHMin = CStr(dm.WHMin)
End Function

Private Function GetKeyRota(dm As DeviceMode) As String
    GetKeyRota = dm.DisplayOrientation_ToStr
End Function

Private Function GetKeyBpPx(dm As DeviceMode) As String
    GetKeyBpPx = dm.BitsPerPixel ' & " Bit"
End Function

Private Function GetKeyFreq(dm As DeviceMode) As String
    GetKeyFreq = dm.DisplayFrequency '& " Hz"
End Function

Public Property Let AllFilter(WMax, HMin, Rota, BpPx, Freq)
    m_FilterWMax = IIf(IsMissing(WMax), vbNullString, WMax)
    m_FilterHMin = IIf(IsMissing(HMin), vbNullString, WMax)
    m_FilterRota = IIf(IsMissing(Rota), vbNullString, Rota)
    m_FilterBpPx = IIf(IsMissing(BpPx), vbNullString, BpPx)
    m_FilterFreq = IIf(IsMissing(Freq), vbNullString, Freq)
    ApplyFilter
End Property

Private Sub ApplyFilter()
    Set m_DMFiltered = New Collection
    Dim dm As DeviceMode
    Dim i As Long, c As Long: c = m_DeviceModes.Count
    Dim fWMax As Long: If Len(m_FilterWMax) Then fWMax = CLng(m_FilterWMax)
    Dim fHMin As Long: If Len(m_FilterHMin) Then fHMin = CLng(m_FilterHMin)
    Dim fRota As Long: If Len(m_FilterRota) Then fRota = CLng(m_FilterRota)
    Dim fBpPx As Long: If Len(m_FilterBpPx) Then fBpPx = CLng(m_FilterBpPx)
    Dim fFreq As Long: If Len(m_FilterFreq) Then fFreq = CLng(m_FilterFreq)
    Dim bAdd As Boolean
            
    For i = 1 To c
        Set dm = m_DeviceModes.Item(i)
        If Len(m_FilterWMax) Then
            If dm.WHMax = fWMax Then bAdd = True
        End If
        If bAdd And Len(m_FilterHMin) Then
            If dm.WHMin = fHMin Then bAdd = True Else bAdd = False
        End If
        If bAdd And Len(m_FilterRota) Then
            If dm.DisplayOrientation = fRota Then bAdd = True Else bAdd = False
        End If
        If bAdd And Len(m_FilterBpPx) Then
            If dm.BitsPerPixel = fBpPx Then bAdd = True Else bAdd = False
        End If
        If bAdd And Len(m_FilterFreq) Then
            If dm.DisplayFrequency = fFreq Then bAdd = True Else bAdd = False
        End If
        If bAdd Then
            m_DMFiltered.Add dm
        End If
    Next
End Sub

Public Property Get FilterWMax() As String
    FilterWMax = m_FilterWMax
End Property
Public Property Let FilterWMax(ByVal Value As String)
    m_FilterWMax = Value
End Property

Public Property Get FilterHMin() As String
    FilterHMin = m_FilterHMin
End Property
Public Property Let FilterHMin(ByVal Value As String)
    m_FilterHMin = Value
End Property

Public Property Get FilterRota() As String
    FilterRota = m_FilterRota
End Property
Public Property Let FilterRota(ByVal Value As String)
    m_FilterRota = Value
End Property

Public Property Get FilterBpPx() As String
    FilterBpPx = m_FilterBpPx
End Property
Public Property Let FilterBpPx(ByVal Value As String)
    m_FilterBpPx = Value
End Property

Public Property Get FilterFreq() As String
    FilterFreq = m_FilterFreq
End Property
Public Property Let FilterFreq(ByVal Value As String)
    m_FilterFreq = Value
End Property


Public Property Get AllWMax() As Collection
    'MPtr.Col_Sort m_FilterSize
    Set AllWMax = m_AllWMax
End Property

Public Property Get AllHMin() As Collection
    'MPtr.Col_Sort m_FilterSize
    Set AllHMin = m_AllHMin
End Property

Public Property Get AllRota() As Collection
    'MPtr.Col_Sort m_FltRotatio
    Set AllRota = m_AllRota
End Property

Public Property Get AllBpPx() As Collection
    'MPtr.Col_Sort m_FilterBpPx
    Set AllBpPx = m_AllBpPx
End Property

Public Property Get AllFreq() As Collection
    'MPtr.Col_Sort m_FilterFreq
    Set AllFreq = m_AllFreq
End Property
