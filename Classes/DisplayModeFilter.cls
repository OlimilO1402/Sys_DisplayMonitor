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
Private m_FilterSize As Collection '
Private m_FltRotatio As Collection
Private m_FilterBpPx As Collection
Private m_FilterFreq As Collection

Private Sub Class_Initialize()
    Set m_FilterSize = New Collection 'Of String
    Set m_FltRotatio = New Collection 'Of String
    Set m_FilterBpPx = New Collection 'Of String
    Set m_FilterFreq = New Collection 'Of String
End Sub

Friend Sub New_(List As Collection)  'Of DeviceMedia
    Dim v, dm As DeviceMode
    Dim lkey As Long, skey As String
    For Each v In List
        Set dm = v
        key = GetKeySize(dm): If Not FltSize_Contains(key) Then m_FilterSize.Add key, key
        key = GetKeyRota(dm): If Not FltRota_Contains(key) Then m_FltRotatio.Add key, key
        key = GetKeyBpPx(dm): If Not FltBpPx_Contains(key) Then m_FilterBpPx.Add key, key
        key = GetKeyFreq(dm): If Not FltFreq_Contains(key) Then m_FilterFreq.Add key, key
    Next
End Sub

Private Function FltSize_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_FilterSize(key)) Then: 'DoNothing
    FltSize_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function FltRota_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_FltRotatio(key)) Then: 'DoNothing
    FltRota_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function FltBpPx_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_FilterBpPx(key)) Then: 'DoNothing
    FltBpPx_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function FltFreq_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_FilterFreq(key)) Then: 'DoNothing
    FltFreq_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GetKeySize(dm As DeviceMode) As String
    Dim w As Long: w = dm.PixelsWidth
    Dim h As Long: h = dm.PixelsHeight
    Dim t As Long: t = w
    w = Max(w, h)
    h = Min(t, h)
    GetKeySize = CStr(w) & " x " & CStr(h)
End Function

Private Function GetKeyRota(dm As DeviceMode) As String
    GetKeyRota = dm.DisplayOrientation_ToStr
End Function

Private Function GetKeyBpPx(dm As DeviceMode) As String
    GetKeyBpPx = dm.BitsPerPixel & " Bit"
End Function

Private Function GetKeyFreq(dm As DeviceMode) As String
    GetKeyFreq = dm.DisplayFrequency & " Hz"
End Function

Public Property Get FilterSize() As Collection
    Set FilterSize = m_FilterSize
End Property

Public Property Get FilterRotation() As Collection
    Set FilterRotation = m_FltRotatio
End Property

Public Property Get FilterBitsPerPixel() As Collection
    Set FilterBitsPerPixel = m_FilterBpPx
End Property

Public Property Get FilterFrequency() As Collection
    Set FilterFrequency = m_FilterFreq
End Property
