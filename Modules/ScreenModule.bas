Attribute VB_Name = "ScreenModule"
Option Explicit

'
' (c)2002 Roeland Kluit
' y2kfixx@hotmail.com
'
' Used various code to build this!
'

Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH  As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const CDS_FORCE     As Long = &H80000000
Private Const HORZRES       As Long = 8
Private Const VERTRES       As Long = 10
Private Const BITSPIXEL     As Long = 12
Private Const VREFRESH      As Long = 116

Private Type DEVMODE
    dmDeviceName    As String * 32
    dmSpecVersion   As Integer
    dmDriverVersion As Integer
    dmSize          As Integer
    dmDriverExtra   As Integer
    dmFields        As Long
    dmOrientation   As Integer
    dmPaperSize     As Integer
    dmPaperLength   As Integer
    dmPaperWidth    As Integer
    dmScale         As Integer
    dmCopies        As Integer
    dmDefaultSource As Integer
    dmPrintQuality  As Integer
    dmColor         As Integer
    dmDuplex        As Integer
    dmYResolution   As Integer
    dmTTOption      As Integer
    dmCollate       As Integer
    dmFormName      As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel    As Long
    dmPelsWidth     As Long
    dmPelsHeight    As Long
    dmDisplayFlags  As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal modeIndex As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private lpDevMode() As DEVMODE
Private CurrentIndex As Long

Public Sub GetDisplaySettings(displayDescr() As String)
    Dim index As Long
    Dim displayCount As Long
    Dim colorDescr As String
    
    ' set the DEVMODE flags and structure size
    
    ReDim lpDevMode(0 To 1) As DEVMODE
    lpDevMode(0).dmSize = Len(lpDevMode(0))
    lpDevMode(0).dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    
    ' count how many display settings are there
    Do While EnumDisplaySettings(0, displayCount, lpDevMode(0)) > 0
        displayCount = displayCount + 1
    Loop
    
    ' now displayCount holds the number of display settings
    ' and we can DIMension the result arrays
    ReDim displayDescr(0 To displayCount) As String
    ReDim lpDevMode(0 To displayCount) As DEVMODE
    
    For index = 0 To displayCount
        ' retrieve info on the index-th display mode
        EnumDisplaySettings 0, index, lpDevMode(index)
        
        Select Case lpDevMode(index).dmBitsPerPel
            Case 4
                colorDescr = "16 colors"
            Case 8
                colorDescr = "256 colors"
            Case Is <= 24
                colorDescr = "24bit color"
            Case Else
                colorDescr = "32bit color"
        End Select
        
        displayDescr(index) = lpDevMode(index).dmPelsWidth & " x " & lpDevMode(index).dmPelsHeight & ", " & colorDescr
        If lpDevMode(index).dmDisplayFrequency > 1 Then
            displayDescr(index) = displayDescr(index) & ", " & lpDevMode(index).dmDisplayFrequency & " Hz"
        Else
            displayDescr(index) = displayDescr(index) & ", (Hardware default)"
        End If
    Next

End Sub

Public Function ChangeScreenResolution(ByRef index As Long) As Boolean
    If ChangeDisplaySettings(lpDevMode(index), CDS_FORCE) = 0 Then _
        ChangeScreenResolution = True
End Function

Public Function lookupCurrent() As Long

   Dim currHRes  As Long
   Dim currVRes  As Long
   Dim currBPP   As Long
   Dim currVFreq As Long
   Dim sBPPtype  As String
   Dim sFreqtype As String
   Dim hDC       As LongPtr
   Dim i         As Long
   
   lookupCurrent = -1
   
   hDC = GetDC(0)
   
   'get the system settings
  
   currHRes  = GetDeviceCaps(hDC, HORZRES)
   currVRes  = GetDeviceCaps(hDC, VERTRES)
   currBPP   = GetDeviceCaps(hDC, BITSPIXEL)
   currVFreq = GetDeviceCaps(hDC, VREFRESH)
   
   Call DeleteDC(hDC)
   
   For i = 0 To UBound(lpDevMode) - 1
   
    If lpDevMode(i).dmPelsWidth = currHRes Then
        If (lpDevMode(i).dmPelsHeight = currVRes) Then
            If (lpDevMode(i).dmBitsPerPel = currBPP) Then
                If (lpDevMode(i).dmDisplayFrequency = currVFreq) Then
                    lookupCurrent = i
                    Exit Function
                End If
            End If
        End If
    End If
   Next
      
End Function

Public Function FillList(List As Object)
    Dim stringList() As String
    Dim i As Long

    List.Clear

    Call GetDisplaySettings(stringList)

    For i = 0 To UBound(stringList) - 1
       List.AddItem stringList(i)
    Next 

    CurrentIndex = lookupCurrent()
    If CurrentIndex <> -1 Then
        List.ListIndex = CurrentIndex
    Else
        MsgBox "Error: Could not read current settings!", vbCritical
    End If

End Function

Public Property Get CurrentResolution() As Long
    CurrentResolution = CurrentIndex
End Property

Public Property Let CurrentResolution(ByVal vNewValue As Long)
    MsgBox "Your screen settings will change, please hit escape if your system is unable to display the new setting.", vbInformation
    If ChangeScreenResolution(vNewValue) Then
        If Not (MsgBox("Keep current setting?", vbYesNoCancel + vbQuestion) = vbYes) Then
            Call ChangeScreenResolution(CurrentIndex)
        Else
            CurrentIndex = vNewValue
        End If
    End If
End Property
