Attribute VB_Name = "Monitors"
Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type


'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaymonitors
'BOOL EnumDisplayMonitors([in] HDC hdc, [in] LPCRECT lprcClip, [in] MONITORENUMPROC lpfnEnum, [in] LPARAM dwData);
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As LongPtr, ByVal lprcClip As LongPtr, ByVal lpfnEnum As LongPtr, ByVal dwData As LongPtr) As Long
'[in] hdc:      Ein Handle für einen Anzeigegerätekontext, der den sichtbaren Bereich von Interesse definiert. Wenn dieser Parameter NULL ist, ist der hdcMonitor-Parameter, der
'               an die Rückruffunktion übergeben wird, NULL, und der sichtbare Bereich von Interesse ist der virtuelle Bildschirm, der alle Anzeigen auf dem Desktop umfasst.
'[in] lprcClip: Ein Zeiger auf eine RECT-Struktur , die ein Abschneiderechteck angibt. Der bereich von Interesse ist die Schnittmenge des Abschneiderechtecks mit dem sichtbaren Bereich,
'               der von hdc angegeben wird. Wenn hdc nicht NULL ist, sind die Koordinaten des Abschneiderechtecks relativ zum Ursprung des hdc. Wenn hdcNULL ist, sind die Koordinaten
'               virtuelle Bildschirmkoordinaten. Dieser Parameter kann NULL sein, wenn Sie die von hdc angegebene Region nicht ausschneiden möchten.
'[in] lpfnEnum: Ein Zeiger auf eine anwendungsdefinierte Rückruffunktion von MonitorEnumProc .
'[In] dwData:   Anwendungsdefinierte Daten, die EnumDisplayMonitors direkt an die MonitorEnumProc-Funktion übergibt.

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nc-winuser-monitorenumproc
'MONITORENUMPROC Monitorenumproc;
'BOOL Monitorenumproc( HMONITOR unnamedParam1, HDC unnamedParam2, LPRECT unnamedParam3, LPARAM unnamedParam4) {...}

Private m_hDC  As LongPtr
Private m_Clip As RECT
Private m_Monitors As Collection
'Private m_Count As Long

Public Sub Init()
    'm_Count = 10
    Set m_Monitors = New Collection
    Dim lprcClip As LongPtr
    Dim clip As Boolean
    If clip Then lprcClip = VarPtr(m_Clip)
    Dim hr As Long: hr = EnumDisplayMonitors(m_hDC, lprcClip, FncPtr(AddressOf MonitorEnumProc), 0)
    'Debug.Print hr
End Sub

Public Property Get hDC() As LongPtr
    hDC = m_hDC
End Property
Public Property Let hDC(ByVal Value As LongPtr)
    m_hDC = Value
End Property

Public Function Trim0(ByVal s As String) As String
    Dim p0 As Long: p0 = InStr(1, s, vbNullChar)
    If p0 Then Trim0 = Left(s, p0 - 1)
End Function

Private Function MonitorEnumProc(ByVal HMonitor_Param1 As LongPtr, ByVal hDC_Param2 As LongPtr, ByVal lpRECT_Param3 As LongPtr, ByVal Param4 As LongPtr) As Long
    If m_Monitors Is Nothing Then Set m_Monitors = New Collection
    Dim aMonitor As Monitor: Set aMonitor = MNew.Monitor(HMonitor_Param1, hDC_Param2, lpRECT_Param3)
    If aMonitor.Handle <> 0 Then
        m_Monitors.Add aMonitor
        MonitorEnumProc = 1
        Exit Function
    End If
End Function

Public Property Get Count() As Long
    Count = m_Monitors.Count
End Property

Public Property Get Item(ByVal Index As Long) As Monitor
    If Index < Count Then Set Item = m_Monitors.Item(Index + 1)
End Property

Public Property Get ItemByHandle(ByVal HMONITOR As LongPtr) As Monitor
    Dim m As Monitor
    For Each m In m_Monitors
        If m.Handle = HMONITOR Then
            Set ItemByHandle = m
            Exit Property
        End If
    Next
End Property
