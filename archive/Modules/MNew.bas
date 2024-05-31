Attribute VB_Name = "MNew"
Option Explicit

Public Function DisplayDevice(Optional ByVal DevName As String = vbNullString, Optional ByRef i_inout As Long = 0) As DisplayDevice
    Set DisplayDevice = New DisplayDevice: DisplayDevice.New_ DevName, i_inout
End Function

Public Function DisplayConnector(ByVal DevName As String, ByRef i_inout As Long) As DisplayConnector
    Set DisplayConnector = New DisplayConnector: DisplayConnector.New_ DevName, i_inout
End Function


Public Function Monitor(ByVal HMONITOR As LongPtr, ByVal hDC As LongPtr, ByVal lpClipRECT As LongPtr) As Monitor
    Set Monitor = New Monitor: Monitor.New_ HMONITOR, hDC, lpClipRECT
End Function

Public Function MonitorP(ByVal X As Long, ByVal Y As Long, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest) As Monitor
    Set MonitorP = New Monitor: MonitorP.NewP X, Y, aFlag
End Function

Public Function MonitorR(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest) As Monitor
    Set MonitorR = New Monitor: MonitorR.NewR X, Y, Width, Height, aFlag
End Function

Public Function MonitorW(ByVal hWnd As LongPtr, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest) As Monitor
    Set MonitorW = New Monitor: MonitorW.NewW hWnd, aFlag
End Function

'
'Public Function DisplayDevices(ByVal DevName As String) As DisplayDevices
'    Set DisplayDevices = New DisplayDevices: DisplayDevices.New_ DevName
'End Function
'
'Public Function DeviceMode(ByVal DeviceName As String) As DeviceMode
'    Set DeviceMode = New DeviceMode: DeviceMode.New_ DeviceName
'End Function

Public Function DisplayModeFilter(ListOfDisplayModes As Collection) As DisplayModeFilter
    Set DisplayModeFilter = New DisplayModeFilter: DisplayModeFilter.New_ ListOfDisplayModes
End Function
