Attribute VB_Name = "MNew"
Option Explicit

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

