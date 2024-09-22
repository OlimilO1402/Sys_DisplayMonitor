VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrReset 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   4440
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Testen"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Zur�cksetzen"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Neue Aufl�sung"
      Height          =   855
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   6015
      Begin VB.Label lblNewResolution 
         Caption         =   "Nicht ausgew�hlt"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.ListBox lstResolutions 
      Height          =   1620
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Unterst�tzte Aufl�sungen"
      Height          =   2175
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Aktuelle Aufl�sung"
      Height          =   855
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      Begin VB.Label lblInfo 
         Caption         =   "Bitte warten..."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Liefert Informationen �ber die Bildschirmaufl�sung(en)
Private Declare Function EnumDisplaySettings Lib "user32" _
        Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName _
        As Long, ByVal iModeNum As Long, lpDevMode As Any) _
        As Boolean
        
'�ndert die Bildschirmeinstellungen
Private Declare Function ChangeDisplaySettings Lib "user32" _
        Alias "ChangeDisplaySettingsA" (lpDevMode As Any, _
        ByVal dwFlags As Long) As Long

Const CCDEVICENAME As Long = 32&
Const CCFORMNAME As Long = 32&

'Flags, die angeben, welche Werte ge�ndert werden sollen
Const DM_BITSPERPEL As Long = &H40000           'Farbtiefe
Const DM_PELSWIDTH As Long = &H80000            'Breite
Const DM_PELSHEIGHT As Long = &H100000          'H�he
Const DM_DISPLAYFREQUENCY As Long = &H400000    'Wiederholfrequenz

Const ENUM_CURRENT_SETTINGS As Long = -1&       'Flag f�r aktuelle Einstellungen

'Flags, die angeben, welcher Modus verwendet werden soll
Const CDS_UPDATEREGISTRY = &H1  '�nderungen in Registry eintragen (�nderung nur f�r aktuellen Benutzer)
Const CDS_TEST = &H2            'Nur testen, nicht �ndern
Const CDS_FULLSCREEN = &H4      'Vollbildanwendung (tempor�re �nderung)
Const CDS_GLOBAL = &H8          'Global (�nderungen f�r alle Nutzer)
Const CDS_RESET = &H40000000    'Neuzuweisung, selbst wenn sich die Aufl�sung nicht ge�ndert hat

'R�ckgabewerte von ChangeDisplaySettings
Private Const DISP_CHANGE_SUCCESSFUL As Long = 0&
Private Const DISP_CHANGE_RESTART    As Long = 1&
Private Const DISP_CHANGE_FAILED     As Long = -1&
Private Const DISP_CHANGE_BADMODE    As Long = -2&
Private Const DISP_CHANGE_NOTUPDATED As Long = -3&
Private Const DISP_CHANGE_BADFLAGS   As Long = -4&
Private Const DISP_CHANGE_BADPARAM   As Long = -5&

'R�ckgabe der aktuellen Einstellungen
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'Enum f�r Funktion "SetScreen"
Private Enum enmChangeMode
    Temporary = 0
    CurrentUser = 1
    Systemwide = 2
End Enum

Dim uOldResolution As DEVMODE   'Beinhaltet die alten Einstellungen (f�r Reset)
Dim bReset As Byte              'Speichert den Wert f�r den Countdown (Reset)
Dim Firststart As Boolean       'Speichert, ob es sich um die erste Ausf�hrung handelt

Private Sub Form_Load()
    Dim Result As Long
    Dim Dev As DEVMODE
    Dim Counter As Integer
    
    Result = -1 'Nicht gesetzt
    Counter = 0
    
    lstResolutions.Clear
    lblInfo.Caption = vbNullString
    
    Do While Result <> 0
        'Werte erhalten
        Result = EnumDisplaySettings(0&, Counter, Dev)
           
        If Not Result = 0 Then
            'In Liste einf�gen
            lstResolutions.AddItem Dev.dmPelsWidth & "x" & Dev.dmPelsHeight & " " & Dev.dmBitsPerPel & " Bit @ " & Dev.dmDisplayFrequency
            
            Counter = Counter + 1
        Else
            'aktuelle Aufl�sung in Erfahrung bringen
            Result = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, Dev)
            
            'Zwischenspeichern, falls Erststart
            If Not Firststart Then
                uOldResolution = Dev
                Firststart = True
            End If
            
            'Anzeigen und ausw�hlen
            lblInfo.Caption = Dev.dmPelsWidth & "x" & Dev.dmPelsHeight & " " & Dev.dmBitsPerPel & " Bit @ " & Dev.dmDisplayFrequency
            lstResolutions.Text = lblInfo.Caption
            
            Exit Do
        End If
        
        DoEvents
    Loop
End Sub

Private Sub lstResolutions_Click()
    'In Label eintragen
    lblNewResolution.Caption = lstResolutions.Text
End Sub

Private Sub cmdReset_Click()
    'Alten Wert zur�ckschreiben
    SetScreen uOldResolution.dmPelsWidth, uOldResolution.dmPelsHeight, uOldResolution.dmDisplayFrequency, _
        uOldResolution.dmBitsPerPel, False, Systemwide
        
    'Timer deaktivieren und zur�cksetzen
    tmrReset.Enabled = False
    bReset = 0
    cmdTest.Caption = "Testen"
End Sub

Private Sub cmdTest_Click()
    Dim uDev As DEVMODE
    
    'Aus dem String die Daten r�ckgewinnen
    uDev = ParseResolutionInfo(lblNewResolution.Caption)
    
    If Not cmdTest.Caption = "Testen" Then  'Falls nicht getestet wird -> Einstellungen beibehalten
        tmrReset.Enabled = False
        cmdTest.Caption = "Testen"
        bReset = 0
    Else    'Falls zuerst getestet werden soll
        If SetScreen(uDev.dmPelsWidth, uDev.dmPelsHeight, uDev.dmDisplayFrequency, uDev.dmBitsPerPel, True) = True Then
            'Falls der Test erfolgreich war, so kann die Einstellung nun ge�ndert werden. Fehler sollten dabei keine auftreten
            Call SetScreen(uDev.dmPelsWidth, uDev.dmPelsHeight, uDev.dmDisplayFrequency, uDev.dmBitsPerPel, False, Temporary)
            
            'Reset-Timer aktivieren
            tmrReset.Enabled = True
        End If
    End If
End Sub

Private Sub tmrReset_Timer()
    If bReset = 10 Then
        'Zur�cksetzen
        cmdTest.Caption = "Testen"
        Call cmdReset_Click
        
        tmrReset.Enabled = False
        bReset = 0
    Else
        'Countdown anzeigen
        cmdTest.Caption = "�bernehmen (" & 10 - bReset & ")"
        bReset = bReset + 1
    End If
End Sub

' Funktion SetScreen
'
' Erwartet:
'   X           Neue Aufl�sung in der Horizontalen
'   Y           Neue Aufl�sung in der Vertikalen
'   NewFreq     Neue Bildwiederholfrequenz
'   ColorDepth  Die neue Farbtiefe
'   IsTestMode  Gibt an, ob nur getestet werden soll
'   ChangeMode  Gibt an, auf welche Art die Aufl�sung ge�ndert werden soll
'
'               Parameter:
'
'               Temporary
'                   Es handelt sich um eine Vollbildanwendung. Die Aufl�sung wird automatisch zur�ckgestellt,
'                   sobald die Anwendung beendet wurde
'
'               Current User
'                   Die �nderungen sollen nur f�r den aktuell angemeldeten Benutzer �bernommen werden.
'                   Andere Benutzerkonten sind von dieser �nderung nicht betroffen.
'
'               Systemwide
'                   Die �nderungen werden Global f�r alle Nutzer dauerhaft �bernommen
'
' R�ckgabe:
'   - True, falls die Einstellungen ge�ndert wurden
'   - False, falls die Einstellungen nicht ge�ndert wurden
'     Es wird ein betreffender Hinweis ausgegeben
Private Function SetScreen(ByVal X As Long, ByVal Y As Long, ByVal NewFreq As Byte, ByVal ColorDepth As Byte, _
    ByVal IsTestMode As Boolean, Optional ByVal ChangeMode As enmChangeMode) As Boolean
    
    Dim Result As Long
    Dim Dev As DEVMODE
    Dim NewFlags As Long
    
    Result = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, Dev)    'Aktuellen Einstellungen erfahren
    If Result <> 0 Then
        'Neue Werte festlegen
        Dev.dmDisplayFrequency = NewFreq
        Dev.dmBitsPerPel = ColorDepth
        Dev.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_DISPLAYFREQUENCY Or DM_BITSPERPEL    'siehe Deklarationsteil
        Dev.dmPelsWidth = X
        Dev.dmPelsHeight = Y
   
        Result = ChangeDisplaySettings(Dev, CDS_TEST)   'Neuen Grafikmodus testen
        
        'Ergebnis auswerten
        If Result = DISP_CHANGE_FAILED Or Result = DISP_CHANGE_BADMODE Then
            'Hardwarefehler
            MsgBox "Der ausgw�hlte Modus wird nicht unterst�tzt!", vbExclamation, "Fehlgeschlagen"
        ElseIf Result = DISP_CHANGE_BADFLAGS Or Result = DISP_CHANGE_BADPARAM Then
            'Softwarefehler
            MsgBox "Ung�ltige Parameter", vbExclamation, "Fehler"
        ElseIf Result = DISP_CHANGE_SUCCESSFUL Then
            'Erfolgreich getestet. Falls nicht im Testmodus, fortsetzen
            If Not IsTestMode Then
                'Neue Flags zusammensetzen
                Select Case ChangeMode
                    Case 0:
                        NewFlags = CDS_RESET Or CDS_FULLSCREEN
                    Case 1:
                        NewFlags = CDS_RESET Or CDS_UPDATEREGISTRY
                    Case 2:
                        NewFlags = CDS_RESET Or CDS_GLOBAL Or CDS_UPDATEREGISTRY
                End Select
                        
                'Neue Werte zuweisen
                Result = ChangeDisplaySettings(Dev, NewFlags)
                
                'Ergebnis auswerten
                If Result = DISP_CHANGE_RESTART Then
                    MsgBox "Es ist ein Neustart erforderlich!", vbInformation, "Neustart"
                    
                    Exit Function
                ElseIf Result = DISP_CHANGE_NOTUPDATED Then
                    MsgBox "Fehler beim Schreiben in die Registry!", vbExclamation, "Fehler"
                    
                    Exit Function
                ElseIf Not Result = 0 Then
                    MsgBox "Es trat ein unbekannter Fehler auf!"
                    
                    Exit Function
                End If
                
                Call Form_Load
            End If
            
            SetScreen = True
        Else
            MsgBox "Unbekannter Fehler!", vbCritical, "Fehler"
            Exit Function
        End If
    End If
End Function

'Zerlegt den String (z.B. 1024x786 32 Bit @ 60 Hz) in die einzelnen Werte
Private Function ParseResolutionInfo(ByVal strResolution As String) As DEVMODE
    Dim Result As DEVMODE
    Dim strarr1() As String, strarr2() As String
    
    strarr1 = Split(strResolution, " ")
    strarr2 = Split(strarr1(0), "x")
        
    Result.dmPelsWidth = strarr2(0)
    Result.dmPelsHeight = strarr2(1)
    Result.dmBitsPerPel = strarr1(1)
    Result.dmDisplayFrequency = strarr1(4)
    
    ParseResolutionInfo = Result
End Function

