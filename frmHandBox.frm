VERSION 5.00
Begin VB.Form frmHandBox 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Skybadger"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   1980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox MSComm1 
      Height          =   480
      Left            =   1440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   14
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdTraffic 
      Caption         =   "Traffic"
      Height          =   330
      Left            =   1080
      TabIndex        =   13
      ToolTipText     =   "Display debugging traffic"
      Top             =   4140
      Width           =   825
   End
   Begin VB.CommandButton cmdPark 
      Caption         =   "Park"
      Height          =   450
      Left            =   1080
      TabIndex        =   12
      ToolTipText     =   "Move Dome to Park"
      Top             =   3540
      Width           =   825
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Sync:"
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Enter Azimuth to Synchronise Dome to"
      Top             =   1800
      Width           =   765
   End
   Begin VB.CommandButton cmdSlew 
      Caption         =   "step"
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   10
      ToolTipText     =   "Step Anti-Clockwise"
      Top             =   2880
      Width           =   540
   End
   Begin VB.CommandButton cmdSlew 
      Caption         =   "step"
      Height          =   255
      Index           =   4
      Left            =   1275
      TabIndex        =   9
      ToolTipText     =   "Step Clockwise"
      Top             =   2880
      Width           =   540
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Goto:"
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Enter Azimuth to move to "
      Top             =   1440
      Width           =   765
   End
   Begin VB.TextBox txtNewAz 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Width           =   810
   End
   Begin VB.CommandButton cmdConnectDome 
      Caption         =   "Connect Dome"
      Height          =   450
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Connect to Dome controllers"
      Top             =   3540
      Width           =   825
   End
   Begin VB.Timer timer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   360
      Top             =   480
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup"
      Height          =   330
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Setup Dome parameters"
      Top             =   4140
      Width           =   825
   End
   Begin VB.CommandButton cmdSlew 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   5
      Left            =   780
      TabIndex        =   3
      ToolTipText     =   "Halt Everything"
      Top             =   2400
      Width           =   420
   End
   Begin VB.CommandButton cmdSlew 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   165
      TabIndex        =   0
      ToolTipText     =   "Run AntiClockwise"
      Top             =   2325
      Width           =   540
   End
   Begin VB.CommandButton cmdSlew 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   3
      Left            =   1275
      TabIndex        =   1
      ToolTipText     =   "Run Clockwise"
      Top             =   2310
      Width           =   540
   End
   Begin VB.Shape shpPark 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1620
      Shape           =   3  'Circle
      Top             =   3300
      Width           =   255
   End
   Begin VB.Shape ShpError 
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Image imgBrewster 
      Height          =   555
      Left            =   405
      MouseIcon       =   "frmHandBox.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmHandBox.frx":0152
      ToolTipText     =   "Click to go to astro.brewsters.net "
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label txtDomeAz 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "---.-"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   900
      TabIndex        =   5
      Top             =   885
      Width           =   915
   End
   Begin VB.Label lblDomeAz 
      BackColor       =   &H00000000&
      Caption         =   "Dome Az:"
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   240
      TabIndex        =   4
      Top             =   780
      Width           =   570
   End
End
Attribute VB_Name = "frmHandBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
' Copyright © 2000-2002 SPACE.com Inc., New York, NY
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". SPACE.COM, INC. MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'   ==============
'   FRMHANDBOX.FRM
'   ==============
'
' CyberDrive hand box form
'
' Written:  28-Jun-00   Robert B. Denny <rdenny@dc3.com>
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 20-Jun-03 jab     Initial edit
' -----------------------------------------------------------------------------
Option Explicit

Private BtnState As Integer

' ======
' EVENTS
' ======

Private Sub Form_Load()

    BtnState = 0
    FloatWindow Me.hwnd, False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Unload frmSetup
    If Not g_show Is Nothing Then _
        Unload g_show
    DoShutdown
    
End Sub

Private Sub cmdConnectDome_Click()

    If g_bConnected Then
        HW_Park
        HW_Shutdown
    Else
        HW_Init
    End If
    
End Sub

Private Sub cmdGoto_Click(index As Integer)
    Dim Az As Double
    
    Az = INVALID_COORDINATE
    On Error Resume Next
    Az = CDbl(txtNewAz.Text)
    On Error GoTo 0
    If Az < -360 Or Az > 360 Then
        MsgBox "Input value must be between" & _
            vbCrLf & "+/- 360", vbExclamation
        Exit Sub
    End If
    
    Az = AzScale(Az)
    
    If index = 0 Then
        HW_Move Az
    Else
        HW_Sync Az
    End If
    
End Sub

Private Sub cmdPark_Click()

    If g_bAtPark Then
        MsgBox "Already parked.", vbExclamation
        Exit Sub
    End If
        
    If g_dSetPark < -360 Or g_dSetPark > 360 Then
        MsgBox "Park location must be between +/- 360." & vbCrLf & _
            "Click on [Setup] to change it.", vbExclamation
        Exit Sub
    End If

    HW_Park
    
End Sub

Private Sub cmdSetup_Click()

    DoSetup                         ' May change our topmost state
    
End Sub

Private Sub cmdSlew_MouseDown(index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)

    If g_bConnected Then _
        BtnState = index

End Sub

Private Sub cmdTraffic_Click()

    If g_show Is Nothing Then _
        Set g_show = New frmShow
    
    g_show.Caption = "Skybadger ASCOM Traffic"
    g_show.Show
    
End Sub

Private Sub imgBrewster_Click()

    DisplayWebPage "http://www.skybadger.net/"
    
End Sub

Private Sub timer_Timer()

    timer_tick
    
End Sub

' =================
' PUBLIC PROPERTIES
' =================

Public Property Get ButtonState() As Integer

    ButtonState = BtnState
    BtnState = 0
        
End Property

Public Property Let DomeAz(Az As Double)

    If Az = INVALID_COORDINATE Then
        txtDomeAz.Caption = "---.-"
    Else
        Az = AzScale(Az)
        txtDomeAz.Caption = Format$(Az, "000.0")
    End If

End Property

' ==============
' LOCAL ROUTINES
' ==============

Public Sub LabelButtons()
  
    If g_bConnected Then
        If g_dSetPark = INVALID_COORDINATE Then
            cmdPark.Enabled = False
            cmdPark.Caption = "Park"
        Else:
            cmdPark.Enabled = True
            cmdPark.Caption = "Park: " & Format$(g_dSetPark, "000.0") & "°"
        End If
    Else
        cmdPark.Enabled = False
        cmdPark.Caption = "Park"
    End If
    
    cmdGoto(0).Enabled = g_bConnected
    cmdGoto(1).Enabled = g_bConnected
    txtNewAz.Enabled = g_bConnected
            
    If g_bConnected Then
        cmdConnectDome.Caption = "Disconnect Dome"
    Else
        cmdConnectDome.Caption = "Connect Dome"
    End If
        
End Sub

Public Sub RefreshLEDs()

    shpPark.FillColor = IIf(g_bAtPark, &HFF&, &H0&)

End Sub

Public Sub ErrorLED()

    ShpError.FillColor = &HFF&

End Sub

