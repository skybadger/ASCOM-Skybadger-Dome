VERSION 5.00
Begin VB.Form frmSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cyberdrive Setup"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3987.738
   ScaleMode       =   0  'User
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Dome Connection"
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1935
      Begin VB.ComboBox cbDomeCom 
         Height          =   315
         ItemData        =   "frmSetup.frx":0000
         Left            =   1080
         List            =   "frmSetup.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblCom 
         BackColor       =   &H00000000&
         Caption         =   "COM port for Dome:"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   165
         TabIndex        =   8
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3900
      TabIndex        =   0
      Top             =   3060
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   3900
      TabIndex        =   1
      Top             =   3540
      Width           =   930
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Motion Control"
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   2295
      Begin VB.TextBox txtStepSize 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   780
         Width           =   765
      End
      Begin VB.TextBox txtSlewSpeed 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   315
         Width           =   765
      End
      Begin VB.TextBox txtPark 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label lblParkPosition 
         BackColor       =   &H00000000&
         Caption         =   "Park Position:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblSlewSpeed 
         BackColor       =   &H00000000&
         Caption         =   "Slew Speed:"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   165
         TabIndex        =   4
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblStepSize 
         BackColor       =   &H00000000&
         Caption         =   "Step Size:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   840
         Width           =   765
      End
   End
   Begin VB.Label lblLastModified 
      BackColor       =   &H00000000&
      Caption         =   "<last modified>"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Image imgBrewster 
      Height          =   555
      Left            =   1882
      MouseIcon       =   "frmSetup.frx":004C
      MousePointer    =   99  'Custom
      Picture         =   "frmSetup.frx":019E
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblDriverInfo 
      BackColor       =   &H00000000&
      Caption         =   "<version, etc.>"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
   End
End
Attribute VB_Name = "frmSetup"
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
'   ============
'   FRMSETUP.FRM
'   ============
'
' CyberDrive setup form
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

Private m_bResult As Boolean
Private m_bAllowUnload As Boolean



' ======
' EVENTS
' ======

Private Sub Form_Load()
    Dim tzName As String
    Dim l As Long
    Dim fs, F
    Dim DLM As String
    
    
    FloatWindow Me.hwnd, True                       ' Setup window always floats
    m_bAllowUnload = True                           ' Start out allowing unload
    
    lblDriverInfo = App.FileDescription & " Version " & _
        App.Major & "." & App.Minor & "." & App.Revision
        
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFile(App.Path & "\Skybadgerdome.exe")
    DLM = F.DateLastModified
    
    lblLastModified = "Last Modified " & DLM
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Me.Hide                                     ' Assure we don't unload
    Cancel = Not m_bAllowUnload                 ' Unless our flag permits it
    
End Sub

Private Sub cmdCancel_Click()

    m_bResult = False
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    m_bResult = True
    Me.Hide

End Sub

Private Sub imgBrewster_Click()

    DisplayWebPage "http://www.skybadger.net/"
    
End Sub

' =================
' PUBLIC PROPERTIES
' =================

Public Property Let AllowUnload(b As Boolean)

    m_bAllowUnload = b
    
End Property

Public Property Let DomeCom(newVal As Integer)
    Dim index As Integer
    
    If newVal < 0 Then
        index = 0
    ElseIf newVal > 10 Then
        index = 10
    Else
        index = newVal
    End If
    
    cbDomeCom.ListIndex = index
        
End Property

Public Property Get DomeCom() As Integer

    DomeCom = cbDomeCom.ItemData(cbDomeCom.ListIndex)

End Property


Public Property Let Park(newVal As Double)

    If newVal < -360 Or newVal > 360 Then
        txtPark.Text = "000.0"
    Else
        If newVal < 0 Then _
            newVal = newVal + 360
        txtPark.Text = Format$(newVal, "000.0")
    End If
    
End Property

Public Property Get Park() As Double

    Park = 180
    On Error Resume Next
    Park = CDbl(txtPark.Text)
    On Error GoTo 0
    
    Park = AzScale(Park)
    
End Property

Public Property Get result() As Boolean

    result = m_bResult              ' Set by OK or Cancel button
    
End Property

Public Property Let SlewSpeed(newVal As Double)
    
    txtSlewSpeed.Text = Format$(newVal, "0.0")
        
End Property

Public Property Get SlewSpeed() As Double

    ' error check ???
    SlewSpeed = CDbl(txtSlewSpeed.Text)
    If (SlewSpeed < 40) Then SlewSpeed = 40
    If (SlewSpeed > 127) Then SlewSpeed = 127

End Property

Public Property Let StepSize(newVal As Double)
    Dim index As Integer
    
    txtStepSize.Text = Format$(newVal, "0.0")
    
End Property

Public Property Get StepSize() As Double

    StepSize = 1
    On Error Resume Next
    StepSize = CDbl(txtStepSize.Text)
    On Error GoTo 0
    
    If StepSize < 1 Then _
        StepSize = 1
    If StepSize > 90 Then _
        StepSize = 90

End Property

'
' LOCAL UTILITIES
'

Private Sub lblDriverInfo_Click()

End Sub
