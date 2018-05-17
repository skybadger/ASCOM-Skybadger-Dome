VERSION 5.00
Begin VB.Form frmShow 
   BackColor       =   &H00000000&
   Caption         =   "ASCOM Traffic"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShow.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6615
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBytes 
      BackColor       =   &H00000000&
      Caption         =   "Monitor Bytes (I/O)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   2685
   End
   Begin VB.CheckBox chkEnc 
      BackColor       =   &H00000000&
      Caption         =   "Monitor Encoder"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   2385
   End
   Begin VB.CheckBox chkSlew 
      BackColor       =   &H00000000&
      Caption         =   "Slews, Syncs, Park/Unpark"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Value           =   1  'Checked
      Width           =   4785
   End
   Begin VB.CheckBox chkSlewing 
      BackColor       =   &H00000000&
      Caption         =   "Monitor Slewing"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2205
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   1260
      Width           =   915
   End
   Begin VB.CheckBox chkOther 
      BackColor       =   &H00000000&
      Caption         =   "Other ASCOM Calls"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2445
   End
   Begin VB.CheckBox chkHW 
      BackColor       =   &H00000000&
      Caption         =   "HW Calls"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   1545
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   1620
      Width           =   915
   End
   Begin VB.TextBox txtTraffic 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   4485
      Left            =   98
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1920
      Width           =   5115
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'   ===========
'   FRMSHOW.FRM
'   ===========
'
' ASCOM Jornaling form
'
' Written:  31-May-03   Jon Brewster
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 31-may-03 jab     Initial edit
' 17-Jun-17 jab     Modified for general use (LOGLENGTH, TextOffset)
' 10-Sep-03 jab     make resize more robust
' -----------------------------------------------------------------------------

Option Explicit

Private Const LOGLENGTH = 2000
Private m_bDisable As Boolean
Private m_bBOL As Boolean

Private Sub cmdClear_Click()

  m_bBOL = True
  txtTraffic.Text = ""

End Sub

Private Sub Form_Load()

    cmdClear_Click
    m_bDisable = False
    Me.cmdDisable.Caption = "Disable"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not g_show Is Nothing Then _
        Set g_show = Nothing
    
End Sub

Private Sub Form_Resize()
    Dim TextOffset As Integer
        
    If WindowState = vbNormal Or WindowState = vbMaximized Then
        TextOffset = txtTraffic.Top + 600
        
        If Width < 5415 Then _
            Width = 5415
        If Height < TextOffset + 400 Then _
            Height = TextOffset + 400
            
        txtTraffic.Height = Height - TextOffset
        txtTraffic.Width = Width - 300
    End If
  
End Sub

Private Sub cmdDisable_Click()

    m_bDisable = Not m_bDisable
    Me.cmdDisable.Caption = IIf(m_bDisable, "Enable", "Disable")
    
End Sub

Public Sub TrafficChar(val As String)

    If m_bDisable Then _
        Exit Sub

    If m_bBOL Then
        m_bBOL = False
        txtTraffic.Text = Right(txtTraffic.Text & val, LOGLENGTH)
    Else
        txtTraffic.Text = Right(txtTraffic.Text & " " & val, LOGLENGTH)
    End If
     
    txtTraffic.SelStart = Len(txtTraffic.Text)
    
End Sub

Public Sub TrafficLine(val As String)

    If m_bDisable Then _
        Exit Sub
        
    If m_bBOL Then
        val = val & vbCrLf
    Else
        val = vbCrLf & val & vbCrLf
    End If
    
    m_bBOL = True
     
    txtTraffic.Text = Right(txtTraffic.Text & val, LOGLENGTH)
    txtTraffic.SelStart = Len(txtTraffic.Text)
    
End Sub

Public Sub TrafficStart(val As String)

    If m_bDisable Then _
        Exit Sub
        
    If Not m_bBOL Then _
        val = vbCrLf & val
   
    m_bBOL = False
     
    txtTraffic.Text = Right(txtTraffic.Text & val, LOGLENGTH)
    txtTraffic.SelStart = Len(txtTraffic.Text)
    
End Sub

Public Sub TrafficEnd(val As String)

    If m_bDisable Then _
        Exit Sub
    
    val = val & vbCrLf
   
    m_bBOL = True
     
    txtTraffic.Text = Right(txtTraffic.Text & val, LOGLENGTH)
    txtTraffic.SelStart = Len(txtTraffic.Text)
    
End Sub

