VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Screen"
   ClientHeight    =   3825
   ClientLeft      =   1755
   ClientTop       =   2145
   ClientWidth     =   5505
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   5505
   WindowState     =   2  'Maximized
   Begin VB.Label lblGraphic 
      Alignment       =   2  'Center
      Caption         =   "Galactic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4812
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Starmap"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
    InitializeToolBar
End Sub

Public Sub InitializeToolBar()
    frmWinGal.cmdToolBar(0).Caption = "Starmaps" + Chr$(13) + Chr$(10) + "{F1}"
    frmWinGal.cmdToolBar(1).Caption = "About" + Chr$(13) + Chr$(10) + "{F2}"
    frmWinGal.cmdToolBar(2).Caption = "Options" + Chr$(13) + Chr$(10) + "{F3}"
    frmWinGal.cmdToolBar(3).Caption = ""
    frmWinGal.cmdToolBar(4).Caption = ""
    frmWinGal.cmdToolBar(5).Caption = ""
    frmWinGal.cmdToolBar(6).Caption = ""
    frmWinGal.cmdToolBar(7).Caption = "Quit" + Chr$(13) + Chr$(10) + "{Esc}"
End Sub

Public Sub KeyRoute(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyF1, vbTB0
        frmGalaxies.Show
        Me.Hide
    Case vbKeyF2, vbTB1
        MsgBox "0"
    Case vbKeyF3, vbTB2
        MsgBox "0"
    Case vbKeyEscape, vbTB7
        End
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyRoute KeyCode
End Sub

Private Sub Form_Resize()
    lblGraphic.Top = Me.ScaleTop
    lblGraphic.Left = Me.ScaleLeft
    lblGraphic.Width = Me.ScaleWidth
    lblGraphic.Height = Me.ScaleHeight
End Sub

