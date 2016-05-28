VERSION 5.00
Begin VB.MDIForm frmWinGal 
   BackColor       =   &H8000000C&
   Caption         =   "Galactic for Windows"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "frmWinGal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      Height          =   612
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8412
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   7
         Left            =   6720
         TabIndex        =   8
         Top             =   0
         Width           =   950
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   6
         Left            =   5760
         TabIndex        =   7
         Top             =   0
         Width           =   950
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   5
         Left            =   4800
         TabIndex        =   6
         Top             =   0
         Width           =   950
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   4
         Left            =   3840
         TabIndex        =   5
         Top             =   0
         Width           =   950
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   3
         Left            =   2880
         TabIndex        =   4
         Top             =   0
         Width           =   950
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   2
         Left            =   1920
         TabIndex        =   3
         Top             =   0
         Width           =   950
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   0
         Width           =   950
      End
      Begin VB.CommandButton cmdToolBar 
         Caption         =   "Command1"
         Height          =   492
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   950
      End
   End
End
Attribute VB_Name = "frmWinGal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdToolBar_Click(Index As Integer)
    If Not Me.ActiveForm Is Nothing Then
        Me.ActiveForm.KeyRoute 10000 + Index
        'Me.ActiveForm.SetFocus
    End If
End Sub

Private Sub MDIForm_Load()
    Init
End Sub
Public Sub Init()
    frmGal.Show
End Sub

Private Sub MDIForm_Resize()
'    frmMain.Left = 0
'    frmMain.Top = 0
'    frmMain.Width = Me.ScaleWidth
'    frmMain.Height = Me.ScaleHeight
End Sub
