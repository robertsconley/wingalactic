VERSION 5.00
Begin VB.Form frmGalaxies 
   Caption         =   "Galaxies"
   ClientHeight    =   6960
   ClientLeft      =   8775
   ClientTop       =   5415
   ClientWidth     =   9600
   ControlBox      =   0   'False
   Icon            =   "frmGalaxies.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   9600
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtEG 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmGalaxies.frx":0442
      Top             =   4560
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.PictureBox picEG 
      BackColor       =   &H00800000&
      Height          =   3972
      Left            =   3600
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   240
      Width           =   3972
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2052
      End
   End
   Begin VB.ListBox lstGalaxy 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2532
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
      Begin VB.Menu mnuToolsEditPrev 
         Caption         =   "Edit Gal.lst {<}"
      End
      Begin VB.Menu mnuToolsEditCurrent 
         Caption         =   "Edit Galaxy .lst {>}"
      End
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
Attribute VB_Name = "frmGalaxies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyGalaxies As New colGalaxy
Dim MenuDir As String
Sub ReadGalaxyList()
    Dim FileNum As Integer
    Dim TempS As String
    FileNum = FreeFile
    Open App.Path + "\gals\gal.lst" For Input As FileNum
    Line Input #FileNum, TempS
    Line Input #FileNum, TempS
    MyGalaxies.Clear
    Do Until EOF(FileNum)
        Line Input #FileNum, TempS
        TempS = RTrim$(TempS)
        If TempS <> "" Then
            'MyGalaxies.Add LCase$(RTrim$(Left$(TempS, 8))), Right$(TempS, Len(TempS) - 13)
        End If
    Loop
    Close #FileNum
End Sub
Private Sub Form_Activate()
    InitializeToolBar
    ReadGalaxyList
    MyGalaxies.FillList Me.lstGalaxy
End Sub
Public Sub InitializeToolBar()
    frmWinGal.cmdToolBar(0).Caption = "Select" + Chr$(13) + Chr$(10) + "{Enter}"
    frmWinGal.cmdToolBar(1).Caption = "Create" + Chr$(13) + Chr$(10) + "{Ins}"
    frmWinGal.cmdToolBar(2).Caption = "Delete" + Chr$(13) + Chr$(10) + "{Del}"
    frmWinGal.cmdToolBar(3).Caption = "Rd Notes" + Chr$(13) + Chr$(10) + "{F1}"
    frmWinGal.cmdToolBar(4).Caption = "" + Chr$(13) + Chr$(10) + ""
    frmWinGal.cmdToolBar(5).Caption = "Sectors" + Chr$(13) + Chr$(10) + "{F5}"
    frmWinGal.cmdToolBar(6).Caption = "" + Chr$(13) + Chr$(10) + ""
    frmWinGal.cmdToolBar(7).Caption = "Exit" + Chr$(13) + Chr$(10) + "{Esc}"
End Sub

Public Sub KeyRoute(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyReturn, vbTB0
        MsgBox "0"
    Case vbKeyInsert, vbTB1
        MsgBox "0"
    Case vbKeyDelete, vbTB2
        MsgBox "0"
    Case vbKeyF1, vbTB3
        If Me.lstGalaxy.ListIndex < 0 Then
            MsgBox "Need to Select a Galaxy"
        Else
            ReadGalaxyMenu
        End If
    Case vbKeyF2
    Case vbKeyF5, vbTB4
        MsgBox "0"
    Case vbKeyI
        MsgBox "0"
    Case vbKeyX
    MsgBox "0"
    Case vbKeyEscape, vbTB7
        frmMain.Show
        Me.Hide
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyRoute KeyCode
End Sub

Private Sub Form_Resize()
    Me.lstGalaxy.Top = 0
    Me.lstGalaxy.Left = 0
    Me.lstGalaxy.Height = Me.Height
    Me.picEG.Top = 0
    Me.picEG.Height = Me.Height
    Me.picEG.Left = lstGalaxy.Left + lstGalaxy.Width
    Me.picEG.Width = Me.Width - Me.picEG.Left
    Me.txtEG.Top = 0
    Me.txtEG.Height = Me.Height
    Me.txtEG.Left = lstGalaxy.Left + lstGalaxy.Width
    Me.txtEG.Width = Me.Width - Me.picEG.Left
End Sub

Private Sub ReadGalaxyMenu()
    Dim FileNum As Integer
    Dim TempS As String
    If Me.lstGalaxy.ListIndex < 0 Then
        Exit Sub
    End If
    TempS = MyGalaxies(lstGalaxy.ItemData(lstGalaxy.ListIndex)).Directory
    'frmNotes.Directory = App.Path & "\gals\" & TempS & "\gen\"
    'frmNotes.MenuName = "Galaxy.mnu"
    'Set frmNotes.Source = Me
    'frmNotes.Show
    Me.Hide
End Sub
