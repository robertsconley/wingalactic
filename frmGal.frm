VERSION 5.00
Begin VB.Form frmGal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Galactic"
   ClientHeight    =   6975
   ClientLeft      =   300
   ClientTop       =   1290
   ClientWidth     =   10560
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picLetter 
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   0
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   1
      Top             =   0
      Width           =   75
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1050
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   7785
   End
End
Attribute VB_Name = "frmGal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Inkey = Convert(KeyCode, Shift)
End Sub

Function Convert(KeyCode As Integer, Shift As Integer) As String
    Dim Code As Integer
    Dim DoExtended As Boolean
    Select Case KeyCode
    Case vbKeyF1 To vbKeyF12
        Code = (KeyCode - 112) + 59
        DoExtended = True
    Case vbKeyLeft
        Code = 75
        DoExtended = True
    Case vbKeyRight
        Code = 77
        DoExtended = True
    Case vbKeyUp
        Code = 72
        DoExtended = True
    Case vbKeyHome
        Code = 71
        DoExtended = True
    Case vbKeyEnd
        Code = 79
        DoExtended = True
    Case vbKeyDown
        Code = 80
        DoExtended = True
    Case vbKeyPageUp
        Code = 73
        DoExtended = True
    Case vbKeyPageDown
        Code = 81
        DoExtended = True
    Case vbKeyInsert
        Code = 82
        DoExtended = True
    Case vbKeyDelete
        Code = 83
        DoExtended = True
    Case 9
        If Shift Mod 2 = 1 Then
            Code = 15
            DoExtended = True
        Else
            KeyCode = Code
        End If
    Case vbKeyA To vbKeyZ
        If Shift > 1 And Shift <> 4 Then
            Code = (KeyCode - 65) + 1
        Else
            Code = KeyCode
        End If
    Case 188
        If Shift Mod 2 = 1 Then
            Code = 60
        Else
            Code = KeyCode
        End If
    Case 190
        If Shift Mod 2 = 1 Then
            Code = 62
        Else
            Code = KeyCode
        End If
    Case 191
        If Shift Mod 2 = 1 Then
            Code = 63
        Else
            Code = KeyCode
        End If
    Case Else
        Code = KeyCode
    End Select
    If DoExtended Then
        Convert = Chr$(0) & Chr$(Code)
    Else
        Convert = Chr$(Code)
    End If
End Function

Private Sub Form_Resize()
    'Me.Scale (0, 0)-(639, 479)
    Me.txtInput.Left = Me.ScaleLeft
    Me.txtInput.Width = Me.ScaleWidth
    Me.picLetter.Left = Me.ScaleWidth - 5
    Me.picLetter.Top = Me.ScaleHeight - 5
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
