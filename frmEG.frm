VERSION 5.00
Begin VB.Form frmEG 
   BackColor       =   &H00000000&
   Caption         =   "EG"
   ClientHeight    =   5085
   ClientLeft      =   240
   ClientTop       =   2145
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmEG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6585
End
Attribute VB_Name = "frmEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim p(50)
'Dim F$(50)
'Dim back(50)
'Dim L$(25)
'Dim R$(25)
'Dim Row(25)
'Dim ColName$(16)
'Dim hip
'Dim Old
'Dim topp
'Dim Reader$
'Dim InputText As String
'
'Private Sub Form_Activate()
'    Call Main
'End Sub
'
'Private Sub Main(Cmd As String)
'    Dim N As Integer
'    Dim k1 As Integer
'    Dim k2 As Integer
'    Dim ext As String
'    Call Init(Cmd)
'
'    N = 1
'
'    Call DrawScreen(N)
'
'    Do
'        Call ReadKeyboard(k1, k2)
'        If k2 = 0 And k1 = 27 Then ' REM esc
'            If back(N) = 0 Then
'              Call ExitProgram
'            Else
'              N = back(N)
'              Call DrawScreen(N)
'            End If
'        End If
'
'        If k2 = 0 And (k1 = 81 Or k1 = 113) Then ' REM Q/q
'            If back(N) = 0 Then
'              Call ExitProgram
'            Else
'              N = back(N)
'              Call DrawScreen(N)
'            End If
'
'        End If
'        If k2 = 1 And (k1 = 72 Or k1 = 80) Then Call ProcessArrowKeys(k1, k2, N): Rem up/down
'        If k2 = 0 And (k1 = 74 Or k1 = 106) Then Call ProcessArrowKeys(k1, k2, N): Rem J/j=down
'        If k2 = 0 And (k1 = 75 Or k1 = 107) Then Call ProcessArrowKeys(k1, k2, N): Rem K/k=up
'        If k2 = 1 And (k1 = 75 Or k1 = 77) Then Call ProcessArrowKeys(k1, k2, N): Rem left/right
'        If k2 = 1 And (k1 = 73 Or k1 = 81) Then Call ProcessArrowKeys(k1, k2, N): Rem pgup/pgdn
'        If k2 = 1 And (k1 = 71 Or k1 = 79) Then Call ProcessArrowKeys(k1, k2, N): Rem home/end
'        If k2 = 0 And k1 = 60 Then ' REM <
'            Call CheckMenu(Reader$, F$(N))
'            Call DrawScreen(N)
'        End If
'
'        If k2 = 0 And k1 = 62 Then ' REM >
'            ext$ = Right$(R$(p(N)), 3)
'            If ext$ = "txt" Then
'                Call SelectEntry(N)
'            ElseIf ext$ = "mnu" Then
'                Call CheckMenu(Reader$, R$(p(N)))
'            End If
'            Call DrawScreen(N)
'        End If
'
'        If k2 = 0 And k1 = 13 Then ' REM enter
'            Call SelectEntry(N)
'            Call DrawScreen(N)
'        End If
'
'        If k2 = 0 And k1 = 32 Then ' REM space
'            Call SelectEntry(N)
'            Call DrawScreen(N)
'        End If
'
'        If k2 = 0 And (k1 = 72 Or k1 = 104 Or k1 = 63) Then
'            Call PrintHelp
'            Call DrawScreen(N)
'        End If
'
'        If k2 = 0 And (k1 = 67 Or k1 = 99) Then ' REM C/c
'            Call ChangeBrowser(Reader$)
'            Call DrawScreen(N)
'        End If
'    Loop
'
'End Sub
'
'Sub Init(Cmd As String)
'    Dim Flag1 As Integer
'    Dim Flag2 As Integer
'    Dim SP As Integer
'    Dim I As Integer
'    Dim A As String
'    If Cmd$ <> "" Then
'        If InStr(Cmd$, ".MNU") > 0 Then Flag1 = 1
'        If Flag1 = 0 Then Flag2 = 1
'        If InStr(Cmd$, " ") > 0 Then Flag2 = 1
'        If Flag1 = 0 And Flag2 = 1 Then
'            Reader$ = Cmd$
'        ElseIf Flag1 = 1 And Flag2 = 0 Then
'            F$(1) = Cmd$
'        ElseIf Flag1 = 1 And Flag2 = 1 Then
'            SP = InStr(Cmd$, " ")
'            F$(1) = Left$(Cmd$, SP - 1)
'            Reader$ = Mid$(Cmd$, SP + 1, Len(Cmd$) - SP)
'        Else
'            Print "Error in Command Line": End
'        End If
'    End If
'
'    Open "E:\Microsoft Visual Studio\VB98\xtra\colors.dat" For Input As #1
'
'    For I = 0 To 15
'        Line Input #1, A$
'        ColName$(I) = LCase$(Right$(A$, Len(A$) - 4))
'    Next I
'    Close #1
'
'    If Reader$ = "" Then Reader$ = "less "
'
'    p(1) = 1
'
'    If F$(1) = "" Then F$(1) = "main.mnu"
'
'    back(1) = 0
'
'    topp = 1: Rem number of menus encountered
'
'End Sub
'
'Sub DrawScreen(N)
'    Call ReadMenu(F$(N))
'    Call DrawHelp(Old)
'    Call MakeArrow(Row(p(N)))
'End Sub
'
'Sub ReadMenu(F As String)
''SUB ReadMenu (F AS STRING, ColName() AS STRING, hip, l() AS STRING, R() AS STRING, Row())
'    Rem read menu file
'    Dim DirName As String
'    Dim Exist As Integer
'
'    DirName$ = ""
'
'    Rem see if file exists
'    Open F For Random As #1
'        Exist = LOF(1)
'    Close #1
'
'    If Exist = 0 Then
'        Kill F$
'        CLS
'        Color 14
'        Locate 8
'        Print "OOOPSY..."
'        Print
'        Color 11
'        Print F$;
'        Color 14
'        Print " doesn't exist"
'        Print
'        Print "You may want to check this out  ";
'        Color 13
'        Print ":-)"
'        Call ReadKeyboard(0, 0)
'        ExitProgram
'    End If
'
'    'ShellPrg "call xtra\egacolor 00 10"
'    'ShellPrg "call xtra\egacolor 07 46"
'    Dim Q As Integer
'    Dim Count As Integer
'    Dim A As String
'    Dim I As Integer
'    Dim aL As Integer
'    Dim Y As Integer
'
'    Open F For Input As #1
'        CLS
'        Q = 0
'        Count = 0
'        Do Until EOF(1)
'            Count = Count + 1
'            Line Input #1, A$
'            I = InStr(A$, "@")
'            If I > 1 Then I = InStr(A$, " @")
'            If I > 1 Then I = I + 1
'            aL = Len(A$)
'            If I = 1 Then
'                If LCase$(Mid$(A$, 2, 3)) = "dir" Then
'                    DirName$ = Right$(A$, aL - 5)
'                    If Right$(DirName$, 1) <> "\" Then DirName$ = DirName$ + "\"
'                End If
'                For Y = 0 To 15
'                    If LCase$(Right$(A$, aL - 1)) = ColName$(Y) Then Color Y
'                    If LCase$(Right$(A$, aL - 1)) = ColName$(Y) + " blinking" Then
'                        Color Y + 16
'                    End If
'                Next Y
'                Count = Count - 1
'            End If
'            If I > 1 Then
'                Q = Q + 1
'                L$(Q) = Left$(A$, I - 1)
'                R$(Q) = DirName$ + Right$(A$, aL - I)
'                Row(Q) = Count
'                A$ = L$(Q)
'            End If
'            If I <> 1 Then Print A$
'        Loop
'    Close #1
'
'    hip = Q: Rem the number of possible choices
'
'End Sub
'
'Sub ReadKeyboard(k1 As Integer, k2 As Integer)
'    Rem read keyboard
'    Dim K As String
'
'    k1 = 0
'    k2 = 0
'    Do
'        K$ = InputKey
'    Loop While K$ = ""
'    k1 = Asc(K$)
'    If k1 = 13 Then
''        Beep
'    End If
'    If k1 = 0 Then
'        k1 = Asc(Right$(K$, 1))
'        k2 = 1
'    End If
'End Sub
'
'Sub ExitProgram()
'    Me.CLS
'    Print
'    Print
'    Print "Have a nice day..."
'    Print
'    Call ReadKeyboard(0, 0)
'    End
'End Sub
'
'Sub ProcessArrowKeys(k1 As Integer, k2 As Integer, N As Integer)
''    Rem up & down arrow movement
''    LOCATE Row(p(n)), 2
''    Print "  "
''    If k2 = 1 Then
''    If k1 = 72 Then p(n) = p(n) - 1: Rem up
''    If k1 = 80 Then p(n) = p(n) + 1: Rem down
''    If k1 = 75 Then p(n) = p(n) - 1: Rem left
''    If k1 = 77 Then p(n) = p(n) + 1: Rem right
''    If k1 = 73 Or k1 = 71 Then p(n) = 1: Rem pgup/home
''    If k1 = 81 Or k1 = 79 Then p(n) = hip: Rem pgdn/end
''    End If
''    If k2 = 0 Then
''    If k1 = 74 Or k1 = 106 Then p(n) = p(n) + 1: Rem J/j
''    If k1 = 75 Or k1 = 107 Then p(n) = p(n) - 1: Rem K/k
''    End If
''    If p(n) > hip Then p(n) = 1
''    If p(n) < 1 Then p(n) = hip
''    Call MakeArrow(Row(p(n)))
'End Sub
'
'Sub CheckMenu(Reader As String, F As String)
''    Rem < check out this file directly
''    Color 7: Cls: Print "."
''    e$ = "call " + Reader$ + " " + F$
''    ShellPrg "call xtra\egacolor 00 10"
''    ShellPrg "call xtra\egacolor 07 46"
''    ShellPrg  e$
'End Sub
'
'Sub SelectEntry(N As Integer)
''    Rem select
''    ext$ = Right$(R$(p(n)), 3)
''    If ext$ = "exe" Or ext$ = "com" Or ext$ = "bat" Then
''       e$ = "call " + R$(p(n))
''       ShellPrg  e$
''    End If
''    If ext$ = "txt" Or ext$ = "dat" Or ext$ = "bas" Then
''       Color 7: Cls: Print "."
''       e$ = "call " + Reader$ + " " + R$(p(n))
''       ShellPrg "call xtra\egacolor 00 10"
''       ShellPrg "call xtra\egacolor 07 46"
''       ShellPrg  e$
''    End If
''    If ext$ = "asc" Then
''       Color 7: Cls: Print ".": Cls
''       ShellPrg "call xtra\egacolor 00 10"
''       ShellPrg "call xtra\egacolor 07 46"
''       tmp$ = "type"
''       If Reader$ = "EDIT" Then tmp$ = "edit"
''       ShellPrg  tmp$ + " " + R$(p(n))
''       If tmp$ = "type" Then Call ReadKeyboard(k1, k2)
''    End If
''    If ext$ = "mnu" Then
''
''        Rem determine if we've been there before
''        there = 0
''        For z = 1 To topp
''           If R$(p(n)) = F$(z) Then there = z
''        Next z
''        Rem yes we have
''        If there <> 0 Then n = there
''        Rem no we haven't, so let's make a new place
''        If there = 0 Then
''           topp = topp + 1
''           back(topp) = n
''           F$(topp) = R$(p(n))
''           p(topp) = 1
''           n = topp
''        End If
''    End If
'
'End Sub
'
'Sub PrintHelp()
''CLS:      Color 11: Print
''    Print "Publish your own electronic books with the..."
''    Print: Color 14
''    Print "   Electric Guildsman Menuing System v2.11"
''    Print "   Public Domain 1997 Jim Vassilakos"
''    Print: Color 12
''    Print "Command Line: eg [<file.mnu>] [<textbrowser>]"
''    Print "   default file.mnu = main.mnu"
''    Print "   default textbrowser = less"
''    Print: Color 13
''    Print "Commands:  < : Browse This Menu      > : Browse That Menu"
''    Print "           C : Change Browser        Q : Quit"
''    Print "           Arrows : Move Pointer     Enter : Select Item"
''    Print
''    Print "Recognized File Extensions:"
''    Print "   txt/dat/bas: Vanilla Text     asc: Ascii (Non-Vanilla)"
''    Print "   bat/com/exe: Executable       mnu: Menu"
''    Print: Color 11
''    Print "For help & stuff, send email to:"
''    Print "   jimv@empirenet.com               jimv@cs.ucr.edu"
''    Print "   jimv@silver.lcs.mit.edu          jimv@wizards.com"
''    Call ReadKeyboard(k1, k2)
'End Sub
'
'Sub ChangeBrowser(Reader As String)
''CLS:      Print: Color 11
''    Print "This program is just a menuing system. The actual text"
''    Print "reading/editing is done by another program, called the"
''    Print "text browser. You can use any text browsers you want."
''    Print "Some examples are:"
''    Print: Color 10
''    Print "  less      Read-only, easy to use, ignores tilde (~)"
''    Print " browse     Read-only, also easy to use"
''    Print "  ted       Editor, easy to use and light-weight"
''    Print "  edit      Editor, easy to use but heavy on memory"
''    Print "   vi       Editor, difficult for beginners (see vi.hlp)"
''    Print: Color 11
''    Print "You can even use a wordprocessor such as 'wp' if you"
''    Print "like. Just make sure you save your work as text, not"
''    Print "in wp-format."
''    Print: Color 13
''    Print "Your current text brower is "; Reader$
''    Print: Color 12
''    INPUT "Enter name of new text browser: ", nr$
''    If nr$ <> "" Then Reader$ = nr$
'
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    InputText = InputText & Chr$(KeyAscii)
'End Sub
'
'Private Function InputKey() As String
'    InputKey = InputText
'    InputText = ""
'    DoEvents
'End Function
'
'Private Sub Color(N As Integer)
'    If N > 16 Then
'        N = N - 16
'        Me.FontItalic = True
'    Else
'        Me.FontItalic = False
'    End If
'    Me.ForeColor = QBColor(N)
'End Sub
'
'Public Sub Locate(Row As Integer, Optional Col As Integer = 0)
'    Dim I As Integer
'    CurrentX = 0
'    CurrentY = 0
'    For I = 1 To Row - 1
'        Print ""
'    Next I
'    Print Space$(Col);
'End Sub
'
'Sub DrawHelp(Old)
'    If Old = 0 Then
'       Locate 25, 75: Color 12
'       Print "H=Help";
'       Old = 1
'    End If
'End Sub
'
'Sub MakeArrow(Row)
'    Rem make arrow
'    Locate Row, 2
'    Color 12
'    Print Chr$(196); Chr$(26);
'End Sub
'
