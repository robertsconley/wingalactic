Attribute VB_Name = "modEg"
Option Explicit
Dim p(50) As Integer
Dim F$(50)
Dim back(50) As Integer
Dim L$(25)
Dim R$(25)
Dim Row(25) As Integer
Dim ColName$(16)
Dim hip As Integer
Dim Old As Integer
Dim topp As Integer
Dim Reader$
Dim InputText As String

Public Sub EgMain(Cmd As String)
    Dim N As Integer
    Dim k1 As Integer
    Dim k2 As Integer
    Dim Ext As String
    
    Reader$ = ReaderStr
    Call Init(Cmd)

    N = 1

    Call DrawScreen(N)

    Do
        Call ReadKeyboard(k1, k2)
        If k2 = 0 And k1 = 27 Then ' REM esc
            If back(N) = 0 Then
              Call ExitProgram
              Exit Sub
            Else
              N = back(N)
              Call DrawScreen(N)
            End If
        End If

        If k2 = 0 And (k1 = 81 Or k1 = 113) Then ' REM Q/q
            If back(N) = 0 Then
              Call ExitProgram
              Exit Sub
            Else
              N = back(N)
              Call DrawScreen(N)
            End If

        End If
        If k2 = 1 And (k1 = 72 Or k1 = 80) Then Call ProcessArrowKeys(k1, k2, N): Rem up/down
        If k2 = 0 And (k1 = 74 Or k1 = 106) Then Call ProcessArrowKeys(k1, k2, N): Rem J/j=down
        If k2 = 0 And (k1 = 75 Or k1 = 107) Then Call ProcessArrowKeys(k1, k2, N): Rem K/k=up
        If k2 = 1 And (k1 = 75 Or k1 = 77) Then Call ProcessArrowKeys(k1, k2, N): Rem left/right
        If k2 = 1 And (k1 = 73 Or k1 = 81) Then Call ProcessArrowKeys(k1, k2, N): Rem pgup/pgdn
        If k2 = 1 And (k1 = 71 Or k1 = 79) Then Call ProcessArrowKeys(k1, k2, N): Rem home/end
        If k2 = 0 And k1 = 60 Then ' REM <
            Call CheckMenu(Reader$, F$(N))
            Call DrawScreen(N)
        End If

        If k2 = 0 And k1 = 62 Then ' REM >
            Ext$ = Right$(R$(p(N)), 3)
            If Ext$ = "txt" Then
                Call SelectEntry(N)
            ElseIf Ext$ = "mnu" Then
                Call CheckMenu(Reader$, R$(p(N)))
            End If
            Call DrawScreen(N)
        End If

        If k2 = 0 And k1 = 13 Then ' REM enter
            Call SelectEntry(N)
            Call DrawScreen(N)
        End If

        If k2 = 0 And k1 = 32 Then ' REM space
            Call SelectEntry(N)
            Call DrawScreen(N)
        End If

        If k2 = 0 And (k1 = 72 Or k1 = 104 Or k1 = 63) Then
            Call CurrentForm.PrintHelp
            Call DrawScreen(N)
        End If

        If k2 = 0 And (k1 = 67 Or k1 = 99) Then ' REM C/c
            Call ChangeBrowser(Reader$)
            Call DrawScreen(N)
        End If
    Loop

End Sub

Sub Init(ByVal Cmd As String)
    Dim Flag1 As Integer
    Dim Flag2 As Integer
    Dim SP As Integer
    Dim I As Integer
    Dim A As String
    Cmd$ = UCase$(Cmd$)
    If Cmd$ <> "" Then
        If InStr(Cmd$, ".MNU") > 0 Then Flag1 = 1
        If Flag1 = 0 Then Flag2 = 1
        If InStr(Cmd$, " ") > 0 Then Flag2 = 1
        If Flag1 = 0 And Flag2 = 1 Then
            Reader$ = Cmd$
        ElseIf Flag1 = 1 And Flag2 = 0 Then
            F$(1) = Cmd$
        ElseIf Flag1 = 1 And Flag2 = 1 Then
            SP = InStr(Cmd$, " ")
            F$(1) = Left$(Cmd$, SP - 1)
            Reader$ = Mid$(Cmd$, SP + 1, Len(Cmd$) - SP)
        Else
            CurrentForm.Print "Error in Command Line": End
        End If
    End If

    Open "xtra\colors.dat" For Input As #1
    
    For I = 0 To 15
        Line Input #1, A$
        ColName$(I) = LCase$(Right$(A$, Len(A$) - 4))
    Next I
    Close #1

    If Reader$ = "" Then Reader$ = "less "

    p(1) = 1

    If F$(1) = "" Then F$(1) = "main.mnu"

    back(1) = 0

    topp = 1: Rem number of menus encountered

End Sub

Sub DrawScreen(N)
    Call ReadMenu(F$(N))
    Call DrawHelp(Old)
    Call MakeArrow(Row(p(N)))
End Sub

Sub ReadMenu(F As String)
'SUB ReadMenu (F AS STRING, ColName() AS STRING, hip, l() AS STRING, R() AS STRING, Row())
    Rem read menu file
    Dim DirName As String
    Dim Exist As Integer
    
    DirName$ = ""

    Rem see if file exists
    Open F For Random As #1
        Exist = LOF(1)
    Close #1

    If Exist = 0 Then
        Kill F$
        CLS
        Color 14
        Locate 8
        CurrentForm.Print "OOOPSY..."
        CurrentForm.Print
        Color 11
        CurrentForm.Print F$;
        Color 14
        CurrentForm.Print " doesn't exist"
        CurrentForm.Print
        CurrentForm.Print "You may want to check this out  ";
        Color 13
        CurrentForm.Print ":-)"
        Call ReadKeyboard(0, 0)
        ExitProgram
        Exit Sub
    End If

    'ShellPrg "call xtra\egacolor 00 10"
    'ShellPrg "call xtra\egacolor 07 46"
    Dim Q As Integer
    Dim Count As Integer
    Dim A As String
    Dim I As Integer
    Dim aL As Integer
    Dim y As Integer
    
    Open F For Input As #1
        CLS
        Q = 0
        Count = 0
        Do Until EOF(1)
            Count = Count + 1
            Line Input #1, A$
            I = InStr(A$, "@")
            If I > 1 Then I = InStr(A$, " @")
            If I > 1 Then I = I + 1
            aL = Len(A$)
            If I = 1 Then
                If LCase$(Mid$(A$, 2, 3)) = "dir" Then
                    DirName$ = Right$(A$, aL - 5)
                    If Right$(DirName$, 1) <> "\" Then DirName$ = DirName$ + "\"
                End If
                For y = 0 To 15
                    If LCase$(Right$(A$, aL - 1)) = ColName$(y) Then Color y
                    If LCase$(Right$(A$, aL - 1)) = ColName$(y) + " blinking" Then
                        Color y + 16
                    End If
                Next y
                Count = Count - 1
            End If
            If I > 1 Then
                Q = Q + 1
                L$(Q) = Left$(A$, I - 1)
                R$(Q) = DirName$ + Right$(A$, aL - I)
                Row(Q) = Count
                A$ = L$(Q)
            End If
            If I <> 1 Then CurrentForm.Print A$
        Loop
    Close #1

    hip = Q: Rem the number of possible choices

End Sub

Sub ReadKeyboard(k1 As Integer, k2 As Integer)
   Call GetKeyChar(k1, k2)
End Sub

Sub ExitProgram()
    CLS
    CurrentForm.Print
    CurrentForm.Print
    CurrentForm.Print "Have a nice day..."
    CurrentForm.Print
    Call ReadKeyboard(0, 0)
End Sub

Sub ProcessArrowKeys(k1 As Integer, k2 As Integer, N As Integer)
    Rem up & down arrow movement
    Locate Row(p(N)), 2
    Color 0
    CurrentForm.Print Chr$(196); Chr$(26);
    If k2 = 1 Then
        If k1 = 72 Then p(N) = p(N) - 1: Rem up
        If k1 = 80 Then p(N) = p(N) + 1: Rem down
        If k1 = 75 Then p(N) = p(N) - 1: Rem left
        If k1 = 77 Then p(N) = p(N) + 1: Rem right
        If k1 = 73 Or k1 = 71 Then p(N) = 1: Rem pgup/home
        If k1 = 81 Or k1 = 79 Then p(N) = hip: Rem pgdn/end
    End If
    If k2 = 0 Then
        If k1 = 74 Or k1 = 106 Then p(N) = p(N) + 1: Rem J/j
        If k1 = 75 Or k1 = 107 Then p(N) = p(N) - 1: Rem K/k
    End If
    If p(N) > hip Then p(N) = 1
    If p(N) < 1 Then p(N) = hip
    Call MakeArrow(Row(p(N)))
End Sub

Sub CheckMenu(Reader As String, F As String)
    Rem < check out this file directly
    Dim E As String
    Color 7
    CLS
    CurrentForm.Print "."
    E$ = "call " + Reader$ + " " + F$
'    ShellPrg "call xtra\egacolor 00 10"
'    ShellPrg "call xtra\egacolor 07 46"
    ShellPrg E$
End Sub

Sub SelectEntry(N As Integer)
    Rem select
    Dim Ext As String
    Dim E As String
    Dim Tmp As String
    Dim There As Integer
    Dim Z As Integer
    
    Ext$ = Right$(R$(p(N)), 3)
    If Ext$ = "exe" Or Ext$ = "com" Or Ext$ = "bat" Then
       E$ = R$(p(N))
       ShellPrg E$, vbMaximizedFocus
    End If
    If Ext$ = "txt" Or Ext$ = "dat" Or Ext$ = "bas" Then
       Color 7: CLS: CurrentForm.Print "."
       E$ = Reader$ + " " + R$(p(N))
       'ShellPrg "call xtra\egacolor 00 10"
       'ShellPrg "call xtra\egacolor 07 46"
       ShellPrg E$, vbMaximizedFocus
    End If
    If Ext$ = "asc" Then
       Color 7: CLS: CurrentForm.Print ".": CLS
       'ShellPrg "call xtra\egacolor 00 10"
       'ShellPrg "call xtra\egacolor 07 46"
       Tmp$ = "type"
       If Reader$ = "EDIT" Then Tmp$ = "edit"
       ShellPrg Tmp$ + " " + R$(p(N))
       If Tmp$ = "type" Then Call ReadKeyboard(0, 0)
    End If
    If Ext$ = "mnu" Then

        Rem determine if we've been there before
        There = 0
        For Z = 1 To topp
           If R$(p(N)) = F$(Z) Then There = Z
        Next Z
        
        If There = 0 Then
           topp = topp + 1
           back(topp) = N
           F$(topp) = R$(p(N))
           p(topp) = 1
           N = topp
        Else
            N = There
        End If
    End If
End Sub

Sub PrintHelp()
    CLS
    Color 11
    CurrentForm.Print
    CurrentForm.Print "Publish your own electronic books with the..."
    CurrentForm.Print: Color 14
    CurrentForm.Print "   Electric Guildsman Menuing System v2.11"
    CurrentForm.Print "   Public Domain 1997 Jim Vassilakos"
    CurrentForm.Print: Color 12
    CurrentForm.Print "Command Line: eg [<file.mnu>] [<textbrowser>]"
    CurrentForm.Print "   default file.mnu = main.mnu"
    CurrentForm.Print "   default textbrowser = less"
    CurrentForm.Print: Color 13
    CurrentForm.Print "Commands:  < : Browse This Menu      > : Browse That Menu"
    CurrentForm.Print "           C : Change Browser        Q : Quit"
    CurrentForm.Print "           Arrows : Move Pointer     Enter : Select Item"
    CurrentForm.Print
    CurrentForm.Print "Recognized File Extensions:"
    CurrentForm.Print "   txt/dat/bas: Vanilla Text     asc: Ascii (Non-Vanilla)"
    CurrentForm.Print "   bat/com/exe: Executable       mnu: Menu"
    CurrentForm.Print: Color 11
    CurrentForm.Print "For help & stuff, send email to:"
    CurrentForm.Print "   jimv@empirenet.com               jimv@cs.ucr.edu"
    CurrentForm.Print "   jimv@silver.lcs.mit.edu          jimv@wizards.com"
    Call ReadKeyboard(0, 0)
End Sub

Sub ChangeBrowser(Reader As String)
    Dim NR As String
    CLS
    CurrentForm.Print
    Color 11
    CurrentForm.Print "This program is just a menuing system. The actual text"
    CurrentForm.Print "reading/editing is done by another program, called the"
    CurrentForm.Print "text browser. You can use any text browsers you want."
    CurrentForm.Print "Some examples are:"
    CurrentForm.Print: Color 10
    CurrentForm.Print "  less      Read-only, easy to use, ignores tilde (~)"
    CurrentForm.Print " browse     Read-only, also easy to use"
    CurrentForm.Print "  ted       Editor, easy to use and light-weight"
    CurrentForm.Print "  edit      Editor, easy to use but heavy on memory"
    CurrentForm.Print "   vi       Editor, difficult for beginners (see vi.hlp)"
    CurrentForm.Print: Color 11
    CurrentForm.Print "You can even use a wordprocessor such as 'wp' if you"
    CurrentForm.Print "like. Just make sure you save your work as text, not"
    CurrentForm.Print "in wp-format."
    CurrentForm.Print: Color 13
    CurrentForm.Print "Your current text brower is "; Reader$
    CurrentForm.Print: Color 12
    NR$ = InputBox("Enter name of new text browser: ")
    If NR$ <> "" Then Reader$ = NR$
    ReaderStr = Reader$

End Sub


'Private Function InputKey() As String
'    InputKey = InputText
'    InputText = ""
'    DoEvents
'End Function

'Public Sub Locate(Row As Integer, Optional Col As Integer = 0)
'    Dim I As Integer
'    CurrentX = 0
'    CurrentY = 0
'    For I = 1 To Row - 1
'        CurrentForm.Print ""
'    Next I
'    CurrentForm.Print Space$(Col);
'End Sub

Sub DrawHelp(Old)
    If Old = 0 Then
       Locate 25, 75: Color 12
       CurrentForm.Print "H=Help";
       Old = 1
    End If
End Sub

Sub MakeArrow(Row)
    Rem make arrow
    Locate Row, 2
    Color 12
    CurrentForm.Print Chr$(196); Chr$(26);
End Sub

Private Sub Locate(Optional y, Optional x)
    Gal3.Locate y - 1, x
End Sub


