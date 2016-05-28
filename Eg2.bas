Attribute VB_Name = "EgMod"
Rem Electric Guildsman Menuing System 2.2
Rem Public Domain 1998 Jim Vassilakos

Dim p(50)
Dim F$(50)
Dim back(50)
Dim LL$(25)
Dim R$(25)
Dim Row(25)
Dim ColName$(16)
Dim Flag1
Dim Flag2
Dim Reader$
Dim topp
Dim N
Dim hip

Public Sub Eg(Directory, Menu)
    Dim FileNum As Integer
    Dim TempS As String
    frmNotes.Directory = App.Path & "\" & Directory
    frmNotes.MenuName = Menu
    Set frmNotes.Source = CurrentForm
    frmNotes.Show
    CurrentForm.Hide
    Exit Sub
    If Command <> "" Then
    
        If InStr(Command, ".MNU") > 0 Then Flag1 = 1
        If Flag1 = 0 Then Flag2 = 1
        If InStr(Command, " ") > 0 Then Flag2 = 1
        If Flag1 = 0 And Flag2 = 1 Then
            Reader$ = Command
        ElseIf Flag1 = 1 And Flag2 = 0 Then
            F$(1) = Command$
        ElseIf Flag1 = 1 And Flag2 = 1 Then
            SP = InStr(Command$, " ")
            F$(1) = Left$(Command$, SP - 1)
            Reader$ = Mid$(Command$, SP + 1, Len(Command$) - SP)
        Else
            CurrentForm.Print "Error in Command Line"
            Exit Sub
        End If
    End If
    Open "xtra\colors.dat" For Input As #1
    For A = 0 To 15
        Line Input #1, TempS
        ColName$(A) = LCase$(Right$(TempS, Len(TempS) - 4))
    Next A
    Close #1
    
    If Reader$ = "" Then Reader$ = "notepad "
    p(1) = 1
    If F$(1) = "" Then F$(1) = "main.mnu"
    back(1) = 0
    topp = 1: Rem number of menus encountered
    Do
        Call ReadFile
        
        If Old = 0 Then
           Locate 25, 75: Color 12
           CurrentForm.Print "H=Help";
           Old = 1
        End If
    
        MakeArrow Row(p(N)), 2
        
        Do
            Call GetKeyChar(Key1, Key2)
            If Key2 = 0 And Key1 = 27 Then Exit Do
            If Key2 = 0 And (Key1 = 81 Or Key1 = 113) Then Exit Do: Rem Q/q
            If Key2 = 1 And (Key1 = 72 Or Key1 = 80) Then Call MoveArrow(Key1, Key2): Rem up/down
            If Key2 = 0 And (Key1 = 74 Or Key1 = 106) Then Call MoveArrow(Key1, Key2): Rem J/j=down
            If Key2 = 0 And (Key1 = 75 Or Key1 = 107) Then Call MoveArrow(Key1, Key2): Rem K/k=up
            If Key2 = 1 And (Key1 = 75 Or Key1 = 77) Then Call MoveArrow(Key1, Key2): Rem left/right
            If Key2 = 1 And (Key1 = 73 Or Key1 = 81) Then Call MoveArrow(Key1, Key2): Rem pgup/pgdn
            If Key2 = 1 And (Key1 = 71 Or Key1 = 79) Then Call MoveArrow(Key1, Key2): Rem home/end
            '->If Key2 = 0 And Key1 = 60 Then GoTo 30: Rem <
            '->If Key2 = 0 And Key1 = 62 Then GoTo 35: Rem >
            If Key2 = 0 And Key1 = 13 Then Call SelectFile: Rem enter
            If Key2 = 0 And Key1 = 32 Then Call SelectFile: Rem space
            'If Key2 = 0 And (Key1 = 72 Or Key1 = 104 Or Key1 = 63) Then GoTo 70
            'If Key2 = 0 And (Key1 = 67 Or Key1 = 99) Then GoTo 80: Rem C/c
        Loop
        If back(N) = 0 Then N = back(N)
    Loop While back(N) = 0
    Color 7
    CLS
    CurrentForm.Print: CurrentForm.Print
    CurrentForm.Print "Have a nice day..."
    CurrentForm.Print
End Sub

Sub ReadFile()
    N = 1
    MyDir$ = ""
    Rem see if file exists
    Open F$(N) For Random As #1
        Exist = LOF(1)
    Close 1
    If Exist = 0 Then Exit Sub
'    ShellPrg "call xtra\egacolor 00 10"
'    ShellPrg "call xtra\egacolor 07 46"
    Open F$(N) For Input As #1
    CLS
    Q = 0
    Count = 0
    Do Until EOF(1)
       Count = Count + 1
       Line Input #1, TempS
       I = InStr(TempS, "@")
       If I > 1 Then I = InStr(TempS, " @")
       If I > 1 Then I = I + 1
       L = Len(TempS)
       If I = 1 Then
          If LCase$(Mid$(TempS, 2, 3)) = "dir" Then
                MyDir$ = Right$(TempS, L - 5)
                If Right$(MyDir$, 1) <> "\" Then MyDir$ = MyDir$ + "\"
          End If
          For y = 0 To 15
            If LCase$(Right$(TempS, L - 1)) = ColName$(y) Then Color y
            If LCase$(Right$(TempS, L - 1)) = ColName$(y) + " blinking" Then
                '->Color Y + 16
            End If
          Next y
          Count = Count - 1
       End If
       If I > 1 Then
          Q = Q + 1
          LL$(Q) = Left$(TempS, I - 1)
          R$(Q) = MyDir$ + Right$(TempS, L - I)
          Row(Q) = Count
          TempS = LL$(Q)
       End If
       If I <> 1 Then CurrentForm.Print TempS
    Loop
    Close #1
    hip = Q: Rem the number of possible choices
End Sub

Sub MakeArrow(y, x)
    Locate y, x
    Color 12
    CurrentForm.Print Chr$(196); Chr$(26);
End Sub

Sub EraseArrow(y, x)
    Locate y, x
    If ClrMode = 2 Or ClrMode = 4 Then
        Color 15
    Else
        Color 0
    End If
    CurrentForm.Print Chr$(196); Chr$(26);
End Sub

Sub MoveArrow(Key1, Key2)
    EraseArrow Row(p(N)), 2
    If Key2 = 1 Then
        If Key1 = 72 Then p(N) = p(N) - 1: Rem up
        If Key1 = 80 Then p(N) = p(N) + 1: Rem down
        If Key1 = 75 Then p(N) = p(N) - 1: Rem left
        If Key1 = 77 Then p(N) = p(N) + 1: Rem right
        If Key1 = 73 Or Key1 = 71 Then p(N) = 1: Rem pgup/home
        If Key1 = 81 Or Key1 = 79 Then p(N) = hip: Rem pgdn/end
    End If
    If Key2 = 0 Then
        If Key1 = 74 Or Key1 = 106 Then p(N) = p(N) + 1: Rem J/j
        If Key1 = 75 Or Key1 = 107 Then p(N) = p(N) - 1: Rem K/k
    End If
    If p(N) > hip Then p(N) = 1
    If p(N) < 1 Then p(N) = hip
    MakeArrow Row(p(N)), 2
End Sub

Sub SelectFile()
    Ext$ = Right$(R$(p(N)), 3)
    If Ext$ = "exe" Or Ext$ = "com" Or Ext$ = "bat" Then
       E$ = R$(p(N))
       ShellPrg E$, vbNormalFocus
    End If
    If Ext$ = "txt" Or Ext$ = "dat" Or Ext$ = "bas" Then
       Color 7: CLS: CurrentForm.Print "."
       E$ = Reader$ + " " + R$(p(N))
       ShellPrg E$, vbMaximizedFocus
    End If
    If Ext$ = "asc" Then
       Color 7: CLS: CurrentForm.Print ".": CLS
       ShellPrg "notepad" + " " + R$(p(N)), vbMaximizedFocus
       If Tmp$ = "type" Then Call GetKeyChar(Key1, Key2)
    End If
    If Ext$ <> "mnu" Then Exit Sub
    Rem determine if we've been there before
    There = 0
    For Z = 1 To topp
       If R$(p(N)) = F$(Z) Then There = Z
    Next Z
    Rem yes we have
    If There <> 0 Then N = There
    Rem no we haven't, so let's make a new place
    If There = 0 Then
       topp = topp + 1
       back(topp) = N
       F$(topp) = R$(p(N))
       p(topp) = 1
       N = topp
    End If
End Sub

