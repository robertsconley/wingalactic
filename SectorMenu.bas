Attribute VB_Name = "modSectorMenu"
Private sec$(200, 2)
Private Col(200)
Private Editor$
Private Reader$
Private NumSec
Private Devel
Private Gal0$
Private CurrentPosition

Public Sub SectorMenu(Command As String)
    Dim Key1 As Integer
    Dim Key2 As Integer
    C$ = LCase$(LTrim$(RTrim$(Command$)))
    If C$ <> "" Then
        I = InStr(C$, "\")
        i2 = InStr(I + 1, C$, "\")
        Gal0$ = Mid$(C$, I + 1, i2 - I - 1)
        SecList$ = C$
        
        Call ReadGalCfg
        Call ReadSectorList(SecList$)
        
        CurrentPosition = 1
        help = 1
        Devel = 0
        delfiles = 0
        Call DrawScreen
    End If
    DrawArrow CurrentPosition
    Do
        Call GetKeyChar(Key1, Key2)
        If Key2 = 0 And Key1 = 27 Then Exit Do: Rem esc
        If Key2 = 0 And Key1 = 81 Then Exit Do: Rem Q
        If Key2 = 1 And (Key1 = 75 Or Key1 = 77) Then Call ShiftPosition(Key1): Rem left/right
        If Key2 = 1 And (Key1 = 72 Or Key1 = 80) Then Call ShiftPosition(Key1): Rem up/down
        If Key2 = 1 And (Key1 = 73 Or Key1 = 81) Then Call ShiftPosition(Key1): Rem pgup/pgdn
        If Key2 = 0 And Key1 = 13 Then Call ProcessEnter(Key1): Rem enter
        If Key2 = 0 And Key1 = 68 Then Call DevelopmentToggle: Rem D
        If Key2 = 0 And Key1 = 69 Then Call ProcessEnter(Key1): Rem E
        If Key2 = 0 And (Key1 = 63 Or Key1 = 72) Then Call PrintCommands: Rem H/?
    Loop
    
'10000 Rem end
    Color 7
    If delfiles = 1 Then
        ShellPrg "del tmp1.tmp"
        ShellPrg "del tmp2.tmp"
    End If
    CLS
    CurrentForm.Print
    CurrentForm.Print "For Assistance and/or Snide Remarks:"
    CurrentForm.Print
    CurrentForm.Print "     Email:  jimv@empirenet.com"
    CurrentForm.Print "             jimvassila@aol.com"
    CurrentForm.Print
    CurrentForm.Print "  Homepage:  http://members.aol.com/jimvassila"
End Sub

Private Sub ReadGalCfg()
    Dim TempS As String
    Open "gal.cfg" For Input As #1
    Line Input #1, TempS: Editor$ = Right$(TempS, Len(TempS) - 7)
    Line Input #1, TempS: Reader$ = Right$(TempS, Len(TempS) - 7)
    Close #1
End Sub

Private Sub ReadSectorList(SecList As String)
'700 Rem read list of sectors
    Dim TempS As String
    A = 0
    Open SecList$ For Input As #1
    Input #1, TempS
    Input #1, TempS
    Do Until EOF(1)
        Line Input #1, TempS
        TempS = LTrim$(RTrim$(TempS))
        If TempS <> "" Then
            If Left$(TempS, 1) <> "#" Then
                coltmp = Asc(Mid$(TempS, 61, 1)) - 65
                If coltmp <> 8 Then
                    A = A + 1
                    Col(A) = coltmp
                    sec$(A, 1) = RTrim$(Left$(TempS, 8))
                    sec$(A, 2) = LTrim$(RTrim$(Mid$(TempS, 14, 37)))
                End If
            End If
        End If
    Loop
    Close #1
    NumSec = A
End Sub

Public Sub DrawArrow(CurrentPosition)
'14 Rem arrow
    ar = CurrentPosition Mod 20
    If ar = 0 Then ar = 20
    Locate ar + 2, 2
    Color 12
    CurrentForm.Print Chr$(196); Chr$(26)
End Sub

Public Sub EraseArrow(CurrentPosition)
'14 Rem arrow
    ar = CurrentPosition Mod 20
    If ar = 0 Then ar = 20
    Locate ar + 2, 2: Color 0
    CurrentForm.Print Chr$(196); Chr$(26)
End Sub

Private Sub DevelopmentToggle()
    If Devel = 0 Then Devel = 2
    If Devel = 1 Then Devel = 0
    If Devel = 2 Then Devel = 1
    Call DrawScreen
    Call DrawArrow(CurrentPosition)
End Sub

Private Sub DrawScreen()
    Dim Page
    Screen 0
    
    CLS
    Color 14
    Page = Int((CurrentPosition - 1) / 20) + 1
    CurrentForm.Print "Directory of Sectors     -     Page"; Page;
    If Devel = 1 Then
       Locate , 50: CurrentForm.Print "Gen";
       Locate , 60: CurrentForm.Print "Loc";
    End If
    CurrentForm.Print
    CurrentForm.Print "-------------------------------------";
    If Page = 10 Then CurrentForm.Print "-";
    If Devel = 1 Then
       CurrentForm.Print "--------------------------";
    End If
    CurrentForm.Print: CurrentForm.Print
    Add = (Page - 1) * 20
    For A = 1 + Add To 20 + Add
        Color Col(A)
        CurrentForm.Print "     "; sec$(A, 2);
        If Devel = 1 And sec$(A, 2) <> "" Then
            delfiles = 1
            ShellPrg "dir /s/b gals\" + Gal0$ + "\" + sec$(A, 1) + "\gen > tmp1.tmp"
            ShellPrg "dir /s/b gals\" + Gal0$ + "\" + sec$(A, 1) + "\loc > tmp2.tmp"
            
            Open "tmp1.tmp" For Input As #1: b1 = 0
            Do Until EOF(1)
                Line Input #1, tt$
                b1 = b1 + 1
            Loop
            Close #1
            Open "tmp2.tmp" For Input As #1: b2 = 0
            Do Until EOF(1)
                Line Input #1, tt$
                b2 = b2 + 1
            Loop
            Close 1
            Locate , 50: CurrentForm.Print b1 - 1;
            Locate , 60: CurrentForm.Print b2;
        End If
        CurrentForm.Print
    Next A
    If help = 1 Then
       Color 13: Locate 24, 70
       CurrentForm.Print "(? = Help)";
    End If
    help = 0
End Sub

Private Sub ShiftPosition(Key1 As Integer)
'17 Rem up/down/pgup/pgdn
    oldp = CurrentPosition
    Call EraseArrow(CurrentPosition)
    If Key1 = 72 Or Key1 = 75 Then CurrentPosition = CurrentPosition - 1
    If Key1 = 80 Or Key1 = 77 Then CurrentPosition = CurrentPosition + 1
    If Key1 = 73 Then CurrentPosition = CurrentPosition - 20
    If Key1 = 81 Then CurrentPosition = CurrentPosition + 20
    If CurrentPosition > NumSec Then CurrentPosition = NumSec
    If CurrentPosition < 1 Then CurrentPosition = 1
    If Int((CurrentPosition - 1) / 20) <> Int((oldp - 1) / 20) Then Call DrawScreen
    Call DrawArrow(CurrentPosition)
End Sub

Private Sub ProcessEnter(Key1 As Integer)
    '18 Rem enter
    'If Key1 = 69 Then App$ = Editor$
    Call Eg("gals\" + Gal0$ + "\" + sec$(CurrentPosition, 1) + "\gen", "sector.mnu")
    Call DrawScreen
    Call DrawArrow(CurrentPosition)
End Sub

Private Sub PrintCommands()
'19 Rem commands list
    CLS
    Color 10
    CurrentForm.Print "List of Commands"
    CurrentForm.Print "----------------"
    Color 11
    CurrentForm.Print
    CurrentForm.Print "   <Arrows>   Move Arrow"
    CurrentForm.Print "   <Enter>    View Sector Notes"
    CurrentForm.Print "     <E>      Edit Sector Notes"
    CurrentForm.Print "     <D>      Toggle Development Rating"
    CurrentForm.Print "   <Esc>/Q    Return to Galaxy Map"
    CurrentForm.Print: CurrentForm.Print: Color 10
    CurrentForm.Print "Notes:"
    CurrentForm.Print: Color 11
    CurrentForm.Print "A sector's development rating is a set of two values,"
    CurrentForm.Print "'Gen' & 'Loc'. The former represents the number of"
    CurrentForm.Print "general information files the sector contains. The"
    CurrentForm.Print "latter represents the number of location-specific"
    CurrentForm.Print "files. Hexworld and Star System maps are not included"
    CurrentForm.Print "within either count and are essentially ignored in the"
    CurrentForm.Print "development rating. The whole purpose of this is to"
    CurrentForm.Print "give you a rough idea of the extent to which each"
    CurrentForm.Print "sector has been developed and to point to those which"
    CurrentForm.Print "need further work."
    Call GetKeyChar(Key1, Key2)
    Call DrawScreen
    Call DrawArrow(CurrentPosition)
End Sub





