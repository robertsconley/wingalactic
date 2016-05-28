Attribute VB_Name = "modSecMap"
Dim SubName$(16, 2)
Dim world$(80)
Dim basalg$(2, 50, 2)
Dim basalgn(2)
Dim algncol(50)
Dim subtxt%(16)

Public Sub SecMap(Command$)
    If Command$ = "" Then
        CLS
        Color 11
        CurrentForm.Print "This program is not meant to be run as a stand-alone"
        CurrentForm.Print "piece of software. It is part of the 'Galactic' package,"
        CurrentForm.Print "and it meant to be executed from gal.exe."
        CurrentForm.Print
        CurrentForm.Print "Hit any key to Exit"
        Call GetKeyChar(k1, k2)
        GoTo 10000
    End If
4:   '
    tog1 = 1
    tog2 = 1
    tog3 = 1
5 Rem startover
    GoSub 5700: Rem read configuration
    GoSub 5200: Rem calculate video stuff
    GoSub 2200: Rem read sector file
    Screen 12
    GoSub 5800: Rem set palette
    GoTo 6000
'
        Call GetKeyChar(k1, k2)
        GoTo 10000


1100 Rem convert real hex to map hex
Rem in -> rhex
Rem out -> hex
    t1 = Int(rHex / 100)
    t2 = rHex - t1 * 100
1105     If t1 <= 8 Then GoTo 1110
    t1 = t1 - 8
    GoTo 1105
1110     If t2 <= 10 Then GoTo 1115
    t2 = t2 - 10
    GoTo 1110
1115     mHex = t1 * 100 + t2
    Return

2200     Rem get sector data
    F$ = cDir$ + SecFile$
    Open F$ For Input As #1
    Line Input #1, SecName$
    Line Input #1, temp$
    Open GalDir$ + "\" + galdir0$ + ".lst" For Input As #2
    Input #2, T$
    Input #2, T$
    Do Until EOF(2)
        Line Input #2, T$
        If RTrim$(Left$(T$, 8)) = dir0$ Then
            wherex = Val(Mid$(T$, 51, 4))
            wherey = Val(Mid$(T$, 56, 4))
        End If
    Loop
    Close 2

    Open GalDir$ + "\where.dat" For Output As #2
    Print #2, wherex
    Print #2, wherey
    Close 2

    north$ = "<nil>"
    south$ = "<nil>"
    east$ = "<nil>"
    west$ = "<nil>"

    Open GalDir$ + "\" + galdir0$ + ".lst" For Input As #2
    Input #2, T$
    Input #2, T$
    Do Until EOF(2)
        Line Input #2, T$
        tmpx = Val(Mid$(T$, 51, 4))
        tmpy = Val(Mid$(T$, 56, 4))
        dude$ = RTrim$(Left$(T$, 8))
        If Mid$(T$, 61, 1) = "I" Then dude$ = "<nil>"
        If tmpx = wherex + 1 And tmpy = wherey Then east$ = dude$
        If tmpx = wherex - 1 And tmpy = wherey Then west$ = dude$
        If tmpx = wherex And tmpy = wherey + 1 Then south$ = dude$
        If tmpx = wherex And tmpy = wherey - 1 Then north$ = dude$
    Loop
    Close 2

    For I = 1 To 16
        Line Input #1, temp$
        t9 = Len(temp$)
        subtxt%(I) = 0
        If Mid$(temp$, 50, 1) = "f" Then subtxt%(I) = 1
        If Mid$(temp$, 50, 1) = "m" Then subtxt%(I) = 2
        SubName$(I, 1) = RTrim$(Mid$(temp$, 4, 26))
        SubName$(I, 2) = RTrim$(Mid$(temp$, 30, 12))
    Next I
    j = 1
    Line Input #1, temp$
44 Rem get base & status info
    I = 0
    Line Input #1, basalg$(j, I, 0)
    I = 1
45     If EOF(1) = -1 Then GoTo 46
    Line Input #1, temp$
    If temp$ = "" Then GoTo 46
    t9 = Len(temp$)
    If j = 1 Then
        basalg$(1, I, 1) = Left$(temp$, 1)
        basalg$(1, I, 2) = RTrim$(Right$(temp$, t9 - 4))
    End If
    If j = 2 Then
        algncol(I) = Val(Left$(temp$, 2))
        basalg$(2, I, 1) = Mid$(temp$, 4, 2)
        basalg$(2, I, 2) = RTrim$(Right$(temp$, t9 - 8))
    End If
    I = I + 1
    GoTo 45
46     basalgn(j) = I - 1
    If j = 2 Then GoTo 48
    j = j + 1
    GoTo 44
48     Close 1
    Return

5200 Rem calculate video stuff
    xMost = 640
    yMost = 500
Rem screen aspect ratio
Rem scar = 2.4
    Scar = (3 / 4) * (xMost / yMost)
Rem hex radii
    HexA = 25
    HexB = Int((HexA ^ 2 - (0.5 * HexA) ^ 2) ^ 0.5)
Rem corrected for aspect ratio
    cHexA = HexA * Scar
    cHexB = HexB
Rem yank right
    Yank = xMost - Int(cHexA * 1.5 * 8.3333) - 22
    zI = 1
    zJ = 1
    Return

5500 Rem egacolor gold on blue
'    'ShellPrg "xtra\egacolor 00 10"
'    'ShellPrg "xtra\egacolor 07 46"
Return
'
5700 Rem read in settings
    Open "gal.cfg" For Input As #1
        Input #1, T$: Editor$ = Right$(T$, Len(T$) - 7)
        Input #1, T$: Reader$ = Right$(T$, Len(T$) - 7)
        Input #1, T$: ClrMode = Val(Right$(T$, 1))
    Close 1
    C$ = LCase$(Command$)
    C$ = Right$(C$, Len(C$) - 5)
    A = InStr(C$, "\")
    lc = Len(C$)
    galdir0$ = Left$(C$, A - 1)
    GalDir$ = "gals\" + galdir0$
    L = Len(GalDir$)
    Open "gals\gal.lst" For Input As #2
        Input #2, T$
        Input #2, T$
        Do Until EOF(2)
            Input #2, T$
            T$ = RTrim$(T$)
            If T$ = "" Then GoTo 5705
            If Left$(T$, L + 1) = galdir0$ + " " Then
                L = Len(T$)
                GalName$ = Right$(T$, L - 13)
            End If
5705     Loop
    Close 2
    cDir$ = Right$(C$, lc - A)
    NoSector = 0
    If cDir$ = "" Then NoSector = 1
5710     Rem read in sector particulars
    dir0$ = cDir$
    cDir$ = GalDir$ + "\" + cDir$ + "\"
    mapdir$ = cDir$ + "map\"
    locdir$ = cDir$ + "loc\"
    gendir$ = cDir$ + "gen\"
    hexdir$ = cDir$ + "hex\"
    SecFile$ = dir0$ + ".dat"
    Close 1
    L = Len(dir0$)
    If NoSector = 0 Then
        Open GalDir$ + "\" + galdir0$ + ".lst" For Input As #1
            Input #1, T$
            Input #1, T$
            Do Until EOF(1)
                Input #1, T$
                T$ = RTrim$(Left$(T$, 50))
                If T$ <> "" Then
                    If Left$(T$, L + 1) = dir0$ + " " Then
                        L = Len(T$)
                        secname2$ = RTrim$(Mid$(T$, 14, 37))
                    End If
                End If
             Loop
        Close 1
    End If
    Return

5800 Rem set palette
'    Select Case ClrMode
'    Case 1
'    Rem colors on black
'   PALETTE
'Case 2
'   Rem colors on white
'PALETTE:    PALETTE 0, 4144959: PALETTE 15, 0
'Case 3
'   Rem white on black
'   T = 4144959
'   PALETTE 0, 0: PALETTE 1, T: PALETTE 2, T: PALETTE 3, T
'   PALETTE 4, T: PALETTE 5, T: PALETTE 6, T: PALETTE 7, T
'   PALETTE 8, T: PALETTE 9, T: PALETTE 10, T: PALETTE 11, T
'   PALETTE 12, T: PALETTE 13, T: PALETTE 14, T: PALETTE 15, T
'Case 4
'   Rem black on white
'   PALETTE 0, 4144959: PALETTE 1, 0: PALETTE 2, 0: PALETTE 3, 0
'   PALETTE 4, 0: PALETTE 5, 0: PALETTE 6, 0: PALETTE 7, 0
'   PALETTE 8, 0: PALETTE 9, 0: PALETTE 10, 0: PALETTE 11, 0
'   PALETTE 12, 0: PALETTE 13, 0: PALETTE 14, 0: PALETTE 15, 0
'End Select
Return

6000 Rem graphical sector
    CLS
    Rem plot hexes
    Color 1
    For c1 = 1 To 32
        For c2 = 1 To 40
            GoSub 6200: Rem center
            GoSub 6250: Rem hex
        Next c2
    Next c1

Rem plot jump routes
    If tog1 = 0 Then GoTo 6060
    For A = 1 To 16
        fsub$ = mapdir$ + SubName$(A, 2)
        Open fsub$ For Input As #1
            Do Until EOF(1)
                Line Input #1, T$
                If Len(T$) = 0 Then GoTo 6050
                If Left$(T$, 1) <> "$" Then GoTo 6050
                rHex = Val(Mid$(T$, 2, 4)): GoSub 1100
                j1 = mHex
                rHex = Val(Mid$(T$, 7, 4)): GoSub 1100
                j2 = mHex
                j3 = Val(Mid$(T$, 12, 2))
                place = 14
                If j3 = -1 Then place = place + 1
                j4 = Val(Mid$(T$, place, 2))
                place = place + 2
                If j4 = -1 Then place = place + 1
                j5 = Val(Mid$(T$, place, 2))
                If j5 = 0 Then j5 = 11
                c1 = Int(j1 / 100) + (((A - 1) Mod 4) * 8)
                c2 = j1 - (Int(j1 / 100) * 100) + (Int((A - 1) / 4) * 10)
                GoSub 6200: Rem center
                d1 = Cent1: d2 = Cent2
                c1 = Int(j2 / 100) + (((A - 1) Mod 4) * 8) + (j3 * 8)
                c2 = j2 - (Int(j2 / 100) * 100) + (Int((A - 1) / 4) * 10) + (j4 * 10)
                GoSub 6200: Rem center
                d3 = Cent1 + 1: d4 = Cent2 + 1
                CurrentForm.DrawStyle = vbDash
                CurrentForm.Line (d1, d2)-(d3, d4), QBColor(j5)
6050         Loop
        Close 1
    Next A

6060 Rem plot stars
    For A = 1 To 16
        fsub$ = mapdir$ + SubName$(A, 2)
        Open fsub$ For Input As #1
            Do Until EOF(1)
                Line Input #1, T$
                If Len(T$) = 0 Then GoTo 6090
                If InStr("@#$", Left$(T$, 1)) <> 0 Then GoTo 6090
                C$ = Mid$(T$, 15, 4)
                c1 = Val(Left$(C$, 2))
                c2 = Val(Right$(C$, 2))
                s$ = Mid$(T$, 56, 2)
                For i2 = 1 To basalgn(2)
                    If s$ = basalg$(2, i2, 1) Then All = algncol(i2)
                Next i2
                GoSub 6200: Rem center
                tmp14 = 0: tmp15 = 0
                If Mid$(T$, 65, 1) = "h" Then tmp14 = 1
                If Mid$(T$, 65, 1) = "H" Then tmp14 = 2
                If Mid$(T$, 20, 1) = "*" Then tmp15 = Val(Mid$(T$, 21, 1))
                If tmp15 = 0 Or tmp14 < 2 Then CurrentForm.PSet (Cent1 + 1, Cent2), QBColor(All)
                If tmp15 = 0 Then CurrentForm.Circle (Cent1 + 1, Cent2), 1, QBColor(All)
                If tmp15 = 4 And tmp14 < 2 Then
                    Color All
                    x = Cent1 + 1: y = Cent2
                    CurrentForm.PSet (x - 2, y - 2): CurrentForm.PSet (x - 2, y + 2)
                    CurrentForm.PSet (x + 2, y - 2): CurrentForm.PSet (x + 2, y + 2)
                End If
                If Mid$(T$, 63, 1) = "f" Or Mid$(T$, 63, 1) = "m" Then
                    If tog2 = 1 Then
                        Color 4: GoSub 6250: Rem hex
                    End If
                End If
                If tog3 = 0 Then GoTo 6090
                If tmp14 = 2 And tmp15 > 0 Then GoTo 6090
                If Mid$(T$, 49, 1) = "A" Then CurrentForm.Circle (Cent1 + 1, Cent2), 3, QBColor(14)
                If Mid$(T$, 49, 1) = "R" Then CurrentForm.Circle (Cent1 + 1, Cent2), 3, QBColor(12)
                If Mid$(T$, 49, 1) = "B" Then CurrentForm.Circle (Cent1 + 1, Cent2), 3, QBColor(11)
6090:     '
            Loop
        Close 1
    Next A
    Color 14
    Locate 2, 40: CurrentForm.Print "Sector: "; SecName$
    Color 11
    For A = 1 To 16
        Locate 3 + A, 43
        CurrentForm.Print "Subsector ";
        CurrentForm.Print Chr$(A + 64); ": "; SubName$(A, 1)
    Next A
    Color 13
    Locate 21, 40: CurrentForm.Print "Jump Routes:      ";
    If tog1 = 0 Then CurrentForm.Print "Not ";
    CurrentForm.Print "Showing"
    Locate 22, 40: CurrentForm.Print "Red Hexes:        ";
    If tog2 = 0 Then CurrentForm.Print "Not ";
    CurrentForm.Print "Showing"
    Locate 23, 40: CurrentForm.Print "Red/Amber Zones:  ";
    If tog3 = 0 Then CurrentForm.Print "Not ";
    CurrentForm.Print "Showing"
    Locate 30, 70: Color 12: CurrentForm.Print "? = Help";
6095: ' GoSub 900
    Call GetKeyChar(k1, k2)
    If k2 = 0 And (k1 >= 97 And k1 <= 122) Then k1 = k1 - 32
    If k2 = 1 And k1 = 72 Then GoSub 6110: Rem up
    If k2 = 0 And k1 = 75 Then GoSub 6110: Rem k
    If k2 = 1 And k1 = 80 Then GoSub 6120: Rem down
    If k2 = 0 And k1 = 74 Then GoSub 6120: Rem j
    If k2 = 1 And k1 = 75 Then GoSub 6130: Rem left
    If k2 = 0 And k1 = 72 Then GoSub 6130: Rem h
    If k2 = 1 And k1 = 77 Then GoSub 6140: Rem right
    If k2 = 0 And k1 = 76 Then GoSub 6140: Rem l
    If k2 = 0 And k1 = 27 Then GoTo 10000: Rem esc
    If k2 = 0 And k1 = 81 Then GoTo 10000: Rem q
    If k2 = 0 And k1 = 63 Then GoTo 6150: Rem ?
    If k2 = 0 And k1 = 49 Then GoTo 6100: Rem 1
    If k2 = 0 And k1 = 50 Then GoTo 6102: Rem 2
    If k2 = 0 And k1 = 51 Then GoTo 6104: Rem 3
    If k2 = 1 And k1 = 68 Then GoTo 6300: Rem F10
GoTo 6095
6100 Rem 1 / toggle jumproutes
    tog1 = tog1 + 1: If tog1 = 2 Then tog1 = 0
GoTo 6000
6102 Rem 2 / toggle redhexes
    tog2 = tog2 + 1: If tog2 = 2 Then tog2 = 0
GoTo 6000
6104 Rem 3 / toggle zonage
    tog3 = tog3 + 1: If tog3 = 2 Then tog3 = 0
GoTo 6000
6110 Rem up
    If north$ = "<nil>" Then GoTo 6095
    cDir$ = north$: wherey = wherey - 1: GoSub 5710: GoSub 2200: GoTo 6000
6120 Rem down
    If south$ = "<nil>" Then GoTo 6095
    cDir$ = south$: wherey = wherey + 1: GoSub 5710: GoSub 2200: GoTo 6000
6130 Rem left
    If west$ = "<nil>" Then GoTo 6095
    cDir$ = west$: wherex = wherex - 1: GoSub 5710: GoSub 2200: GoTo 6000
6140 Rem right
    If east$ = "<nil>" Then GoTo 6095
    cDir$ = east$: wherex = wherex + 1: GoSub 5710: GoSub 2200: GoTo 6000
6150 Rem graphical sector page commands list
    Screen 0: Color 7: CLS: CurrentForm.Print "."
    GoSub 5500: Rem egacolor
    ShellPrg Reader$ + " help\gsec-cmd.txt"
    Screen 12
GoTo 6000

6200 Rem center
    Cent1 = c1 * 8 + Int((c1 - 1) / 8)
    Cent2 = c2 * 10 + Int((c2 - 1) / 10)
    If c1 Mod 2 = 0 Then Cent2 = Cent2 + 5
Return

6250 Rem plot hex from one o'clock clockwise
    tmp1 = Cent1 + 4: tmp2 = Cent2 - 5
    tmp3 = Cent1 + 6: tmp4 = Cent2
    CurrentForm.Line (tmp1, tmp2)-(tmp3, tmp4)
    CurrentForm.Line -Step(-2, 5)
    CurrentForm.Line -Step(-6, 0)
    CurrentForm.Line -Step(-2, -5)
    CurrentForm.Line -Step(2, -5)
    CurrentForm.Line -Step(6, 0)
Return

6300 Rem save screen
'    outfile$ = dir0$ + ".bmp"
'    horiz = 18 - (Len(secname$) / 2)
'    If horiz < 1 Then horiz = 1
'    Locate 28, horiz: Color 11
'    CurrentForm.Print secname$;
'    savescr outfile$, 0, 0, 269, 449, 4, 0
'    Screen 0: CLS: Locate 5: Color 11
'    CurrentForm.Print
'    CurrentForm.Print "Sector map saved to "; outfile$
'    CurrentForm.Print
'    CurrentForm.Print "Hit any key to continue..."
'    CurrentForm.Print
'    GoSub 900
'    'Screen 12
GoTo 6000
'
10000 Rem end
    Color 7
    CLS
    CurrentForm.Print
    CurrentForm.Print "For Assistance and/or Snide Remarks:"
    CurrentForm.Print
    CurrentForm.Print "     Email: jimv@empirenet.com"
    CurrentForm.Print "            JimVassila@aol.com"
    Exit Sub

End Sub

