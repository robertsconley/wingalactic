Attribute VB_Name = "Gal3"

Rem Galactic v2.4
Rem Public Domain 1997 Jim Vassilakos
Rem San Bernardino, California
Rem jimv@empirenet.com
Rem Written in Microsoft QuickBasic version 4.5

Dim SubNameStr(16, 2)
Dim WorldStr(80)
Dim BasAlgStr(2, 50, 2)
Dim basalgn(2)
Dim algncol(50)
Dim LetImage(13, 44) As Integer
Dim LetPicture As Collection
Dim Length(44)
Dim Jump%(300, 5)
Dim subtxt%(16)
Dim ColorsStr(15)
Dim UWPstatStr(15): Rem stat names
Dim UWPnum%(15): Rem number of definitions per stat
Dim UWPpos%(15, 3): Rem positions (main & mod) & color
Dim UWPdefStr(15, 40): Rem value definitions
Dim UWPval%(15, 40, 2): Rem low & high values for each def
Dim nHook%(3)
Dim Hooke%(3, 20)
Dim HookAc%(3, 20)
Dim HookComStr(3, 20)
Dim HookSpecStr(3, 20)
Dim HookVarStr(9)
Dim HookValStr(9)
Dim GalaxyList As colGalaxy
Dim xx(200)
Dim yy(200)
Dim Column(200)
Dim SectorList(200, 2)
Dim PRStr(3)
Dim Scar
Dim HexA
Dim HexB
Dim cHexA
Dim cHexB
Dim Yank
Dim zI
Dim zJ
Dim zK
Public EditorStr As String
Public ReaderStr As String
Dim ClrMode
Dim MainMenuPointer
Dim OldMainMenuPointer
Dim CurrentGalaxy As String
Dim GalDirStr
Dim CurrentGalaxyIndex As Long
Dim HelpToG
Dim SecListStr
Dim xNow
Dim yNow
Dim NumSec
Dim CurrentSector
Dim SecDirStr
Dim MapDirStr
Dim LocDirStr
Dim GenDirStr
Dim HexDirStr
Dim SysDirStr
Dim SecFileStr
Dim SecFPStr
Dim SecOCC
Dim MoveSec
Dim MvSecX
Dim MvSecY
Dim SecNameStr
Dim NumStats
Dim NoSector
Dim NorthStr
Dim SouthStr
Dim WestStr
Dim EastStr
Dim HHex
Dim Trip
Dim fSubStr
Dim NewJX
Dim NewJY
Dim Jumps
Dim rHex
Dim Ref
Dim Cent1
Dim Cent2
Dim NonWorld
Dim NoName
Dim Hidden
Dim HiddenStr
Dim SubLetStr
Dim CoordStr
Dim Trip2
Dim HWStr
Dim UWPtog
Dim NewJumpStr
Dim NewJFStr
Dim HHStr
Public CurrentForm As Form
Dim Hex
Public Inkey As String

Sub Main()
    Set FSO = New FileSystemObject
    ChDrive D
    ChDir App.Path
    Set CurrentForm = frmGal
    frmGal.Show
    Screen 0
    'Exit Sub
    On Error GoTo 9900
    Randomize Int(Timer / 2)
    Set GalaxyList = New colGalaxy
    MainMenuPointer = 1:  OldMainMenuPointer = 1
    HookVarStr(1) = "<reader>"
    HookVarStr(2) = "<editor>"
    HookVarStr(3) = "<galdir>"
    HookVarStr(4) = "<secdir>"
    HookVarStr(5) = "<seclist>"
    HookVarStr(6) = "<secfilepath>"
    HookVarStr(7) = "<subfilepath>"
    HookVarStr(8) = "<sublet>"
    HookVarStr(9) = "<coord>"
    
    Call VideoInit: Rem calculate video stuff
    Call ReadColors: Rem read names of 15 colors for insert jumproute query
    Call ReadHooks: Rem read hooks
    
    Call ReadGalCfg: Rem read gal.cfg
    Call ReadLetterData: Rem read letters.dat -> image data
        
    Call DrawFrontPage
    Call LocateArrow(MainMenuPointer)
    Do
        Call GetKeyChar(Key1, Key2)
        If Key2 = 0 And (Key1 = 27 Or Key1 = 81) Then ' Rem esc
            Call EndProgram
            Exit Do
        End If
        If Key2 = 1 And (Key1 = 75 Or Key1 = 77) Then Call ProcessArrow(Key1): Rem left/right
        If Key2 = 1 And (Key1 = 72 Or Key1 = 80) Then Call ProcessArrow(Key1): Rem up/down
        If Key2 = 1 And (Key1 = 73 Or Key1 = 81) Then Call ProcessArrow(Key1): Rem pgup/pgdn
        If Key2 = 0 And Key1 = 13 Then ' Rem enter
            If MainMenuPointer = 1 Then SelectGalaxy
            If MainMenuPointer = 2 Then Call ExecuteHelp1
            If MainMenuPointer >= 3 And MainMenuPointer <= 5 Then Call ExecuteChangeSetting(MainMenuPointer)
            If MainMenuPointer = 6 Then
                Call EndProgram
                Exit Do
            End If
        End If
    Loop
    End
    
9900     Rem error
    Color 7: CLS
    CurrentForm.Print "Error #"; LTrim$(Str$(Err.Number)); " detected around line ";
    CurrentForm.Print LTrim$(Str$(Erl)); ". Now Exiting Galactic."
    CurrentForm.Print
    CurrentForm.Print "Please report this error to jimv@empirenet.com"
    CurrentForm.Print "Let me know where it occured so I can examine the"
    CurrentForm.Print "code and see if there's a bug that needs squashing."
    CurrentForm.Print
    CurrentForm.Print "For the latest patches to this software, see my"
    CurrentForm.Print "homepage at http://members.aol.com/jimvassila, and"
    CurrentForm.Print "for the latest information about Galactic, join our"
    CurrentForm.Print "mailing list (see help\gml\signon.txt)."
    CurrentForm.Print
    Call GetKeyChar(Key1, Key2)
    
    End
    
    Rem data

End Sub

Function xStr$(f21)
    xStr$ = LTrim$(Str$(f21))
End Function

Function HexToNumber(f13Str)
   f13Str = RTrim$(LTrim$(f13Str))
   If Len(f13Str) = 1 Then
      Select Case Asc(f13Str)
         Case 48 To 57
            f12 = Val(f13Str)
         Case 65 To 90
            f12 = Asc(f13Str) - 55
         Case Else
            f12 = Asc(f13Str) + 100
      End Select
   End If
   If Len(f13Str) = 2 Then
      f13aStr = Left$(f13Str, 1)
      f13bStr = Right$(f13Str, 1)
      f12 = (Asc(f13aStr) * 256) + Asc(f13bStr)
   End If
   If Len(f13Str) <> 1 And Len(f13Str) <> 2 Then
      f12 = 0
   End If
   HexToNumber = f12
End Function

Function FNuwpStr(f2)
    f3len = 1
    If UWPval%(f2, 1, 1) > 1000 Then f3len = 2
    f3aStr = Mid$(WorldStr(Ref), UWPpos%(f2, 1), f3len)
    If f3aStr = "" Or f3aStr = "  " Then f3aStr = " "
    f8 = Len(WorldStr(Ref))
    If UWPpos%(f2, 2) <= f8 And UWPpos%(f2, 2) <> 0 Then
       f3bStr = Mid$(WorldStr(Ref), UWPpos%(f2, 2), 1)
    Else
       f3bStr = " "
    End If
    f4a = HexToNumber(f3aStr)
    f6Str = ""
    For f5 = 1 To UWPnum%(f2)
       If f4a >= UWPval%(f2, f5, 1) And f4a <= UWPval%(f2, f5, 2) Then
          f6Str = UWPdefStr(f2, f5)
       End If
    Next f5
    If f6Str = "" Then f6Str = "Unknown"
    If f3aStr = " " Then f6Str = "Empty"
    If f3bStr <> " " Then
       f7 = InStr(f6Str, "?")
       f8 = Len(f6Str)
       If f7 = 1 Then f6Str = f3bStr + Right$(f6Str, f8 - 1)
       If f7 > 1 And f8 = f7 Then f6Str = Left$(f6Str, f8 - 1) + f3bStr
       If f7 > 1 And f8 <> f7 Then
          f6aStr = Left$(f6Str, f7 - 1)
          f6bStr = Right$(f6Str, f8 - f7)
          f6Str = f6aStr + f3bStr + f6bStr
       End If
    End If
    FNuwpStr = f6Str
End Function

Function FNhookStr(f21Str)
    HookValStr(1) = ReaderStr
    HookValStr(2) = EditorStr
    HookValStr(3) = GalDirStr
    HookValStr(4) = SecDirStr
    HookValStr(5) = SecListStr
    HookValStr(6) = SecFPStr
    HookValStr(7) = fSubStr
    HookValStr(8) = SubLetStr
    HookValStr(9) = CoordStr
    For f5 = 1 To 9
       L = Len(f21Str)
       f22 = InStr(f21Str, HookVarStr(f5))
       If f22 > 0 Then
          l2 = Len(HookVarStr(f5))
          f23Str = Left$(f21Str, f22 - 1)
          f24Str = Right$(f21Str, L - f22 - l2 + 1)
          f21Str = f23Str + HookValStr(f5) + f24Str
       End If
    Next f5
    FNhookStr = f21Str
End Function

Function FNsec(f3x, f3y)
    f30 = 0
    For f31 = 1 To NumSec
       If f3x = xx(f31) And f3y = yy(f31) Then f30 = f31
    Next f31
    FNsec = f30
End Function

Sub VideoInit()
    Dim xMost
    Dim yMost
    Rem calculate video stuff
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
End Sub

Sub ReadColors()
    Rem read 15 colors
    Open "xtra\colors.dat" For Input As #1
    Line Input #1, tmpStr
    For A = 1 To 15
       Line Input #1, tmpStr
       ColorsStr(A) = Right$(tmpStr, Len(tmpStr) - 4)
    Next A
    Close 1
End Sub

Sub ReadHooks()
Rem hooks
    Open "data\hooks.dat" For Input As #1
    Do Until EOF(1)
        Line Input #1, tStr
        tStr = RTrim$(tStr)
        If tStr <> "" Then
            Select Case tStr
            Case "Galaxy Map:"
                A = 1
                nHook%(1) = 0
            Case "Text Sector Map:"
                A = 2
                nHook%(2) = 0
            Case "Subsector Map:"
                A = 3
                nHook%(3) = 0
            Case Else
                If Left$(tStr, 1) <> "#" Then
                    L = Len(tStr)
                    nHook%(A) = nHook%(A) + 1
                    If Mid$(tStr, 10, 1) = "E" Then Hooke%(A, nHook%(A)) = 1
                    If Mid$(tStr, 10, 1) = "-" Then Hooke%(A, nHook%(A)) = 0
                    HookAc%(A, nHook%(A)) = Val(Mid$(tStr, 11, 3))
                    If L > 59 Then
                       HookComStr(A, nHook%(A)) = RTrim$(Mid$(tStr, 15, 45))
                    Else
                       HookComStr(A, nHook%(A)) = RTrim$(Right$(tStr, L - 14))
                    End If
                    If L > 64 Then
                       HookSpecStr(A, nHook%(A)) = RTrim$(Mid$(tStr, 60, 5))
                    ElseIf L > 59 And L <= 64 Then
                       HookSpecStr(A, nHook%(A)) = RTrim$(Right$(tStr, L - 59))
                    ElseIf L <= 59 Then
                       HookSpecStr(A, nHook%(A)) = ""
                    End If
                End If
            End Select
        End If
    Loop
    Close 1
End Sub

Sub ReadGalCfg()
Rem read configuration info
    Open "gal.cfg" For Input As #1
    Line Input #1, tStr: EditorStr = Right$(tStr, Len(tStr) - 7)
    Line Input #1, tStr: ReaderStr = Right$(tStr, Len(tStr) - 7)
    Line Input #1, tStr: ClrMode = Val(Right$(tStr, 1))
    Close 1
End Sub

Sub ReadLetterData()
Rem get letter image data
    Dim lRowStr(5)
    Dim p As StdPicture
    Set LetPicture = New Collection
    Screen 12
    Call DisplayLoadingMsg
    Open "data\letters.dat" For Input As #1
    For Z = 1 To 44
        Input #1, tStr
        For A = 1 To 5
            Input #1, lRowStr(A)
        Next A
        Rem cls
        Color 15
        Call ConvertLetterData(lRowStr)
        LetPicture.Add ToMono(frmGal)
        '->GET (1, 1)-(l, 5), letimage(1, z)
        Color 0
        Call ConvertLetterData(lRowStr)
    Next Z
    
    Close 1
    Screen 0
End Sub

Sub ConvertLetterData(lRowStr())
    For A = 1 To 5
        Length(Z) = Len(lRowStr(A)) - 2
        L = Length(Z)
        For B = 1 To L
            If Mid$(lRowStr(A), B + 1, 1) = "*" Then CurrentForm.PSet (B - 1, A - 1)
        Next B
    Next A
End Sub

Sub DisplayLoadingMsg()
Rem loading data message
    CLS
    Color 14
    Locate 8, 3
    CurrentForm.Print "Loading Data..."
End Sub

Sub GetKeyChar(Key1, Key2)
    Key1 = 0
    Key2 = 0
    Do
        kStr = Inkey$
        Inkey$ = ""
        DoEvents
    Loop While kStr = ""
    Key1 = Asc(kStr)
    If Key1 = 0 Then
        Key1 = Asc(Right$(kStr, 1))
        Key2 = 1
        If Key2 = 0 And (Key1 >= 97 And Key1 <= 122) Then Key1 = Key1 - 32
        DoEvents
    End If
End Sub

Public Sub EndProgram()
Rem end
    Color 7
    CLS
    CurrentForm.Print
    CurrentForm.Print "For Assistance and/or Snide Remarks:"
    CurrentForm.Print
    CurrentForm.Print "     Email:  jimv@empirenet.com"
    CurrentForm.Print "             jimvassila@aol.com"
    CurrentForm.Print
    CurrentForm.Print "  Homepage:  http://members.aol.com/jimvassila"
    CurrentForm.Print
End Sub

Sub LocateArrow(Location)
    ErasePointer MainSelectionToRow(OldMainMenuPointer), 3
    PrintPointer MainSelectionToRow(Location), 3, 12
    OldMainMenuPointer = Location
End Sub

Function MainSelectionToRow(ByVal Selection As Integer) As Integer
    Dim Row As Integer
    If Selection = 1 Then Row = 1
    If Selection = 2 Then Row = 2
    If Selection = 3 Then Row = 4
    If Selection = 4 Then Row = 5
    If Selection = 5 Then Row = 6
    If Selection = 6 Then Row = 8
    MainSelectionToRow = Row
End Function

Sub ErasePointer(ByVal y As Double, ByVal x As Double)
    Dim OldColor As Long
    OldColor = CurrentForm.ForeColor
    Select Case ClrMode
     Case 1:
        CurrentForm.ForeColor = vbBlack
     Case 2:
        CurrentForm.ForeColor = vbWhite
     Case 3:
        CurrentForm.ForeColor = vbBlack
     Case 4:
        CurrentForm.ForeColor = vbWhite
    End Select
    Locate y, x
    CurrentForm.Print Chr$(196); Chr$(26)
    CurrentForm.ForeColor = OldColor
End Sub

Sub PrintPointer(ByVal y As Double, ByVal x As Double, Optional ByVal PointerColor)
    If Not IsMissing(PointerColor) Then Color PointerColor
    Locate y, x
    CurrentForm.Print Chr$(196); Chr$(26)
End Sub

Sub ProcessArrow(Key1)
    Rem up/down/pgup/pgdn
    If Key1 = 72 Or Key1 = 75 Then MainMenuPointer = MainMenuPointer - 1
    If Key1 = 80 Or Key1 = 77 Then MainMenuPointer = MainMenuPointer + 1
    If Key1 = 73 Then MainMenuPointer = 1
    If Key1 = 81 Then MainMenuPointer = 0
    If MainMenuPointer = 0 Then MainMenuPointer = 6
    If MainMenuPointer = 7 Then MainMenuPointer = 1
    '->SOUND 10000, 0.1
    Call LocateArrow(MainMenuPointer)
End Sub

Sub DrawFrontPage()
    Screen 0
    CLS
    Color 14
    'ShellPrg "sub\funky data\title.dat"
    
    CurrentForm.Print
    Color 11
    CurrentForm.Print "      Starmaps"
    CurrentForm.Print "      About this Program"
    CurrentForm.Print
    CurrentForm.Print "      Select Text Editor     ";
    Color 13: CurrentForm.Print "("; EditorStr; ")"
    Color 11: CurrentForm.Print "      Select Text Reader     ";
    Color 13: CurrentForm.Print "("; ReaderStr; ")"
    Color 11: CurrentForm.Print "      Foreground/Background  ";
    Color 13
    Select Case ClrMode
     Case 1:
      CurrentForm.Print "(colors on black)"
     Case 2:
      CurrentForm.Print "(colors on white)"
     Case 3:
      CurrentForm.Print "(white on black)"
     Case 4:
      CurrentForm.Print "(black on white)"
    End Select
    Color 11
    CurrentForm.Print
    CurrentForm.Print "      Quit"
    Color 9
    CurrentForm.Print
    CurrentForm.Print "Traveller is a registered trademark of Far Future Enterprises."
    CurrentForm.Print "Portions of this software are Copyright 1977-1998 Far Future ";
    CurrentForm.Print "Enterprises."
End Sub

Public Sub ExecuteHelp1()
    Rem Help #1
    EgMain "help\help.mnu"
End Sub

Public Sub ExecuteChangeSetting(MainMenuPointer)
    Rem change settings
    If MainMenuPointer = 3 Then Cfg 1
    If MainMenuPointer = 4 Then Cfg 2
    If MainMenuPointer = 5 Then Cfg 3
    Call ReadGalCfg: Rem read gal.cfg
    Call ReadLetterData: Rem read letters.dat -> image data
    Call DrawFrontPage
    Call LocateArrow(MainMenuPointer)
End Sub

Sub ReadGalaxyList()
    Rem read galaxy list
    Open "gals\gal.lst" For Input As #1
    Line Input #1, tStr
    Line Input #1, tStr
    GalaxyList.Clear
    Do Until EOF(1)
        Line Input #1, tStr
        tStr = RTrim$(tStr)
        If tStr <> "" Then
            GalaxyList.Add
            GalaxyList.Last.Directory = LCase$(RTrim$(Left$(tStr, 8)))
            GalaxyList.Last.GalaxyName = Right$(tStr, Len(tStr) - 13)
        End If
    Loop
    Close 1
End Sub

Public Sub ReadGalaxyWhereDat()
    Rem read gals\where.dat
    Open "gals\where.dat" For Input As #1
    Line Input #1, CurrentGalaxy
    tmpStr = LCase$(LTrim$(RTrim$(CurrentGalaxy)))
    Close 1
    CurrentGalaxyIndex = 1
    For Count = 1 To GalaxyList.Count
        If GalaxyList(Count).Directory = tmpStr Then CurrentGalaxyIndex = Count
    Next Count
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    GalDirStr = "gals\" + CurrentGalaxy
End Sub

Sub WriteGalaxyWhereDat()
    Rem write gals\where.dat
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    Open "gals\where.dat" For Output As #1
    Print #1, CurrentGalaxy
    Close 1
End Sub

Public Sub DrawSelectGalaxy()
    Screen 0
    CLS
    Color 10
    CurrentForm.Print "Galaxy Menu"
    CurrentForm.Print
    Color 11
    If GalaxyList.Count > 0 Then
        For I = 1 To GalaxyList.Count
            CurrentForm.Print "    "; GalaxyList(I).GalaxyName
        Next I
    Else
        CurrentForm.Print "     <no galaxies present>"
    End If
    Color 13
    If HelpToG = 1 Then
        Locate 3, 40: CurrentForm.Print "<Arrows>  Move Arrow"
        Locate 4, 40: CurrentForm.Print "<Enter>   Select Galaxy"
        Locate 5, 40: CurrentForm.Print " <Ins>    Create New Galaxy"
        Locate 6, 40: CurrentForm.Print " <Del>    Delete Galaxy"
        Locate 7, 40: CurrentForm.Print " F1/F2    Read/Edit Galaxy Notes"
        Locate 8, 40: CurrentForm.Print "   F5     Sector Menu"
        Locate 9, 40: CurrentForm.Print "   I      Import Galaxy from .gar file"
        Locate 10, 40: CurrentForm.Print "   X      Export Galaxy to .gar file"
        Locate 11, 40: CurrentForm.Print "   <      Edit gals\gal.lst"
        Locate 12, 40: CurrentForm.Print "   >      Edit gals\<galaxy>\<galaxy>.lst"
        Locate 13, 40
    Else
        Locate 3, 40
    End If
    CurrentForm.Print "   ?      Help Toggle"
End Sub

Sub DrawSelectGalaxyArrow()
    Locate CurrentGalaxyIndex + 1, 2
    Color 12
    CurrentForm.Print Chr$(196); Chr$(26)
End Sub

Sub ProcessSelectGalaxyArrow(Key1)
    Locate CurrentGalaxyIndex + 1, 2
    Color 0
    CurrentForm.Print Chr$(196); Chr$(26)
    If Key1 = 72 Or Key1 = 75 Then CurrentGalaxyIndex = CurrentGalaxyIndex - 1
    If Key1 = 80 Or Key1 = 77 Then CurrentGalaxyIndex = CurrentGalaxyIndex + 1
    If Key1 = 73 Then CurrentGalaxyIndex = 1
    If Key1 = 81 Then CurrentGalaxyIndex = GalaxyList.Count
    If CurrentGalaxyIndex = 0 Then CurrentGalaxyIndex = GalaxyList.Count
    If CurrentGalaxyIndex > GalaxyList.Count Then CurrentGalaxyIndex = 1
    '->SOUND 10000, 0.1
    DrawSelectGalaxyArrow
End Sub

Sub DisplayEGAColor()
    Rem egacolor gold on blue
    'Color 7
    'CLS
    'CurrentForm.Print "."
    'ShellPrg "xtra\egacolor 00 10"
    'ShellPrg "xtra\egacolor 07 46"
End Sub

Sub EditGalLst()
    Call DisplayEGAColor: Rem egacolor
    ShellPrg EditorStr + " gals\gal.lst", vbMaximizedFocus
    Screen 12
    Call ReadGalaxyList
    Call ReadGalaxyWhereDat
    Call DrawSelectGalaxy
    Call DrawSelectGalaxyArrow
End Sub

Sub GalaxyHelpToggle()
    Rem help toggle
    If HelpToG = 0 Then HelpToG = 2
    If HelpToG = 1 Then HelpToG = 0
    If HelpToG = 2 Then HelpToG = 1
    Call DrawSelectGalaxy
    Call DrawSelectGalaxyArrow
End Sub

Sub CreateGalaxy()
    Rem create new galaxy
    'ShellPrg "sub\makegal"
    Call MakeGalaxy
    Call ReadGalaxyList
    Call ReadGalaxyWhereDat
    Call DrawSelectGalaxy
    Call DrawSelectGalaxyArrow
End Sub

Sub ExportToArchive()
    Rem export to archive
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    ShellPrg "sub\arch " + CurrentGalaxy
End Sub

Sub ImportFromArchive()
    Rem import from archive
    Do
        ShellPrg "xtra\ls *.gar > tmp.tmp"
        Open "tmp.tmp" For Input As #1
        L = LOF(1)
        Close 1
        If L = 0 Then
           ShellPrg "del tmp.tmp"
           Color 11: CLS
           CurrentForm.Print "There are no .gar files available for importing"
           Call GetKeyChar(Key1, Key2)
           Exit Sub
        Else
            Color 11: CLS: CurrentForm.Print "Files available for importing:": CurrentForm.Print
            ShellPrg "xtra\ls *.gar"
            Color 10: CurrentForm.Print
            tStr = InputText("Name of file to be imported: ")
            tStr = LTrim$(RTrim$(LCase$(tStr)))
            If tStr = "" Then
                ShellPrg "del tmp.tmp"
                Exit Sub
            Else
                If Len(tStr) < 4 Then tStr = tStr + ".gar"
                If Right$(tStr, 4) <> ".gar" Then tStr = tStr + ".gar"
                Ok = 0
                Open "tmp.tmp" For Input As #1
                Do Until EOF(1)
                    Line Input #1, tmpStr
                    If tStr = tmpStr Then Ok = 1
                Loop
                Close 1
                ShellPrg "del tmp.tmp"
            End If
        End If
    Loop Until Ok = 1
    ShellPrg "sub\arch -d " + tStr
    Call ReadGalaxyList
    Call ReadGalaxyWhereDat
    Call DrawSelectGalaxy
    Call DrawSelectGalaxyArrow
End Sub

Public Sub EditGalaxyLst()
    Rem edit gals\<galaxy>\<galaxy>.lst
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    GalDirStr = "gals\" + CurrentGalaxy
    SecListStr = GalDirStr + "\" + CurrentGalaxy + ".lst"
    Call DisplayEGAColor: Rem egacolor
    ShellPrg EditorStr + " " + SecListStr, vbMaximizedFocus
    Screen 12
End Sub

Public Sub GalaxyReadMenu()
    Rem F1
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    EgMain "gals\" + CurrentGalaxy + "\gen\galaxy.mnu"
End Sub

Public Sub GalaxyEditMenu()
    Rem F2
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    EgMain "gals\" + CurrentGalaxy + "\gen\galaxy.mnu"
End Sub

Public Sub GalaxySectorMenu()
    Rem sector menu
    Dim SecListStr  As String
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    GalDirStr = "gals\" + CurrentGalaxy
    SecListStr = GalDirStr + "\" + CurrentGalaxy + ".lst"
    Call SectorMenu(SecListStr)
    Call DrawSelectGalaxy
    Call DrawSelectGalaxyArrow
End Sub

Public Sub DeleteGalaxy()
    Rem delete galaxy
    CLS
    Locate 5
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    GalDirStr = "gals\" + CurrentGalaxy
    CLS
    Color 11: Locate 5
    CurrentForm.Print "Are you sure you want to delete ";
    CurrentForm.Print GalaxyList(CurrentGalaxyIndex).GalaxyName; "? (y/n)"
    ynStr = InputYN
    If ynStr = "y" Then
        CLS
        FSO.DeleteFolder GalDirStr, True
        'ShellPrg "xtra\rm -r " + GalDirStr
        Open "gals\gal.lst" For Input As #1
        Open "gals\tmp.tmp" For Output As #2
        Do Until EOF(1)
           Line Input #1, tStr
           If RTrim$(Left$(tStr, 8)) <> CurrentGalaxy Then Print #2, tStr
        Loop
        Close 1
        Close 2
        FSO.DeleteFile "gals\gal.lst", True
        FSO.CopyFile "gals\tmp.tmp", "gals\gal.lst"
        FSO.DeleteFile "gals\tmp.tmp", True
        Rem modify gals\where.dat
        If CurrentGalaxyIndex = GalaxyList.Count Then CurrentGalaxyIndex = CurrentGalaxyIndex - 1
        If CurrentGalaxyIndex < 0 Then CurrentGalaxyIndex = 1
        Call WriteGalaxyWhereDat: Rem write gals\where.dat
        Call ReadGalaxyList
        Call ReadGalaxyWhereDat
        Call DrawSelectGalaxy
        Call DrawSelectGalaxyArrow
    End If
End Sub

Function InputYN()
    Rem y/n
    Do
        ynStr = ""
        Call GetKeyChar(Key1, Key2)
        If Key2 = 0 And Key1 = 78 Then ynStr = "n": Rem N
        If Key2 = 0 And Key1 = 89 Then ynStr = "y": Rem Y
    Loop While ynStr = ""
    InputYN = ynStr
End Function

Sub ReadSectorWhereDat()
    Rem read gals\<sector>\where.dat
    Open GalDirStr + "\where.dat" For Input As #1
    Input #1, xNow
    Input #1, yNow
    Close 1
End Sub

Sub WriteSectorWhereDat()
    Rem write gals\<sector>\where.dat
    Open GalDirStr + "\where.dat" For Output As #1
    Print #1, xNow
    Print #1, yNow
    Close 1
End Sub

Sub ReadSectorList()
    Rem read list of sectors
    For A = 1 To 200
        xx(A) = 0
        yy(A) = 0
        Column(A) = 0
        SectorList(A, 1) = ""
        SectorList(A, 2) = ""
    Next A
    A = 0
    SecListStr = GalDirStr + "\" + CurrentGalaxy + ".lst"
    Open SecListStr For Input As #1
    Input #1, tStr
    Input #1, tStr
    Do Until EOF(1)
        Line Input #1, tStr
        tStr = LTrim$(RTrim$(tStr))
        If tStr <> "" And Left$(tStr, 1) <> "#" Then
            A = A + 1
            xx(A) = Val(Mid$(tStr, 51, 4))
            yy(A) = Val(Mid$(tStr, 56, 4))
            Column(A) = Asc(Mid$(tStr, 61, 1)) - 65
            SectorList(A, 1) = RTrim$(Left$(tStr, 8))
            SectorList(A, 2) = LTrim$(RTrim$(Mid$(tStr, 14, 37)))
        End If
    Loop
    Close 1
    NumSec = A
End Sub

Sub ReadSectorPaths()
    Rem determine directory/file paths
    Rem CurrentGalaxy & CurrentSector are known
    GalDirStr = "gals\" + CurrentGalaxy
    SecListStr = GalDirStr + "\" + CurrentGalaxy + ".lst"
    SecDirStr = GalDirStr + "\" + CurrentSector
    MapDirStr = SecDirStr + "\" + "map\"
    LocDirStr = SecDirStr + "\" + "loc\"
    GenDirStr = SecDirStr + "\" + "gen\"
    HexDirStr = SecDirStr + "\" + "hex\"
    SysDirStr = SecDirStr + "\" + "sys\"
    SecFileStr = CurrentSector + ".dat"
    SecFPStr = SecDirStr + "\" + SecFileStr
End Sub

Sub DrawSectorMap()
    Screen 12: CLS
    For A = xNow - 3 To xNow + 3
        For B = yNow - 3 To yNow + 3
            loc1 = (A - xNow + 4) * 11 - 10
            loc2 = (B - yNow + 4) * 4 - 3
            hit = FNsec(A, B)
            Locate loc2, loc1: Color 9
            CurrentForm.Print Chr$(197); String$(10, 196); Chr$(197);
            If hit = 0 Then
               For D = 1 To 3
                   Locate loc2 + D, loc1: Color 9
                   CurrentForm.Print Chr$(179); String$(10, 32); Chr$(179);
                Next D
            Else
               For C = 1 To 3
                  PRStr(C) = ""
               Next C
               C = 1: D = 0
               For E = 1 To Len(SectorList(hit, 2))
                    If Mid$(SectorList(hit, 2), E, 1) = " " Then
                        C = C + 1: D = 0
                    Else
                        D = D + 1
                        If C <= 3 And D <= 10 Then
                            PRStr(C) = PRStr(C) + Mid$(SectorList(hit, 2), E, 1)
                        End If
                    End If
               Next E
               For E = 1 To 3
                    PRStr(E) = Space$(Int(10 - Len(PRStr(E))) / 2 - 0.1) + PRStr(E)
                    PRStr(E) = PRStr(E) + Space$(10 - Len(PRStr(E)))
               Next E
               If LTrim$(PRStr(2)) = "" Then Swap PRStr(1), PRStr(2)
               For D = 1 To 3
                    Locate loc2 + D, loc1
                    Color 9: CurrentForm.Print Chr$(179);
                    Color Column(hit): CurrentForm.Print PRStr(D);
                    Color 9: CurrentForm.Print Chr$(179);
               Next D
            End If
        Next B
    Next A
    For D = 1 To 7
        Locate 29, D * 11 - 10: Color 9
        CurrentForm.Print Chr$(197); String$(10, 196); Chr$(197);
    Next D
    Locate 13, 34: Color 11
    CurrentForm.Print Chr$(201); String$(10, 205); Chr$(187);
    For D = 14 To 16
        Locate D, 34: CurrentForm.Print Chr$(186);
        Locate D, 45: CurrentForm.Print Chr$(186)
    Next D
    Locate 17, 34
    CurrentForm.Print Chr$(200); String$(10, 205); Chr$(188);
    If MoveSec <> 0 Then
        If Abs(MvSecX - xNow) < 4 And Abs(MvSecY - yNow) < 4 Then
           loc1 = (MvSecX - xNow + 4) * 11 - 10
           loc2 = (MvSecY - yNow + 4) * 4 - 3
           Color 13
           Locate loc2, loc1
           CurrentForm.Print Chr$(218); String$(10, 196); Chr$(191);
           For D = 1 To 3
                For d2 = 0 To 11 Step 11
                    Locate loc2 + D, loc1 + d2
                    CurrentForm.Print Chr$(179);
                Next d2
           Next D
           Locate loc2 + 4, loc1
           CurrentForm.Print Chr$(192); String$(10, 196); Chr$(217);
        End If
    End If
        
    If MoveSec <> 0 Then
        Locate 29, 10: Color 13
        CurrentForm.Print "To move sector: move to empty cell and hit <F6> again";
    End If
    Locate 29, 70: Color 12: CurrentForm.Print "?=Help";
    SecOCC = FNsec(xNow, yNow)
    If SecOCC <> 0 Then
        CurrentSector = SectorList(SecOCC, 1)
        Call ReadSectorPaths: Rem assign dirs
    End If
End Sub

Sub SectorMapUp()
    yNow = yNow - 1
    DrawSectorMap
End Sub

Sub SectorMapDown()
    yNow = yNow + 1
    DrawSectorMap
End Sub

Sub SectorMapLeft()
    xNow = xNow - 1
    DrawSectorMap
End Sub

Sub SectorMapRight()
    xNow = xNow + 1
    DrawSectorMap
End Sub

Sub ShiftOfSector()
    Rem shift of sector
    Rem we know CurrentGalaxy & galdirStr
    Rem we know xnow & ynow
    CurrentSector = SectorList(FNsec(xNow, yNow), 1)
    Call ReadSectorPaths: Rem file/directory paths
End Sub

Sub AddSectorNameToMap()
    Rem add sector name to map
    Tmp = FNsec(xNow, yNow)
    If Tmp > 0 Then
        CLS
        Color 11
        Do
            tmp2Str = GetSectorName
            If tmp2Str = "" Then Exit Do
            tmp1Str = GetDirectoryName
        Loop While tmp1Str = ""
        If tmp2Str <> "" And tm12Str <> "" Then
            seclineStr = tmp1Str + Space$(13 - Len(tmp1Str))
            seclineStr = seclineStr + tmp2Str + Space$(37 - Len(tmp2Str))
            tmp3Str = xStr$(xNow)
            tmp3Str = tmp3Str + Space$(4 - Len(tmp3Str))
            tmp4Str = xStr$(yNow)
            tmp4Str = tmp4Str + Space$(4 - Len(tmp4Str))
            seclineStr = seclineStr + tmp3Str + " " + tmp4Str + " I"
            Open SecListStr For Append As #1
            Print #1, seclineStr
            Close 1
            Call ReadSectorList: Rem read list of sectors
        End If
    End If
    Call DrawSectorMap
End Sub

Function GetSectorName()
    Do
        Locate 5
        CurrentForm.Print Space$(79)
        Locate 5
        tmp2 = 0
        tmp2Str = InputText("Sector Name (35 chars max): ")
        tmp2Str = LTrim$(RTrim$(tmp2Str))
        If tmp2Str = "" Then
            Exit Do
        Else
            If Len(tmp2Str) > 35 Then tmp2 = 1
            For tmp1 = 1 To NumSec
               If SectorList(tmp1, 2) = tmp2Str Then tmp2 = 2
            Next tmp1
            Locate 7: CurrentForm.Print Space$(79): Locate 7
            If tmp2 = 1 Then CurrentForm.Print "(35 chararcters max)"
            If tmp2 = 2 Then CurrentForm.Print "(sector name already taken)"
            If tmp2 = 0 Then Exit Do
        End If
    Loop
    GetSectorName = tmp2Str
End Function

Function GetDirectoryName()
    Do
        Locate 10
        CurrentForm.Print Space$(79)
        Locate 10
        tmp2 = 0
        tmp1Str = InputText("Sector Directory (8 chars max): ")
        tmp1Str = LTrim$(RTrim$(LCase$(tmp1Str)))
        If tmp1Str = "" Then
            Exit Do
        Else
            If Len(tmp1Str) > 8 Then tmp2 = 1
            For tmp1 = 1 To NumSec
               If SectorList(tmp1, 1) = tmp1Str Then tmp2 = 2
            Next tmp1
            If InStr(tmp1Str, " ") > 0 Then tmp2 = 3
            Locate 12: CurrentForm.Print Space$(79): Locate 12
            If tmp2 = 1 Then CurrentForm.Print "(8 characters max)"
            If tmp2 = 2 Then CurrentForm.Print "(directory name already taken)"
            If tmp2 = 3 Then CurrentForm.Print "(no spaces allowed)"
            If tmp2 > 0 Then Exit Do
        End If
    Loop
    GetDirectoryName = tmp1Str
End Function

Sub ImportSectorFromArchive()
    Rem import from archive
    Do
        CLS
        Color 11
        ShellPrg "xtra\ls *.sar > tmp.tmp"
        Open "tmp.tmp" For Input As #1
        L = LOF(1)
        Close 1
        If L = 0 Then
           ShellPrg "del tmp.tmp"
           CLS
           CurrentForm.Print "There are no .sar files available for importing"
           Call GetKeyChar(Key1, Key2)
           Exit Do
        End If
        If L = 0 Then
            Exit Do
        Else
            CLS
            CurrentForm.Print "Files available for importing:": CurrentForm.Print
            ShellPrg "xtra\ls *.sar"
            Color 10: CurrentForm.Print
            tStr = InputText("Name of file to be imported: ")
            tStr = LTrim$(RTrim$(LCase$(tStr)))
            If tStr = "" Then
                ShellPrg "del tmp.tmp"
                Exit Do
            Else
                If Len(tStr) < 4 Then tStr = tStr + ".sar"
                If Right$(tStr, 4) <> ".sar" Then tStr = tStr + ".sar"
                Ok = 0
                Open "tmp.tmp" For Input As #1
                Do Until EOF(1)
                 Line Input #1, tmpStr
                 If tStr = tmpStr Then Ok = 1
                Loop
                Close 1
                ShellPrg "del tmp.tmp"
                If Ok <> 0 Then
                    CLS
                    Screen 0
                    ShellPrg "sub\arch -d " + tStr
                    Call ReadSectorList: Rem read list of sectors
                End If
            End If
        End If
    Loop Until Ok = 1
    Call DrawSectorMap
End Sub

Sub ExportSectorToArchive()
    Rem export to archive
    Call WriteGalaxyWhereDat: Rem write gals\where.dat
    CurrentSector = SectorList(FNsec(xNow, yNow), 1)
    CLS
    Screen 0
    ShellPrg "sub\arch " + CurrentGalaxy + "\" + CurrentSector
    Call DrawSectorMap
End Sub

Sub DeleteSector()
Rem delete sector
    CLS
    Locate 5
    CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
    GalDirStr = "gals\" + CurrentGalaxy
    Tmp = FNsec(xNow, yNow)
    If Tmp <> 0 Then
        CurrentSector = SectorList(Tmp, 1)
        SecDirStr = GalDirStr + "\" + CurrentSector
        listfileStr = GalDirStr + "\" + CurrentGalaxy + ".lst"
        CLS
        Color 11: Locate 5
        If Column(Tmp) = 8 Then
           CurrentForm.Print "Are you sure you want to remove ";
        Else
           CurrentForm.Print "Are you sure you want to delete ";
        End If
        CurrentForm.Print SectorList(FNsec(xNow, yNow), 2);
        If Column(Tmp) = 8 Then CurrentForm.Print " entirely";
        CurrentForm.Print "? (y/n)"
        ynStr = InputYN
        If ynStr = "y" Then
            CLS
            If Column(Tmp) <> 8 Then ShellPrg "xtra\rm -r " + SecDirStr
            Open listfileStr For Input As #1
            Open "tmp.tmp" For Output As #2
            Do Until EOF(1)
               Line Input #1, tStr
               If RTrim$(Left$(tStr, 8)) <> CurrentSector Then
                  Print #2, tStr
               Else
                  If Column(Tmp) <> 8 Then Print #2, Left$(tStr, Len(tStr) - 1) + "I"
               End If
            Loop
            Close 1
            Close 2
            ShellPrg "del " + listfileStr
            ShellPrg "copy tmp.tmp " + listfileStr
            ShellPrg "del tmp.tmp"
            Call ReadSectorList: Rem read list of sectors
        End If
    End If
    Call DrawSectorMap
End Sub

Sub CreateSector()
    Rem insert/create sector
    Tmp = FNsec(xNow, yNow)
    If Column(Tmp) = 8 Then
        If Tmp = 0 Then
            CurrentSector = ""
        Else
            CurrentSector = SectorList(Tmp, 1)
        End If
        Call WriteGalaxyWhereDat: Rem write gals\where.dat
        Call WriteSectorWhereDat: Rem write gals\<galaxy>\where.dat
        ShellPrg "sub\makesec " + CurrentSector
        Open "secgen.tmp" For Input As #1
        Line Input #1, tStr
        Close 1
        ShellPrg "del secgen.tmp"
        If tStr = "full steam ahead" Then ShellPrg "sub\secgen 000"
        Call ReadSectorList: Rem read list of sectors
    End If
    Call DrawSectorMap
End Sub

Sub MoveSector()
    Rem move sector
    Tmp = FNsec(xNow, yNow)
    If MoveSec = 0 Then
        Rem picking sector up
        If Tmp = 0 Then
            Call DrawSectorMap
            Exit Sub
        End If
        MoveSec = Tmp
        MvSecX = xNow
        MvSecY = yNow
    Else
        Rem putting sector down
        If xNow = MvSecX And yNow = MvSecY Then MoveSec = 0
        If Tmp = 0 Then
            Tmp = MoveSec
            xx(Tmp) = xNow
            yy(Tmp) = yNow
            listfileStr = GalDirStr + "\" + CurrentGalaxy + ".lst"
            ShellPrg "copy " + listfileStr + " tmp.tmp"
            Open "tmp.tmp" For Input As #1
            Open listfileStr For Output As #2
            Line Input #1, tStr
            Print #2, tStr
            Line Input #1, tStr
            Print #2, tStr
            Do Until EOF(1)
                Line Input #1, tStr
                If RTrim$(Left$(tStr, 8)) = SectorList(Tmp, 1) Then
                    t2Str = Left$(tStr, 50)
                    nowatxStr = xStr$(xNow)
                    t2Str = t2Str + nowatxStr + Space$(5 - Len(nowatxStr))
                    nowatyStr = xStr$(yNow)
                    t2Str = t2Str + nowatyStr + Space$(5 - Len(nowatyStr))
                    t2Str = t2Str + Mid$(tStr, 61, 1)
                    Print #2, t2Str
                Else
                    Print #2, tStr
                End If
            Loop
            MoveSec = 0
            Close
            ShellPrg "del tmp.tmp"
            Call ReadSectorList
        End If
    End If
    Call DrawSectorMap
End Sub

Sub CaughtSectorHook(catch)
    Rem hook caught
    If InStr(HookSpecStr(1, catch), "1") > 0 And SecOCC = 0 Then Exit Sub
    If InStr(HookSpecStr(1, catch), "0") > 0 And SecOCC > 0 Then Exit Sub
    CLS
    Screen 0
    If InStr(HookSpecStr(1, catch), "E") > 0 Then
       Call DisplayEGAColor
    End If
    Call WriteSectorWhereDat: Rem write current position
    CurrentSector = SectorList(FNsec(xNow, yNow), 1)
    Call ReadSectorPaths: Rem file/directory paths
    ShellPrg FNhookStr(HookComStr(1, catch)), vbMaximizedFocus
    Call ReadSectorWhereDat: Rem read new position
    Rem IF INSTR(hookspecStr(1, catch), "U") > 0 THEN do nothing
    If InStr(HookSpecStr(1, catch), "P") > 0 Then Call ReloadPosition
    Call ReadSectorList
    Call DrawSectorMap
End Sub

Sub ReloadPosition()
    Rem reload positional info
    Call ReadSectorWhereDat: Rem read gals\<sector>\where.dat
    Call ReadSectorList: Rem read list of sectors
    CurrentSector = SectorList(FNsec(xNow, yNow), 1)
    Call ReadSectorPaths: Rem file/directory paths
End Sub

Sub ReadSectorData()
'2200
Rem get sector data
    Call ReadGalUWPStats: Rem read galaxy's uwp stats
    fStr = SecFPStr
    Open fStr For Input As #1
    Line Input #1, SecNameStr
    Line Input #1, tempStr
    Call DetermineNSWE: Rem determine n/s/e/w
    For I = 1 To 16
     Line Input #1, tempStr
     t9 = Len(tempStr)
     subtxt%(I) = 0
     If Mid$(tempStr, 50, 1) = "f" Then subtxt%(I) = 1
     If Mid$(tempStr, 50, 1) = "m" Then subtxt%(I) = 2
     SubNameStr(I, 1) = RTrim$(Mid$(tempStr, 4, 26))
     SubNameStr(I, 2) = RTrim$(Mid$(tempStr, 30, 12))
    Next I
    j = 1
    Line Input #1, tempStr
    Do
        Rem get base & status info
        I = 0
        Line Input #1, BasAlgStr(j, I, 0)
        I = 1
        Do
            If EOF(1) = -1 Then Exit Do
            Line Input #1, tempStr
            If tempStr = "" Then Exit Do
            t9 = Len(tempStr)
            If j = 1 Then
               BasAlgStr(1, I, 1) = Left$(tempStr, 1)
               BasAlgStr(1, I, 2) = RTrim$(Right$(tempStr, t9 - 4))
            End If
            If j = 2 Then
               algncol(I) = Val(Left$(tempStr, 2))
               BasAlgStr(2, I, 1) = Mid$(tempStr, 4, 2)
               BasAlgStr(2, I, 2) = RTrim$(Right$(tempStr, t9 - 8))
            End If
            I = I + 1
        Loop
        basalgn(j) = I - 1
        If j = 2 Then Exit Do
        j = j + 1
    Loop
    Close 1
End Sub

Sub ReadGalUWPStats()
'5450     Rem read in galaxy's uwp codes
    Open GalDirStr + "\uwp.dat" For Random As #1
    Exist = LOF(1)
    Close 1
    If Exist = 0 Then
       ShellPrg "copy data\uwp.dat " + GalDirStr + "\uwp.dat"
    End If
    Open GalDirStr + "\uwp.dat" For Input As #1
    tStr = ""
    A = 0: B = 0
    Do Until EOF(1)
       Line Input #1, tStr
       If LTrim$(tStr) <> "" And Left$(tStr, 1) <> "#" Then
          C = InStr(tStr, "=")
          If C = 0 Then
             If A > 0 Then UWPnum%(A) = B
             A = A + 1: B = 0
             UWPstatStr(A) = tStr
          End If
          If Left$(tStr, 4) = "pos=" Then
             UWPpos%(A, 1) = Val(Mid$(tStr, 5, 2))
             If Len(tStr) > 6 Then
                UWPpos%(A, 2) = Val(Mid$(tStr, 8, 2))
             Else
                UWPpos%(A, 2) = 0
             End If
             If Len(tStr) > 9 Then
                UWPpos%(A, 3) = Val(Mid$(tStr, 11, 2))
             Else
                UWPpos%(A, 3) = 0
             End If
          End If
          If C <> 0 And Left$(tStr, 4) <> "pos=" Then
             B = B + 1
             tmpStr = RTrim$(LTrim$(Left$(tStr, C - 1)))
             If Len(tmpStr) < 3 Then
                UWPval%(A, B, 1) = HexToNumber(tmpStr)
                UWPval%(A, B, 2) = UWPval%(A, B, 1)
             Else
                UWPval%(A, B, 1) = HexToNumber(Left$(tmpStr, 1))
                UWPval%(A, B, 2) = HexToNumber(Right$(tmpStr, 1))
             End If
             UWPdefStr(A, B) = RTrim$(LTrim$(Right$(tStr, Len(tStr) - C)))
          End If
       End If
    Loop
    UWPnum%(A) = B
    NumStats = A
    Close 1
'5460     Rem read in sector's uwp amendments
    Rem check if file exists
    If NoSector = 1 Then Exit Sub
    Open SecDirStr + "\uwp.dat" For Random As #1
    Exist = LOF(1)
    Close 1
    
    If Exist = 0 Then Kill SecDirStr + "\uwp.dat"
    If Exist = 0 Then Exit Sub
    Open SecDirStr + "\uwp.dat" For Input As #1
    tStr = ""
    First = 1
    Do Until EOF(1)
       Line Input #1, tStr
       If LTrim$(tStr) <> "" And Left$(tStr, 1) <> "#" Then
          C = InStr(tStr, "=")
          If C = 0 Then
             If First = 0 Then UWPnum%(A) = B
             If First = 1 Then First = 0
             A = 0
             For b1 = 1 To NumStats
                If tStr = UWPstatStr(b1) Then A = b1
             Next b1
             If A = 0 Then
                CLS
                CurrentForm.Print "Error: ";
                CurrentForm.Print tStr; " not defined in "; GalDirStr + "\uwp.dat"
                End
             End If
             B = UWPnum%(A)
          End If
          If C <> 0 Then
             B = B + 1
             tmpStr = RTrim$(LTrim$(Left$(tStr, C - 1)))
             If Len(tmpStr) < 3 Then
                UWPval%(A, B, 1) = HexToNumber(tmpStr)
                UWPval%(A, B, 2) = UWPval%(A, B, 1)
             Else
                UWPval%(A, B, 1) = HexToNumber(Left$(tmpStr, 1))
                UWPval%(A, B, 2) = HexToNumber(Right$(tmpStr, 1))
             End If
             UWPdefStr(A, B) = RTrim$(LTrim$(Right$(tStr, Len(tStr) - C)))
          End If
       End If
    Loop
    UWPnum%(A) = B
    Close 1
End Sub

Sub DetermineNSWE()
'3700     Rem determine n/s/e/w
    NorthStr = "<nil>"
    SouthStr = "<nil>"
    EastStr = "<nil>"
    WestStr = "<nil>"
    For A = 1 To NumSec
        dudeStr = SectorList(A, 1)
        If Column(A) = 8 Then dudeStr = "<nil>"
        If xx(A) = xNow + 1 And yy(A) = yNow Then EastStr = dudeStr
        If xx(A) = xNow - 1 And yy(A) = yNow Then WestStr = dudeStr
        If xx(A) = xNow And yy(A) = yNow + 1 Then SouthStr = dudeStr
        If xx(A) = xNow And yy(A) = yNow - 1 Then NorthStr = dudeStr
    Next A
End Sub

Sub DrawSubSectorOuter()
    Screen 0
    Color 2
    HHex = 101
    Trip = 1
    CLS
    Locate 1, 1: CurrentForm.Print Chr$(201);
    Locate 21, 61: CurrentForm.Print Chr$(188);
    Locate 1, 61: CurrentForm.Print Chr$(187);
    Locate 21, 1: CurrentForm.Print Chr$(200);
    For I = 16 To 46 Step 15
        Locate 1, I: CurrentForm.Print Chr$(203);
        Locate 21, I: CurrentForm.Print Chr$(202);
    Next I
    For I = 6 To 16 Step 5
        Locate I, 1: CurrentForm.Print Chr$(204);
        Locate I, 61: CurrentForm.Print Chr$(185);
    Next I
    For I = 16 To 46 Step 15
        For j = 6 To 16 Step 5
            Locate j, I: CurrentForm.Print Chr$(206);
        Next j
    Next I
    For I = 1 To 46 Step 15
        For j = 1 To 14
            For k = 1 To 21 Step 5
                Locate k, I + j: CurrentForm.Print Chr$(205);
            Next k
        Next j
    Next I
    For I = 1 To 16 Step 5
        For j = 1 To 4
            For k = 1 To 61 Step 15
                Locate I + j, k: CurrentForm.Print Chr$(186);
            Next k
        Next j
    Next I
    
    Rem fill in subsector names
    Color 14
    For I = 1 To 4
        For j = 1 To 4
            k = j + ((I - 1) * 4)
            lname = Len(SubNameStr(k, 1))
            csubnameStr = SubNameStr(k, 1)
            If lname > 14 Then csubnameStr = Left$(SubNameStr(k, 1), 14)
            If lname > 14 Then lname = 14
            Locate ((I - 1) * 5) + 3, ((j - 1) * 15) + 8 - Int((lname - 1) / 2)
            CurrentForm.Print csubnameStr;
        Next j
    Next I
    Locate 23, 31 - Int(Len(SecNameStr) / 2)
    CurrentForm.Print SecNameStr;
    Locate 25, 70: Color 12: CurrentForm.Print "? = Help";
End Sub

Sub DrawSubSectorInner()
'60     Rem make inner box and get instructions
    tx = ((zJ - 1) * 15) + 2
    ty = ((zI - 1) * 5) + 2
    Color 13
    Locate ty, tx: CurrentForm.Print Chr$(218);
    Locate ty, tx + 13: CurrentForm.Print Chr$(191);
    Locate ty + 2, tx: CurrentForm.Print Chr$(192);
    Locate ty + 2, tx + 13: CurrentForm.Print Chr$(217);
    For I = tx + 1 To tx + 12
        Locate ty, I: CurrentForm.Print Chr$(196);
        Locate ty + 2, I: CurrentForm.Print Chr$(196);
    Next I
    k = zJ + ((zI - 1) * 4)
    lname = Len(SubNameStr(k, 1))
    If lname < 13 Then
       Locate ty + 1, tx
       CurrentForm.Print Chr$(179);
    End If
    If lname < 14 Then
       Locate ty + 1, tx + 13
       CurrentForm.Print Chr$(179);
    End If
End Sub

Sub DrawSubSectorInstructions()
    tx = ((zJ - 1) * 15) + 2
    ty = ((zI - 1) * 5) + 2
    Locate ty, tx: CurrentForm.Print Chr$(32);
    Locate ty, tx + 13: CurrentForm.Print Chr$(32);
    Locate ty + 2, tx: CurrentForm.Print Chr$(32);
    Locate ty + 2, tx + 13: CurrentForm.Print Chr$(32);
    For I = tx + 1 To tx + 12
        Locate ty, I: CurrentForm.Print Chr$(32);
        Locate ty + 2, I: CurrentForm.Print Chr$(32);
    Next I
    k = zJ + ((zI - 1) * 4)
    lname = Len(SubNameStr(k, 1))
    If lname < 13 Then
       Locate ty + 1, tx
       CurrentForm.Print Chr$(32);
    End If
    If lname < 14 Then
       Locate ty + 1, tx + 13
       CurrentForm.Print Chr$(32);
    End If
End Sub

Sub CaughtSubHook(catch)
'62     Rem hook caught
    zK = zJ + ((zI - 1) * 4)
    fSubStr = MapDirStr + SubNameStr(zK, 2)
    CLS
    If InStr(HookSpecStr(2, catch), "E") > 0 Then Call DisplayEGAColor
    Call WriteSectorWhereDat: Rem write position
    ShellPrg FNhookStr(HookComStr(2, catch))
    If InStr(HookSpecStr(2, catch), "U") > 0 Then Call ReadGalUWPStats
    If InStr(HookSpecStr(2, catch), "P") > 0 Then Call ReloadPosition
    Screen 12
    HHex = Hex
    Trip = 1
    Call ReadSectorData
    Call DrawSubSectorOuter
End Sub

Sub SubSectorUp()
    '65 Rem up
    zI = zI - 1
    If zI = 0 Then
       If NorthStr = "<nil>" Then zI = 1
       If NorthStr <> "<nil>" Then
          zI = 4: NewJY = NewJY + 4: yNow = yNow - 1
          Call ShiftOfSector
          Call ReadSectorData
          Call DrawSubSectorOuter
       End If
    End If
    Call DrawSubSectorOuter
End Sub

Sub SubSectorDown()
    '70 Rem down
    zI = zI + 1
    If zI = 5 Then
       If SouthStr = "<nil>" Then zI = 4
       If SouthStr <> "<nil>" Then
          zI = 1: NewJY = NewJY - 4: yNow = yNow + 1
          Call ShiftOfSector
          Call ReadSectorData
          Call DrawSubSectorOuter
       End If
    End If
    Call DrawSubSectorOuter
End Sub

Sub SubSectorLeft()
    zJ = zJ - 1
    If zJ = 0 Then
       If WestStr = "<nil>" Then zJ = 1
       If WestStr <> "<nil>" Then
          zJ = 4: NewJX = NewJX + 4: xNow = xNow - 1
          Call ShiftOfSector
          Call ReadSectorData
          Call DrawSubSectorOuter
       End If
    End If
    Call DrawSubSectorOuter
End Sub

Sub SubSectorRight()
    '80 Rem right
    zJ = zJ + 1
    If zJ = 5 Then
       If EastStr = "<nil>" Then zJ = 4
       If EastStr <> "<nil>" Then
          zJ = 1: NewJX = NewJX - 4: xNow = xNow + 1
          Call ShiftOfSector
          Call ReadSectorData
          Call DrawSubSectorOuter
       End If
    End If
    Call DrawSubSectorOuter
End Sub

Sub GotoSubScreen()
    zK = zJ + ((zI - 1) * 4)
    fSubStr = MapDirStr + SubNameStr(zK, 2)
End Sub

Sub ReadSubsector()
    Jumps = 0
    For I = 1 To 80
        WorldStr(I) = ""
    Next I
    Open fSubStr For Input As #1
    Do
        Line Input #1, t4Str
        t4Str = LTrim$(RTrim$(t4Str))
        If t4Str <> "" Then
            t3Str = Left$(t4Str, 1)
            If Not (t3Str = "@" Or t3Str = "#") Then
                If t3Str = "$" Then
                   Jumps = Jumps + 1
                   rHex = Val(Mid$(t4Str, 2, 4)): Call ConvertRealHexToMap
                   Jump%(Jumps, 1) = Hex
                   rHex = Val(Mid$(t4Str, 7, 4)): Call ConvertRealHexToMap
                   Jump%(Jumps, 2) = Hex
                   Jump%(Jumps, 3) = Val(Mid$(t4Str, 12, 2))
                   place = 14
                   If Jump%(Jumps, 3) = -1 Then place = place + 1
                   Jump%(Jumps, 4) = Val(Mid$(t4Str, place, 2))
                   place = place + 2
                   If Jump%(Jumps, 4) = -1 Then place = place + 1
                   Jump%(Jumps, 5) = Val(Mid$(t4Str, place, 2))
                   If Jump%(Jumps, 5) = 0 Then Jump%(Jumps, 5) = 11
                Else
                    rHex = Val(Mid$(t4Str, 15, 4))
                    Call ConvertRealHexToMap
                    Call ConvertMapHexToRef
                    WorldStr(Ref) = t4Str
                End If
            End If
        End If
        If EOF(1) = -1 Then Exit Do
    Loop
    Close 1
End Sub

Sub ConvertRealHexToMap()
'1100 Rem convert real hex to map hex
Rem in -> rhex
Rem out -> hex
    t1 = Int(rHex / 100)
    t2 = rHex - t1 * 100
    Do
        If t1 <= 8 Then Exit Do
        t1 = t1 - 8
    Loop
    Do
        If t2 <= 10 Then Exit Do
        t2 = t2 - 10
    Loop
    Hex = t1 * 100 + t2
End Sub

Sub ConvertMapHexToRef()
    '1150 Rem convert map hex to ref#
    Rem in -> hex
    Rem out -> ref
    t1 = Int(Hex / 100)
    t2 = Hex - t1 * 100
    Ref = (t1 - 1) * 10 + t2
End Sub

Sub ConvertRefToMapHex()
    '1175 Rem convert ref# to map hex
    Rem in -> ref
    Rem out -> hex
    t1 = Int((Ref + 9) / 10)
    t2 = Ref Mod 10
    If t2 = 0 Then t2 = 10
    Hex = t1 * 100 + t2
End Sub

Sub SetPalette()
'->'6000     Rem set palette
'    Select Case ClrMode
'    Case 1
'       Rem colors on black
'       Palette
'    Case 2
'       Rem colors on white
'       Palette
'       Palette 0, 4144959: Palette 15, 0
'    Case 3
'       Rem white on black
'       t = 4144959
'       Palette 0, 0: Palette 1, t: Palette 2, t: Palette 3, t
'       Palette 4, t: Palette 5, t: Palette 6, t: Palette 7, t
'       Palette 8, t: Palette 9, t: Palette 10, t: Palette 11, t
'       Palette 12, t: Palette 13, t: Palette 14, t: Palette 15, t
'    Case 4
'       Rem black on white
'       Palette 0, 4144959: Palette 1, 0: Palette 2, 0: Palette 3, 0
'       Palette 4, 0: Palette 5, 0: Palette 6, 0: Palette 7, 0
'       Palette 8, 0: Palette 9, 0: Palette 10, 0: Palette 11, 0
'       Palette 12, 0: Palette 13, 0: Palette 14, 0: Palette 15, 0
'    End Select
End Sub

Sub InitGraphics()
    CLS
    Screen 12
    Call SetPalette: Rem set palette according to clrmode
End Sub

Sub DrawSubGraphic()
'151     Rem draw subsector map
    CLS
    Rem draw blue hexs
    Color 1
    For A = 1 To 8
        For B = 1 To 10
            Hex = (A * 100) + B
            Call PlotHex
        Next B
    Next A
    Rem draw red hexs
    Color 4
    For A = 1 To 8
        For B = 1 To 10
            Hex = (A * 100) + B
            C = ((A - 1) * 10) + B
            tmpStr = Mid$(WorldStr(C), 63, 1)
            tmp2Str = Mid$(WorldStr(C), 64, 1)
            If LTrim$(tmpStr + tmp2Str) <> "" Then Call PlotHex
        Next B
    Next A
    Rem draw jumproutes
    If Jumps > 0 Then
       For A = 1 To Jumps
          Hex = Jump%(A, 1): Call FindCenterPt
          P1 = Cent1 + Yank: p2 = Cent2
          Hex = Jump%(A, 2): Call FindCenterPt
          p3 = Cent1 + Yank: p4 = Cent2
          If Jump%(A, 3) <> 0 Then
             p3 = p3 + (Jump%(A, 3) * cHexA * 12)
          End If
          If Jump%(A, 4) <> 0 Then
             p4 = p4 + (Jump%(A, 4) * cHexB * 20)
          End If
          CurrentForm.Line (P1, p2)-(p3, p4), QBColor(Jump%(A, 5)) ', , &H8888
       Next A
    End If
    Rem fill in map
    For Ref = 1 To 80
        If Len(WorldStr(Ref)) <> 0 Then
            L = Asc(Mid$(WorldStr(Ref), 20, 1)) - 64
            If L = 24 Then L = 6
            Call ConvertRefToMapHex
            Call DrawStarPort
        End If
    Next Ref
    Ref = 80
    Rem subsector/sector name
    Color 11
    bottomStr = SubNameStr(zK, 1) + "/" + SecNameStr
    L = Len(bottomStr)
    Locate 29, 60 - Int(L / 2)
    CurrentForm.Print bottomStr;
End Sub

Sub PlotHex()
'500     Rem plot a hex
    Call FindCenterPt
    Rem known=> cent1, cent2
    Rem from one o'clock clockwise
    tmp1 = Cent1 + 0.5 * cHexA + Yank
    tmp2 = Cent2 - HexB
    tmp3 = Cent1 + cHexA + Yank
    tmp4 = Cent2
    CurrentForm.Line (tmp1, tmp2)-(tmp3, tmp4)
    CurrentForm.Line -Step(-0.5 * cHexA, cHexB)
    CurrentForm.Line -Step(-cHexA, 0)
    CurrentForm.Line -Step(-0.5 * cHexA, -cHexB)
    CurrentForm.Line -Step(0.5 * cHexA, -cHexB)
    CurrentForm.Line -Step(cHexA, 0)
End Sub

Sub FindCenterPt()
'600     Rem find center point of hex
    Rem in=> hex
    Rem out=> cent1, cent2
    Cent1 = HexA
    Cent2 = HexB
    thex = Hex - 101
    Do
        If thex < 99 Then Exit Do
        Rem going across
        Cent1 = Cent1 + 1.5 * HexA
        thex = thex - 100
    Loop
    Do
        If thex = 0 Then Exit Do
        Rem going down
        Cent2 = Cent2 + 2 * HexB
        thex = thex - 1
    Loop
    If Not (Int(Hex / 100) Mod 2 = 1) Then
        Cent2 = Cent2 + HexB
    End If
    Cent1 = Cent1 * Scar
End Sub

Sub MakeCircle()
'800     Rem make circle
    Call FindCenterPt
    CurrentForm.Circle (Cent1 + Yank, Cent2), HexA / 2, QBColor(13)
End Sub

Sub DrawHex()
    '210
    Call ConvertMapHexToRef: Rem get ref#
    Color 14
    HWStr = Mid$(WorldStr(Ref), 64, 1)
    HiddenStr = Mid$(WorldStr(Ref), 65, 1)
    Hidden = 0: NonWorld = 0: NoName = 0
    If HiddenStr = "h" Then Hidden = 1
    If HiddenStr = "H" Then Hidden = 2
    If Mid$(WorldStr(Ref), 20, 1) = "*" Then NonWorld = 1
    If RTrim$(Left$(WorldStr(Ref), 14)) = "." Then NoName = 1
    If Len(WorldStr(Ref)) = 0 Or (Hidden = 2 And NonWorld = 1) Then
       Call DrawEmptyHex: Rem empty hex
    Else
       Call DrawNonEmptyHex: Rem non-empty hex
    End If
End Sub

Sub DrawEmptyHex()
'1200     Rem empty hex
    Call ClearTextHalf
    CurrentForm.Print "Empty Parsec"
    Rem CurrentForm.Print "Subsector: "; subnameStr(zk, 1)
    Rem CurrentForm.Print "Sector: "; secnameStr
    tmphex = Hex
    Call ConvertRefToMapHex
    hStr = xStr$(Hex)
    hStr = "0" + hStr
    h1 = Val(Left$(hStr, 2))
    h2 = Val(Right$(hStr, 2))
    h1 = h1 + (zJ - 1) * 8
    h2 = h2 + (zI - 1) * 10
    SubLetStr = Chr$(zK + 64)
    hStr = SubLetStr + hStr
    HHStr = xStr$(h1 * 100 + h2)
    If Len(HHStr) = 3 Then HHStr = "0" + HHStr
    CurrentForm.Print "Coordinates: "; HHStr; " ("; hStr; ")"
    CoordStr = HHStr
    Hex = tmphex
    Call DisplaySubMenu
End Sub

Sub DrawNonEmptyHex()
'1250     Rem non-empty hex
    Call ClearTextHalf
    If NoName = 0 Then
       If Hidden < 2 Then
          If NonWorld = 0 Then CurrentForm.Print "World: ";
          CurrentForm.Print RTrim$(Left$(WorldStr(Ref), 14))
       Else
          If NonWorld = 0 Then CurrentForm.Print "World: ";
          CurrentForm.Print "Unknown"
       End If
    End If
    Rem CurrentForm.Print "Subsector: "; subnameStr(zk, 1)
    Rem CurrentForm.Print "Sector: "; secnameStr
    HHStr = Mid$(WorldStr(Ref), 15, 4)
    CoordStr = HHStr
    CurrentForm.Print "Coordinates: "; HHStr;
    SubLetStr = Chr$(zK + 64)
    CurrentForm.Print " (" + SubLetStr + "0" + Right$(Str$(Hex), 3) + ")"
    
    If Hidden > 0 Then
       CurrentForm.Print
       Color 11
       If NonWorld = 0 Then CurrentForm.Print "World Statistics Unknown"
    End If
    
    If Hidden > 0 Then
        Call DisplaySubMenu
        Exit Sub
    End If
    
    If NonWorld = 1 Then
       tmpStr = RTrim$(Mid$(WorldStr(Ref), 23, 25))
       If tmpStr <> "" Then
          CurrentForm.Print: CurrentForm.Print tmpStr
       End If
    End If
    If NonWorld = 1 Then
        Call DisplaySubMenu
        Exit Sub
    End If
    
    CurrentForm.Print "UWP: "; Mid$(WorldStr(Ref), 20, 9);
    extraStr = LTrim$(RTrim$(Mid$(WorldStr(Ref), 58, 5)))
    If extraStr <> "" Then
        CurrentForm.Print "-"; extraStr
    Else
        CurrentForm.Print
    End If
    
    CurrentForm.Print
    For Tmp = 1 To NumStats
       uwptmpStr = FNuwpStr(Tmp)
       If uwptmpStr <> "Empty" Then
          If UWPpos%(Tmp, 3) = 0 Then
             Color 11
          Else
             Color UWPpos%(Tmp, 3)
          End If
          CurrentForm.Print UWPstatStr(Tmp); ": "; uwptmpStr
       End If
    Next Tmp
    CurrentForm.Print: Color 10
    CurrentForm.Print BasAlgStr(2, 0, 0); " "; GetAllegiance
    CurrentForm.Print BasAlgStr(1, 0, 0); " "; GetBases
    CurrentForm.Print "Trade Codes: "; GetTradeCodes
    CurrentForm.Print "Travel Zone: "; GetZone
    Call DisplaySubMenu
End Sub

Sub ClearTextHalf()
'1180     Rem clear text half
    Dim TextHeight As Double
    Dim TextWidth As Double
    TextHeight = CurrentForm.TextHeight("W")
    TextWidth = CurrentForm.TextWidth("W")
    If ClrMode = 2 Or ClrMode = 4 Then
        CurrentForm.Line (0, 0)-(TextWidth * 35, CurrentForm.ScaleHeight), vbWhite, BF
    Else
        CurrentForm.Line (0, 0)-(TextWidth * 35, CurrentForm.ScaleHeight), vbBlack, BF
    End If
    Rem draw little sector box
    lsb1 = 10
    lsb2 = 50
    lsb3 = 260
    lsb4 = 300
    For I = lsb1 To lsb2 Step (lsb2 - lsb1) / 4
        CurrentForm.Line (lsb3, I)-(lsb4, I), QBColor(1)
    Next I
    For I = lsb3 To lsb4 Step (lsb4 - lsb3) / 4
        CurrentForm.Line (I, lsb1)-(I, lsb2), QBColor(1)
    Next I
    lsb5 = lsb1 + ((zI - 1) * ((lsb2 - lsb1) / 4))
    lsb6 = lsb5 + ((lsb2 - lsb1) / 4)
    lsb7 = lsb3 + ((zJ - 1) * ((lsb4 - lsb3) / 4))
    lsb8 = lsb7 + ((lsb4 - lsb3) / 4)
    CurrentForm.Line (lsb7, lsb5)-(lsb8, lsb6), QBColor(12)
    CurrentForm.Line (lsb7, lsb6)-(lsb8, lsb5), QBColor(12)
    Locate 1, 1
End Sub

Sub DisplaySubMenu()
'1255
    tmpStr = Mid$(WorldStr(Ref), 63, 1)
    CurrentForm.Print
    If (tmpStr = "f" Or tmpStr = "m") And NoName = 0 Then
       Color 13
       If NonWorld = 0 Then CurrentForm.Print "<Enter> = World Description"
       If NonWorld = 1 Then CurrentForm.Print "<Enter> = Site Description"
    End If
    If subtxt%(zK) <> 0 Then
       Color 13: CurrentForm.Print "F1 = Subsector Description"
    End If
    If (HWStr = "x" Or HWStr = "z") And Hidden = 0 Then
       Color 13: CurrentForm.Print "F4 = World Map"
    End If
    If (HWStr = "y" Or HWStr = "z") And Hidden = 0 Then
       Color 13: CurrentForm.Print "F8 = Star System Map"
    End If
    Color 13: CurrentForm.Print "? = Help";

End Sub

Function GetAllegiance()
'1620 Rem allegiance
    sStr = Mid$(WorldStr(Ref), 56, 2)
    For m = 1 To basalgn(2)
        If sStr = BasAlgStr(2, m, 1) Then GetAllegiance = BasAlgStr(2, m, 2)
    Next m
End Function

Function GetBases()
    '1590 Rem bases
    sStr = Mid$(WorldStr(Ref), 31, 1)
    For m = 1 To basalgn(1)
        If sStr = BasAlgStr(1, m, 1) Then s1Str = BasAlgStr(1, m, 2)
    Next m
    GetBases = s1%
End Function

Function GetTradeCodes()
    '1595 Rem trade codes
    If Mid$(WorldStr(Ref), 33, 1) = " " Then
       s1Str = "None"
    Else
       s1Str = Mid$(WorldStr(Ref), 33, 16)
    End If
    GetTradeCodes = s1Str
End Function

Function GetZone()
    '1610 Rem zone
    sStr = Mid$(WorldStr(Ref), 49, 1)
    If sStr = " " Or sStr = "G" Then s1Str = "Green"
    If sStr = "A" Then s1Str = "Amber"
    If sStr = "R" Then s1Str = "Red"
    If sStr = "B" Then s1Str = "Blue"
    GetZone = s1Str
End Function

Sub CatchSubGraphicHook(catch)
'205 Rem hook caught
    If InStr(HookSpecStr(3, catch), "1") > 0 And WorldStr(Ref) = "" Then Exit Sub
    If InStr(HookSpecStr(3, catch), "0") > 0 And WorldStr(Ref) <> "" Then Exit Sub
    If InStr(HookSpecStr(3, catch), "E") > 0 Then
       Screen 0: Call DisplayEGAColor
    End If
    Call WriteSectorWhereDat: Rem write position
    ShellPrg FNhookStr(HookComStr(3, catch))
    If InStr(HookSpecStr(3, catch), "U") > 0 Then Call ReadGalUWPStats
    If InStr(HookSpecStr(3, catch), "C") > 0 Then
       Call ReloadPosition: Rem reload positional info
       Call ReadSectorData: Rem reload sector
    End If
    Screen 12: HHex = Hex: Trip = 1
    Call GotoSubScreen
    Call ReadSubsector
'    Call InitGraphic
    Call DrawSubGraphic
    Call DrawInitialCircle
End Sub

Sub DrawInitialCircle()
    If Trip = 1 Then
       Hex = HHex
       Trip = 0
    End If
    Call MakeCircle
    Call DrawHex
End Sub

Sub DrawMoreInfo(Key1)
'2000 Rem more information
    If WorldStr(Ref) = "" Or Hidden > 0 Or NoName = 1 Then Exit Sub
    ssStr = Mid$(WorldStr(Ref), 63, 1)
    AppStr = ReaderStr
    If Key1 = 62 Then AppStr = EditorStr
    If ssStr <> "f" And ssStr <> "m" Then
        tmpStr = "star system"
        Call DrawMiniMenu1: Rem mini-menu
        Do
            Call GetKeyChar(Key1, Key2)
            If Key2 = 0 And Key1 = 70 Then
                Call Attach("f"): Rem f
                Exit Do
            End If
            If Key2 = 0 And Key1 = 77 Then
                Call Attach("m"): Rem m
                Exit Do
            End If
            If Key2 = 0 And Key1 = 27 Then Exit Do: Rem esc
            If Key2 = 0 And Key1 = 81 Then Exit Do: Rem q
        Loop
    End If
    Screen 0
    fStr = LocDirStr + "p" + Mid$(WorldStr(Ref), 15, 4)
    If ssStr = "f" Then fStr = fStr + ".txt"
    If ssStr = "m" Then fStr = fStr + ".mnu"
    Call DisplayEGAColor: Rem egacolor
    If ssStr = "f" Then ShellPrg AppStr + " " + fStr
    If ssStr = "m" Then EgMain fStr + " " + AppStr
    Screen 12
    HHex = Hex
    Trip = 1
    If Trip2 = 1 Then
        Trip2 = 0
    Else
        Call GotoSubScreen
    End If
    Call ReadSubsector
'    Call InitGraphic
    Call DrawSubGraphic
    Call DrawInitialCircle
End Sub

Sub DrawMiniMenu1()
'5600 Rem mini-menu #1 (attach file/menu query)
    CLS
    Color 14
    Locate 5, 20
    CurrentForm.Print Chr$(218); String$(38, 196); Chr$(191)
    For Tmp = 6 To 14
        Locate Tmp, 20
        CurrentForm.Print Chr$(179); String$(38, 32); Chr$(179)
    Next Tmp
    Locate 15, 20
    CurrentForm.Print Chr$(192); String$(38, 196); Chr$(217)
    Color 11
    Locate 7, 22
    CurrentForm.Print "There is no file or menu currently"
    Locate 8, 22
    CurrentForm.Print "attached to this "; tmpStr; ". Your"
    Locate 9, 22
    CurrentForm.Print "options are:"
    Color 10
    Locate 11, 27
    CurrentForm.Print "(F) Attach a File"
    Locate 12, 27
    CurrentForm.Print "(M) Attach a Menu"
    Locate 13, 27
    CurrentForm.Print "(Q) Get Me Outta Here"
End Sub

Sub DrawMiniMenu2()
'5620     Rem mini-menu #2 (delete star query)
    CLS
    Color 14
    Locate 5, 17
    CurrentForm.Print Chr$(218); String$(44, 196); Chr$(191)
    For Tmp = 6 To 12
        Locate Tmp, 17
        CurrentForm.Print Chr$(179); String$(44, 32); Chr$(179)
    Next Tmp
    Locate 13, 17
    CurrentForm.Print Chr$(192); String$(44, 196); Chr$(217)
    Color 11
    Locate 7, 21
    CurrentForm.Print "Are you sure you really want to delete"
    Locate 8, 21
    CurrentForm.Print "this place?"
    Color 10
    Locate 10, 19
    CurrentForm.Print "(Y) Ya dang right I do! I hate those guys!"
    Locate 11, 19
    CurrentForm.Print "(N) Nah, we'll have pity on the poor sots."
End Sub

Sub DrawMiniMenu3()
'5640 Rem mini-menu #3 (sector's uwp.dat query)
    CLS
    Color 14
    Locate 5, 17
    CurrentForm.Print Chr$(218); String$(44, 196); Chr$(191)
    For Tmp = 6 To 12
        Locate Tmp, 17
        CurrentForm.Print Chr$(179); String$(44, 32); Chr$(179)
    Next Tmp
    Locate 13, 17
    CurrentForm.Print Chr$(192); String$(44, 196); Chr$(217)
    Color 11
    Locate 7, 21
    CurrentForm.Print "This sector has no uwp.dat file. Would"
    Locate 8, 21
    CurrentForm.Print "you like me to copy the one from the"
    Locate 9, 21
    CurrentForm.Print "classic\beyond directory?"
    Color 10: Locate 11, 29
    CurrentForm.Print "(Y) Yep     (N) Nope"
End Sub

Sub DrawMiniMenu4()
'5650     Rem mini-menu #4 (jumproute color query)
    CLS
    Color 14: Locate 3, 17
    CurrentForm.Print Chr$(218); String$(44, 196); Chr$(191)
    For Tmp = 4 To 22
        Locate Tmp, 17
        CurrentForm.Print Chr$(179); String$(44, 32); Chr$(179)
    Next Tmp
    Locate 23, 17
    CurrentForm.Print Chr$(192); String$(44, 196); Chr$(217)
    Color 11: Locate 5, 21
    CurrentForm.Print "Choose a color for the jumproute:"
    For Tmp = 1 To 15
        Locate Tmp + 6, 21: Color Tmp
        CurrentForm.Print "("; Chr$(64 + Tmp); ") "; ColorsStr(Tmp)
    Next Tmp
End Sub

Sub Attach(ssStr)
'     If Key1 = 70 Then ssStr = "f"
'     If Key1 = 77 Then ssStr = "m"
     Rem fix subsector file
     ShellPrg "copy " + fSubStr + " data\tmpsub.dat"
     Open "data\tmpsub.dat" For Input As #1
     Open fSubStr For Output As #2
     Do Until EOF(1)
        Line Input #1, tsubStr
        If WorldStr(Ref) = tsubStr Then
            lline = Len(tsubStr)
            If lline < 62 Then
                tsubStr = tsubStr + Space$(62 - lline) + ssStr + HWStr + HiddenStr
            Else
                tsubStr = Left$(tsubStr, 62) + ssStr + HWStr + HiddenStr
            End If
        End If
        Print #2, tsubStr
     Loop
     Close 1
     Close 2
     
     ShellPrg "erase data\tmpsub.dat"
     AppStr = EditorStr
     Trip2 = 1
     If ssStr = "m" Then
        fStr = LocDirStr + "p" + Mid$(WorldStr(Ref), 15, 4) + ".mnu"
        f1Str = "p" + Mid$(WorldStr(Ref), 15, 4)
        Rem check if file fStr exists
        Open fStr For Random As #1
        Exist = LOF(1)
        Close 1
        If Exist = 0 Then
            ShellPrg "erase " + fStr
            Open fStr For Output As #1
            Open "data\starter.mnu" For Input As #2
            Print #1, "@dir=" + LocDirStr
            Print #1, "@Light Yellow"
            Print #1, ""
            Print #1, "          ";
            Print #1, RTrim$(Left$(WorldStr(Ref), 14)) + " " + "System"
            tmpcount = 0
            Do Until EOF(2)
                Line Input #2, tStr
                lstart = Len(tStr)
                If lstart = 1 Then tStr = ""
                If lstart > 1 Then tStr = Right$(tStr, lstart - 1)
                Print #1, tStr;
                If Right$(tStr, 1) = "@" Then
                    tmpcount = tmpcount + 1
                    If tmpcount = 1 Then Print #1, f1Str + ".txt"
                    If tmpcount = 2 Then Print #1, f1Str + "b.txt"
                Else
                    Print #1, ""
                End If: Rem @
            Loop
            Close 1
            Close 2
        End If: Rem exist=0
     End If: Rem ssStr="m"
     Rem write system stats to text file if file doesn't already exist
     Rem check if file fStr exists
     fStr = LocDirStr + "p" + Mid$(WorldStr(Ref), 15, 4) + ".txt"
     Open fStr For Random As #1
     Exist = LOF(1)
     Close 1
     If Exist = 0 Then
        ShellPrg "erase " + fStr
        Open fStr For Output As #1
        If NonWorld = 0 Then Print #1, "World: ";
        Print #1, RTrim$(Left$(WorldStr(Ref), 14))
        Print #1, "Subsector: "; SubNameStr(zK, 1)
        Print #1, "Sector: "; SecNameStr
        Print #1, "Coordinates: "; Mid$(WorldStr(Ref), 15, 4);
        Print #1, " (" + Chr$(zK + 64) + "0" + Right$(Str$(Hex), 3) + ")"
        If NonWorld = 0 Then
            Print #1, "UWP: "; Mid$(WorldStr(Ref), 20, 9);
            extraStr = LTrim$(RTrim$(Mid$(WorldStr(Ref), 58, 5)))
            If extraStr <> "" Then
                Print #1, "-"; extraStr
            Else
                Print #1, ""
            End If
        End If
        Print #1, ""
        If NonWorld = 1 Then
            tmpStr = RTrim$(Mid$(WorldStr(Ref), 23, 25))
            If tmpStr <> "" Then Print #1, tmpStr
        Else
            For Tmp = 1 To NumStats
                uwptmpStr = FNuwpStr(Tmp)
                If uwptmpStr <> "Empty" Then
                    Print #1, UWPstatStr(Tmp); ": "; uwptmpStr
                End If
            Next Tmp
            Print #1, ""
            Print #1, BasAlgStr(2, 0, 0); " "; GetAllegiance
            Print #1, BasAlgStr(1, 0, 0); " "; GetBases
            Print #1, "Trade Codes: "; GetTradeCodes
            Print #1, "Travel Zone: "; GetZone
        End If
        Close 1
    End If: Rem exist=0
End Sub

Sub DrawStarPort()
'700     Rem plot starport
    If Len(WorldStr(Ref)) = 0 Then Exit Sub
    tmp3Str = Mid$(WorldStr(Ref), 65, 1)
    tmp4 = 0: tmp5 = 0: NoName = 0
    If tmp3Str = "h" Then tmp4 = 1
    If tmp3Str = "H" Then tmp4 = 2
    If RTrim$(Left$(WorldStr(Ref), 14)) = "." Then NoName = 1
    NonWorld = 0
    If Mid$(WorldStr(Ref), 20, 1) = "*" Then
       tmp5 = Val(Mid$(WorldStr(Ref), 21, 1))
       NonWorld = 1
    End If
    Call FindCenterPt: Rem centerpoint
    Rem find allegiance
    sStr = Mid$(WorldStr(Ref), 56, 2)
    For i2 = 1 To basalgn(2)
       If sStr = BasAlgStr(2, i2, 1) Then Alle = algncol(i2)
    Next i2
    Rem draw marker
    Rem asteroids or circle
    If tmp5 = 0 Then
       sStr = Mid$(WorldStr(Ref), 21, 1)
       If (sStr = "R" Or sStr = "0") And tmp4 = 0 Then
          x = Cent1 + Yank: y = Cent2
          Color Alle
          CurrentForm.PSet (x, y): CurrentForm.PSet (x + 2, y + 2): CurrentForm.PSet (x - 2, y - 2)
          CurrentForm.PSet (x + 2, y - 2): CurrentForm.PSet (x - 2, y + 2)
       Else
          CurrentForm.PSet (Cent1 + Yank, Cent2), QBColor(0)
          Rem fill circle if there's any water
          If Mid$(WorldStr(Ref), 23, 1) <> "0" Or tmp4 <> 0 Then
            CurrentForm.FillStyle = vbFSSolid
            CurrentForm.FillColor = QBColor(Alle)
          End If
          CurrentForm.Circle (Cent1 + Yank, Cent2), 3, QBColor(Alle)
          CurrentForm.FillStyle = vbFSTransparent
       End If
       Rem gas giant & bases marker
       If tmp4 = 0 Then
          If Mid$(WorldStr(Ref), 54, 1) <> "0" Then
             CurrentForm.Circle (Cent1 + Yank + 8, Cent2 - 4), 1, QBColor(Alle)
             CurrentForm.PSet (Cent1 + Yank + 8, Cent2 - 4), QBColor(Alle)
          End If
          If Mid$(WorldStr(Ref), 31, 1) <> " " Then
             CurrentForm.Circle (Cent1 + Yank - 8, Cent2 - 4), 1, QBColor(Alle)
             CurrentForm.PSet (Cent1 + Yank - 8, Cent2 - 4), QBColor(Alle)
          End If
       End If
    End If
    Rem nonworld markers
    Rem triangle
    If tmp5 = 1 And tmp4 < 2 Then
       x = Cent1 + Yank: y = Cent2
       CurrentForm.PSet (x, y), QBColor(0)
       CurrentForm.Line (x, y - 2)-(x + 2, y + 2), QBColor(Alle)
       CurrentForm.Line -Step(-4, 0), QBColor(Alle)
       CurrentForm.Line -Step(2, -4), QBColor(Alle)
       '->PAINT (x, y), alle
    End If
    Rem square
    If tmp5 = 2 And tmp4 < 2 Then
       x = Cent1 + Yank: y = Cent2
       CurrentForm.PSet (x, y), QBColor(0)
       CurrentForm.FillStyle = vbFSSolid
       CurrentForm.FillColor = QBColor(Alle)
       CurrentForm.Line (x - 2, y - 2)-(x + 2, y + 2), QBColor(Alle), BF
       CurrentForm.FillStyle = vbFSTransparent
    End If
    Rem diamond
    If tmp5 = 3 And tmp4 < 2 Then
       x = Cent1 + Yank: y = Cent2
       CurrentForm.PSet (x, y), QBColor(0)
       CurrentForm.Line (x, y - 2)-(x + 2, y), QBColor(Alle)
       CurrentForm.Line -Step(-2, 2), QBColor(Alle)
       CurrentForm.Line -Step(-2, -2), QBColor(Alle)
       CurrentForm.Line -Step(2, -2), QBColor(Alle)
       '->PAINT (x, y), alle
    End If
    Rem lotsadots
    If tmp5 = 4 And tmp4 < 2 Then
       x = Cent1 + Yank: y = Cent2
       Color Alle
       For X2 = -16 To 16 Step 4
          For Y2 = -16 To 16 Step 4
             If Abs(X2 * Y2) < 192 Then CurrentForm.PSet (x + X2, y + Y2)
          Next Y2
       Next X2
    End If
    Rem zone
    If tmp4 < 2 Or tmp5 = 0 Then
       sStr = Mid$(WorldStr(Ref), 49, 1)
       If sStr = "A" Then
            CurrentForm.Circle (Cent1 + Yank, Cent2), (HexA / 2) + 1, QBColor(14)
       End If
       If sStr = "R" Then
            CurrentForm.Circle (Cent1 + Yank, Cent2), (HexA / 2) + 1, QBColor(12)
       End If
       If sStr = "B" Then
            CurrentForm.Circle (Cent1 + Yank, Cent2), (HexA / 2) + 1, QBColor(11)
       End If
    End If
    Rem newjump circle
    If NewJumpStr = Mid$(WorldStr(Ref), 15, 4) And NewJFStr = fSubStr Then
       CurrentForm.Circle (Cent1 + Yank, Cent2), (HexA / 2), QBColor(5)
       CurrentForm.Circle (Cent1 + Yank, Cent2), (HexA / 2) + 4, QBColor(5)
    End If
    Rem write name
    wordStr = RTrim$(Left$(WorldStr(Ref), 14))
    If UWPtog = 1 Then wordStr = Mid$(WorldStr(Ref), 20, 9)
    lenword = Len(wordStr)
    x = Cent1 + Yank - (lenword * 3) + 1
    y = Cent2 + 10
    If tmp4 < 2 And NoName = 0 Then Call PrintWord(wordStr, x, y)
    Rem write starport
    wordStr = Mid$(WorldStr(Ref), 20, 1)
    x = Cent1 + Yank - 2: y = Cent2 - 5
    If tmp4 = 0 And wordStr <> "*" Then Call PrintWord(wordStr, x, y)
End Sub

Sub PrintWord(wordStr, x, y)
    Dim p As Picture
    Dim A As Integer
'5100     Rem subroutine: Print word
    y = y - 4
    L = Len(wordStr)
    If L = 0 Then Exit Sub
    For A = -1 To 5
       CurrentForm.Line (x - 1, y + A)-(x + L * 6 - 1, y + A), QBColor(0)
    Next A
    xunder = x: yunder = y + 6
    wordStr = UCase$(wordStr)
    For A = 1 To L
        letStr = Mid$(wordStr, A, 1)
        Rem convert letStr to letnum
        Select Case Asc(letStr)
           Case Is > 64
                letnum = Asc(letStr) - 64
           Case 48 To 57
                letnum = Asc(letStr) - 21
           Case 32
                letnum = 39
           Case 39
                letnum = 38
           Case 43
                letnum = 40
           Case 45
                letnum = 37
           Case 46
                letnum = 42
           Case 58
                letnum = 41
           Case 33
                letnum = 43
           Case 63
                letnum = 44
        End Select
    
        If x <= 635 And y <= 479 Then
            'frmGal.picLetter.Visible = True
            Set frmGal.picLetter.Picture = LetPicture(letnum)
            BitBlt frmGal.hDC, x, y, 5, 5, frmGal.picLetter.hDC, 0, 0, SRCCOPY
            frmGal.picLetter.Picture = Nothing
        End If
        x = x + 5 + 1
    Next A
    Rem underline if hi pop world
    hipop = HexToNumber(Mid$(WorldStr(Ref), 24, 1))
    If hipop >= 9 And L > 1 And NonWorld = 0 Then
       CurrentForm.Line (xunder, yunder)-(x - 2, yunder), QBColor(15)
    End If
    
End Sub

Sub EraseCircle()
'850     Rem erase circle
    Call FindCenterPt
    CurrentForm.Circle (Cent1 + Yank, Cent2), HexA / 2, QBColor(0)
    Call DrawStarPort
End Sub

Sub SubGraphicPgUpPgDn(Key1)
'1010     Rem pgup/pgdn
    Call EraseCircle
    If Key1 = 73 Then Hex = Hex - ((Hex - 1) Mod 10)
    If Key1 = 81 Then Hex = Hex + (9 - (Hex - 1) Mod 10)
    tmp1 = Int(Hex / 100)
    If Key1 = 15 Then Hex = (Hex Mod 100) + (tmp1 - ((tmp1 - 1) Mod 8)) * 100
    If Key1 = 9 Then Hex = (Hex Mod 100) + (tmp1 + (7 - (tmp1 - 1) Mod 8)) * 100
    Call MakeCircle
    Call DrawHex
End Sub

Sub SubGraphicUp()
'1000 Rem up
    Call EraseCircle
    Hex = Hex - 1
    If Right$(Str$(Hex), 1) = "0" Then
        If zI = 1 Then
            If NorthStr <> "<nil>" Then
                zI = 4: HHex = Hex + 10
                NewJY = NewJY + 4: yNow = yNow - 1
                Call ShiftOfSector
                Call ReadSectorData
                Trip = 1
                Call GotoSubScreen
                Call ReadSubsector
                'Call InitGraphic
                Call DrawSubGraphic
                Call DrawInitialCircle
            Else
                Hex = Hex + 1
                Call MakeCircle
                Call DrawHex
            End If
        Else
            zI = zI - 1
            HHex = Hex + 10
            Trip = 1
            Call GotoSubScreen
            Call ReadSubsector
            'Call InitGraphic
            Call DrawSubGraphic
            Call DrawInitialCircle
        End If
    Else
        Call MakeCircle
        Call DrawHex
    End If
End Sub

Sub SubGraphicDown()
'1020     Rem down
    Call EraseCircle
    Hex = Hex + 1
    If Right$(Str$(Hex), 1) = "1" Then
        If zI <> 4 Then
            zI = zI + 1
            HHex = Hex - 10
            Trip = 1
            Call GotoSubScreen
            Call ReadSubsector
            'Call InitGraphic
            Call DrawSubGraphic
            Call DrawInitialCircle
        Else
            If SouthStr <> "<nil>" Then
                zI = 1: HHex = Hex - 10
                NewJY = NewJY - 4: yNow = yNow + 1
                Call ShiftOfSector
                Call ReadSectorData
                Trip = 1
                Call GotoSubScreen
                Call ReadSubsector
                'Call InitGraphic
                Call DrawSubGraphic
                Call DrawInitialCircle
            Else
                Hex = Hex - 1
                Call MakeCircle
                Call DrawHex
            End If
        End If
    Else
        Call MakeCircle
        Call DrawHex
    End If
End Sub

Sub SubGraphicLeft()
'1040     Rem left
    Call EraseCircle
    Hex = Hex - 100
    If Hex < 100 Then
        If zJ <> 1 Then
            zJ = zJ - 1
            HHex = Hex + 800
            Trip = 1
            Call GotoSubScreen
            Call ReadSubsector
            'Call InitGraphic
            Call DrawSubGraphic
            Call DrawInitialCircle
        Else
            If WestStr <> "<nil>" Then
                zJ = 4: HHex = Hex + 800
                NewJX = NewJX + 4: xNow = xNow - 1
                Call ShiftOfSector
                Call ReadSectorData
                Trip = 1
                Call GotoSubScreen
                Call ReadSubsector
                'Call InitGraphic
                Call DrawSubGraphic
                Call DrawInitialCircle
            Else
                Hex = Hex + 100
                Call MakeCircle
                Call DrawHex
            End If
        End If
    Else
        Call MakeCircle
        Call DrawHex
    End If
End Sub

Sub SubGraphicRight()
'1060     Rem right
    Call EraseCircle
    Hex = Hex + 100
    If Hex > 900 Then
        If zJ <> 4 Then
            zJ = zJ + 1
            HHex = Hex - 800
            Trip = 1
            Call GotoSubScreen
            Call ReadSubsector
            'Call InitGraphic
            Call DrawSubGraphic
            Call DrawInitialCircle
        Else
            If EastStr <> "<nil>" Then
                zJ = 1: HHex = Hex - 800
                NewJX = NewJX - 4: xNow = xNow + 1
                Call ShiftOfSector
                Call ReadSectorData
                Trip = 1
                Call GotoSubScreen
                Call ReadSubsector
                'Call InitGraphic
                Call DrawSubGraphic
                Call DrawInitialCircle
            Else
                Hex = Hex - 100
                Call MakeCircle
                Call DrawHex
            End If
        End If
    Else
        Call MakeCircle
        Call DrawHex
    End If
End Sub

Sub EditSectorDat()
'315     Rem sector-level uwp.dat
    tmpStr = SecDirStr + "\uwp.dat"
    Open tmpStr For Random As #1
    Exist = LOF(1)
    Close 1
    If Exist <> 0 Then
        Screen 0
        Call DisplayEGAColor: Rem egacolor
        ShellPrg EditorStr + " " + tmpStr
        Call ReadGalUWPStats: Rem read uwp.dat
    Else
        Call DrawMiniMenu3
        Do
            Call GetKeyChar(Key1, Key2)
            If Key2 <> 1 Then
                If Key1 = 27 Or Key1 = 81 Or Key1 = 78 Then Exit Do: Rem esc/q/n
                If Key1 = 89 Then ' Rem y
                    ShellPrg "copy classic\beyond\uwp.dat " + tmpStr
                    Screen 0
                    Call DisplayEGAColor: Rem egacolor
                    ShellPrg EditorStr + " " + tmpStr
                    Call ReadGalUWPStats: Rem read uwp.dat
                    Exit Do
                End If
            End If
        Loop
    End If
    Screen 12
    Trip = 1
    HHex = Hex
    Call DrawSubGraphic
    Call DrawInitialCircle
End Sub

Sub ReDrawSubGraphic()
'290 Rem redraw subsector
    Trip = 1
    HHex = Hex
    Call DrawSubGraphic
    Call DrawInitialCircle
End Sub

Sub ToggleUWP()
'295 Rem uwp toggle
    If UWPtog = 0 Then UWPtog = 1 Else UWPtog = 0
    Call ReDrawSubGraphic
End Sub

Sub HexWorld()
'425     Rem hexworld
    If WorldStr(Ref) = "" Or NoName = 1 Then Exit Sub
    If Hidden > 0 Or NonWorld = 1 Then Exit Sub
    tmpStr = HexDirStr + "p" + Mid$(WorldStr(Ref), 15, 4)
    tmp2Str = RTrim$(Left$(WorldStr(Ref), 14))
    If HWStr <> "x" And HWStr <> "z" Then
       Rem fix subsector file
       If HWStr = "y" Then newhwStr = "z" Else newhwStr = "x"
       CLS
       Color 11
       Locate 5
       CurrentForm.Print "Creating a blank world..."
       ShellPrg "copy " + fSubStr + " data\tmpsub.dat"
       Open "data\tmpsub.dat" For Input As #1
       Open fSubStr For Output As #2
       Do Until EOF(1)
          Line Input #1, tsubStr
          If WorldStr(Ref) = tsubStr Then
             lline = Len(tsubStr)
             If lline < 63 Then
                tsubStr = tsubStr + Space$(63 - lline) + newhwStr + HiddenStr
             Else
                tsubStr = Left$(tsubStr, 63) + newhwStr + HiddenStr
             End If
          End If
          Print #2, tsubStr
       Loop
       Close 1
       Close 2
       ShellPrg "erase data\tmpsub.dat"
    End If
    Screen 0
    ShellPrg "sub\hexworld " + tmpStr + " " + tmp2Str
    Screen 12: HHex = Hex: Trip = 1
    Call ReadSubsector
    'Call InitGraphic
    Call DrawSubGraphic
    Call DrawInitialCircle
End Sub

Sub InsertJumpRoute()
'460 Rem insert jump route
    If NewJumpStr <> "" Then
        Call SecondJumpPt
    Else
        NewJumpStr = Mid$(WorldStr(Ref), 15, 4)
        NewJX = zJ: NewJY = zI: NewJFStr = fSubStr
        Call ReDrawSubGraphic
    End If
End Sub

Sub SecondJumpPt()
'465 Rem 1st world already selected
    If NewJumpStr <> Mid$(WorldStr(Ref), 15, 4) Then
        Call ConnectJumpPts
    Else
        NewJumpStr = ""
        Call ReDrawSubGraphic
    End If
End Sub

Sub ConnectJumpPts()
'470 Rem 1st & 2nd world different
    If Not (Abs(zJ - NewJX) > 1 Or Abs(zI - NewJY) > 1) Then
        Call DrawMiniMenu4: Rem get color
        Do
            Call GetKeyChar(Key1, Key2)
            If Key2 = 0 And Key1 >= 65 And Key1 <= 79 Then
                tmp1Str = LTrim$(Str$(Key1 - 64))
                Rem from 1st to 2nd
                tmp2Str = "$" + NewJumpStr + " " + Mid$(WorldStr(Ref), 15, 4)
                tmp2Str = tmp2Str + " " + LTrim$(Str$(zJ - NewJX))
                tmp2Str = tmp2Str + " " + LTrim$(Str$(zI - NewJY))
                tmp2Str = tmp2Str + " " + tmp1Str
                Rem from 2nd to 1st
                tmp3Str = "$" + Mid$(WorldStr(Ref), 15, 4) + " " + NewJumpStr
                tmp3Str = tmp3Str + " " + LTrim$(Str$(NewJX - zJ))
                tmp3Str = tmp3Str + " " + LTrim$(Str$(NewJY - zI))
                tmp3Str = tmp3Str + " " + tmp1Str
                tmpfileStr = NewJFStr: tmp4Str = tmp2Str: Call SaveJumpRoute(tmpfileStr, tmp4Str)
                If Abs(zJ - NewJX) = 1 Or Abs(zI - NewJY) = 1 Then
                   tmpfileStr = fSubStr: tmp4Str = tmp3Str: Call SaveJumpRoute(tmpfileStr, tmp4Str)
                End If
                Exit Do
            End If
            If Key2 = 0 And Key1 = 27 Or Key1 = 81 Then
                Exit Do
            End If
        Loop
    End If
    NewJumpStr = "": HHex = Hex: Trip = 1
    Call ReadSubsector
    'Call InitGraphic
    Call DrawSubGraphic
    Call DrawInitialCircle
End Sub

Sub SaveJumpRoute(tmpfileStr, tmp4Str)
    ShellPrg "copy " + tmpfileStr + " data\tmpsub.dat"
    Open "data\tmpsub.dat" For Input As #1
    Open tmpfileStr For Output As #2
    Do Until EOF(1)
        Line Input #1, tmpStr
        If Left$(tmpStr, 10) = "#--------1" Then Print #2, tmp4Str
        Print #2, tmpStr
    Loop
    Close 1
    Close 2
    ShellPrg "erase data\tmpsub.dat"
End Sub

Sub SystemMap()
'430     Rem sysmap
    If WorldStr(Ref) = "" Or NoName = 1 Then Exit Sub
    If Hidden > 0 Or NonWorld = 1 Then Exit Sub
    tmpStr = SysDirStr + "s" + Mid$(WorldStr(Ref), 15, 4)
    tmp2Str = RTrim$(Left$(WorldStr(Ref), 14))
    If HWStr <> "y" And HWStr <> "z" Then
       Rem fix subsector file
       If HWStr = "x" Then newhwStr = "z" Else newhwStr = "y"
       CLS
       Color 11
       Locate 5
       CurrentForm.Print "Creating a star system..."
       ShellPrg "copy " + fSubStr + " data\tmpsub.dat"
       Open "data\tmpsub.dat" For Input As #1
       Open fSubStr For Output As #1
       Do Until EOF(1)
          Line Input #1, tsubStr
          If WorldStr(Ref) = tsubStr Then
             lline = Len(tsubStr)
             If lline < 63 Then
                tsubStr = tsubStr + Space$(63 - lline) + newhwStr + HiddenStr
             Else
                tsubStr = Left$(tsubStr, 63) + newhwStr + HiddenStr
             End If
          End If
          Print #2, tsubStr
       Loop
       Close 1
       Close 2
       ShellPrg "erase data\tmpsub.dat"
    End If
    Screen 0
    ShellPrg "sub\sysmap " + tmpStr + " " + tmp2Str
    Screen 12: HHex = Hex: Trip = 1
    Call ReadSubsector
    'Call InitGraphic
    Call DrawSubGraphic
    Call DrawInitialCircle
End Sub

Sub DeleteWorld()
'440     Rem delete world (or non-world)
    If Len(WorldStr(Ref)) = 0 Then
        Call ReDrawSubGraphic
        Exit Sub
    End If
    Call DrawMiniMenu2: Rem insert star query
    Do
        Call GetKeyChar(Key1, Key2)
        If Key2 = 0 And Key1 = 89 Then ' Rem y
            Screen 0
            ShellPrg "sub\insdel (del) " + HHStr + " " + fSubStr
            Screen 12: HHex = Hex: Trip = 1
            Call ReadSubsector
            'Call InitGraphic
            Call DrawSubGraphic
            Call DrawInitialCircle
            Exit Do
        End If
        If Key2 = 0 And Key1 = 78 Then ' Rem n
            Call ReDrawSubGraphic
            Exit Do
        End If
        If Key2 = 0 And Key1 = 27 Then ' Rem esc
            Call ReDrawSubGraphic
            Exit Do
        End If
        If Key2 = 0 And Key1 = 81 Then ' Rem q
            Call ReDrawSubGraphic
            Exit Do
        End If
    Loop
End Sub

Sub EditWriteUp(Key1)
'330     Rem view/edit subsector write-up
    If subtxt%(zK) = 0 And Key1 = 59 Then Exit Sub
    newfile = 0
    If Key1 = 59 Then AppStr = ReaderStr
    If Key1 = 60 Then AppStr = EditorStr
    txtsubStr = LocDirStr + "sub-" + Chr$(zK + 64)
    If subtxt%(zK) = 0 Then
        newfile = 1
        tmpStr = "subsector"
        Call DrawMiniMenu1: Rem mini-menu
        Do
            Call GetKeyChar(Key1, Key2)
            If Key2 = 0 And Key1 = 70 Then ' Rem f
                subtxt%(zK) = 1: abvStr = "f"
                Exit Do
            End If
            If Key2 = 0 And Key1 = 77 Then ' Rem m
                subtxt%(zK) = 2: abvStr = "m"
                Exit Do
            End If
            If Key2 = 0 And Key1 = 27 Then ' Rem esc
                Call ReDrawSubGraphic
                Exit Sub
            End If
            If Key2 = 0 And Key1 = 81 Then ' Rem q
                Call ReDrawSubGraphic
                Exit Sub
            End If
        Loop
    End If
    If subtxt%(zK) = 1 Then txtsubStr = txtsubStr + ".txt"
    If subtxt%(zK) = 2 Then txtsubStr = txtsubStr + ".mnu"
    Screen 0
    Call DisplayEGAColor: Rem egacolor
    If subtxt%(zK) = 1 And newfile = 0 Then
       ShellPrg AppStr + " " + txtsubStr
    End If
    If subtxt%(zK) = 2 And newfile = 0 Then
       '->Eg txtsubStr + " " + AppStr
    End If
    If subtxt%(zK) = 1 And newfile = 1 Then
       Call InsertAbv(abvStr): Rem insert abvStr into sector file
       Rem set up initial subsector text file
       fStr = txtsubStr
       Rem check if file fStr exists
       Open fStr For Random As #1
       Exist = LOF(1)
       Close 1
       If Exist = 0 Then
          Open txtsubStr For Output As #1
          tmpStr = SubNameStr(zK, 1) + " Subsector"
          Print #1, ""
          Print #1, tmpStr
          Print #1, String$(Len(tmpStr), "-")
          Print #1, ""
          Close 1
       End If: Rem exist=0
       Rem go into subsector text file
       ShellPrg AppStr + " " + txtsubStr
    Else
       If subtxt%(zK) = 2 And newfile = 1 Then
            Call InsertAbv(abvStr): Rem insert abvStr into sector file
            Rem set up initial subsector menu file
            fStr = txtsubStr
            f1Str = "sub-" + Chr$(zK + 64)
            Rem check if file fStr exists
            Open fStr For Random As #1
            Exist = LOF(1)
            Close 1
            If Exist = 0 Then
               ShellPrg "erase " + fStr
               Open fStr For Output As #1
               Open "data\starter2.mnu" For Input As #2
               Print #1, "@dir=" + LocDirStr
               Print #1, "@Light Yellow"
               Print #1, ""
               Print #1, "          ";
               Print #1, SubNameStr(zK, 1) + " Subsector"
               tmpcount = 0
               Do Until EOF(2)
                  Line Input #2, tStr
                  lstart = Len(tStr)
                  If lstart = 1 Then tStr = ""
                  If lstart > 1 Then tStr = Right$(tStr, lstart - 1)
                  Print #1, tStr;
                  If Right$(tStr, 1) = "@" Then
                     tmpcount = tmpcount + 1
                     If tmpcount = 1 Then Print #1, f1Str + ".txt"
                     If tmpcount = 2 Then Print #1, f1Str + "2.txt"
                  Else
                     Print #1, ""
                  End If: Rem @
               Loop
               Close 1
               Close 2
            End If: Rem exist=0
            Rem go into subsector menu file
            Screen 0
            Call DisplayEGAColor: Rem egacolor
            '->Eg txtsubStr + " " + AppStr
        End If
    End If
    Screen 12
    Call ReDrawSubGraphic
End Sub

Sub InsertAbv(abvStr)
'400     Rem insert abvStr into sector file
    fStr = SecFPStr
    ShellPrg "copy " + fStr + " sector.tmp"
    Open "sector.tmp" For Input As #1
    Open fStr For Output As #2
    For I = 1 To 7
       Line Input #1, tStr
       Print #2, tStr
    Next I
    For I = 1 To 16
       Line Input #1, tStr
       L = Len(tStr)
       If I = zK Then
          If L < 49 Then tStr = tStr + String$(49 - L, " ") + abvStr
          If L >= 49 Then tStr = Left$(tStr, 49) + abvStr
       End If
       Print #2, tStr
    Next I
    Do Until EOF(1)
       Line Input #1, tStr
       Print #2, tStr
    Loop
    Close 1
    Close 2
    ShellPrg "erase sector.tmp"
End Sub

Sub DrawSubSectorScreen()
'90     Rem goto subsector screen
    Call GotoSubScreen
    Call ReadSubsector
    'Call InitGraphic
    Call DrawSubGraphic
    Call DrawInitialCircle
    Do
        Call GetKeyChar(Key1, Key2)
        If Key2 = 1 And Key1 = 72 Then Call SubGraphicUp: Rem up
        If Key2 = 0 And Key1 = 75 Then Call SubGraphicUp: Rem k
        If Key2 = 1 And (Key1 = 73 Or Key1 = 81) Then Call SubGraphicPgUpPgDn(Key1): Rem pgup/pgdn
        If Key2 = 0 And Key1 = 9 Then Call SubGraphicPgUpPgDn(Key1): Rem fwd tab
        If Key2 = 1 And Key1 = 15 Then Call SubGraphicPgUpPgDn(Key1): Rem back tab
        If Key2 = 1 And Key1 = 80 Then Call SubGraphicDown: Rem down
        If Key2 = 0 And Key1 = 74 Then Call SubGraphicDown: Rem j
        If Key2 = 1 And Key1 = 75 Then Call SubGraphicLeft: Rem left
        If Key2 = 0 And Key1 = 72 Then Call SubGraphicLeft: Rem h
        If Key2 = 1 And Key1 = 77 Then Call SubGraphicRight: Rem right
        If Key2 = 0 And Key1 = 76 Then Call SubGraphicRight: Rem l
        If Key2 = 0 And Key1 = 27 Then Exit Do: Rem esc
        If Key2 = 0 And Key1 = 81 Then Exit Do: Rem q
        If Key2 = 0 And Key1 = 13 Then Call DrawMoreInfo(Key1): Rem enter
        If Key2 = 0 And Key1 = 32 Then Call DrawMoreInfo(Key1): Rem space
        If Key2 = 0 And Key1 = 62 Then Call DrawMoreInfo(Key1): Rem >
        If Key2 = 0 And Key1 = 85 Then Call EditSectorDat: Rem u
        If Key2 = 0 And Key1 = 12 Then Call ReDrawSubGraphic: Rem ctrl-l
        If Key2 = 0 And Key1 = 21 Then Call ToggleUWP: Rem ctrl-u
        If Key2 = 1 And (Key1 = 59 Or Key1 = 60) Then Call EditWriteUp(Key1): Rem F1/F2
        If Key2 = 1 And Key1 = 62 Then Call HexWorld: Rem F4
        If Key2 = 1 And Key1 = 64 Then Call InsertJumpRoute: Rem F6
        If Key2 = 1 And Key1 = 66 Then Call SystemMap: Rem F8
        If Key2 = 1 And Key1 = 83 Then Call DeleteWorld: Rem delete
        catch = 0: Rem hook catching
        For Tmp = 1 To nHook%(3)
           If Key2 = Hooke%(3, Tmp) And Key1 = HookAc%(3, Tmp) Then catch = Tmp
        Next Tmp
        If catch > 0 Then Call CatchSubGraphicHook(catch)
        Call DrawHex
    Loop
    Call DrawSubSectorOuter
End Sub

Sub GetSectorData()
'40
    'If NoSector = 1 Then GoTo 10
    Call ReadSectorData
    Call DrawSubSectorOuter
    
    Do
        Call DrawSubSectorInner
        Call GetKeyChar(Key1, Key2)
        Call DrawSubSectorInstructions
        If Key2 = 1 And Key1 = 72 Then Call SubSectorUp: Rem up
        If Key2 = 0 And Key1 = 75 Then Call SubSectorUp: Rem k
        If Key2 = 1 And Key1 = 80 Then Call SubSectorDown: Rem down
        If Key2 = 0 And Key1 = 74 Then Call SubSectorDown: Rem j
        If Key2 = 1 And Key1 = 75 Then Call SubSectorLeft: Rem left
        If Key2 = 0 And Key1 = 72 Then Call SubSectorLeft: Rem h
        If Key2 = 1 And Key1 = 77 Then Call SubSectorRight: Rem right
        If Key2 = 0 And Key1 = 76 Then Call SubSectorRight: Rem l
        If Key2 = 1 And Key1 = 73 Then zI = 1: Rem pgup
        If Key2 = 1 And Key1 = 81 Then zI = 4: Rem pgdn
        If Key2 = 0 And Key1 = 9 Then zJ = 4: Rem fwd tab
        If Key2 = 1 And Key1 = 15 Then zJ = 1: Rem back tab
        If Key2 = 0 And Key1 = 27 Then Exit Do: Rem esc
        If Key2 = 0 And Key1 = 81 Then Exit Do: Rem q
        If Key2 = 0 And Key1 = 13 Then Call DrawSubSectorScreen: Rem enter
        If Key2 = 0 And Key1 = 32 Then Call DrawSubSectorScreen: Rem space
        catch = 0: Rem hook catching
        For Tmp = 1 To nHook%(2)
           If Key2 = Hooke%(2, Tmp) And Key1 = HookAc%(2, Tmp) Then catch = Tmp
        Next Tmp
        If catch > 0 Then Call CaughtSubHook(catch)
    Loop
    Call DrawSectorMap
End Sub

Sub SelectGalaxy()
'5760 Rem select galaxy
    Call ReadGalaxyList
    Call ReadGalaxyWhereDat
    Call DrawSelectGalaxy
    Call DrawSelectGalaxyArrow
    Do
        Call GetKeyChar(Key1, Key2)
        If Key2 = 0 And Key1 = 27 Then Exit Do: Rem esc
        If Key2 = 0 And Key1 = 81 Then Exit Do: Rem Q 6,1
        If Key2 = 0 And Key1 = 60 Then Call EditGalLst: Rem <
        If Key2 = 0 And (Key1 = 63 Or Key1 = 72) Then
            Call GalaxyHelpToggle: Rem ?/H
        End If
        If Key2 = 1 And Key1 = 82 Then Call CreateGalaxy: Rem insert (create)
        If Key2 = 0 And Key1 = 73 Then Call ImportFromArchive: Rem I (import)
        If GalaxyList.Count > 0 Then
            If Key2 = 1 And (Key1 = 75 Or Key1 = 77) Then Call ProcessSelectGalaxyArrow(Key1): Rem left/right
            If Key2 = 1 And (Key1 = 72 Or Key1 = 80) Then Call ProcessSelectGalaxyArrow(Key1): Rem up/down
            If Key2 = 1 And (Key1 = 73 Or Key1 = 81) Then Call ProcessSelectGalaxyArrow(Key1): Rem pgup/pgdn
            If Key2 = 0 And Key1 = 13 Then ' Rem enter
                CurrentGalaxy = GalaxyList(CurrentGalaxyIndex).Directory
                GalDirStr = "gals\" + CurrentGalaxy
                Call WriteGalaxyWhereDat: Rem write gals\where.dat
                Call ReadSectorWhereDat: Rem read gals\<sector>\where.dat
                Call ReadSectorList: Rem read list of sectors
                CurrentSector = SectorList(FNsec(xNow, yNow), 1)
                Call ReadSectorPaths: Rem file/directory paths
                Call DrawLocalMap: Rem galaxy map
            End If
            If Key2 = 0 And Key1 = 62 Then Call EditGalaxyLst: Rem >
            If Key2 = 1 And Key1 = 59 Then Call GalaxyReadMenu: Rem F1
            If Key2 = 1 And Key1 = 60 Then Call GalaxyEditMenu: Rem F2
            If Key2 = 1 And Key1 = 63 Then Call GalaxySectorMenu: Rem F5
            If Key2 = 1 And Key1 = 83 Then Call DeleteGalaxy: Rem delete
            If Key2 = 0 And Key1 = 88 Then Call ExportToArchive: Rem X (export)
        End If
    Loop
    Call WriteGalaxyWhereDat: Rem write gals\where.dat
    Call DrawFrontPage
    Call LocateArrow(MainMenuPointer)
End Sub

Sub DrawLocalMap()
'3000     Rem get local map data
    Call DrawSectorMap
    Do
        Call GetKeyChar(Key1, Key2)
        If Key2 = 1 And Key1 = 72 Then Call SectorMapUp: Rem up
        If Key2 = 0 And Key1 = 75 Then Call SectorMapUp: Rem k
        If Key2 = 1 And Key1 = 80 Then Call SectorMapDown: Rem down
        If Key2 = 0 And Key1 = 74 Then Call SectorMapDown: Rem j
        If Key2 = 1 And Key1 = 75 Then Call SectorMapLeft: Rem left
        If Key2 = 0 And Key1 = 72 Then Call SectorMapLeft: Rem h
        If Key2 = 1 And Key1 = 77 Then Call SectorMapRight: Rem right
        If Key2 = 0 And Key1 = 76 Then Call SectorMapRight: Rem l
        If Key2 = 0 And Key1 = 27 Then Exit Do: Rem esc
        If Key2 = 0 And Key1 = 81 Then Exit Do: Rem q
        If Key2 = 0 And Key1 = 13 Then Call SelectSector: Rem enter
        If Key2 = 0 And Key1 = 32 Then Call SelectSector: Rem space
        If Key2 = 0 And Key1 = 78 Then Call AddSectorNameToMap: Rem n
        If Key2 = 0 And Key1 = 73 Then Call ImportSectorFromArchive: Rem i
        If Key2 = 0 And Key1 = 88 Then Call ExportSectorToArchive: Rem x
        If Key2 = 1 And Key1 = 83 Then Call DeleteSector: Rem delete
        If Key2 = 1 And Key1 = 82 Then Call CreateSector: Rem insert (create)
        If Key2 = 1 And Key1 = 64 Then Call MoveSector: Rem F6
        catch = 0: Rem hook catching
        For Tmp = 1 To nHook%(1)
           If Key2 = Hooke%(1, Tmp) And Key1 = HookAc%(1, Tmp) Then catch = Tmp
        Next Tmp
        If catch > 0 Then Call CaughtSectorHook(catch)
    Loop
    Call WriteSectorWhereDat: Rem write gals\<sector>\where.dat
    Color 7
    CurrentForm.Print "."
    Call DrawSelectGalaxy
    Call DrawSelectGalaxyArrow
End Sub

Sub SelectSector()
'3550     Rem enter
    Tmp = FNsec(xNow, yNow)
    If Tmp = 0 Then Call DrawLocalMap
    If Column(Tmp) = 8 Then
        CreateSector
        Call DrawSectorMap
    Else
        Color 7
        CurrentForm.Print "."
        Call ShiftOfSector: Rem shift of sector
        Call GetSectorData
    End If
End Sub

Public Sub Color(RHS)
    Select Case ClrMode
     Case 1
        CurrentForm.ForeColor = QBColor(RHS)
        frmGal.BackColor = vbBlack
        frmGal.picLetter.BackColor = vbBlack
     Case 2:
        CurrentForm.ForeColor = QBColor(RHS)
        frmGal.BackColor = vbWhite
        frmGal.picLetter.BackColor = vbWhite
     Case 3:
        frmGal.ForeColor = vbWhite
        frmGal.BackColor = vbBlack
        frmGal.picLetter.BackColor = vbBlack
     Case 4:
        frmGal.BackColor = vbWhite
        frmGal.picLetter.BackColor = vbWhite
    End Select
End Sub

Public Sub CLS()
    CurrentForm.CLS
End Sub

Public Sub Screen(Mode As Integer)
'    If Mode = 0 Then
'        CurrentForm.Scale (0, 0)-(80, 50)
'    ElseIf Mode = 12 Then
'        CurrentForm.Scale (0, 0)-(800, 600)
'    End If
End Sub

Public Sub Locate(Optional y, Optional x)
    Dim TextHeight As Double
    Dim TextWidth As Double
    TextHeight = CurrentForm.TextHeight("W")
    TextWidth = CurrentForm.TextWidth("W")
    If Not IsMissing(x) Then CurrentForm.CurrentX = x * TextWidth
    If Not IsMissing(y) Then CurrentForm.CurrentY = y * TextHeight
End Sub

Sub Swap(A, B)
    Dim C
    C = B
    B = A
    A = C
End Sub

Sub DrawTest()
    Screen 0
    CLS
    Color 14
    'ShellPrg "sub\funky data\title.dat"
    
    CurrentForm.Print
    Color 11
    CurrentForm.Print "      Alpha"
    CurrentForm.Print "      Beta"
    CurrentForm.Print
    CurrentForm.Print "      Gamma     ";
    Color 13: CurrentForm.Print "("; EditorStr; ")"
    Color 11: CurrentForm.Print "      Select Text Reader     ";
    Color 13: CurrentForm.Print "("; ReaderStr; ")"
    Color 11: CurrentForm.Print "      Foreground/Background  ";
    Color 13
    Select Case ClrMode
     Case 1:
      CurrentForm.Print "(colors on black)"
     Case 2:
      CurrentForm.Print "(colors on white)"
     Case 3:
      CurrentForm.Print "(white on black)"
     Case 4:
      CurrentForm.Print "(black on white)"
    End Select
    Color 11
    CurrentForm.Print
    CurrentForm.Print "      Quit"
    Color 9
    CurrentForm.Print
    CurrentForm.Print "1"
    CurrentForm.Print "2";
    CurrentForm.Print "3"
End Sub

Public Function InputText(Caption As String) As String
    Dim TempY As Double
    TempY = CurrentForm.CurrentY
    CurrentForm.Print Caption
    CurrentForm.txtInput.Top = CurrentForm.CurrentY - CurrentForm.txtInput.Height * 0.75
    CurrentForm.txtInput.Left = CurrentForm.TextWidth(Caption & " ")
    CurrentForm.txtInput.Text = ""
    Select Case ClrMode
     Case 1:
        CurrentForm.txtInput.BackColor = vbBlack
        CurrentForm.txtInput.ForeColor = CurrentForm.ForeColor
     Case 2:
        CurrentForm.txtInput.BackColor = vbWhite
        CurrentForm.txtInput.ForeColor = CurrentForm.ForeColor
     Case 3:
        CurrentForm.txtInput.BackColor = vbBlack
        CurrentForm.txtInput.ForeColor = vbWhite
     Case 4:
        CurrentForm.txtInput.BackColor = vbWhite
        CurrentForm.txtInput.ForeColor = vbBlack
    End Select
    CurrentForm.txtInput.Visible = True
    CurrentForm.txtInput.SetFocus
    Do
        Call GetKeyChar(Key1, Key2)
        If Key2 = 0 And Key1 = 27 Then Exit Do: Rem esc
        If Key2 = 0 And Key1 = 13 Then
            InputText = CurrentForm.txtInput.Text
            Exit Do
        End If
    Loop
    CurrentForm.txtInput.Visible = False
    CurrentForm.CurrentX = CurrentForm.txtInput.Left
    CurrentForm.CurrentY = TempY
    CurrentForm.Print CurrentForm.txtInput.Text
End Function

Public Sub ReadMenu(filename As String)
'    Dim FileNum As Integer
'    Dim TempS As String
'    If Me.lstGalaxy.ListIndex < 0 Then
'        Exit Sub
'    End If
'    TempS = MyGalaxies(lstGalaxy.ItemData(lstGalaxy.ListIndex)).Directory
'    frmNotes.Directory = App.Path & "\gals\" & TempS & "\gen\"
'    frmNotes.MenuName = "Galaxy.mnu"
'    Set frmNotes.Source = Me
'    frmNotes.Show
'    Me.Hide
End Sub

Public Sub ShellPrg(Execute As String, Optional Focus As VbAppWinStyle = vbMaximizedFocus)
    Dim V As Variant
    If InStr(UCase$(Execute), "SECMAP") > 0 Then
        V = Split(Execute, " ", 2)
        If UBound(V) = 0 Then
            SecMap ""
        Else
            SecMap CStr(V(1))
        End If
    ElseIf InStr(UCase$(Execute), "GMAP") > 0 Then
        V = Split(Execute, " ", 2)
        If UBound(V) = 0 Then
            GMap ""
        Else
            GMap CStr(V(1))
        End If
    ElseIf Execute <> "" Then
        Shell Execute, Focus
    End If
End Sub


