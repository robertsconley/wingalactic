Attribute VB_Name = "modGMap"
Dim GalDir$
Dim GalLst$
Dim SecName$(64)
Dim SecFile$(64)
Dim SecPlot$(64)
Dim SecQueries%(64)
Dim SecX%(64)
Dim SecY%(64)
Dim NumSectors%
Dim QueryLog%
Dim CacheSize%
Dim FocusX%
Dim FocusY%
Dim FocNum%
Dim cHexA
Dim cHexB
Dim Yank
Dim zI
Dim zJ
Dim zK

' global variables
Dim SubName$(16, 2)
Dim world$(80)
Dim basalg$(2, 50, 2)
Dim basalgn(2)
Dim algncol(50)
Dim subtxt%(16)
Dim mapscale%

Public Sub GMap(Cmd As String)
    On Error GoTo Test
    If Cmd = "" Then
        CurrentForm.CLS
        Color 11
        CurrentForm.Print "This program is not meant to be run as a stand-alone"
        CurrentForm.Print "piece of software. It is part of the 'Galactic' package,"
        CurrentForm.Print "and is meant to be executed from gal.exe."
        CurrentForm.Print
        CurrentForm.Print "Hit any key to Exit"
        GetKeyChar k1, k2
        ExitGMap
    Else
        NumSectors% = 0
        QueryLog% = 1
        CacheSize% = 64
        mapscale% = 1
        
        ReadSettings Cmd
        VideoSettings
        
        Screen 12
    
        Do    ' graphical sector
            CurrentForm.CLS
             ' plot stars
            If mapscale% = 1 Then
              PlotSectorGrid FocusX%, FocusY%, 1, 1, 6, 3, 1
            Else
              PlotSectorGrid FocusX%, FocusY%, 1, 1, 18, 9, 2
            End If
            Color 11
            Locate 24, 3
            CurrentForm.Print "Centerpoint: " + FindSecName$(FocusX%, FocusY%)
            Locate 25, 70
            Color 13
            CurrentForm.Print "? = Help"
            Color 11
            Do
                GetKeyChar k1, k2
                If k2 = 1 And k1 = 72 Then ' up
                    FocusY% = FocusY% - 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 0 And k1 = 75 Then ' k
                    FocusY% = FocusY% - 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 1 And k1 = 80 Then ' down
                    FocusY% = FocusY% + 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 0 And k1 = 74 Then ' j
                    FocusY% = FocusY% + 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 1 And k1 = 75 Then ' left
                    FocusX% = FocusX% - 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 0 And k1 = 72 Then ' h
                    FocusX% = FocusX% - 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 1 And k1 = 77 Then ' right
                    FocusX% = FocusX% + 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 0 And k1 = 76 Then ' l
                    FocusX% = FocusX% + 1
                    FocNum% = FindSector(FocusX%, FocusY%)
                    Exit Do
                End If
                If k2 = 0 And k1 = 27 Then ' esc
                    ExitGMap
                    Exit Sub
                End If
                If k2 = 0 And k1 = 81 Then ' q
                    ExitGMap
                    Exit Sub
                End If
                If k2 = 0 And k1 = 83 Then ' s
                    If mapscale% = 1 Then
                      mapscale% = 2
                    Else
                      mapscale% = 1
                    End If
                    Exit Do
                End If
                If k2 = 0 And k1 = 63 Then ' ?
                    Screen 0
                    Color 7
                    CurrentForm.CLS
                    CurrentForm.Print "."
                    ShellPrg Reader$ + " help\gmap.txt", vbMaximizedFocus
                    Screen 12
                    Exit Do
                End If
                If k2 = 1 And k1 = 68 Then ' F10 'Save Screen
                End If
            Loop
       Loop
    End If
    Exit Sub
Test:
    Resume Next
End Sub

Public Sub ExitGMap()
    ' end
    ' save where
    fd% = FreeFile
    Open GalDir$ + "\where.dat" For Output As fd%
        Print #fd%, Str$(FocusX%)
        Print #fd%, Str$(FocusY%)
    Close #fd%
    Color 7
    CurrentForm.CLS
    CurrentForm.Print
    CurrentForm.Print "For Assistance and/or Snide 'arks:"
    CurrentForm.Print
    CurrentForm.Print "     Email: jaymin@maths.tcd.ie (Jo)"
    CurrentForm.Print "         or jimv@empirenet.com (Jim)"
End Sub

Public Sub ReadSettings(Cmd As String)
    ' read in settings
    ' Global settings
    fd% = FreeFile
    Open "gal.cfg" For Input As fd%
        Input #fd%, T$
        L = Len(T$)
        Editor$ = Right$(T$, L - 7)
        Input #fd%, T$
        L = Len(T$)
        Reader$ = Right$(T$, L - 7)
    Close fd%
    ' Command line contains path to galaxy
    GalDir$ = LCase$(Cmd)
    ' work out galaxy.lst file
    o% = InStr(GalDir$, "\")
    GalLst$ = GalDir$ + "\" + Mid$(GalDir$, o% + 1) + ".lst"
    ' Current Focus settings
    fd% = FreeFile
    Open GalDir$ + "\where.dat" For Input As fd%
        Input #fd%, T$
        FocusX% = Val(T$)
        Input #fd%, T$
        FocusY% = Val(T$)
    Close #fd%
    FocNum% = FindSector(FocusX%, FocusY%)
End Sub

Public Sub VideoSettings()
    xMost = 640
    yMost = 500
    ' screen aspect ratio
    ' scar = 2.4
    Scar = (3 / 4) * (xMost / yMost)
    ' hex radii
    HexA = 25
    HexB = Int((HexA ^ 2 - (0.5 * HexA) ^ 2) ^ 0.5)
    ' corrected for aspect ratio
    cHexA = HexA * Scar
    cHexB = HexB
    ' yank right
    Yank = xMost - Int(cHexA * 1.5 * 8.3333) - 22
    zI = 1
    zJ = 1
End Sub

Function FindSecName$(x%, y%)
  For I% = 0 To NumSectors% - 1
    If SecX%(I%) = x% And SecY%(I%) = y% Then
      FindSecName$ = SecName$(I%)
      Exit Function
    End If
  Next I%
  FindSecName$ = Str$(x%) + "," + Str$(y%)
End Function

Function FindSector%(x%, y%)
  Rem This is the base function for retreiving sectors. For speed we
  Rem cache a number of sectors. The first thing we do is search the
  Rem cache for the sector. If found, we return immediately. If not
  Rem we check to see if there is any space left in the cache. If so,
  Rem we read the new sector in. If not, we search the cache for the
  Rem oldest referenced sector, and overwrite it.
  Rem Each time we reference a sector we update the query number to
  Rem a sequentially increasing value. We can then use this later to
  Rem determine the sector that hasn't been referenced the longest by
  Rem just looking for the lowest number.

  Rem first see if already loaded
  For I% = 0 To NumSectors% - 1
    If SecX%(I%) = x% And SecY%(I%) = y% Then
      SecQueries%(I%) = QueryLog%
      QueryLog% = QueryLog% + 1
      FindSector% = I%
      Exit Function
    End If
  Next I%
  Rem Second read into cache if there is space
  If NumSectors% <= CacheSize% Then
    bestsec% = ReadSector(x%, y%, NumSectors%)
    If bestsec% >= 0 Then NumSectors% = NumSectors% + 1
    FindSector% = bestsec%
    Exit Function
  End If
  Rem Thrid Find least recently used and zap
  bestsec% = 0
  For I% = 1 To NumSectors% - 1
    If SecQueries%(I%) < SecQueries%(bestsec%) Then
      bestsec% = I%
    End If
  Next I%
  Rem Read into oldest
  bestsec% = ReadSector(x%, y%, bestsec%)
  If bestsec% >= 0 Then
    SecQueries%(bestsec%) = QueryLog%
    QueryLog% = QueryLog% + 1
  End If
  FindSector% = bestsec%
End Function

Function IsPlot%(plot$, x%, y%)
  Rem The Plot$ is used to maintain a list of chars that is, in fact, a
  Rem bitfield containing the on-off state for all hexes in a sector.
  Rem There wasn't enough string space to do it on a byte level so we
  Rem had to do it this way.
  Rem Here, we check to see if hex x,y is set and return 0 or 1 accordingly.
  Dim TempS As String
  offset% = (x% - 1) + (y% - 1) * 32
  byteoffset% = Int(offset% / 8) + 1
  bitoffset% = (offset% Mod 8)
  TempS = Mid$(plot$, byteoffset%, 1)
  mask% = pow2(bitoffset%)
  bval% = Asc(TempS)
  If (bval% And mask%) = 0 Then
    IsPlot% = 0
  Else
    IsPlot% = 1
  End If
End Function

Function LookupSector$(x%, y%)
  LookupSector$ = ""
  fd% = FreeFile
  Open GalLst$ For Input As fd%
  Input #fd%, T$
  Input #fd%, T$
  Do While Not EOF(fd%)
    Input #fd%, T$
    xx% = Val(Mid$(T$, 51, 4))
    yy% = Val(Mid$(T$, 56, 4))
    seccol% = Asc(Mid$(T$, 61, 1)) - 65
    If xx% = x% And yy% = y% And seccol% <> 8 Then
      LookupSector$ = RTrim$(Left$(T$, 12))
      Exit Do
    End If
  Loop
  Close fd%
End Function

Sub PlotSector(ox%, oy%, sec%, scalesec%)
  Rem Given the offset and sector reference, we get the sector
  Rem and draw it on the screen.
  plot$ = SecPlot$(sec%)
  Color 15
  For y% = 1 To 40
    For x% = 1 To 32
      If IsPlot%(plot$, x%, y%) <> 0 Then
        If scalesec% = 1 Then
          X1% = x% * 3 + ox%: Y1% = y% * 3 + oy%
          If x% Mod 2 = 0 Then Y1% = Y1% + 1
          CurrentForm.PSet (X1%, Y1%)
          CurrentForm.PSet (X1% + 1, Y1%)
          CurrentForm.PSet (X1%, Y1% + 1)
          CurrentForm.PSet (X1% + 1, Y1% + 1)
        ElseIf scalesec% = 2 Then
          X1% = x% + ox%: Y1% = y% + oy%
          CurrentForm.PSet (X1%, Y1%)
        End If
      End If
    Next x%
  Next y%
End Sub

Sub PlotSectorGrid(xx%, yy%, ox%, oy%, gx%, gy%, scalesec%)
  Rem This does the whole business of working out the grid of sectors
  Rem (for the given size) and drawing all of them on the screen.
  Dim SecGrid$()
  ReDim SecGrid$(gx%, gy%)
  basex% = -(gx% / 2)
  basex% = basex% + xx%
  basey% = -(gy% / 2)
  basey% = basey% + yy%
  For y% = 1 To gy%
    For x% = 1 To gx%
      o% = FindSector(basex% + x%, basey% + y%)
      If o% >= 0 Then
        If scalesec% = 1 Then
          PlotSector ox% + 32 * 3 * (x% - 1), oy% + (40 * 3 - 1) * (y% - 1), o%, scalesec%
        Else
          PlotSector ox% + 32 * (x% - 1), oy% + 40 * (y% - 1), o%, scalesec%
        End If
      End If
    Next x%
  Next y%
End Sub

Function pow2%(expn%)
  Rem Claculate 2 to the power of expn% For masking,
  Select Case expn%
    Case 0
      pow2% = 1
    Case 1
      pow2% = 2
    Case 2
      pow2% = 4
    Case 3
      pow2% = 8
    Case 4
      pow2% = 16
    Case 5
      pow2% = 32
    Case 6
      pow2% = 64
    Case 7
      pow2% = 128
    End Select
End Function

Function ReadSector%(x%, y%, into%)
  Dim SubName$
  On Error GoTo Test
  SecFil$ = LookupSector(x%, y%)
  If SecFil$ = "" Then
    ReadSector% = -1
    Exit Function
  End If
  SecX%(into%) = x%
  SecY%(into%) = y%
  Rem Read a sector from the disk into the given place in the cache.
  SecFile$(into%) = SecFil$
  tDir$ = GalDir$ + "\" + SecFil$ + "\"
  F$ = tDir$ + SecFil$ + ".dat"
  fd% = FreeFile
  Open F$ For Input As fd%
  Line Input #fd%, SecName$(into%)
  Line Input #fd%, temp$
  plot$ = String$(160, Chr$(0))
  For I = 1 To 16
    Line Input #fd%, temp$
    SubName$ = RTrim$(Mid$(temp$, 30, 12))
    ReadSubsector GalDir$ + "\" + SecFil$ + "\MAP\" + SubName$, plot$
  Next I
  SecPlot$(into%) = plot$
  Close fd%
  ReadSector% = into%
  Exit Function
Test:
    Resume Next
End Function

Sub ReadSubsector(fsub$, plot$)
  Rem Read an individual subsector and set the values in the plot
  Rem accordingly.
  On Error GoTo Test
  fd% = FreeFile
  Open fsub$ For Input As fd%
  Do Until EOF(fd%)
    Line Input #fd%, T$
    T$ = LTrim$(RTrim$(T$))
    If T$ <> "" Then
     If (InStr("@#$", Left$(T$, 1)) = 0) Then
      C$ = Mid$(T$, 15, 4)
      c1% = Val(Left$(C$, 2))
      c2% = Val(Right$(C$, 2))
      If (c1% < 1) Or (c1% > 32) Or (c2% < 1) Or (c2% > 40) Then
        CurrentForm.Print "Erronious line in " + fsub$
        CurrentForm.Print T$
      End If
      SetPlot plot$, c1%, c2%
     End If
    End If
  Loop
  Close fd%
  Exit Sub
Test:
    Resume Next
End Sub

Sub savescr(filename$, sx, sy, ex, ey, nbits, imgnum)
''SAVESCR V0.6 - Screen Capture Function for Qbasic.
''By: Aaron Zabudsky <zabudsk@ecf.utoronto.ca>
''Date: July 17, 1997
''Free - Comments welcome.
''
''Usage: filename$ - Name of the file you want to capture to. Overwrites any
''                   old image that may be under that name.
''       sx        - Starting X coordinate
''       sy        - Starting Y coordinate
''       ex        - Ending X coordinate
''       ey        - Ending Y coordinate
''       nbits     - Number of bits you want in your bitmap. Use 1, 4 or 8.
''                   Use nbits=1 for SCREEN 11
''                   Use nbits=4 for SCREEN 12
''                   Use nbits=8 for SCREEN 13
''       imgnum    - The current number of the image you are saving to.
''                   This can be anything if you have specified a filename
''                   If you have specified a blank filename (""), Autonumbering
''                   is enabled and if you specify a number here, it will save
''                   the image as CAP0.BMP, CAP1.BMP,...,CAP1000.BMP,etc.
''                   If you leave a variable in this spot when you call the
''                   capture function, the function will automatically increment
''                   the variable, so you can "auto-capture" a series of
''                   pictures without worrying about numbers.
''
'' e.g. savescr "test.bmp",0,0,639,479,4,0
''      will capture the entire SCREEN 12 screen with 16 colours and save it
''      to test.bmp.
''      savescr "",0,0,319,199,8,t
''      will capture the entire SCREEN 13 screen with 256 colours and save it
''      as CAP#.BMP, where # is the current value of t, it will then increment
''      t.
''      savescr "",0,0,639,479,1,(t)
''      will capture the entire SCREEN 11 screen with 2 colours and save it
''      as CAP#.BMP as in the previous example, but this time t will not be
''      incremented.
'
'
'If filename$ = "" Then
'   filename$ = "CAP" + LTrim$(RTrim$(Str$(imgnum))) + ".BMP"
'   imgnum = imgnum + 1
'End If
'Open filename$ For Binary As #1
'If LOF(1) <> 0 Then
'   'Alter this code here if you don't want it to overwrite existing files.
'   Close 1
'   Kill filename$
'   Open filename$ For Binary As #1
'End If
'
'va = &H3C7 'VGA Palette Read Address Register
'vd = &H3C9 'VGA Palette Data Register
'
'zero$ = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
'
''Check extents to order points.
'If sx > ex Then Swap sx, ex
'If sy > ey Then Swap sy, ey
'
''Use Windows BMP Header. Size=40
'headersize = 40
'
''Calculate Picture width,height
'picwidth = ex - sx + 1
'picheight = ey - sy + 1
'
''Set Colour Information
''Planes [W] - Must be 1
'nplanes = 1
'
''Calculate offset [LW] to start of data
'If nbits = 1 Or nbits = 4 Or nbits = 8 Then
'   offset = 14 + headersize + 4 * (2 ^ nbits)
'Else
'   offset = 14 + headersize
'End If
'
''Type of file [W] (Should be BM)
'ft$ = "BM"
'
''File Size [LW] (excluding header)
'If nbits = 1 Then
'   If (picwidth Mod 32) <> 0 Then
'      filesize = 4 * (Int(picwidth / 32) + 1) * picheight
'   Else
'      filesize = (picwidth / 8) * picheight
'   End If
'ElseIf nbits = 4 Then
'   If (picwidth Mod 8) <> 0 Then
'      filesize = 4 * (Int(picwidth / 8) + 1) * picheight
'   Else
'      filesize = (picwidth / 2) * picheight
'   End If
'ElseIf nbits = 8 Then
'   If (picwidth Mod 4) <> 0 Then
'      filesize = 4 * (Int(picwidth / 4) + 1) * picheight
'   Else
'      filesize = picwidth * picheight
'   End If
'ElseIf nbits = 24 Then
'   If (3 * picwidth Mod 4) <> 0 Then
'      filesize = 4 * (Int(3 * picwidth / 4) + 1) * picheight
'   Else
'      filesize = 3 * picwidth * picheight
'   End If
'End If
'
''Set reserved values [W] (both must be zero)
'r1 = 0
'r2 = 0
'
''Compression type [LW] - None
'comptype = 0
'
''Image Size [LW]; Scaling Factors xsize, ysize unused.
'imagesize = offset + filesize
'xsize = 0
'ysize = 0
'
''Assume all colours used [LW] - 0 means all colours.
'coloursused = 0
'neededcolours = 0
'
'header$ = ft$ + MKL$(filesize) + MKI$(r1) + MKI$(r2) + MKL$(offset)
'infoheader$ = MKL$(headersize) + MKL$(picwidth)
'infoheader$ = infoheader$ + MKL$(picheight) + MKI$(nplanes)
'infoheader$ = infoheader$ + MKI$(nbits) + MKL$(comptype) + MKL$(imagesize)
'infoheader$ = infoheader$ + MKL$(xsize) + MKL$(ysize) + MKL$(coloursused)
'infoheader$ = infoheader$ + MKL$(neededcolours)
'
''Write headers to BMP File.
'Put #1, 1, header$
'Put #1, , infoheader$
'
''Add palette - Get colours (Write as B0G0R0(0),B1G1R1(0),...)
'If nbits = 1 Or nbits = 4 Or nbits = 8 Then
'   palet$ = ""
'   OUT va, 0
'   For Count = 1 To 2 ^ nbits
'      zr = INP(vd) * 4
'      zg = INP(vd) * 4
'      zb = INP(vd) * 4
'      palet$ = palet$ + Chr$(zb) + Chr$(zg) + Chr$(zr) + Chr$(0)
'   Next Count
'   Put #1, , palet$
'   palet$ = "" 'Save some memory
'End If
'
'
'stpoint = Point(sx, ey + 1)
'
''BMPs are arranged with the top of the image at the bottom of the file.
''Get points off the screen and pack into bytes depending on the number of
''bits used. Deal with unused bits at the end of the line.
''Check for invalid range.
'For count2 = ey To sy Step -1
'   Lin$ = ""
'   If nbits = 1 Then
'      count1 = sx
'      While count1 <= ex
'         If count1 + 7 > ex Then
'            T = 0
'            For count0 = 0 To 7
'               p = Point(count1 + count0, count2)
'               If p < 0 Then p = 0
'               T = T + (2 ^ (7 - count0)) * (p Mod 2)
'            Next count0
'            t2 = ex - count1 + 1
'            T = T And ((2 ^ t2) - 1) * (2 ^ (8 - t2))
'            Lin$ = Lin$ + Chr$(T)
'         Else
'            T = 0
'            For count0 = 0 To 7
'               p = Point(count1 + count0, count2)
'               If p < 0 Then p = 0
'               T = T + (2 ^ (7 - count0)) * (p Mod 2)
'            Next count0
'            Lin$ = Lin$ + Chr$(T)
'         End If
'         count1 = count1 + 8
'      Wend
'   ElseIf nbits = 4 Then
'      count1 = sx
'      While count1 <= ex
'         If count1 = ex Then
'            p = Point(count1, count2)
'            If p < 0 Then p = 0
'            Lin$ = Lin$ + Chr$((p Mod 16) * 16)
'         Else
'            p = Point(count1, count2)
'            p2 = Point(count1 + 1, count2)
'            If p < 0 Then p = 0
'            If p2 < 0 Then p2 = 0
'            Lin$ = Lin$ + Chr$((p Mod 16) * 16 + p2)
'         End If
'         count1 = count1 + 2
'      Wend
'   ElseIf nbits = 8 Then
'      For count1 = sx To ex
'         p = Point(count1, count2)
'         If p < 0 Then p = 0
'         Lin$ = Lin$ + Chr$(p)
'      Next count1
'   ElseIf nbits = 24 Then
'      'I'm not sure what to put here. QBasic doesn't support truecolour
'      'Unused for now.
'   End If
'
'   'Pad line to LongWord boundary
'   If (Len(Lin$) Mod 4) <> 0 Then
'      Lin$ = Lin$ + Mid$(zero$, 1, 4 - (Len(Lin$) Mod 4))
'   End If
'
'   'Indicate our status
'   PSet (sx, count2 + 1), stpoint
'   stpoint = Point(sx, count2)
'   If nbits = 8 Then
'      PSet (sx, count2), 255 - stpoint
'   ElseIf nbits = 4 Then
'      PSet (sx, count2), 15 - stpoint
'   ElseIf nbits = 1 Then
'      PSet (sx, count2), 1 - stpoint
'   End If
'
'   'Write the current line to the BMP file
'   Put #1, , Lin$
'
'Next count2
'
''Save some memory
'Lin$ = ""
'
'PSet (sx, count2 + 1), stpoint
'
''Close the file
'Close

End Sub

Sub SetPlot(plot$, x%, y%)
  Rem The Plot$ is used to maintain a list of chars that is, in fact, a
  Rem bitfield containing the on-off state for all hexes in a sector.
  Rem There wasn't enough string space to do it on a byte level so we
  Rem had to do it this way.
  Rem Here, we set the bit for hex x,y
  Dim TempS As String
  offset% = (x% - 1) + (y% - 1) * 32
  byteoffset% = Int(offset% / 8) + 1
  bitoffset% = (offset% Mod 8)
  TempS = Mid$(plot$, byteoffset%, 1)
  mask% = pow2(bitoffset%)
  bval% = Asc(TempS)
  bval% = bval% + mask%
  plot$ = Left$(plot$, byteoffset% - 1) + Chr$(bval%) + Mid$(plot$, byteoffset% + 1)
End Sub

Public Sub GMap2(Cmd As String)

End Sub

