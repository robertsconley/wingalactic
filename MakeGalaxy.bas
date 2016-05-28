Attribute VB_Name = "modMakeGalaxy"
Rem Galaxy Maker 1.0
Rem Public Domain 1998 Jim Vassilakos
Public Sub MakeGalaxy()
    Dim Gal(15) As String
    Dim NumGals As Long
    Dim YN As String
    Dim GalName As String
    Dim T As String
    Dim L As Long
    Dim Ok As Integer
    Dim A As Long
    
    Open "gals\gal.lst" For Input As #1
    Line Input #1, T$
    Line Input #1, T$
    Do Until EOF(1)
        Line Input #1, T$
        NumGals = NumGals + 1
        Gal$(NumGals) = LCase$(RTrim$(Left$(T$, 8)))
    Loop
    Close 1

    CLS
    Color 11
    CurrentForm.Print "The program will automatically create a directory for"
    CurrentForm.Print "your new galaxy to be placed into (which you must name)."
    CurrentForm.Print
    Color 10
    CurrentForm.Print "Are you sure you really want to create a new galaxy?";
'    Do
'        Call GetKeyChar(Key1, Key2)
'    '950     Rem y/n
'        YN = ""
'        If Key2 = 0 And (Key1 >= 97 And Key1 <= 122) Then Key1 = Key1 - 32
'        If Key2 = 0 And Key1 = 78 Then YN$ = "n": Rem N
'        If Key2 = 0 And Key1 = 89 Then YN$ = "y": Rem Y
'        If YN = "n" Then Exit Sub
'        If YN = "y" Then Exit Do
'    Loop
    YN = InputYN()
    Color 12
    CurrentForm.Print "   Yes"
    CurrentForm.Print
    Color 11
    CurrentForm.Print "Enter the Name of the Galaxy (40 characters max):"
    CurrentForm.Print "   (For example: The Milky Way)";
    Do
        Locate 9, 1
        CurrentForm.Print Space$(70);
        Locate 9, 1
        T$ = InputText("--->")
        If T$ <> "" Then
            T$ = LTrim$(RTrim$(T$))
            If Len(T$) <= 40 Then Exit Do
        End If
    Loop
    CurrentForm.Print
    GalName = T$
    CurrentForm.Print "You galaxy will be placed under the 'gals' directory."
    CurrentForm.Print "Please provide a directory name (8 characters or less):"
    CurrentForm.Print
    Do
        Locate 14, 1
        CurrentForm.Print Space$(70)
        Locate 14, 1
        T$ = InputText("--->")
        T$ = LCase$(LTrim$(RTrim$(T$)))
        If T$ <> "" Then
            L = Len(T$)
            If Right$(T$, 1) = "\" Then
                L = L - 1
                T$ = Left$(T$, L)
            End If
            Ok = 1
            For A = 1 To NumGals
                If T$ = Gal$(A) Then Ok = 2
            Next A
            If InStr(T$, ".") > 0 Then Ok = 4
            If L > 8 Then Ok = 3
            If Ok = 2 Then CurrentForm.Print "   <directory already exists, try again>"
            If Ok = 3 Then CurrentForm.Print "   <8 characters maximum, try again>    "
            If Ok = 4 Then CurrentForm.Print "   <no extension required, try again>   "
            If Ok = 1 Then Exit Do
        End If
    Loop
    Gal0$ = T$
    Path$ = "gals\" + T$

Rem update gal.lst
    L = Len(Gal0$)
    T$ = Gal0$ + Space$(13 - L) + GalName$
    Open "gals\gal.lst" For Append As #1
    Print #1, T$
    Close #1

Rem create galaxy's directories
    FSO.CreateFolder Path$
    FSO.CreateFolder Path$ & "\gen"
Rem create galaxy's list file
    Open Path$ + "\" + Gal0$ + ".lst" For Output As #1
    Print #1, "Directory of Sectors"
    Print #1, "-------------------------------------------------------------"
    Close 1

Rem create general info menu
    Open Path$ + "\gen\galaxy.mnu" For Output As #1
    Print #1, "@dir="; Path$; "\gen"
    Print #1, "@Light Yellow"
    Print #1, ""
    Print #1, GalName$; " / General Information"
    Print #1, ""
    Print #1, ""
    Print #1, "@Light Cyan"
    Print #1, "      Topics & Methods       @ideas.txt"
    Close 1

Rem create ideas.txt, uwp.dat, where.dat
    FSO.CopyFile "help\ideas.txt", Path$ + "\gen\ideas.txt"
    FSO.CopyFile "data\uwp.dat", Path$ + "\uwp.dat"
    Open Path$ + "\where.dat" For Output As #1
    Print #1, "0"
    Print #1, "0"
    Close 1

    Color 7
    CLS
    CurrentForm.Print "Later..."
End Sub
