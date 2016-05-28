Attribute VB_Name = "Utility"
Option Explicit
Public FSO As FileSystemObject
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Sub OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As IID, ByVal fPictureOwnsHandle As Long, ipic As IPicture)
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
    cbSizeofstruct As Long
    picType As Long
    hgdiobj As Long
    hPalOrXYExt As Long
End Type

Public Function ToMono(pbSrc As Form) As StdPicture
    Dim hdcMono As Long
    Dim hbmpMono As Long
    Dim hbmpOld As Long
    Dim P As StdPicture
    Dim P1 As IPicture
    Dim xBlt As Long
    Dim yBlt As Long
    Dim oWidth As Long
    Dim oHeight As Long
    Set P = New StdPicture
    ' Create memory device context
    hdcMono = CreateCompatibleDC(CurrentForm.hDC)
    ' Create monochrome bitmap and select it into DC
    'pbSrc.BorderStyle = 0
    xBlt = 5 'pbSrc.ScaleWidth
    yBlt = 5 'pbSrc.ScaleHeight
    hbmpMono = CreateCompatibleBitmap(CurrentForm.hDC, xBlt, yBlt)
    hbmpOld = SelectObject(hdcMono, hbmpMono)
     ' Copy color bitmap6 to DC to create mono mask
    BitBlt hdcMono, 0, 0, xBlt, yBlt, CurrentForm.hDC, 0, 0, SRCCOPY
    Set P = BitmapToPicture(hbmpMono)
    Set ToMono = P
    ' Copy mono memory mask to visible picture box
    Call SelectObject(hdcMono, hbmpOld)
    Call DeleteDC(hdcMono)
End Function

Public Function BitmapToPicture(ByVal hBmp As Long, Optional ByVal hPal As Long = 0) As IPicture
    ' Fill picture description
    Dim ipic As IPicture, picdes As PICTDESC, iidIPicture As IID
    picdes.cbSizeofstruct = Len(picdes)
    picdes.picType = vbPicTypeBitmap
    picdes.hgdiobj = hBmp
    picdes.hPalOrXYExt = hPal
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    iidIPicture.Data1 = &H7BF80980
    iidIPicture.Data2 = &HBF32
    iidIPicture.Data3 = &H101A
    iidIPicture.Data4(0) = &H8B
    iidIPicture.Data4(1) = &HBB
    iidIPicture.Data4(2) = &H0
    iidIPicture.Data4(3) = &HAA
    iidIPicture.Data4(4) = &H0
    iidIPicture.Data4(5) = &H30
    iidIPicture.Data4(6) = &HC
    iidIPicture.Data4(7) = &HAB
    ' Create picture from bitmap handle
    OleCreatePictureIndirect picdes, iidIPicture, True, ipic
    ' Result will be valid Picture or Nothing-either way set it
    Set BitmapToPicture = ipic
End Function

