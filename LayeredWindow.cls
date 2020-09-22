VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayeredWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ULW_OPAQUE As Long = &H4
Private Const ULW_COLORKEY As Long = &H1
Private Const ULW_ALPHA As Long = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER As Byte = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'API Calls for GDI+
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Private mDC As Long  ' Memory hDC
Private MainBitmap As Long ' Memory Bitmap
Private blendFunc32bpp As BLENDFUNCTION
Private Token As Long ' Needed to close GDI+
Private OldBitmap As Long

'Use this when the form unloads
'If you don't use this, VB can CRASH(!!!)
Public Sub UnloadPNGForm()

    Call GdiplusShutdown(Token)

    SelectObject mDC, OldBitmap
    DeleteObject MainBitmap
    DeleteObject OldBitmap
    DeleteDC mDC

End Sub

'Initialize GDI+
Private Sub Class_Initialize()

  ' Start up GDI+
  
  Dim GpInput As GdiplusStartupInput

    GpInput.GdiplusVersion = 1
    If GdiplusStartup(Token, GpInput) <> 0 Then
        MsgBox "Error loading GDI+!", vbCritical, "GDI Plus Error"
    End If

End Sub

'Use a PNG file on a form
Public Function MakeTrans(pngPath As String, Frm As Object) As Boolean

  Dim tempBI As BITMAPINFO
  Dim tempBlend As BLENDFUNCTION      ' Used to specify what kind of blend we want to perform
  Dim lngHeight As Long, lngWidth As Long
  Dim curWinLong As Long
  Dim img As Long
  Dim graphics As Long
  Dim winSize As Size
  Dim srcPoint As POINTAPI
   
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32    ' Each pixel is 32 bit's wide
        .biHeight = Frm.ScaleHeight  ' Height of the form
        .biWidth = Frm.ScaleWidth    ' Width of the form
        .biPlanes = 1   ' Always set to 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
    End With 'TEMPBI.BMIHEADER

    mDC = CreateCompatibleDC(Frm.hdc)
    MainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    OldBitmap = SelectObject(mDC, MainBitmap)   ' Select the new bitmap, track the old that was selected
    
    ' GDI Initializations
    Call GdipCreateFromHDC(mDC, graphics)
    Call GdipLoadImageFromFile(StrConv(pngPath, vbUnicode), img)  ' Load Png
    Call GdipGetImageHeight(img, lngHeight)
    Call GdipGetImageWidth(img, lngWidth)
    Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)

    ' Change windows extended style to be used by updatelayeredwindow
    curWinLong = GetWindowLong(Frm.hWnd, GWL_EXSTYLE)
    ' Accidently did This line below which flipped entire form, it's neat so I left it in
    ' Comment out the line above and uncomment line below.
    'curWinLong = GetWindowLong(Frm.hwnd, GWL_STYLE)
    SetWindowLong Frm.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
   
    ' Make the window a top-most window so we can always see the cool stuff
    SetWindowPos Frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
   
    ' Needed for updateLayeredWindow call
    srcPoint.x = 0
    srcPoint.Y = 0
    winSize.cx = Frm.ScaleWidth
    winSize.cy = Frm.ScaleHeight
    
    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA ' 32 bit
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With 'BLENDFUNC32BPP
    
    Call GdipDisposeImage(img)
    Call GdipDeleteGraphics(graphics)
    Call UpdateLayeredWindow(Frm.hWnd, Frm.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)

End Function



