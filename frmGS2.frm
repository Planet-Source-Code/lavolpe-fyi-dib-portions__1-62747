VERSION 5.00
Begin VB.Form frmGS2 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   3795
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "-1"
      ToolTipText     =   "-1 defaults to full height"
      Top             =   555
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "-1"
      ToolTipText     =   "-1 defaults to full width"
      Top             =   555
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   3795
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "Where to start grayscale (Top)"
      Top             =   195
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Where to start grayscale (Left)"
      Top             =   195
      Width           =   450
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   1
      Left            =   105
      ScaleHeight     =   1440
      ScaleWidth      =   2610
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1665
      Width           =   2610
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< GrayScale DC"
      Height          =   495
      Left            =   2985
      TabIndex        =   2
      Top             =   990
      Width           =   1320
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   0
      Left            =   1470
      ScaleHeight     =   1440
      ScaleWidth      =   1245
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   105
      Picture         =   "frmGS2.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   1245
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Also, you can easily grayscale an individual image. Bitmap && Icon examples below:"
      Height          =   810
      Left            =   2820
      TabIndex        =   12
      Top             =   1695
      Width           =   1830
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3885
      Picture         =   "frmGS2.frx":1131
      Top             =   2595
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   3135
      Picture         =   "frmGS2.frx":1573
      Top             =   2475
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "B"
      Height          =   270
      Index           =   3
      Left            =   3645
      TabIndex        =   11
      Top             =   555
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "R"
      Height          =   270
      Index           =   2
      Left            =   2955
      TabIndex        =   10
      Top             =   555
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "T"
      Height          =   270
      Index           =   1
      Left            =   3645
      TabIndex        =   9
      Top             =   195
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "L"
      Height          =   270
      Index           =   0
      Left            =   2955
      TabIndex        =   8
      Top             =   195
      Width           =   180
   End
End
Attribute VB_Name = "frmGS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 3 Oct. Modified
' GrayScaleDC function
' ^^ By using zero as the optional parameter for the Width & Height of the DC to
'    grayscale, there wasn't a way to just grayscale the 1 pixel top
'    edge (0,0)-(width,0) or the 1 pixel far left edge (0,0)-(0,height)
'    Changed the optional parameter value default to -1 vs 0
' ^^ Tweaked grayscale average to favor a slightly lighter shade over a darker shade
' Custom functions always returned False, now they return True if successful

' First off. These routines are extremely fast, cause there is no real overhead
' by creating offscreen bitmaps and DCs separately, manipulation of bytes, and
' the efficiency of the DIB section APIs.

' This sample project's only true purpose is to show how you can extract only
' a section of a DC/image into DIB bytes vs extracting the entire DC/image.

' I haven't seen examples of doing sections of DIBs but am sure there are some
' out there. Why this project? I wanted a way to grayscale my CustomWindows v2
' project without having to extract the entire window DC into a byte array in
' order to grayscale say a 16x16 area.  If the window is maximized (1024x768),
' each pixel being a 4 byte (Long) color value, then I am looking at navigating
' thru 3,145,728 bytes vs just 65,536 bytes: a huge waste of 3,080,192 bytes
' Basically the formula for needed bytes is DC.Width*nrRows*4 bytes where
' 1024*16*4 = 65,536 and 1024*768*4 = 3,145,728

' As an added bonus, I provided another routine where you can grayscale an
' image/picture (stdPicture, handle, etc), into an existing DC.

' Regarding the grayscale method: As you graphics gurus play with this, please
' email me with a more softer/better averaging method. The only other algorithm
' I am familiar with is something like: Avg( r*0.5, g*0.33, b*0.11)
' However, that routine is poor when an image is mostly Blue because
' blue is 0,0,255 and that formula translates grayscale blue to RGB(28,28,28)
' which is very dark gray, nearly black. By contrast, the simple averaging
' technique is Avg(0,0,255) which would translate to RGB(85,85,85); big diff!

' A pleasing alternative I've played with is averaging the Lightest color with
' the average of the darkest 2 colors. I feel that it is a nicer grayscale, but
' am unhappy with having to deal with the additional IFs and divisions to find &
' calculate those averages for each & every pixel. I've also played with allowing
' a shade offset where the gray color could be user-changed -/+ 85.
' This is doable & may be the eventual way I go.


' used for GrayScaleDC routine
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Const OBJ_BITMAP As Long = 7

' used for GrayScaleImage routine
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type


' used for both routines
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

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
    bmiColors(0 To 3) As Byte
End Type

Private Sub Command1_Click()

Picture2(0).Cls
Picture2(1).Cls

Picture2(0).PaintPicture Picture1.Image, 0, 0
GrayScaleDC Picture2(0).hdc, Val(Text1(0)), Val(Text1(1)), Val(Text1(2)), Val(Text1(3))

GrayScaleImage Image1.Picture.Handle, Picture2(1).hdc, 10, 10
GrayScaleImage Image2.Picture.Handle, Picture2(1).hdc, _
        (Picture2(1).Width - Image2.Width) \ Screen.TwipsPerPixelX - 10, _
        (Picture2(1).Height - Image2.Height) \ Screen.TwipsPerPixelY - 10

Picture2(0).Refresh
Picture2(1).Refresh

End Sub

Private Function GrayScaleDC(hDestDC As Long, Optional ByVal X As Long, Optional ByVal Y As Long, _
                        Optional ByVal X1 As Long = -1, Optional ByVal Y1 As Long = -1) As Boolean

    ' Parameters
    ' hDestDC is an existing DC to do the painting
    ' X, Y, X1, & Y1 are rectangle coordinates for grayscaling
    ' ^^ if X1=-1 then the full DC width will be used
    ' ^^ if Y1=-1 then the full DC height will be used

    If hDestDC = 0 Then Exit Function
    
    Dim bmp As BITMAPINFOHEADER
    Dim bmi As BITMAPINFO
    Dim dibArray() As Byte
    Dim Xb As Long, Yb As Long
    Dim vLookup(0 To 255) As Byte
    Dim hBmp As Long
    Dim hHandle As Long
    
    ' get the handle to the bitmap in the DC.
    ' For picture boxes, this is equivalent to Picture1.Image.Handle
    hHandle = GetCurrentObject(hDestDC, OBJ_BITMAP)
    If hHandle = 0 Then Exit Function
    
    ' if next line fails, we can't set up our DIB arrays
    If GetGDIObject(hHandle, Len(bmp), bmp) = 0 Then Exit Function
    ' ensure a non-empty bitmap
    If bmp.biHeight < 1 Or bmp.biWidth < 1 Then Exit Function
    
    With bmi.bmiHeader
        ' setup the UDT
        .biSize = Len(bmi.bmiHeader)
        .biBitCount = 32            ' automatically aligns bytes on DWord boundaries
        .biCompression = 0
        .biHeight = bmp.biHeight
        .biWidth = bmp.biWidth
        .biPlanes = 1
        
        ' because we will be setting up an array, we need to make absolutely
        ' sure, we will not go out of bounds. If we do, the APIs will crash VB
        If X1 > .biWidth - 1 Or X1 = -1 Then X1 = .biWidth - 1
        If Y1 > .biHeight - 1 Or Y1 = -1 Then Y1 = .biHeight - 1
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        If X1 < X Or Y1 < Y Then Exit Function
        ' cache any adjusted height (Y<>0 and/or Y1<>.biHeight)
        bmp.biHeight = Y1 - Y + 1
        
        ' size byte array for the section to be grayscaled
        ReDim dibArray(0 To (.biWidth * 4 * bmp.biHeight) - 1)
        ' get only the section to be grayscaled.
        ' Note funky .biHeight-Y1-1 as the starting Y coordinate. Image is flipped.
        GetDIBits hDestDC, hHandle, .biHeight - Y1 - 1, bmp.biHeight, dibArray(0), bmi, 0&
        ' note that the dibArray array is BRG vs RGB
        
        For Xb = 1 To 255
            ' cache look up vs recalculating for each pixel in the source image
            ' adding 2 to offset integer division: 10,10,190 should =70 but would=69
            ' below, so we will soften it up a bit: 10,10,190 would now be 72 vs 69
            vLookup(Xb) = (((Xb + 2) * 33) \ 100)
            ' note: do not increment 33. Any above total*3>100 may cause Overflow
        Next
        
        ' loop thru the bytes, rows then columns
        For Yb = 0 To UBound(dibArray) Step .biWidth * 4
            ' since we needed to pull bytes starting at column 0, we need
            ' to adjust the startpoint for each row if user passed a
            ' parameter other than X = 0
            For Xb = Yb + (4 * X) To Yb + (4 * X1) Step 4
                ' add up the gray scale bytes and apply to all 3 source bytes
                dibArray(Xb) = vLookup(dibArray(Xb)) + vLookup(dibArray(Xb + 1)) + vLookup(dibArray(Xb + 2))
                dibArray(Xb + 1) = dibArray(Xb)
                dibArray(Xb + 2) = dibArray(Xb)
                'dibArray(xb+3) not used; the 4th byte in a 32 byte color
            Next
        Next
        
        ' adjust the height in the bmi UDT
        .biHeight = bmp.biHeight
        ' simply paste the changes back to the destination dc
        GrayScaleDC = (StretchDIBits(hDestDC, 0, Y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, dibArray(0), bmi, 0, vbSrcCopy) <> 0)
    
    End With
    
        
End Function


Private Function GrayScaleImage(hImage As Long, hDestDC As Long, X As Long, Y As Long, _
                            Optional ByVal Cx As Long, Optional ByVal Cy As Long) As Boolean
                            
    ' Parameters
    ' hImage is the handle to the memory image or [object].Picture.Handle
    ' hDestDC is an existing DC to do the painting
    ' X, Y are left & top coordinates to draw the image
    ' Cx and Cy are the width/height of the drawn image


If hImage = 0 Or hDestDC = 0 Then Exit Function
                            
Dim bmp As BITMAPINFOHEADER
Dim bmi As BITMAPINFO
Dim dibArray() As Byte, Xb As Long, Yb As Long
Dim dibMask() As Byte   ' used for icons only
Dim vLookup(0 To 255) As Byte
Dim iInfo As ICONINFO
Dim hBmp As Long
Dim stepValue As Long

If GetGDIObject(hImage, Len(bmp), bmp) = 0 Then

    ' handle passed is not a bitmap. If not an icon, then abort
    If GetIconInfo(hImage, iInfo) = 0 Then Exit Function
    
    ' it's an icon, but is it valid?
    If iInfo.hbmColor = 0 Then
        ' a black & white icon/cursor. nothing to do as it is already grayscaled
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
        DrawIconEx hDestDC, X, Y, hImage, Cx, Cy, 0, 0, &H3
        Exit Function
    End If
    ' ok, got a color bitmap
    hBmp = iInfo.hbmColor
    ' get bitmap information for the icon
    If GetGDIObject(hBmp, Len(bmp), bmp) = 0 Then
        ' If we have iInfo.hbmColor, then this test should never be triggered
        ' However, for robutsness, simply paint the icon & delete execess memory items
        DeleteObject hBmp
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
        DrawIconEx hDestDC, X, Y, hImage, Cx, Cy, 0, 0, &H3
        Exit Function
    End If

Else
    ' bitmap. Transparent GIFs won't work with this routine.
    hBmp = hImage
End If

' ensure we have a non-empty bitmap
If bmp.biHeight < 1 Or bmp.biWidth < 1 Then Exit Function

With bmi.bmiHeader
    ' setup the UDT
    .biSize = Len(bmi.bmiHeader)
    .biBitCount = 32
    .biCompression = 0
    .biHeight = bmp.biHeight
    .biWidth = bmp.biWidth
    .biPlanes = 1
    
    stepValue = .biWidth * 4
    ReDim dibArray(0 To (stepValue * .biHeight - 1))
    GetDIBits hDestDC, hBmp, 0, .biHeight, dibArray(0), bmi, 0&
    
    For Xb = 1 To 255    ' array in BRG vs RGB
        ' cache look up vs recalculating for each pixel in the source image
        ' adding 2 to offset integer division: 10,10,190 should =70 but would=69
        ' below, so we will soften it up a bit: 10,10,190 would now be 72 vs 69
        vLookup(Xb) = (((Xb + 2) * 33) \ 100)
        ' note: do not increment 33. Any above total*3>100 may cause Overflow
    Next
    
    For Yb = 0 To UBound(dibArray) Step stepValue
        For Xb = Yb To Yb + stepValue - 1 Step 4
            ' add up the gray scale bytes and apply to all 3 source bytes
            dibArray(Xb) = vLookup(dibArray(Xb)) + vLookup(dibArray(Xb + 1)) + vLookup(dibArray(Xb + 2))
            dibArray(Xb + 1) = dibArray(Xb)
            dibArray(Xb + 2) = dibArray(Xb)
        Next
    Next
    
    ' use the image width/height if user passed zeros
    If Cx < 1 Then Cx = .biWidth
    If Cy < 1 Then Cy = .biHeight

    If iInfo.hbmColor Then  ' icon was passed
        dibMask = dibArray
        GetDIBits hDestDC, iInfo.hbmMask, 0, .biHeight, dibMask(0), bmi, 0&
        ' the bitmaps returned by a call to GetIconInfo must be deleted else memory leaks
        DeleteObject hBmp
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
        StretchDIBits hDestDC, X, Y, Cx, Cy, 0, 0, .biWidth, .biHeight, dibMask(0), bmi, 0, vbSrcAnd
        GrayScaleImage = (StretchDIBits(hDestDC, X, Y, Cx, Cy, 0, 0, .biWidth, .biHeight, dibArray(0), bmi, 0, vbSrcPaint) <> 0)
    Else
        GrayScaleImage = (StretchDIBits(hDestDC, X, Y, Cx, Cy, 0, 0, .biWidth, .biHeight, dibArray(0), bmi, 0, vbSrcCopy) <> 0)
    End If
    
End With


End Function
