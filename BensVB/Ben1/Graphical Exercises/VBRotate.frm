VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VBRotate"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      LargeChange     =   5
      Left            =   30
      Max             =   255
      Min             =   1
      TabIndex        =   2
      Top             =   3180
      Value           =   1
      Width           =   3060
   End
   Begin VB.PictureBox Pictbox 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   30
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   60
      Width           =   3060
   End
   Begin VB.PictureBox Pictbuf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   45
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Bitmap rotation demo, by David Brebner Unlimited Realities 98"
      Height          =   405
      Left            =   60
      TabIndex        =   3
      Top             =   3525
      Width           =   3030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Sub zeromemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numbytes As Long)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'the current angle (0 to 255)
'where 255 = 360 degrees
Private angle As Byte

'lookup tables for the current
'angle to save using sin * cos at runtime
Private sin_ang(255) As Double
Private cos_ang(255) As Double

'these will hold the x and y offset
'mapping each pixel to its rotated
'destination
Private xstep_r(255) As Integer
Private xstep_c(255) As Integer
Private ystep_r(255) As Integer
Private ystep_c(255) As Integer




Sub DoImagePro()
' these are used to address the pixel using matrices
Dim pict() As Byte
Dim pict2() As Byte


Dim sa As SAFEARRAY2D, bmp As BITMAP
Dim sa2 As SAFEARRAY2D, bmp2 As BITMAP
Dim r As Integer, c As Integer
' get bitmap info
GetObjectAPI Pictbox.Picture, Len(bmp), bmp 'dest
GetObjectAPI Pictbuf.Picture, Len(bmp2), bmp2 'source
' exit if not a supported bitmap
If bmp.bmBitsPixel <> 8 Then
    MsgBox " 8-bit bitmaps only", vbCritical
    Exit Sub
End If
   
' have the local matrix point to bitmap pixels
With sa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp.bmWidthBytes
    .pvData = bmp.bmBits
End With
CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    
' have the local matrix point to bitmap pixels
With sa2
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp2.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp2.bmWidthBytes
    .pvData = bmp2.bmBits
End With
CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
zeromemory ByVal bmp.bmBits, bmp.bmWidthBytes * bmp.bmHeight
'do the core routine
For c = 0 To UBound(pict2, 1)
    For r = 0 To UBound(pict2, 2)
        'only copy solid information (1 is black, speed routine up)
        If pict2(c, r) > 1 Then
            'copy two pixels one on top of the other.
            'this is a nasty way to cover gaps in the mapping
            pict(xstep_c(c) + ystep_r(r), ystep_c(c) - xstep_r(r)) = pict2(c, r)
            pict(xstep_c(c) + ystep_r(r) + 1, ystep_c(c) - xstep_r(r)) = pict2(c, r)
        End If
    Next
Next


' clear the temporary array descriptor
' without destroying the local temporary array
CopyMemory ByVal VarPtrArray(pict), 0&, 4
CopyMemory ByVal VarPtrArray(pict2), 0&, 4


End Sub




Private Sub Form_Load()
'load the larger blank picture for display
Pictbox.Picture = LoadPicture(App.Path & "\rotate2.gif")
'load the image to rotate into the buffer
Pictbuf.Picture = LoadPicture(App.Path & "\rotate1.gif")

Dim a%
'pre-calculate sin & cos lookup tables
For a% = 0 To 255
    '2Pi is 360 degrees in radians
    sin_ang(a%) = Sin(a% / 255 * 6.283185)
    cos_ang(a%) = Cos(a% / 255 * 6.283185)
Next

dorotate
End Sub

Private Sub HScroll1_Change()
    dorotate
End Sub
Sub dorotate()
Dim a%
Dim xs As Double, ys As Double
Dim wid%, hgt%

    wid% = Pictbuf.ScaleWidth - 1
    hgt% = Pictbuf.ScaleHeight - 1
    
    'rotate the appropriate amount (note 255 = 360 degrees)
    angle = HScroll1.Value
    xs = sin_ang(angle) * wid%
    ys = cos_ang(angle) * hgt%

    'precalculate the x and y steps
    For a% = 0 To wid%
        xstep_c(a%) = (xs / 2) - (xs * (a% / wid%))
        ystep_c(a%) = (ys / 2) - (ys * (a% / wid%)) + (hgt% / 1.5)
    Next
    For a% = 0 To hgt%
        xstep_r(a%) = (xs / 2) - (xs * (a% / hgt%))
        ystep_r(a%) = (ys / 2) - (ys * (a% / hgt%)) + (hgt% / 1.5)
    Next
        
    DoImagePro
    Pictbox.Refresh
End Sub

Private Sub HScroll1_Scroll()
dorotate
End Sub
