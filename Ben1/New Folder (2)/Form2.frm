VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   256
   ScaleMode       =   0  'User
   ScaleWidth      =   320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Command1"
      Height          =   555
      Left            =   3540
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, _
ByVal x As Long, _
ByVal y As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, _
ByVal ySrc As Long, _
ByVal dwRop As Long) As Long
'API calls
'
'Blitting
'
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long
'
'code timer
'
Private Declare Function GetTickCount Lib "kernel32" () As Long
'
'creating buffers/loading sprites
'
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'
'loading sprites
'
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hobject As Long) As Long
'
'cleanup
'
Private Declare Function DeleteObject Lib "gdi32" (ByVal hobject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'
'our buffer's Device Context (DC "address")
'
Public myBackBuffer As Long
Public myBufferBmp As Long
'
'the DC of the sprite
'
Public mySprite As Long
'
'coordinates of sprite
'
Public SpriteX As Long
Public SpriteY As Long
Public Function LoadGraphicDC(sFileName As String) As Long
'error handling
On Error Resume Next
'temporary varible to hold address of graphic
Dim LoadGraphicDCTEMP As Long
'create the DC address compatible with screen's DC
LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))
'load graphic into the DC compatible w/ screen...
SelectObjectLoadGraphicDCTEMP , LoadPicture(sFileName)
' return address of file
LoadGraphicDC = LoadGraphicDCTEMP
End Function



End Function


Private Sub cmdTest_Click()
Dim T1 As Long, T2 As Long

'create a compatable DC for back buffer
myBackBuffer = CreateCompatibleDC(GetDC(0))
'create compatible bitmap surface for DC that's the size of the form (320 * 256)
Mybuffer.Bmp = CreateCompatibleBitmap(GetDC(0), 320, 256)

'final makings of backbuffer
'load the blank bitmap onto backbuffer

SelectObject myBackBuffer, myBufferBmp
'fill back buffer with black
BitBlt myBackBuffer, 0, 0, 320, 256, 0, 0, 0, vbWhiteness
'load sprite
mySprite = LoadGraphicDC(App.Path & "\sprite1.bmp")
cmdTest.Enabled = False
'===<<<MAIN LOOP STARTS HERE>>>===
T2 = GetTickCount
Do
DoEvents 'so mouse doesn't freeze up
T1 = GetTickCount
If (T1 - T2) >= 15 Then
BitBlt myBackBuffer, SpriteX - 1, SpriteY - 1, 32, 32, 0, 0, 0, vbBlackness
BitBlt myBackBuffer, SpriteX, SpriteY, 32, 32, mySprite, 0, 0, vbSrcPaint
SpriteX = SpriteX + 1
SpriteY = SpriteY + 1
T2 = GetTickCount
End If
Loop Until SpriteX = 320

End Sub

Private Sub Form_Unload(Cancel As Integer)
'clear up memory
DeleteObject Mybuffer.Bmp
DeleteDC myBackBuffer
DeleceDC mySprite
End
End Sub
