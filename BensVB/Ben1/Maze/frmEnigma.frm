VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmEnigma.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   60
      Top             =   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   10860
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Memory Device Context
'

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
'****************************************
Dim RoomDC(1 To 5) As Long
Dim RotPos As Integer
Dim Link(1 To 5, 1 To 8) As Integer
Dim LinkX(1 To 5, 1 To 8) As Integer
Dim LinkY(1 To 5, 1 To 8) As Integer
Dim LinkHeight(1 To 5, 1 To 8) As Integer
Dim LinkWidth(1 To 5, 1 To 8) As Integer
Dim LinkEnabled(1 To 5, 1 To 8) As Boolean
Dim Music(1 To 5) As String





Private Sub cmdDraw_Click()

'draw the transparent sprite
BitBlt Me.hdc, 0, 0, SpriteWidth, SpriteHeight, DCMask, 0, 0, vbSrcAnd
BitBlt Me.hdc, 0, 0, SpriteWidth, SpriteHeight, DCSprite, 0, 0, vbSrcPaint

Me.Refresh

End Sub

Private Sub cmdExit_Click()

DeleteGeneratedDC DCMask
DeleteGeneratedDC DCSprite

Unload Me
Set frmMemoryDC = Nothing

End Sub

Private Sub cmdLoad_Click()
'Generate the two DCs, one for the mask and one for the sprite
DCMask = GenerateDC(App.Path & "\RM1.bmp")
If DCMask <= 0 Then
    MsgBox "Device Context Failure -- Graphics files may have been moved.  " & vbLf _
    & "Try moving the files to the same folder as the application.", , "ERROR!"
    Exit Sub
End If

'DCSprite = GenerateDC(App.Path & "\RM2.bmp")
'
'If DCSprite <= 0 Then
'    MsgBox "Failure in creating sprite dc"
'    'Clean up after the created mask
'    DeleteGeneratedDC DCMask
'End If


End Sub

'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

Private Sub Command1_Click()
BitBlt Me.hdc, 0, 0, 500, 400, RoomDC(1), 0, 0, vbSrcCopy
End Sub

Private Sub Form_Load()
RoomDC(1) = GenerateDC(App.Path & "\RM1.bmp")
If RoomDC(1) <= 0 Then
    MsgBox "Device Context Failure -- Graphics files may have been moved.  " & vbLf _
    & "Try moving the files to the same folder as the application.", , "ERROR!"
    Exit Sub
End If

End Sub

Private Sub TmrTime_Timer()
Me.Cls
If RotPos > 1500 Then
BitBlt Me.hdc, 0, 0, 500 - (RotPos + 500 - 2000), 400, RoomDC(1), RotPos, 0, vbSrcPaint
BitBlt Me.hdc, 500 - (RotPos + 500 - 2000), 0, RotPos + 500 - 2000, 400, RoomDC(1), 0, 0, vbSrcPaint
Else
BitBlt Me.hdc, 0, 0, 500, 400, RoomDC(1), RotPos, 0, vbSrcPaint
End If
Refresh
RotPos = RotPos + 10
If RotPos > 2000 Then RotPos = RotPos - 2000
End Sub
