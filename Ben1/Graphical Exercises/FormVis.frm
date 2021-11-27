VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim color(300, 200) As Long
Dim newcolor(300, 200) As Long

Private Sub Form_Click()
Dim u As Integer, v As Integer, u1 As Integer, v1 As Integer
On Error Resume Next
For v = 2 To 200
If Rnd > 0.5 Then
    color(1, v) = RGB(255, 255, 255)
Else
    color(1, v) = RGB(0, 0, 0)
End If
For u = 2 To 300
For v = 2 To 200
For u1 = -1 To 1
For v1 = -1 To 1
newcolor(u, v) = color(u, v) + color(u + u1, v + v1) / 8
Next v1
Next u1
newcolor(u, v) = newcolor(u, v) / 2
PSet (u, v), newcolor(u, v)
Next v
Next u
For u = 1 To 300
For v = 1 To 200
color(u, v) = newcolor(u, v)
Next v
Next u

End Sub

Private Sub Form_Load()
Dim x, y
For x = 1 To 300
For y = 1 To 200
color(x, y) = &H888888
newcolor(x, y) = &H888888
Next y
Next x
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If x > 300 Or y > 200 Then Exit Sub
PSet (x, y), &H0
color(x, y) = RGB(255, 255, 255)

End Sub

