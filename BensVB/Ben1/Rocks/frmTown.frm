VERSION 5.00
Begin VB.Form frmTown 
   BackColor       =   &H00000080&
   Caption         =   "Golddigger 1.0 -- Mining Town"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000040C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4500
      Picture         =   "frmTown.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1380
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000040C0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000040C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1380
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   3240
      TabIndex        =   9
      Text            =   "100"
      Top             =   1140
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtInterest 
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      Text            =   "1"
      Top             =   1860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDuration 
      Height          =   315
      Left            =   3240
      TabIndex        =   7
      Text            =   "1"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkFood 
      BackColor       =   &H00000080&
      Caption         =   "Food ($4.00)"
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CheckBox chkDynamite 
      BackColor       =   &H00000080&
      Caption         =   "Dynamite ($35.00)"
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CheckBox chkWood 
      BackColor       =   &H00000080&
      Caption         =   "Wood ($10.00)"
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdLoan 
      BackColor       =   &H000040C0&
      Caption         =   "Loan $"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdBorrow 
      BackColor       =   &H000040C0&
      Caption         =   "Borrow $"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1140
      Width           =   1695
   End
   Begin VB.CommandButton cmdSell 
      BackColor       =   &H000040C0&
      Caption         =   "Sell"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H000040C0&
      Caption         =   "Buy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1695
   End
   Begin VB.Label lblEndFunds 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "$1,000.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblFunds 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "$1,000.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label lblInterest 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "% Interest:"
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1980
      TabIndex        =   12
      Top             =   1860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblDuration 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Duration (Weeks):"
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Top             =   1500
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblAmount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   1140
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "frmTown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkDynamite_Click()
Select Case Business
Case Is = "Buy"
 Select Case chkDynamite.Value
 Case Is = 0
 Profits = Profits + TNTPrice
 Case Is = 1
 Profits = Profits - TNTPrice
 End Select
Case Is = "Sell"
 Select Case chkDynamite.Value
 Case Is = 0
 Profits = Profits - TNTPrice
 Case Is = 1
 Profits = Profits + TNTPrice
 End Select
 End Select
 EndFunds = Money + Profits
 lblEndFunds.Caption = FormatCurrency(EndFunds)

End Sub

Private Sub chkFood_Click()
Select Case Business
Case Is = "Buy"
 Select Case chkFood.Value
 Case Is = 0
 Profits = Profits + FoodPrice * (5 - Food)
 Case Is = 1
 Profits = Profits - FoodPrice * (5 - Food)
 End Select
Case Is = "Sell"
 Select Case chkFood.Value
 Case Is = 0
 Profits = Profits - FoodPrice
 Case Is = 1
 Profits = Profits + FoodPrice
 End Select
 End Select
 EndFunds = Money + Profits
 lblEndFunds.Caption = FormatCurrency(EndFunds)

End Sub

Private Sub chkWood_Click()
Select Case Business
Case Is = "Buy"
 Select Case chkWood.Value
 Case Is = 0
 Profits = Profits + WoodPrice
 Case Is = 1
 Profits = Profits - WoodPrice
 End Select
Case Is = "Sell"
 Select Case chkWood.Value
 Case Is = 0
 Profits = Profits - WoodPrice
 Case Is = 1
 Profits = Profits + WoodPrice
 End Select
 End Select
 EndFunds = Money + Profits
 lblEndFunds.Caption = FormatCurrency(EndFunds)
End Sub

Private Sub cmdBorrow_Click()
cmdBuy.Enabled = False
cmdSell.Enabled = False
cmdBorrow.Enabled = False
cmdLoan.Enabled = False
cmdBorrow.BackColor = RGB(255, 200, 0)
lblAmount.Visible = True
txtAmount.Visible = True
lblDuration.Visible = True
txtDuration.Visible = True
lblInterest.Visible = True
txtInterest.Visible = True
cmdExit.Visible = False
cmdOK.Visible = True
cmdCancel.Visible = True
Business = "Borrow"
InterestRate = 0.01 + 1.01 ^ Week
End Sub

Private Sub cmdBuy_Click()
cmdBuy.Enabled = False
cmdSell.Enabled = False
cmdBorrow.Enabled = False
cmdLoan.Enabled = False
cmdBuy.BackColor = RGB(255, 200, 0)
chkWood.Value = 0
chkDynamite.Value = 0
chkFood.Value = 0
chkWood.Visible = True
chkDynamite.Visible = True
chkFood.Visible = True
cmdExit.Visible = False
cmdOK.Visible = True
cmdCancel.Visible = True
Business = "Buy"
If CarryWood = True Then
chkWood.Enabled = False
Else
chkWood.Enabled = True
End If
If TNTPlaced = 1 Then
chkDynamite.Enabled = False
Else
chkDynamite.Enabled = True
End If
If Food = 5 Then
chkFood.Enabled = False
Else
chkFood.Enabled = True
End If
WoodPrice = 9 + (1.006 ^ Week)
TNTPrice = 49 + (1.03 ^ Week)
FoodPrice = 3 + (1.003 ^ Week)
chkWood.Caption = "Wood (" & FormatCurrency(WoodPrice) & ")"
chkDynamite.Caption = "Dynamite (" & FormatCurrency(TNTPrice) & ")"
chkFood.Caption = "Food (" & FormatCurrency(FoodPrice) & ")"
lblFunds.Caption = FormatCurrency(Money)
lblEndFunds = FormatCurrency(Money + Profits)
End Sub

Private Sub cmdCancel_Click()
    Business = ""
    chkWood.Visible = False
    chkDynamite.Visible = False
    chkFood.Visible = False
    cmdExit.Visible = True
    lblEndFunds.Caption = FormatCurrency(Money)
    cmdBuy.Enabled = True
    cmdSell.Enabled = True
    cmdBorrow.Enabled = True
cmdLoan.Enabled = True
    cmdBuy.BackColor = &H40C0&
cmdSell.BackColor = &H40C0&
cmdBorrow.BackColor = &H40C0&
cmdLoan.BackColor = &H40C0&

End Sub

Private Sub cmdExit_Click()
frmMine.tmrTime.Enabled = True
Unload frmTown
End Sub

Private Sub cmdOK_Click()
Dim i
Select Case Business
Case Is = "Buy"
    If chkFood.Value = 1 Then
    Food = 5
    For i = 0 To 4
    frmMine.imgFood(i).Visible = True
    Next i
    End If
    If chkWood.Value = 1 Then
    CarryWood = True
    frmMine.imgCarryWood.Visible = True
    End If
    If chkDynamite.Value = 1 Then
    Load frmMine.imgTNT(TNTPlaced - 1)
    frmMine.imgTNT(TNTPlaced - 1).Left = 480 * (TNTPlaced - 1)
    frmMine.imgTNT(TNTPlaced - 1).Top = 3780
    frmMine.imgTNT(TNTPlaced - 1).Visible = True
    TNTPlaced = TNTPlaced - 1
    End If
    Money = Money + Profits
    Profits = 0
    lblFunds.Caption = FormatCurrency(Money)
    lblEndFunds.Caption = FormatCurrency(Money)
    Business = ""
    chkWood.Visible = False
    chkDynamite.Visible = False
    chkFood.Visible = False
    cmdExit.Visible = True
    lblEndFunds.Caption = FormatCurrency(Money)
Case Is = "Sell"
    If chkFood.Value = 1 Then
    Food = Food - 1
    frmMine.imgFood(Food).Visible = False
    End If
    If chkWood.Value = 1 Then
    CarryWood = False
    frmMine.imgCarryWood.Visible = False
    End If
    If chkDynamite.Value = 1 Then
    Unload frmMine.imgTNT(TNTPlaced)
    TNTPlaced = TNTPlaced + 1
    End If
    Money = Money + Profits
    Profits = 0
    lblFunds.Caption = FormatCurrency(Money)
    lblEndFunds.Caption = FormatCurrency(Money)
    Business = ""
    chkWood.Visible = False
    chkDynamite.Visible = False
    chkFood.Visible = False
    cmdExit.Visible = True
    lblEndFunds.Caption = FormatCurrency(Money)
    Case Is = "Borrow"
    If Val(txtAmount.Text) > 0 And Val(txtDuration.Text) > 0 And Val(txtInterest.Text) > 0 Then
    Bonds = Bonds + 1
    ReDim BorrowAmount(1 To Bonds), BorrowDuration(1 To Bonds)
    BorrowAmount(Bonds) = txtAmount.Text + txtInterest.Text * txtAmount
    BorrowDuration(Bonds) = Int(txtDuration.Text)
    Profits = Profits + txtInterest.Text
    Else
    MsgBox "You have angered the yellow submarine.  You shall die."
    End If
End Select
cmdBuy.Enabled = True
cmdSell.Enabled = True
cmdBorrow.Enabled = True
cmdLoan.Enabled = True
cmdBuy.BackColor = &H40C0&
cmdSell.BackColor = &H40C0&
cmdBorrow.BackColor = &H40C0&
cmdLoan.BackColor = &H40C0&
End Sub

Private Sub cmdSell_Click()
cmdBuy.Enabled = False
cmdSell.Enabled = False
cmdBorrow.Enabled = False
cmdLoan.Enabled = False
cmdSell.BackColor = RGB(255, 200, 0)
chkWood.Value = 0
chkDynamite.Value = 0
chkFood.Value = 0
chkWood.Visible = True
chkDynamite.Visible = True
chkFood.Visible = True
cmdExit.Visible = False
cmdOK.Visible = True
cmdCancel.Visible = True
Business = "Sell"
If CarryWood = False Then
chkWood.Enabled = False
Else
chkWood.Enabled = True
End If
If TNTPlaced = 6 Then
chkDynamite.Enabled = False
Else
chkDynamite.Enabled = True
End If
If Food = 0 Then
chkFood.Enabled = False
Else
chkFood.Enabled = True
End If
WoodPrice = 9 + (1.008 ^ Week)
TNTPrice = 49 + (1.03 ^ Week)
FoodPrice = 3 + (1.003 ^ Week)
chkWood.Caption = "Wood (" & FormatCurrency(WoodPrice) & ")"
chkDynamite.Caption = "Dynamite (" & FormatCurrency(TNTPrice) & ")"
chkFood.Caption = "Food (" & FormatCurrency(FoodPrice) & ")"
lblFunds.Caption = FormatCurrency(Money)
lblEndFunds = FormatCurrency(Money + Profits)
End Sub

Private Sub Form_Load()
frmTown.ZOrder
lblFunds.Caption = FormatCurrency(Money)
    lblEndFunds.Caption = FormatCurrency(Money)
End Sub

