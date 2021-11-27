Attribute VB_Name = "Module1"
Option Explicit
Type ShelfType
Product As String * 20
Price As Currency
Qty As Integer
End Type
Public Shelf As ShelfType
Public ShelfNumber As Integer
Sub Read_Record(which)
Get #1, which, Shelf
Form1.txtProduct.Text = Shelf.Product
Form1.txtPrice.Text = FormatNumber(Shelf.Price, 2)

Form1.txtQty.Text = Shelf.Qty
Form1.Caption = "Shelf " & Str(which)
ShelfNumber = which
End Sub
Sub Write_Record(which)
Shelf.Product = Form1.txtProduct.Text
Shelf.Price = Form1.txtPrice.Text
Shelf.Qty = Form1.txtQty.Text
Put #1, which, Shelf
End Sub
