Attribute VB_Name = "Module1"
Option Explicit
Type ShelfType
Product As String * 20
Price As Currency
Qty As Integer
End Type
Public Shelf As ShelfType
Global ShelfNumber As Integer
