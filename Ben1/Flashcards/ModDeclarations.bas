Attribute VB_Name = "Module1"
Type CardData
    Prompt As String
    Answer As String
    Explanation As String
    Correct As Integer
End Type

Public Function Replace(SearchText As String, FindText As String, ReplaceText As String)
Dim i
Replace = SearchText
i = 1
Do While i < Len(SearchText)
If Mid(SearchText, i, Len(FindText)) = FindText Then
SearchText = Mid(SearchText, 1, i - 1) & ReplaceText & Right(SearchText, Len(SearchText) - i - Len(FindText) + 1)
Replace = SearchText
i = Len(SearchText) - Len(Right(SearchText, Len(SearchText) - i - Len(ReplaceText) + 1))
End If
i = i + 1
Loop
End Function
