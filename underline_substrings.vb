Option Explicit
Function Create_Vector(Matrix_Range As Range) As Variant
Dim No_of_Cols As Integer, No_Of_Rows As Integer
Dim i As Integer
Dim j As Integer
Dim Cell

No_of_Cols = Matrix_Range.Columns.Count
No_Of_Rows = Matrix_Range.Rows.Count
ReDim Temp_Array(No_of_Cols * No_Of_Rows)

If Matrix_Range Is Nothing Then Exit Function
If No_of_Cols = 0 Then Exit Function
If No_Of_Rows = 0 Then Exit Function

For j = 1 To No_Of_Rows
    For i = 0 To No_of_Cols - 1
    Temp_Array((i * No_Of_Rows) + j) = Matrix_Range.Cells(j, i + 1)
    Next i
Next j
Create_Vector = Temp_Array
End Function

Sub Underline_SubStrings()

Dim NamesVector
NamesVector = Create_Vector(Sheets("Sheet2").Range("a1:b4"))

Dim cl As Range
Dim SearchText As String
Dim StartPos As Integer
Dim EndPos As Integer
Dim TestPos As Integer
Dim TotalLen As Integer

On Error Resume Next

Dim Name As Variant
For Each Name In NamesVector
  If Not IsEmpty(Name) Then
     For Each cl In Selection
     TotalLen = Len(Name)
     StartPos = InStr(cl.Value2, Name)
     TestPos = 0
       Do While StartPos > TestPos
         cl.Characters(StartPos, TotalLen).Font.Underline = xlUnderlineStyleSingle
         EndPos = StartPos + TotalLen
         TestPos = EndPos + 1
         StartPos = InStr(TestPos, cl.Value2, Name)
      Loop
   Next cl
  End If
Next Name

End Sub

