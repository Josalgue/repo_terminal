Sub GradeBook()

Dim grade As Double

    grade = Range("B2")
    
    If grade >= 90 Then
        Cells(2, 3).Value = "Pass"
        Cells(2, 3).Interior.ColorIndex = 4
        Cells(2, 4).Value = "A"
    
    ElseIf grade >= 80 Then
        Cells(2, 3).Value = "Pass"
        Cells(2, 3).Interior.ColorIndex = 4
        Cells(2, 4).Value = "B"
    
    ElseIf grade >= 70 Then
        Cells(2, 3).Value = "Pass"
        Cells(2, 3).Interior.ColorIndex = 6
        Cells(2, 4).Value = "C"
    
    Else
        Cells(2, 3).Value = "Fail"
        Cells(2, 3).Interior.ColorIndex = 3
        Cells(2, 4).Value = "D"
    
    End If
    
End Sub

Sub reset_button()

    Range("B12").Value = Range("B2").Value
        Cells(12, 3).Value = Cells(2, 3).Value
        Cells(12, 3).Interior.ColorIndex = Cells(2, 3).Interior.ColorIndex
        Cells(12, 4).Value = Cells(2, 4).Value
        
 Range("B2").Value = ""
        Cells(2, 3).Value = ""
        Cells(2, 3).Interior.ColorIndex = 0
        Cells(2, 4).Value = ""
        

End Sub
