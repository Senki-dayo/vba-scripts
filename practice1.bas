Attribute VB_Name = "Module1"
'HelloWorld
 Sub HelloWorld()
    MsgBox ("hello world")
 End Sub
'�Z���ɕ�������������
Sub CellChange()
    Worksheets("Sheet1").Range("A1").Value = "hello"
    Range("A2").Value = "hello2"
    Cells(3, 1).Value = "hello3"
    Cells(3, 1).Offset(1, 0).Value = "hello4"
End Sub
'�Z���ɕ������������� 2
Sub CellChange2()
    Range("A1", "B3").Value = "Thank you"
    Range("A4:C7").Value = "Thank you2"
    Range("4:4").Value = "row 4"
    Range("C:C").Value = "Column C"
End Sub
 ' �Z���ւ̏������݂�S�ď���
Sub CellClear()
    Cells.Clear
End Sub

