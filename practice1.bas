Attribute VB_Name = "Module1"
'HelloWorld
 Sub HelloWorld()
    MsgBox ("hello world")
 End Sub
'セルに文字を書き込む
Sub CellChange()
    Worksheets("Sheet1").Range("A1").Value = "hello"
    Range("A2").Value = "hello2"
    Cells(3, 1).Value = "hello3"
    Cells(3, 1).Offset(1, 0).Value = "hello4"
End Sub
'セルに文字を書き込む 2
Sub CellChange2()
    Range("A1", "B3").Value = "Thank you"
    Range("A4:C7").Value = "Thank you2"
    Range("4:4").Value = "row 4"
    Range("C:C").Value = "Column C"
End Sub
' セルへの書き込みを全て消す
Sub CellClear()
    Cells.Clear
End Sub
'with ステートメントでまとめて書く  ( 前 )
Sub WithTest()
    Range("A1").Value = "hello"
    Range("A1").Font.Bold = True
    Range("A1").Font.Size = 16
    Range("A1").Interior.Color = vbRed
 End Sub
'with ステートメントでまとめて書く
 Sub WithTest2()
    With Range("A2")
        .Value = "hello"
        With .Font
            .Bold = True
            .Size = 16
        End With
         .Interior.Color = vbRed
    End With
 End Sub
 'セルの値を取得する
 Sub GetTest()
    MsgBox (Range("A1").Value)
    MsgBox (Range("A1").Font.Size)
 End Sub
 'メソッドを使う'
Sub MethodTest()
    Range("A1", "B8").Value = "test"
    Range("B2").Clear
    Range("B5").Delete shift:=xlShiftUp
    Worksheets.Add after:=Worksheets("Sheet1"), Count:=2
End Sub
   
