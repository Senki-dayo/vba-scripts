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
'with �X�e�[�g�����g�ł܂Ƃ߂ď���  ( �O )
Sub WithTest()
    Range("A1").Value = "hello"
    Range("A1").Font.Bold = True
    Range("A1").Font.Size = 16
    Range("A1").Interior.Color = vbRed
 End Sub
'with �X�e�[�g�����g�ł܂Ƃ߂ď���
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
 '�Z���̒l���擾����
 Sub GetTest()
    MsgBox (Range("A1").Value)
    MsgBox (Range("A1").Font.Size)
 End Sub
 '���\�b�h���g��'
Sub MethodTest()
    Range("A1", "B8").Value = "test"
    Range("B2").Clear
    Range("B5").Delete shift:=xlShiftUp
    Worksheets.Add after:=Worksheets("Sheet1"), Count:=2
End Sub
'�ϐ����g��
Sub VariableTest()
    Dim x As Integer
    x = 1
    Dim y As Double
     y = 10.5
     Dim s As String
     s = "hello"
     Dim d As Date
     d = "2012/04/23"
     Dim z As Variant
      '�l����͂��Ă���^�����܂�
     Dim f As Boolean
      f = True
     Dim r As Range
     Set r = Range("A1")
    
    '�C�~�f�B�G�C��ʂłŃo�b�N�ł���
    Debug.Print x
    Debug.Print y / 3
    Debug.Print y \ 3
    Debug.Print s & "world"
    Range("A1").Value = x
    r.Value = d + 7
 End Sub
'�z����g��
Sub ArrayTest()
    Dim sales(2) As Integer
    sales(0) = 200
    sales(1) = 150
    sales(2) = 300
    Debug.Print sales(1)
    
    Dim arr As Variant
    arr = Array(10, 20, 30)
    Debug.Print arr(2)
    
End Sub
'������������� (If)
Sub IfTest()
    Range("A1").Value = 70
    '= < > <= >= <>(�������Ȃ�) and not or
    If Range("A1").Value > 80 Then
        Range("A2").Value = "OK!"
    ElseIf Range("A1").Value > 60 Then
        Range("A2").Value = "soso..."
    Else
        Range("A2").Value = "NG!"
    End If
End Sub
 '�������������(Select)
Sub SelectTest()
    Dim signal As String
    signal = Range("A1").Value
    Dim result As Range
    Set result = Range("A2")
    
    Select Case signal
    Case "red"
        result.Value = "Stop!"
    Case "green"
        result.Value = "Go!"
    Case "yelow"
        result.Value = "Caution!"
    Case Else
        result.Value = "n.a"
    End Select
End Sub
'�J��Ԃ�������(while)
Sub WhileTest()
    Dim i As Integer
    i = 1
    
    Do While i < 10
        Cells(i, 1).Value = i
        i = i + 1
    Loop

End Sub
'�J��Ԃ������� (for)
Sub ForTest()
    Dim i As Integer
    
    For i = 1 To 9 Step 2
        Cells(i, 1).Value = i
    Next i

End Sub
'�J��Ԃ������� (each)
Sub EachTest()
    Dim names As Variant
    names = Array("taguchi", "fkoji", "dotinstall")
    
    For Each name In names
        Debug.Print name
    Next name

End Sub
'�O���֐�(Sub:�Ԃ�l�Ȃ�)���Ăяo��
Sub CallSubTest()

    Dim names As Variant
    names = Array("taguchi", "fkoji", "dotinstall")
    
    For Each name In names
        Call SayHiSub(name)
    Next name

End Sub
Sub SayHiSub(ByVal name As String)
    Debug.Print "hi!, " & name
End Sub
'�O���֐�(Function:�Ԃ�l ����)���Ăяo��
Sub CallFuncTest()

    Dim names As Variant
    names = Array("taguchi", "fkoji", "dotinstall")
    
    For Each name In names
        Debug.Print SayHiFunc(name)
    Next name

End Sub
Function SayHiFunc(ByVal name As String)
     SayHiFunc = "hi!, " & name
End Function
'Sample DoWhile
Sub FindLowScores()

    Dim i As Long
    Dim n As Long
    i = 2
    n = 0
    ' (�s�ԍ�,��ԍ�)
    Do While Cells(i, 1).Value <> ""
        If Cells(i, 2).Value < 60 Then
            Cells(i, 2).Interior.Color = vbRed
            n = n + 1
        End If
        i = i + 1
    Loop
    
    MsgBox (n & "���Y�����܂����I")
End Sub
