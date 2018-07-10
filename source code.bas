Attribute VB_Name = "模块1"
Const MainpageSheet As String = "Homepage"
Const BOMInputSheet As String = "Input"
Const BOMOutputSheet As String = "Output"

Function num2asc2(ByVal n As Integer) As String
num2asc2 = Mid(Cells(1, n).Address, 2, IIf(n < 27, 1, 2))
End Function

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'清空 changed表
    Sheets(MainpageSheet).Range("e4") = ""
    Sheets(BOMOutputSheet).Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Sheets(MainpageSheet).Select
    Sheets(MainpageSheet).Range("e4") = "DONE"
    Range("E5:E10").ClearContents
End Sub
Sub Macro2()
'将原始BOM复制到changed表中
    Sheets(MainpageSheet).Range("e5") = ""
    Sheets(BOMInputSheet).Select
    Columns("A:Z").Select
    Selection.Copy
    Sheets(BOMOutputSheet).Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets(MainpageSheet).Select
    Sheets(MainpageSheet).Range("e5") = "DONE"
    Range("E6:E10").ClearContents
End Sub

Sub macro3()
'bom层级数字化
    Sheets(MainpageSheet).Range("e6") = ""
    zuihou = Sheets(BOMOutputSheet).Range("c1000000").End(xlUp).Row
    For i = 2 To zuihou
        Text = Sheets(BOMOutputSheet).Range("c" & i).Value
        Sheets(BOMOutputSheet).Range("c" & i).Value = Val(Replace(Text, ".", ""))
    Next
    Sheets(MainpageSheet).Range("e6") = "DONE"
    Range("E7:E10").ClearContents
End Sub
Sub macro4()
    Sheets(MainpageSheet).Range("e7") = ""
    zuihou = Sheets(BOMOutputSheet).Range("c1000000").End(xlUp).Row
'条件一，不是E或者F
    For i = 2 To zuihou
        If Sheets(BOMOutputSheet).Range("G" & i).Value = "" Then
            Sheets(BOMOutputSheet).Range("A" & i).Value = "X"
        End If
    Next
'条件二，是E
    For i = 2 To zuihou
        If Sheets(BOMOutputSheet).Range("G" & i).Value = "E" Then
            Sheets(BOMOutputSheet).Range("A" & i).Value = "X"
        End If
    Next
'条件三，是F50
    For i = 2 To zuihou
        If Sheets(BOMOutputSheet).Range("G" & i).Value = "F" And Sheets(BOMOutputSheet).Range("H" & i).Value = "50" Then
            Sheets(BOMOutputSheet).Range("A" & i).Value = "X"
        End If
    Next
'条件三，是F30
    For i = 2 To zuihou
        If Sheets(BOMOutputSheet).Range("G" & i).Value = "F" And Sheets(BOMOutputSheet).Range("H" & i).Value = "30" Then
            Sheets(BOMOutputSheet).Range("A" & i).Value = "X"
        End If
    Next
'条件4，是T
    For i = 2 To zuihou
        If Sheets(BOMOutputSheet).Range("E" & i).Value = "T" Then
            Sheets(BOMOutputSheet).Range("A" & i).Value = "X"
        End If
    Next
'条件5，是N
    For i = 2 To zuihou
        If Sheets(BOMOutputSheet).Range("E" & i).Value = "N" Then
            Sheets(BOMOutputSheet).Range("A" & i).Value = "X"
        End If
    Next
    
    Sheets(MainpageSheet).Range("e7") = "DONE"
    Range("E8:E10").ClearContents
End Sub

Sub Macro5()
'
' 插入适当的列
'
    Sheets(MainpageSheet).Range("e8") = ""
    Sheets(BOMOutputSheet).Select
    zuihou = Sheets(BOMOutputSheet).Range("c1000000").End(xlUp).Row
    k = 1
    For i = 2 To zuihou
        If k > Sheets(BOMOutputSheet).Range("c" & i).Value Then
        Else
            k = Sheets(BOMOutputSheet).Range("c" & i).Value
        End If
    Next
    kk = k + 3
   Do While kk > 0
        Sheets(BOMOutputSheet).Columns("G:G").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        kk = kk - 1
    Loop
    Sheets(MainpageSheet).Select
    Sheets(MainpageSheet).Range("e8") = "DONE"
    Range("E9:E10").ClearContents
End Sub

Sub macro6()
'调整bom层级位置
    Sheets(MainpageSheet).Range("e9") = ""
    Sheets(BOMOutputSheet).Select
    zuihou = Sheets(BOMOutputSheet).Range("c1000000").End(xlUp).Row
    For i = 2 To zuihou
        Text = Range("f" & i)
        numb = Range("c" & i) + 8
        Range(num2asc2(numb) & i) = Text
    Next
    k = 1
    For i = 2 To zuihou
        If k > Sheets(BOMOutputSheet).Range("c" & i).Value Then
        Else
            k = Sheets(BOMOutputSheet).Range("c" & i).Value
        End If
    Next
    
    Columns("G:" & num2asc2(k + 6)).Select
    Selection.ColumnWidth = 10
    Sheets(MainpageSheet).Select
    Sheets(MainpageSheet).Range("e9") = "DONE"
    Range("E10:E10").ClearContents
End Sub

Sub macro8()
'计算是否采购
Sheets(MainpageSheet).Range("e10") = ""
    Sheets(BOMOutputSheet).Select
    k = 1
    zuihou = Sheets(BOMOutputSheet).Range("c1000000").End(xlUp).Row
    For i = 2 To zuihou
        If k > Sheets(BOMOutputSheet).Range("c" & i).Value Then
        
        Else
            k = Sheets(BOMOutputSheet).Range("c" & i).Value
        End If
    Next
    kk = 10000
    For i = 2 To zuihou
        If kk < Sheets(BOMOutputSheet).Range("c" & i).Value Then
        
        Else
            kk = Sheets(BOMOutputSheet).Range("c" & i).Value
        End If
    Next
    
    For i = 2 To zuihou
            currentLevel = Sheets(BOMOutputSheet).Range("c" & i).Value
            
            If currentLevel = kk Then
                If Sheets(BOMOutputSheet).Range("a" & i).Value = "X" Then
                    Sheets(BOMOutputSheet).Range("G" & i).Value = "X"
                    Sheets(BOMOutputSheet).Range("b" & i).Value = "Higher level unpurchased"
                Else
                    Sheets(BOMOutputSheet).Range("G" & i).Value = "Y"
                    Sheets(BOMOutputSheet).Range("b" & i).Value = "This level purchased"
                End If
            End If
            
            If currentLevel > kk Then
                tarlevel = currentLevel - 1
                For NN = i - 1 To 2 Step -1
                        findlevel = Sheets(BOMOutputSheet).Range("c" & NN).Value + 0
                        If findlevel = tarlevel Then
'                            MsgBox Sheets("changed").Range("A" & NN).Value
                            If Sheets(BOMOutputSheet).Range("G" & NN).Value = "X" And Sheets(BOMOutputSheet).Range("b" & NN).Value = "Higher level unpurchased" Then
                                                If Sheets(BOMOutputSheet).Range("a" & i).Value = "X" Then
                                                            If Sheets(BOMOutputSheet).Range("c" & (i + 1)).Value <= Sheets(BOMOutputSheet).Range("c" & (i)).Value Then
                                                                    Sheets(BOMOutputSheet).Range("G" & i).Value = "Y"
                                                                    Sheets(BOMOutputSheet).Range("b" & i).Value = "Problem"
                                                            ElseIf Sheets(BOMOutputSheet).Range("c" & (i + 1)).Value > Sheets(BOMOutputSheet).Range("c" & (i)).Value Then
                                                                     Sheets(BOMOutputSheet).Range("G" & i).Value = "X"
                                                                     Sheets(BOMOutputSheet).Range("b" & i).Value = "Higher level unpurchased"
                                                            End If
                                                Else
                                                            Sheets(BOMOutputSheet).Range("G" & i).Value = "Y"
                                                            Sheets(BOMOutputSheet).Range("b" & i).Value = "This level purchased"
                                                            
                                                End If
                                                
                                                
                            End If
                            
                             If Sheets(BOMOutputSheet).Range("G" & NN).Value = "X" And Sheets(BOMOutputSheet).Range("b" & NN).Value = "Higher level purchased" Then
                                                    Sheets(BOMOutputSheet).Range("G" & i).Value = "X"
                                                    Sheets(BOMOutputSheet).Range("b" & i).Value = "Higher level purchased"
                            End If
                            
                            If Sheets(BOMOutputSheet).Range("G" & NN).Value = "Y" And Sheets(BOMOutputSheet).Range("b" & NN).Value = "This level purchased" Then
                                                    Sheets(BOMOutputSheet).Range("G" & i).Value = "X"
                                                    Sheets(BOMOutputSheet).Range("b" & i).Value = "Higher level purchased"
                            End If
                            
                            
                            
                            
                            
                            GoTo aaax:
                        End If
                Next
aaax:
            End If
                
    Next
    Sheets(MainpageSheet).Range("e10") = "DONE"
End Sub
