# VBA 选择EXCEL单元格启动宏

1、在工作表sheet1标签上击右键，查看代码

2、粘贴如下代码

```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Count = 1 Then '选择的单元格数量为1个时
    
        If Target.Column > 2 And Target.Row = 1 And Target.Value <> "" Then
            '单元格是第一行，第二列以后的非空单元格
            
            Call Test '调用指定的过程（宏）
        End If
    End If
End Sub
```


