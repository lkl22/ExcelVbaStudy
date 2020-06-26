# VBA 获取合并单元格的Value

**思路：Select合并格中任意单元格都可以得到合并格的值。**

```vba
﻿Sub Test1()
    Dim ws As Worksheet

    Dim Target1 As Range
    Dim Target2 As Range
    Dim Target3 As Range
    
    Set ws = Worksheets("example")
    
    Set Target1 = ws.Range("H1")
    Set Target2 = ws.Range("I1")
    Set Target3 = ws.Range("J1")

    Target1.Select

    Debug.Print Selection.Cells(1).Value

    Target2.Select

    Debug.Print Selection.Cells(1).Value
    
    Target3.Select
    
    Debug.Print Selection.Cells(1).Value
End Sub
```

上面的例子中，H1、I1、J1三个为一个合并单元格，可以通过拿到任意一个Range，然后`.Select`，就可以通过`Selection.Cells(1).Value`获取到合并单元格的值。