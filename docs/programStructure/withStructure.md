# VBA With 结构

VBA 中，With 结构用于组合同一个对象的多个属性和方法，避免重复写同一个对象名，提高编程和运行效率。

* [With 结构语法](#With结构语法)
* [With 结构实例](#With结构实例)
* [嵌套 With 结构](#嵌套With结构)
* [总结](#总结)

## <a name="With结构语法">With 结构语法</a>

`With` 结构由 `With` 和 `End With` 两个语句构成，对象的属性和方法都写在两者之间。基本语法如下：
```vba
With [对象]
    .[属性] = [数据]
    .[方法]
    '其他属性和方法
End With
```
`With` 结构里，对象的属性和方法均有点 (.)符号开始，后接对象的属性名和方法名。

## <a name="With结构实例">With 结构实例</a>

现在看一个实际的例子，需要将工作簿中 Sheet1 工作表设置新名称，然后设置标签颜色为黑色，最后隐藏工作表。

如果不用 With 结构，代码如下：
```vba
Sub MyCode()

    Worksheets("Sheet1").Name = "新名称"
    Worksheets("Sheet1").Tab.ThemeColor = xlThemeColorLight1
    Worksheets("Sheet1").Visible = xlSheetHidden
    
End Sub
```
可以看到，每个语句都重复写 `Worksheets("Sheet1") `部分。

使用 `With` 结构，可以避免重复写同一个对象名，代码如下：
```vba
Sub MyCode()

    With Worksheets("Sheet1")
        .Name = "新名称"
        .Tab.ThemeColor = xlThemeColorLight1
        .Visible = xlSheetHidden
    End With
    
End Sub
```
## <a name="嵌套With结构">嵌套 With 结构</a>

`With` 结构还能嵌套编写，即一个 `With` 结构中，如果父对象的属性是另一个对象，则针对这个子对象，继续使用 `With` 结构。

在之前的例子中，如果需要将 Sheet1 工作表中，A1:A10 单元格区域设置背景颜色，调整字体和字体大小，可以使用如下代码：
```vba
Sub MyCode()

    With Worksheets("Sheet1")
        .Name = "新名称"
        .Tab.ThemeColor = xlThemeColorLight1
        .Visible = xlSheetHidden
        
        With .Range("A1:A10")
            .Interior.ThemeColor = xlThemeColorAccent1
            .Font.Size = 12
            .Font.Name = "等线"
        End With
        
    End With
    
End Sub
```

## 总结

本篇我们学习了 VBA 程序结构中的 With 结构。With 结构可以将同一个对象的多个属性和方法组合起来，避免重复写对象名。此外，With 结构还能嵌套使用，进一步提高编程效率和程序运行效率。
