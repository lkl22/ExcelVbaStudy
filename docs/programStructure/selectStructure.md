# VBA 程序选择结构

VBA 程序执行三大结构中，选择结构（判断）用于选择性地执行代码。选择结构与 Excel 的 `IF 函数`类似，也是以 `If` 为关键词，按照判断条件的真假，执行不同的操作。但是 VBA 中的 `If` 比 Excel 中的 `IF 函数`更强大。

* [选择结构基础](#选择结构基础)
* [示例数据](#示例数据)
* [If Then 结构](#ifThen)
* [If Else 结构](#ifElse)
* [If ElseIf Else 结构](#IfElseIfElse)
* [Select Case 结构](#SelectCase)
* [总结](#总结)

## 选择结构基础

选择结构，根据提供的条件表达式的值，如果为真（True），则执行选择结构的主体代码，否则跳过。

选择结构的核心是判断条件表达式的真假，这一步理解了，也就理解了选择结构的多种形式。

```flow
st=>start: 开始
op=>operation: 选择结构
主体代码
op1=>operation: 其他代码
cond=>condition: 条件表达式
e=>end: 结束

st->cond
cond(yes)->op->op1
cond(no)->op1->e
```

## 示例数据

本篇我们使用一个班级的考试成绩作为示例数据。

||A|B|C|D|E|
|---|---|---|---|---|---|
|1|学生|成绩|是否及格|评级|绩点|
|2|﻿Emma|45||||
|3|﻿Noah|99||||
|4|﻿Olivia|50||||
|5|﻿Liam|63||||
|6|﻿Ava|78||||
|7|﻿William|71||||
|8|﻿Sophia|91||||
|9|﻿Mason|69||||
|10|﻿Yuli|47||||

## <a name="ifThen">If Then 结构</a>

选择结构中，If Then 结构是最基础的一个。它只有条件表达式真时，执行的代码。

`If Then`结构基本语法如下，其中 `End If`是选择结构的结束标志。
```vba
If 条件表达式 Then
    '表达式为真时，执行的代码
End If
```
现在我们看实际的例子，判断学生是否及格，及格条件是成绩 ≥60。如果及格，在C列对应单元格填写“及格”。具体代码如下：
```vba
Sub MyCode()

    Dim i As Integer
    
    For i = 2 To 10
    
        If Cells(i, "B").Value >= 60 Then
            Cells(i, "C") = "及格"
        End If
        
    Next i

End Sub
```
我们可以看到，我们使用 B 列中的学生成绩与 60 分比较，如果≥60分，就在 C 列填写及格。

条件表达式是 `Cells(i, "B").Value >= 60`，选择性执行的代码部分是 `Cells(i, "C") = "及格"`。

其中，For 语句是表示循环结构，这里只需知道程序从第一个学生循环到最后一个学生，依次判断每个学生的成绩。循环结构将在下一篇中做详细介绍。

将以上代码运行后，可以看到运行结果如下：

||A|B|C|D|E|
|---|---|---|---|---|---|
|1|学生|成绩|是否及格|评级|绩点|
|2|﻿Emma|45||||
|3|﻿Noah|99|及格|||
|4|﻿Olivia|50||||
|5|﻿Liam|63|及格|||
|6|﻿Ava|78|及格|||
|7|﻿William|71|及格|||
|8|﻿Sophia|91|及格|||
|9|﻿Mason|69|及格|||
|10|﻿Yuli|47||||
﻿

## <a name="ifElse">If Else 结构</a>

很多时候，我们根据表达式的真假，真时执行一块代码，假时执行另一块代码。这种需求可以使用 `If Else`结构实现。

`If Else`结构中，条件表达式在真时，执行`Then`后的代码；条件表达式为假时，执行 `Else`后的代码。基本语法如下：
```vba
If 条件表达式 Then
    '真时执行的代码
Else
    '假时执行的代码
End If
```
我们继续看实际的例子。在上一个例子的基础上，这次对不及格的学生，在C列填入不及格。代码如下：
```vba
Sub MyCode()

    Dim i As Integer
    
    For i = 2 To 10
    
        If Cells(i, "B").Value >= 60 Then
            Cells(i, "C") = "及格"
        Else
            Cells(i, "C") = "不及格"
        End If
        
    Next i

End Sub
```
在这个例子中，条件表达式 `Cells(i, "B").Value >= 60`为假时，表示学生成绩低于60分，即不及格。这时就执行 `Else`后的代码。

程序运行结果如下：

||A|B|C|D|E|
|---|---|---|---|---|---|
|1|学生  |成绩 |是否及格   |评级 |绩点|
|2|﻿Emma    |45|不及格|   |   |
|3|﻿Noah    |99|及格  |   |   |
|4|﻿Olivia  |50|不及格|||
|5|﻿Liam    |63|及格|||
|6|﻿Ava     |78|及格|||
|7|﻿William |71|及格|||
|8|﻿Sophia  |91|及格|||
|9|﻿Mason   |69|及格|||
|10|﻿Yuli   |47|不及格|||

## <a name="IfElseIfElse">If ElseIf Else 结构</a>

前面两种结构中，最多有两种选择，即 ≥ 60 和 ＜ 60。有时针对同一个变量，可能存在多种判断标准。例如，对及格的学生，继续评级及格、良和优。



选择结构中，可以使用 `If ElseIf Else`结构，对同一个变量进行多次判断，并且为每一个判断结果编写不同的代码块，达到执行式 n 选 1 的效果。

`If ElseIf Else`结构的基本语法如下：
```vba
If 条件表达式1 Then
    '表达式1真时，执行的代码
ElseIf 条件表达式2 Then
    '表达式2真时，执行的代码
ElseIf 条件表达式3 Then
    '表达式3真时，执行的代码
    ...
ElseIf 条件表达式n Then
    '表达式n真时，执行的代码
Else
    '以上表达式都不为真时，执行的代码
End If
```
这种选择结构需要注意的是：

* 条件表达式是从第一个开始判断。
* 判断过程中，只要有一个表达式结果为真，那么执行对应的代码块，然后退出选择结构，不再继续判断剩下的表达式。
* 当所有的表达式都不为真时，执行 `Else`后的代码块。

根据以上规律，我们写一下判断学生成绩评级的代码。思路是，拿学生成绩，分别于85、75、60分比较，在 D 列填写对应的评级。
```vba
Sub MyCode()

    Dim i As Integer
    
    For i = 2 To 10
    
        If Cells(i, "B").Value >= 85 Then
            Cells(i, "D") = "优"
        ElseIf Cells(i, "B").Value >= 75 Then
            Cells(i, "D") = "良"
        ElseIf Cells(i, "B").Value >= 60 Then
            Cells(i, "D") = "及格"
        Else
            Cells(i, "D") = "不及格"
        End If
        
    Next i

End Sub
```
代码运行结果如下：

||A|B|C|D|E|
|---|---|---|---|---|---|
|1|学生  |成绩 |是否及格   |评级 |绩点|
|2|﻿Emma    |45|不及格| 不及格  |   |
|3|﻿Noah    |99|及格  | 优  |   |
|4|﻿Olivia  |50|不及格|不及格||
|5|﻿Liam    |63|及格|及格||
|6|﻿Ava     |78|及格|良||
|7|﻿William |71|及格|及格||
|8|﻿Sophia  |91|及格|优||
|9|﻿Mason   |69|及格|及格||
|10|﻿Yuli   |47|不及格|不及格||

## <a name="SelectCase">Select Case 结构</a>

`Select Case`结构是对同一个变量进行多次判断的另一种方式。相对于`If ElseIf Else`结构，它把条件表达式中的变量提取出来，使得代码结构更简洁，也更易于阅读。

`Select Case`结构的基本语法如下：
```vba
Select Case 变量
	Case 判断条件 1
    	'条件 1 真时，执行的代码
	Case 判断条件 2
    	'条件 2 真时，执行的代码
	Case 判断条件 3
    	'条件 3 真时，执行的代码
    Case Else
    	'之前的所有条件都不为真时，执行的代码
End Select
```
可以看到，`Select Case`结构把 `If`结构中的条件表达式拆分了，即把变量和判断条件分开了。

我们看前一个例子，使用Select Case结构，代码如下：
```vba
Sub MyCode()

    Dim i As Integer
    
    For i = 2 To 10
    
        Select Case Cells(i, "B").Value
            Case Is >= 85
                Cells(i, "D") = "优"
            Case Is >= 75
                Cells(i, "D") = "良"
            Case Is >= 60
                Cells(i, "D") = "及格"
            Case Else
                Cells(i, "D") = "不及格"
        End Select
        
    Next i

End Sub
```
这一例子中，学生成绩是变量，即 `Cells(i, "B").Value`，判断条件是每个 `Case `语句后的条件。

代码运行结果如下：

||A|B|C|D|E|
|---|---|---|---|---|---|
|1|学生  |成绩 |是否及格   |评级 |绩点|
|2|﻿Emma    |45|不及格| 不及格  |   |
|3|﻿Noah    |99|及格  | 优  |   |
|4|﻿Olivia  |50|不及格|不及格||
|5|﻿Liam    |63|及格|及格||
|6|﻿Ava     |78|及格|良||
|7|﻿William |71|及格|及格||
|8|﻿Sophia  |91|及格|优||
|9|﻿Mason   |69|及格|及格||
|10|﻿Yuli   |47|不及格|不及格||

## 总结

以上就是选择结构的基本用法，以及 4 种选择结构。选择结构的核心是判断条件表达式的真假，从另一个角度看，核心是如何写条件表达式。这一步写好了，下一步就是根据判断结果执行不同的代码块。


