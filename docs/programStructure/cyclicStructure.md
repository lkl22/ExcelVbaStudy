# VBA 程序循环结构

VBA 中，循环结构用于多次重复执行同一段代码。重复次数通过特定数字或特定条件控制。

通过控制循环过程中特定变量，循环结构可执行复杂的重复任务。


* [VBA循环结构类型](#VBA循环结构类型)
* [For循环](#For循环)
  * [For … Next 循环](#ForNext)
  * [For Each 循环](#ForEach)
  * [Exit For 语句](#ExitFor)
* [Do While 循环](#DoWhile)
  * [Do While … Loop 循环](#DoWhileLoop)
  * [Do … Loop While 循环](#DoLoopWhile)
  * [Exit Do 语句](#ExitDo)
* [Do Until 循环](#DoUntil)
  * [Do Until … Loop 循环](#DoUntilLoop)
  * [Do … Loop Until 循环](#DoLoopUntil)
* [总结](#总结)


## <a name="VBA循环结构类型">VBA循环结构类型</a>

VBA 中循环结构有 3 种类型，它们是：

* For 循环
* Do While 循环
* Do Until 循环

## <a name="For循环">For循环</a>

For 循环是最常用的循环类型，它有两种形式：

* For … Next 循环
* For Each 循环

### <a name="ForNext">For … Next 循环</a>
使用 `For ... Next` 循环可以按指定次数，循环执行一段代码。For 循环使用一个数字变量，从初始值开始，每循环一次，变量值增加或减小，直到变量的值等于指定的结束值时，循环结束。

`For ... Next` 循环语法如下：
```vba
For [变量] = [初始值] To [结束值] Step [步长]
    '这里是循环执行的语句
Next
```
其中：

* [变量] 是一个数字类型变量，可在循环执行的语句里使用。
* [初始值] 和 [结束值] 是给定的值；
* [步长] 是每次循环时，变量的增量。如果为正值，变量增大；如果为负值，变量减小。

下面看一个实际的例子，求 1 至 10 数字的累积和。

```vba
Sub MyCode()

    Dim i As Integer
    Dim sum As Integer
    
    For i = 1 To 10 Step 1
        sum = sum + i
    Next
    
End Sub
```
可以看到，For 循环使用 `i` 变量，循环 10 次，`i` 的值从 1 到 10 变化。

**值得注意的是，For 循环的 Step 值如果是 1，则 Step 关键词可省略**。上述过程循环部分可写成如下方式：
```vba
For i = 1 To 10
    sum = sum + i
Next
```

### <a name="ForEach">For Each 循环</a>

`For Each` 循环用于逐一遍历一个数据集合中的所有元素。数据集合包括数组、Excel 对象集合、字典等。

`For Each` 循环不需要一个数字变量，但是需要与数据集合中的元素相同的数据类型变量。其基本语法如下：
```vba
For Each [元素] In [元素集合]
    '循环执行的代码
Next [元素]
```
其中，

* [元素] 是与集合中的元素相同类型的变量，该变量可在循环代码中使用。
* [元素集合]是包括多个元素的集合。

下面看一个实际例子，循环打印出工作簿中所有工作表的名称。
```vba
Sub MyCode()

    Dim sh As Worksheet
    
    For Each sh In Worksheets
        Debug.Print sh.Name
    Next sh

End Sub
```
`sh` 变量就是元素变量，`Worksheets` 是工作簿中所有工作表的集合。

### <a name="ExitFor">Exit For 语句</a>

`Exit For` 语句用于跳出循环过程，一般在提前结束循环时使用，均适用于 `For Next` 循环和 `For Each` 循环。

看一个实际的例子，求 1 – 10 数字的和时，当和大于 30 就停止循环。
```vba
Sub MyCode()

    Dim i As Integer
    Dim sum As Integer
    
    For i = 1 To 10
    
        sum = sum + i
        
        If sum > 30 Then
            Exit For
        End If
        
    Next
    
End Sub
```
在这段代码中，`sum` 变量大于 30 时，循环就停止。

## <a name="DoWhile">Do While 循环</a>

`Do While` 循环用于满足指定条件时循环执行一段代码的情形。循环的指定条件在 `While` 关键词后书写。

`Do While` 循环也有两种形式：

* Do While … Loop 循环
* Do … Loop While 循环

### <a name="DoWhileLoop">Do While … Loop 循环</a>

`Do While … Loop `循环，根据 `While` 关键词后的条件表达式的值，真时执行，假时停止执行。基本语法如下：
```vba
Do While [条件表达式]
    '循环执行的代码
Loop
```
其中，只要 [条件表达式] 为真，将一直循环执行。[条件表达式] 一旦为假，则停止循环，程序执行 `Loop` 关键词后的代码。

看一个实际的例子，还是求 1- 10 累积和。
```vba
Sub MyCode()

    Dim i As Integer
    Dim sum As Integer
    
    i = 1
    Do While i <= 10
        sum = sum + i
        i = i + 1
    Loop
    
End Sub
```
`i` 变量的初始值是 1，根据 `While` 后的条件，只要 `i` 变量小于等于 10，后续的代码就可以一直循环执行。

这里为了演示使用了 `Do While` 循环，实际情况下，这种求和问题，使用 `For` 循环更简洁。

### <a name="DoLoopWhile">Do … Loop While 循环</a>

与上一种 Do 循环不同的是，`Do ... Loop While`循环至少循环执行代码一次后，再判断条件表达式的值。基本语法如下：
```vba
Do
    '循环执行的代码
Loop While [条件表达式]
```
其中，While 和条件表达式写在 Loop 关键词后。

### <a name="ExitDo">Exit Do 语句</a>

与 `Exit For` 语句类似，`Exit Do` 语句用于跳出 `Do While` 循环。

## <a name="DoUntil">Do Until 循环</a>

`Do Until` 循环与 `Do While` 循环类似。不同点在于，`Do While` 在条件表达式为真时，继续执行循环；而 `Do Until` 在条件表达式为真时，停止执行循环。

Do Until 循环也有两种形式：

* Do Until … Loop 循环
* Do … Loop Until 循环

### <a name="DoUntilLoop">Do Until … Loop 循环</a>

循环开始前判断 `Until` 后条件表达式的值，如果是真，停止循环；如果是假，继续执行循环。基本语法如下：
```vba
Do Until [条件表达式]
    '循环执行的代码
Loop
```

### <a name="DoLoopUntil">Do … Loop Until 循环</a>

先运行一次，再判断 `Until` 后条件表达式的值，如果是真，停止循环；如果是假，继续执行循环。基本语法如下：
```vba
Do
    '循环执行的代码
Loop Until [条件表达式]
```
其他使用方法与 Do While 循环一致。

## 总结

本篇文章我们学习了 VBA 中程序循环结构基础，以及多种循环结构形式。包括子类在内，VBA 中常使用的循环结构包括 6 种，它们是：

|循环结构	|说明
|---|---
|For … Next 循环	|按指定次数循环执行
|For Each 循环	|逐一遍历数据集合中的每一个元素
|Do While … Loop 循环	|当条件为真时，循环执行
|Do … Loop While 循环	|当条件为真时，循环执行。无论条件真假，至少运行一次
|Do Until … Loop 循环	|直到条件为真时，循环执行
|Do … Loop Until 循环	|直到条件为真时，循环执行。无论条件真假，至少运行一次

此外，学习了两种跳出循环的语句，它们是：

|跳出语句	|说明
|---|---
|Exit For	|跳出 For 循环
|Exit Do	|跳出 Do While/Until 循环