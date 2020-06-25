# VBA 函数(Function)入门教程和实例

* [VBA函数基础语法](#VBA函数基础语法)
  * [无参数函数](#无参数函数)
  * [有参数函数](#有参数函数)
* [调用函数](#调用函数)
* [提前退出函数](#提前退出函数)
  * [Exit Function 语句](#ExitFunction语句)
  * [End语句](#End语句)
* [总结](#总结)

## VBA函数基础语法

VBA 函数与 VBA 过程很相似，除了使用的关键词外，主要区别是，函数可以返回值。

### 无参数函数

无参数 VBA 函数的基本语法如下：
```vba
Function [函数名]() As [返回值类型]
    语句1
    语句2
    ...
    语句n
    [函数名] = [返回值]
End Function
```
可以看到，函数使用 `Function` 和 `End Function` 语句作为函数的开始和结束。

函数包含的语句中，相比过程，可以看到多一个 [函数名] = [返回值] 语句，这是函数的返回值语句。函数名后制定该函数返回值的类型，语法与声明变量类似。

看一个实际的例子。
```vba
'声明函数，该函数随机返回 true 或 false。函数需指定返回值类型。
Function RandomLogic() As Boolean
    RandomLogic = Rnd() > 0.5
End Function
```
该函数的名称是 `RandomLogic`，返回值类型时 `Boolean` 类型，运行调用后，随机返回一个 `true` 或 `false` 值。实现方法是，使用 VBA 内置函数 `Rnd`（随机产生0-1的数字），随机数与0.5对比大小，产生 true 或 false 值，并把值赋值给函数名。

### 有参数函数

函数与过程一样，也可以接收参数，其语法与过程相同。
```vba
Function [函数名]([变量名1] As [数据类型1],...[变量名n] As [数据类型n]) As [返回值类型]
    语句1
    语句2
    ...
    语句3
    [函数名] = [返回值]
End Function
```
同样，函数接收的参数，在函数主体中使用。

我们看一个实际的例子。
```vba
Function Add2Number(num1 As Double, num2 As Double) As Double
    Add2Number = num1 + num2
End Function
```
上述函数接受2个 `Double` 类型的数字作为参数，两者相加，返回和，其类型也是 `Double` 类型。

## 调用函数

函数与子过程的区别是，函数可以返回值。如果一个函数不返回值，它与子过程并无区别，其中调用方式与子过程相同。

调用有返回值的函数时，一般有两种情形：

* 一是，使用一个变量存储函数返回的值
* 二是，函数返回的值参与其他计算

两种情形调用函数方式相同，无参数函数直接书写，**有参数函数将参数放在括号内。**
```vba
Sub Main()
    '使用变量存储函数返回的值
    Dim result As Double
    result = Add(12, 345)
    
    '函数返回值继续参与计算
    Dim result As Double
    result = RandNum + Add(12, 345)
End Sub

'函数：返回一个随机值
Function RandNum()
    RandNum = Rnd * 100
End Function
'函数：返回两数的和
Function Add(num1 As Double, num2 As Double) As Double
    Add = num1 + num2
End Function
``` 

## 提前退出函数

正常情况下，函数使用 `Function` 和 `End Function` 语句作为函数的开始和结束。但有时根据实际情况，可能需要提前结束并退出函数。VBA 提供 2 种提前退出过程的方法，`Exit Function` 和 `End` 方法。

### <a name="ExitFunction语句">Exit Function 语句</a>

在一个函数中，当程序运行到 `Exit Function` 语句时，立即结束当前函数，提前退出。

这里需要注意的是，`Exit Function` 语句只作用于当前过程，不影响调用它的父过程或函数。

### <a name="End语句">End 语句</a>

在一个函数，当程序运行到 `End` 语句时，**立即结束当前运行的所有 VBA 过程和函数。**

在实际开发中，应谨慎使用 `End` 结束语句。`End` 语句的效果类似于电脑的强制关机命令，立即结束所有程序，不会保存任何值，于 VBA 有以下效果：

* 程序中对象的各类事件不会被触发；
* 任何在运行的 VBA 程序都会停止；
* 对象引用都会失效；
* 任何打开的窗体都被关闭。

## 总结

函数与过程类似，大部分用法相同，主要区别是函数可以返回一个值，而过程不可以。两者均可以接受0个或多个参数，参数可以在过程或函数里使用。调用函数时，参数需要放置在括号内部，接函数名后。

