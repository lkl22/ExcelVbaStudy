# VBA 过程(Sub) 入门教程和实例

VBA 中，过程是一切的开始，几乎所有的代码，都会被写在一个或多个过程里。

实际开发中，通常一个过程，建议只完成一个特定的小目标。因此，我们的程序往往会包含多个过程。这就是 VBA 中过程概念存在的一个原因。

程序中使用过程，可以使程序更简洁、清晰，开发中大型项目更易于管理代码。

* [过程基础语法](#过程基础语法)
  * [无参数过程](#无参数过程)
  * [有参数过程](#有参数过程)
* [调用子过程](#调用子过程)
  * [直接调用](#直接调用)
  * [使用关键词Call调用](#使用关键词Call调用)
* [提前退出过程](#提前退出过程)
  * [Exit Sub 语句](#ExitSub语句)
  * [End 语句](#End语句)
* [总结](#总结)

## 过程基础语法

VBA 过程以 `Sub` 语句开始，以 `End Sub` 语句结束，包含一个或多个语句，完成一个特定的目标。

### 无参数过程

无参数的 VBA 过程的基本语法如下：
```vba
Sub [过程名]()
    语句1
    语句2
    ...
    语句n
End Sub
```
可以看到，过程以 `Sub` 语句开始，以 `End Sub` 语句结束，并且具备一个名称，名称后有括号 `()`。

我们看一个简单的例子。
```vba
Sub SayHello()
    Msgbox "Hello World"
End Sub
```
上述就是一个简单的过程，过程名是 `SayHello`。这个过程只包含一个语句，运行时，弹出对话框显示 `Hello World`。

### 有参数过程

过程还可以接受一个或多个参数，参数可以是常量、变量、表达式，并且每个参数指定其名称。在过程的语句中，接受的参数，以名称指定方式被使用。

接受参数的过程基本语法如下：
```vba
Sub [过程名]([变量名1] As [数据类型1],...[变量名n] As [数据类型n])
    语句1
    语句2
    ...
    语句3
End Sub
```
与无参数过程相比，有参数过程在过程名后的括号 `()` 中，包含一个或多个参数。参数的写法与声明变量语句类似，不同点是在这里不用写 `Dim`。
```vba
[变量名1] As [数据类型1]
```
我们看一个例子。
```vba
'声明一个过程
Sub SayHello(name As String)
    Msgbox "Hello" & name
End Sub

'在另一个过程，调用上述过程，调用时，提供一个实际的 name 参数
Sub MyCode()
    SayHello "World 2"
End Sub
```
我们在运行 `MyCode` 过程时，提供了 `name` 变量，即 `World 2` ，运行时弹出对话框显示 `Hello World 2`。

## 调用子过程

在程序开发中，把代码拆分成多个子过程和函数，可以使项目更容易管理、测试和运行，VBA 中也不例外。

实际开发中，项目通常具备一个主入口过程，或称为父过程。父过程通过调用多个子过程和函数，完成一系列复杂的操作。其中子过程和函数一般只负责一个操作或动作。

下面看一个简单的例子。
```vba
'主入口
Sub Main()
    Dim name As String
    Dim title As String
    
    name = "Zhang san"
    title = "CEO"
    
    WriteInfo name & "," & title
End Sub

'子过程，在工作表A1单元格填写信息
Sub WriteInfo(info As String)
    Range("A1") = info
End Sub
```
以上的例子中，`Main` 过程是一个主入口（父过程），程序从此处开始执行，先是给 `name` 和 `title` 变量赋值，最后调用 `WriteInfo` 子过程，将两个信息合并后写到工作表上的 A1 单元格。

接下来介绍调用子过程和函数的基本语法。

调用子过程有两种方法，**直接调用**和**使用 Call 关键词调用**。两种方法对子过程的参数有不同的要求。

### 直接调用

直接调用，直接写过程名，即可调用过程。
```vba
Sub Main()
    MySub
End Sub

Sub MySub()
    '代码
End Sub
```
如果子过程需要输入参数，多个参数只需用逗号（,）分开即可。
```vba
Sub Main()
    MySub 2019,"年"
End Sub

Sub MySub(val1 As Integer, val2 As String)
    '代码
End Sub
```

### <a name="使用关键词Call调用">使用关键词Call调用</a>

使用 Call 关键词调用时，Call 后接过程名。
```vba
Sub Main()
    Call MySub
End Sub

Sub MySub()
    '代码
End Sub
```
如果子过程需要输入参数，则需要将参数放在括号内。
```vba
Sub Main()
    Call MySub(2019,"年")
End Sub

Sub MySub(val1 As Integer, val2 As String)
    '代码
End Sub
```
> 注：程序角度看，调用过程时，不需要使用 Call 关键字，因此不建议此种方法。

## 提前退出过程

正常情况下，VBA 过程以 `Sub` 语句开始，以 `End Sub` 语句结束。但有时根据实际情况，可能需要提前结束并退出过程。VBA 提供 2 种提前退出过程的方法，`Exit Sub` 和 `End` 方法。

### <a name="ExitSub语句">Exit Sub 语句</a>

在一个过程中，当程序运行到 `Exit Sub` 语句时，立即结束当前过程，提前退出。
```vba
Sub Main()
    Call MySub
    Msgbox "父过程"
End Sub

Sub MySub()
    Exit Sub
    Msgbox "子过程"
End Sub

'运行 Main 过程，返回结果：
=> "父过程"
```
在以上例子中，`Main` 过程调用 `MySub` 子过程，遇到 `Exit Sub` 语句，立即退出子过程，回到父过程 `Main` ，继续运行余下的语句。

> **这里需要注意的是，Exit Sub 语句只作用于当前过程，不影响调用它的父过程。**

### <a name="End语句">End 语句</a>

在一个过程，当程序运行到 `End` 语句时，立即结束当前运行的所有 VBA 过程。
```vba
Sub Main()
    Call MySub
    Msgbox "父过程"
End Sub

Sub MySub()
	End
    Msgbox "子过程"
End Sub

'运行 Main 过程，返回结果：
=> 无返回结果
```
在以上例子中，`Main` 过程调用 `MySub` 子过程，遇到 `End` 语句时，立即结束当前运行的所有过程，包括父过程 `Main`。

在实际开发中，应谨慎使用 `End` 结束语句。`End` 语句的效果类似于电脑的强制关机命令，立即结束所有程序，不会保存任何值，于 VBA 有以下效果：

* 程序中对象的各类事件不会被触发；
* 任何在运行的 VBA 程序都会停止；
* 对象引用都会失效；
* 任何打开的窗体都被关闭。

## 总结

过程是 VBA 的一个核心概念，几乎所有的代码会写在一个或多个过程里。过程可以接受 0 个或多个参数，参数可以在过程或函数里使用。在过程中可以调用其他子过程，把复杂的代码分成若干个过程，使代码易于管理和编写。最后过程可以提前结束，做到不需要运行所有的语句就退出过程。



