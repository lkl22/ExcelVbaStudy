# VBA 中 ByVal 和 ByRef 的基础用法和区别

VBA 中定义过程或函数时，如果需要传递变量，需指定参数的传递类型，包括以下 2 类：

* ByVal：传递参数的值
* ByRef：传递参数的引用

本篇将介绍 2 种方法的用法以及区别。过程和函数传递参数方法基本相同，本篇以过程(Sub)举例说明他们的用法和区别。

* [ByVal和ByRef基础](#ByVal和ByRef基础)
  * [ByVal实例](#ByVal实例)
  * [ByRef实例](#ByRef实例)
* [省略传递类型](#省略传递类型)
* [使用ByVal和ByRef传递对象](#使用ByVal和ByRef传递对象)
* [使用ByVal和ByRef传递数组](#使用ByVal和ByRef传递数组)
* [总结](#总结)

## <a name="ByVal和ByRef基础">ByVal和ByRef基础</a>

在定义过程或函数时，如果需要传递变量，则每个参数需要指定传递类型。传递类型有 2 种，分别是 `ByVal` 和 `ByRef` 。
```vba
'ByVal 传递类型
Sub TestSub1(ByVal msg As String)

End Sub

'ByRef 传递类型
Sub TestSub2(ByRef msg As String)

End Sub
```
针对基础数据类型，例如数字、文本等，两种传递类型的说明和区别如下：

* ByVal：传递变量时，复制一份该变量，传入过程或函数。在过程和函数内部对该变量进行修改，只对该副本有效，对上一级过程（父过程）的变量没有影响。
* ByRef：传递变量时，将该变量的引用地址传入过程或函数。传入引用地址意味着，在过程或函数内部对其修改时，也会影响上一级过程（父过程）中的变量的值。

### <a name="ByVal实例">ByVal实例</a>
通过以下代码测试 ByVal 类型：
```vba
Sub Test()

    Dim msg As String
    msg = "main"
    
    TestSub1 msg
    
    Msgbox msg

End Sub

'ByVal 传递类型
Sub TestSub1(ByVal msg As String)
    msg = "val"
End Sub
```
首先定义一个 `msg` 变量，赋值 `main`，然后调用 `TestSub1` 过程，传入 `msg` 变量，在过程内部对 `msg` 重新赋值 `val`。最后返回上一个过程，显示 `msg` 变量。结果，`msg` 变量的值没有改变。

### <a name="ByRef实例">ByRef实例</a>

通过以下代码测试 ByVal 类型：
```vba
Sub Test()

    Dim msg As String
    msg = "main"
    
    TestSub2 msg
    
    MsgBox msg

End Sub

'ByRef 传递类型
Sub TestSub2(ByRef msg As String)
    msg = "ref"
End Sub
```
首先定义一个 `msg` 变量，赋值 `main`，然后调用 `TestSub2` 过程，传入 `msg` 变量，在过程内部对 `msg` 重新赋值 `ref`。最后返回上一个过程，显示 `msg` 变量。结果，`msg` 变量的值已改变。

## 省略传递类型

默认情况下，当省略传递类型时，默认值是 `ByVal`，因此以下两种写法是等效的。
```vba
'指定 ByVal 传递类型
Sub TestSub1(ByVal msg As String)

End Sub

'省略传递类型
Sub TestSub1(msg As String)

End Sub
```
## <a name="使用ByVal和ByRef传递对象">使用ByVal和ByRef传递对象</a>

在上述介绍中说道，以上机制适用于传递基础类型变量，例如数字、文本、逻辑值等。

使用 ByVal 和 ByRef 传递对象时，情况有些不同。具体用法和不同点将在介绍对象时详细说明。

## <a name="使用ByVal和ByRef传递数组">使用ByVal和ByRef传递数组</a>

过程或函数传递数组时，只能以引用形式传递，即以 ByRef 形式。如果尝试用 ByVal 传递数组，VBA 会提示错误。详细的用法将在介绍数组时详细说明。

## 总结

ByVal 和 ByRef 表示参数传递的类型。针对基础数据类型的变量，ByVal 会创建变量的一个副本，传递给过程或函数，从此之后与父过程的变量没有关系。而 ByRef 方式传递变量的引用，该引用始终会与父过程的变量相连。

因此建议，尽量使用 ByVal 传递类型，防止在子过程或函数中，不小心更改父过程里的变量，导致一些不容易发现的问题。

对象和数组变量的传递，有别于基础类型变量，在相关的教程中详细说明。
