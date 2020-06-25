# VBA 过程和函数：传递参数教程和实例

* [带参数的子过程定义方法](#带参数的子过程定义方法)
* [调用带参数的子过程](#调用带参数的子过程)
* [可选参数的用法](#可选参数的用法)
  * [可选参数语法](#可选参数语法)
  * [设置可选参数的默认值](#设置可选参数的默认值)
  * [可选参数的位置](#可选参数的位置)
* [总结](#总结)

## 带参数的子过程定义方法

子过程可以接受一个或多个参数，参数可以是[常量]、[变量]、[表达式]，并且每个参数指定其名称和数据类型。

看实际的例子，以下代码定义了带两个参数的一个过程，过程名是 `CustomLog` ，参数分别是 `num` 和 `base`。此过程的用途是计算任意底数的对数，`num` 是计算对数的值，`base` 是底数。
```vba
'声明一个带参数的子过程
Sub CustomLog(num As Double, base As Integer)
    Debug.Print Log(num) / Log(base)
End Sub
```
子过程按照这种方法定义后，调用时，VBA 会提示需要提供什么参数以及参数类型。

## 调用带参数的子过程

调用带参数的过程，只需将参数**按定义顺序**书写即可，多个参数使用逗号分开。

以上述过程为例，在一个主过程调用 CustomLog 子过程。
```vba
'主入口
Sub Main()
    CustomLog 100, 10
End Sub
```
除了按顺序书写参数外，也可以按任意顺序书写参数，但是这时需要给出`参数名`。带参数名的传递参数语法如下：
```vba
[参数名]:=[实际参数值]
```
参数名后写冒号等号(:=)，再写需传递的参数值。看实际的例子，以下三种方式是等效的。
```vba
'主入口
Sub Main()
    CustomLog 100, 10 '方式一
    CustomLog num:=100, base:=10 '方式二
    CustomLog base:=10, num:=100 '方式三
End Sub
``` 

## 可选参数的用法

实际开发中，有时子过程的参数可能不是必须的，我们希望根据参数有无情况，执行不同的操作。针对这种情况，VBA 提供了可选参数机制。

### 可选参数语法

可选参数在定义子过程时需要指定，方法是在参数名前添加 `Optional` 关键词。
```vba
Optional [参数名] As [数据类型]
```
还是以 CustomLog 子过程为例，我们把底数 base 设为可选参数。
```vba
'声明一个带可选参数的子过程
Sub CustomLog(num As Double, Optional base As Integer)
    '子过程代码
End Sub
```
调用时，VBA 会提示可选参数，参数放置在中括号中。

### 设置可选参数的默认值

可选参数定以后，如果在子过程中使用，需要判断参数有无提供。否则未提供而直接使用时，程序会出错。

针对这种情况，VBA 提供了默认值机制，即可选参数未提供时，使用预算设置好的默认值。

可选参数默认值，在定义过程时就设置，语法如下：
```vba
Optional [参数名] As [数据类型] = [默认值]
```
还是以 CustomLog 子过程为例，我们把底数 base 设为可选参数，并且默认值设为 10。
```vba
'声明一个带可选参数的子过程
Sub CustomLog(num As Double, Optional base As Integer = 10)
    Debug.Print Log(num) / Log(base)
End Sub
```
调用时，如果提供了 base 底数，则以提供的底数计算；如果未提供 base 底数，则以默认值 10 计算。
```vba
'主入口
Sub Main()

    CustomLog 100, 100 '返回 1
    CustomLog 100 '返回 2

End Sub
```

### 可选参数的位置

当子过程有多个参数时，其中的可选参数需写在参数列表的末尾，否则 VBA 提示错误。

可选参数错误顺序：
```vba
'声明一个带可选参数的子过程
Sub CustomLog(Optional num As Double, base As Integer = 10)
    Debug.Print Log(num) / Log(base)
End Sub
```

可选参数的正确顺序：
```vba
'声明一个带可选参数的子过程
Sub CustomLog(num As Double, Optional base As Integer = 10)
    Debug.Print Log(num) / Log(base)
End Sub
```

## 总结

VBA 过程和函数均可以接受一个或多个参数。当调用它们时，需要注意传入的参数的书写顺序：不写参数名时，按照定义的顺序传递；写参数名时，对书写顺序没有要求。此外，过程和函数可以设置某一个参数是可选的，类似 VLOOKUP 函数的第四个参数，是否精确查找。当设置成可选时，还可以指定可选参数的默认值。

