# VBA 过程或函数作用域

* [1.模块作用域](#模块作用域)
* [2.工程作用域](#工程作用域)
* [3.全局作用域](#全局作用域)

## <a name="模块作用域">1.模块作用域</a>

在模块中，使用 `Private` 关键词声明的过程或函数，具备模块作用域，只能在当前模块中使用。
```vba
Private Sub Test()

End Sub
```

## <a name="工程作用域">2.工程作用域</a>

在模块中，顶部声明 `Option Private Module` 修饰语句，并且直接声明或使用 `Public` 关键词声明的过程或函数，具备工程作用域，在当前工程的所有模块中使用。
```vba
Option Private Module

Sub Test1()

End Sub

Public Sub Test2()

End Sub
```
以上例子中，Test1 过程和 Test2 过程均具备工程作用域。**由于直接声明和使用关键词 Public 是等效的，因此可以省略 Public 关键词**。

## <a name="全局作用域">3.全局作用域</a>

在模块中，直接声明或使用 `Public` 关键词声明的过程或函数，具备全局作用域。例如，
```vba
Sub Test1()

End Sub

Public Sub Test2()

End Sub
```
以上例子中，`Test1` 过程和 `Test2` 过程均具备全局作用域，可以在打开的任何一个工作簿中使用。

此外，它们还能直接在工作簿宏列表中执行。

