# VBA 变量作用域

VBA 中，变量的作用域决定变量在哪里能被获取和使用。根据变量的声明位置和声明方式，变量的作用域有以下 4 种：

* 过程作用域
* 模块作用域
* 工程作用域
* 全局作用域

下面将逐一介绍每一种作用域对应的变量的声明方法以及使用方法。

* [1.过程作用域](#过程作用域)
* [2.模块作用域](#模块作用域)
* [3.工程作用域](#工程作用域)
* [4.全局作用域](#全局作用域)
* [作用域冲突](#作用域冲突)

## <a name="过程作用域">1.过程作用域</a>

在过程或函数内部声明的变量，只有在当前过程或函数内被使用。例如：
```vba
Sub Test()

    Dim name As String
    Dim age As Integer
    
    name = "张三"
    age = 35

End Sub
```
以上代码中，变量 `name` 和 `age` 在 `Test` 过程声明，因此它们只能在该过程中内使用，包括赋值和读取。如果尝试在外部和其他过程中直接使用它们，**VBA 会提示变量未定义错误**。

## <a name="模块作用域">2.模块作用域</a>

一个模块中，在任何一个过程和函数外面，使用关键词 `Private` 或 `Dim` 声明的变量，称之为模块变量，其作用域是当前模块。例如，
```vba
Dim guest As String

Sub Test()

    Dim message As String
    
    guest = "张三"
    message = "你好"
    
    MsgBox message & "！ " & guest

End Sub
```
以上代码中，变量 `guest` 是在过程 `Test` 外面，使用 `Dim` 关键词声明的，称之为模块变量。**模块变量的作用域是当前模块，在模块里面任何过程和函数内均可以使用。**

如前文所述，使用关键词 `Private` 或 `Dim` 声明的变量，都是模块变量，因此以下两种声明方式是等效的。
```vba
Dim guest As String
Private guest As String
```

## <a name="工程作用域">3.工程作用域</a>

Excel VBA 中，一个 Excel 工作簿是一个 VBA 工程。与之对应，**工程作用域表示变量在当前工程中的模块、Excel 对象、用户窗体、类模块中均可以被使用。**

工程级别变量，在所在模块顶部声明 `Option Private Module` 修饰语句前提下，在过程或函数外面，使用关键词 `Public` 声明的变量，其作用域是当前工程。例如，
```vba
Option Private Module

Public guest As String

Sub Test()

    Dim message As String
    
    guest = "张三"
    message = "你好"
    
    MsgBox message & "！ " & guest

End Sub
```
以上例子中，变量 `guest` 是使用 `Public` 关键词声明，是工程级别变量。它在当前工程中其他的模块中也能被使用。

## <a name="全局作用域">4.全局作用域</a>

**全局作用域表示，全局变量在打开的任何一个工作簿都可以被使用**。全局变量的声明方式与工程变量相似，不同点是**不使用模块顶部的 `Option Private Module` 修饰语句**。

## 作用域冲突

当相同名称的变量，多次以不同的作用域声明时，出现作用域冲突。这种情况，**VBA 会自动以就近原则使用变量，即优先使用最近定义的变量**。例如，
```vba
Dim name As String

Sub Test()

    Dim name As String
    
    name = "李四"

End Sub
```
以上例子中，两次声明 name 变量，分别是模块变量和过程变量。根据就近原则，在过程内部使用时，将使用过程变量。