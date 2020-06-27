# VBA 数组

一组顺序索引的元素，这些元素具有相同的内在数据类型。 数组的每个元素具有唯一的识别索引号。 对数组的一个元素进行的更改不会影响其他元素。

* [声明数组](#声明数组)
  * [声明固定数组](#声明固定数组)
  * [声明动态数组](#声明动态数组)
* [使用数组](#使用数组)
  * [更改下限](#更改下限)
  * [在数组中存储Variant values](#在数组中存储Variantvalues)
  * [使用多维数组](#使用多维数组)
* [了解参数数组](#了解参数数组)
* [常用库函数](#常用库函数)
  * [Array 函数](#Array函数)
  * [数组的合并 Join函数](#数组的合并join)
  * [数组的拆分 Split函数](#数组的拆分Split函数)
  * [IsArray 函数](#IsArray函数)
  * [Filter 函数](#Filter函数)
  * [WorksheetFunction.Transpose method (Excel)](#WorksheetFunctionTransposemethod)

## 声明数组

数组的声明方式与其他变量相同，即，使用 `Dim`、`Static`、`Private` 或 `Public` 语句声明。 标准变量（不是数组的变量）和数组变量之间的区别在于您通常必须指定数组的大小。 指定了大小的数组为固定大小的数组。 程序运行时大小可以更改的数组是动态数组。

> 数组是以 0 还是以 1 开始编制索引取决于` Option Base` 语句的设置。 如果 `Option Base 1 `未指定，则所有数组索引都从零开始。

### 声明固定数组

在下面的代码行中，一个固定大小的数组声明为具有 11 行和 11 列的 Integer 数组：
```vba
Dim MyArray(10, 10) As Integer 
```
第一个参数代表行；第二个参数代表列。

> 与其他任何变量声明一样，除非您为数组指定数据类型，否则声明的数组中元素的数据类型为`Variant`。 数组的每个**数值**`Variant`元素都使用16个字节。 每个**字符串**`Variant`元素使用22个字节。 要编写尽可能紧凑的代码，请显式声明数组为`Variant`以外的数据类型。

以下代码行比较了几个数组的大小
```vba
' Integer array uses 22 bytes (11 elements * 2 bytes). 
ReDim MyIntegerArray(10) As Integer 
 
' Double-precision array uses 88 bytes (11 elements * 8 bytes). 
ReDim MyDoubleArray(10) As Double 
 
' Variant array uses at least 176 bytes (11 elements * 16 bytes). 
ReDim MyVariantArray(10) 
 
' Integer array uses 100 * 100 * 2 bytes (20,000 bytes). 
ReDim MyIntegerArray (99, 99) As Integer 
 
' Double-precision array uses 100 * 100 * 8 bytes (80,000 bytes). 
ReDim MyDoubleArray (99, 99) As Double 
 
' Variant array uses at least 160,000 bytes (100 * 100 * 16 bytes). 
ReDim MyVariantArray(99, 99) 
```
数组的最大大小因操作系统和可用内存而异。 使用超过系统中可用 RAM 大小的数组会导致性能下降，因为必须从磁盘读取数据和向其中写入数据。

### 声明动态数组

通过声明动态数据，可以在运行代码时调整数组的大小。 使用 `Static` 、 `Dim` 、 `Private` 或 `Public` 语句声明一个数组，将括号内留空，如下面的示例所示。
```vba
Dim sngArray() As Single 
```

> 您可以使用`ReDim`语句在过程中隐式声明一个数组。 使用`ReDim`语句时，请注意不要拼写错误数组的名称。 即使模块中包含`Option Explicit`语句，也会创建第二个数组。

在数组范围内的过程中，使用 `ReDim` 语句可更改**维度数**、定义**元素数**，以及定义每个维度的**上限**和**下限**。 可以根据需要使用 `ReDim` 语句更改动态数组。 但是，每次这么做时，数组中的现有值都会丢失。 使用 `ReDim Preserve` 可在保留数组中现有值的情况下扩展数组。

例如，下面的语句将数组扩增了 10 个元素，而未丢失原始元素的当前值。
```vba
ReDim Preserve varArray(UBound(varArray) + 10) 
```

> 在对动态数组使用 `Preserve` 关键字时，仅能更改最后一个维度的上限，而不能更改维度数。

## 使用数组

您可以声明一个数组以处理一组相同**数据类型**的值。 数组是具有多个可存储值的隔离舱的单个**变量**，而典型的变量只有一个存储隔离舱，其中只能存储一个值。 您可以在需要引用数组中包含的所有值时将数组作为整体引用，也可以引用其中的单个元素。

例如，若要存储一年中每天的日常开支，您可以声明一个具有 365 个元素的数组变量，而不是声明 365 个变量。 数组中的每个元素包含一个值。 以下语句声明具有 365 个元素的数组变量。 默认情况下，数组的索引从零开始，因此该数组的上限是 364 而不是 365。
```vba
Dim curExpense(364) As Currency 
```
若要设置单个元素的值，您可以指定该元素的索引。 以下示例向该数组中的每个元素均分配一个初始值 20。
```vba
Sub FillArray() 
    Dim curExpense(364) As Currency 
    Dim intI As Integer 
    For intI = 0 to 364 
        curExpense(intI) = 20 
    Next 
End Sub
```
### 更改下限

您可以使用模块顶部的 **Option Base** 语句将第一个元素的默认索引从0更改为1。 在下面的示例中, `Option Base`语句更改第一个元素的索引, `Dim` 语句声明包含365个元素的数组变量。
```vba
Option Base 1 
Dim curExpense(365) As Currency 
```
也可以通过使用 To 子句明确设置数组的下限，如以下示例所示。
```vba
Dim curExpense(1 To 365) As Currency 
Dim strWeekday(7 To 13) As String 
```

### <a name="在数组中存储Variantvalues">在数组中存储Variant values<a/>

有两种方法可以创建 **Variant** values 的数组。 一种是声明 `Variant` 数据类型的数组，如以下示例所示：
```vba
Dim varData(3) As Variant 
varData(0) = "Claudia Bendel" 
varData(1) = "4242 Maple Blvd" 
varData(2) = 38 
varData(3) = Format("06-09-1952", "General Date") 
```
另一种方法是将 `Array` 函数返回的数组分配给 `Variant` 变量，如以下示例所示：
```vba
Dim varData As Variant 
varData = Array("Ron Bendel", "4242 Maple Blvd", 38, _ 
    Format("06-09-1952", "General Date")) 
```
无论使用哪种方法创建数组，均可通过索引识别 `Variant` values的数组中的元素。 例如，以下语句可添加到上述两个示例中的任意一个示例中。
```vba
MsgBox "Data for " & varData(0) & " has been recorded." 
```

### 使用多维数组

在 Visual Basic 中，您可以声明最多包含 60 个维度的数组。 例如，以下语句声明了一个二维、5*10 的数组。
```vba
Dim sngMulti(1 To 5, 1 To 10) As Single 
```
如果将数组看作矩阵，则第一个参数表示行，第二个参数表示列。

使用嵌套 `For...Next` 语句来处理多维数组。 以下过程使用 Single 值填充一个二维度组。
```vba
Sub FillArrayMulti() 
    Dim intI As Integer, intJ As Integer 
    Dim sngMulti(1 To 5, 1 To 10) As Single 
    
    ' Fill array with values. 
    For intI = 1 To 5 
        For intJ = 1 To 10 
            sngMulti(intI, intJ) = intI * intJ 
            Debug.Print sngMulti(intI, intJ) 
        Next intJ 
    Next intI 
End Sub
```

## 了解参数数组

参数数组可用于将参数数组传递给**过程**。 **定义过程时，您不必知道数组中的元素数量**。

使用 `ParamArray` 关键字可表示参数数组。 该数组必须声明为类型为 `Variant` 的数组，并且**它必须是过程定义中的最后一个参数**。

下面的示例演示如何使用参数数组定义过程。
```vba
Sub AnyNumberArgs(strName As String, ParamArray intScores() As Variant) 
    Dim intI As Integer 
    
    Debug.Print strName; " Scores" 
    ' Use UBound function to determine upper limit of array. 
    For intI = 0 To UBound(intScores()) 
        Debug.Print " "; intScores(intI) 
    Next intI 
End Sub
```
下面的示例演示如何调用此过程。
```vba
AnyNumberArgs "Jamie", 10, 26, 32, 15, 22, 24, 16 
 
AnyNumberArgs "Kelly", "High", "Low", "Average", "High" 
```

## 常用库函数

### <a name="Array函数">Array 函数<a/>

返回一个包含数组的 `Variant`。

#### 语法
```vba
Array(arglist)
```
必需的 `arglist` 参数是以逗号分隔的值的列表，这些值将分配给包含在 `Variant` 中的数组的元素。 如果没有指定任何参数，则将创建零长度的数组。

#### Remarks

用来引用数组元素的符号由变量名和括号组成，括号中包含指示所需元素的索引号。

```vba
Dim A As Variant, B As Long, i As Long
A = Array(10, 20, 30)  ' A is a three element list by default indexed 0 to 2
B = A(2)               ' B is now 30
ReDim Preserve A(4)    ' Extend A's length to five elements
A(4) = 40              ' Set the fifth element's value
For i = LBound(A) To UBound(A)
    Debug.Print "A(" & i & ") = " & A(i)
Next i
```
使用 `Array` 函数创建的数组的下限由通过 `Option Base` 语句指定的下限确定，除非使用类型库的名称（如 `VBA.Array`）限定 `Array`。 如果使用类型库名称进行限定，则 `Array` 不受 `Option Base` 的影响。

> 未声明为数组的 `Variant` 仍可包含一个数组。 `Variant` 变量可以包含任何类型的数组（固定长度的字符串和用户定义类型除外）。 虽然从概念上说，包含数组的 `Variant` 与其元素属于类型 `Variant` 的数组不同，但将按照相同的方式访问数组元素。

#### 示例

此示例使用 `Array` 函数返回包含 `Variant` 的数组。
```vba
Dim MyWeek, MyDay
MyWeek = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
' Return values assume lower bound set to 1 (using Option Base
' statement).
MyDay = MyWeek(2)    ' MyDay contains "Tue".
MyDay = MyWeek(4)    ' MyDay contains "Thu".
```

### <a name="数组的合并join">数组的合并 Join函数<a/>

返回通过联接数组中包含的大量子字符串创建的字符串。

#### 语法
```vba
Join(sourcearray, [ delimiter ])
```
Join 函数语法包含以下命名参数：

|参数	|说明
|---|---
|sourcearray	|必需。 一维数组，包含要联接的子字符串。
|分隔符	|可选。 用于分隔返回字符串中子字符串的字符串。 如果省略，将使用空格 ("")。 如果 delimiter 是一个零长度字符串 ("")，将连接列表中的所有项，而不使用分隔符。

### <a name="数组的拆分Split函数">数组的拆分 Split函数<a/>

返回包含指定数目的子字符串的从零开始的一维数组。

#### 语法
```vba
Split(expression, [ delimiter, [ limit, [ compare ]]])
```
Split 函数语法具有以下命名参数：

|命名参数	|说明
|---|---
|expression	|必需。 包含子字符串和分隔符的字符串表达式。 如果 expression 是零长度字符串 ("")，<br>则 Split 返回空数组，即不包括任何元素和数据的数组。
|delimiter 分隔符	|可选。 用于标识子字符串限制的 String 字符。 如果省略，则假定空格符 (" ") 为分隔符。 如果 delimiter 是零长度字符串，则返回包含完整 expression 字符串的只含单一元素的数组。
|limit	|可选。 要返回的子字符串的数目;-1 表示返回所有子字符串。
|compare	|可选。 指示计算子字符串时使用的比较类型的数值。 请参阅“Settings”部分以了解各个值。

#### Settings

compare 参数可以包含以下值：

|常量	|值	|说明
|---|---|---
|vbUseCompareOption	|-1	|使用 [Option Compare](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/option-compare-statement) 语句的设置来执行比较。
|vbBinaryCompare	|0	|执行二进制比较。
|vbTextCompare	|1	|执行文本比较。
|vbDatabaseCompare	|2	|仅用于 Microsoft Access。 根据数据库中的信息执行比较。

#### 示例

本示例演示如何使用Split函数。
```vba
Dim strFull As String
Dim arrSplitStrings1() As Variant
Dim arrSplitStrings2() As Variant
Dim strSingleString1 As String
Dim strSingleString2 As String
Dim strSingleString3 As String
Dim i As Long

strFull = "Some - Old - Hags - Can - Always - Hide - Their - Old - Age"    ' String that will be used. 

arrSplitStrings1 = Split(strFull, "-")      ' arrSplitStrings1 will be an array from 0 To 8. 
                                            ' arrSplitStrings1(0) = "Some " and arrSplitStrings1(1) = " Old ". 
                                            ' The delimiter did not include spaces, so the spaces in strFull will be included in the returned array values. 

arrSplitStrings2 = Split(strFull, " - ")    ' arrSplitStrings2 will be an array from 0 To 8. 
                                            ' arrSplitStrings2(0) = "Some" and arrSplitStrings2(1) = "Old". 
                                            ' The delimiter includes the spaces, so the spaces will not be included in the returned array values. 

'Multiple examples of how to return the value "Can" (array position 3). 

strSingleString1 = arrSplitStrings2(3)      ' strSingleString1 = "Can". 

strSingleString2 = Split(strFull, " - ")(3) ' strSingleString2 = "Can".
                                            ' This syntax can be used if the entire array is not needed, and the position in the returned array for the desired value is known. 

For i = LBound(arrSplitStrings2, 1) To UBound(arrSplitStrings2, 1)
    If InStr(1, arrSplitStrings2(i), "Can", vbTextCompare) > 0 Then
        strSingleString3 = arrSplitStrings2(i)
        Exit For
    End If 
Next i
```

### <a name="IsArray函数">IsArray 函数</a>

返回指定变量是否是数组的 Boolean 值。

#### 语法
```vba
IsArray(varname)
```
必需的 `varname` 参数是指定变量的标识符。

#### Remarks
“`IsArray`”在变量是数组时返回“`True`”；否则返回“`False`”。 “`IsArray`”对包含数组的变量尤其有用。

#### 示例

此示例使用 `IsArray` 函数检查变量是否是数组。
```vba
Dim MyArray(1 To 5) As Integer, YourArray, MyCheck    ' Declare array variables.
YourArray = Array(1, 2, 3)    ' Use Array function.
MyCheck = IsArray(MyArray)    ' Returns True.
MyCheck = IsArray(YourArray)    ' Returns True.
```

### <a name="Filter函数">Filter 函数</a>

返回一个从零开始的数组, 该数组包含基于指定的筛选条件的字符串数组的子集。

#### 语法
```vba
Filter(sourcearray, match, [ include, [ compare ]])
```
Filter函数语法包含以下命名参数:

|命名参数	|说明
|---|---
|sourcearray	|必需。 要搜索的字符串的一维度组。
|match	|必需。 要搜索的字符串。
|include	|可选。 Boolean 值，指示是否返回包括或排除 match 的子字符串。 <br>如果 include 为 True，则 Filter 返回包含 match 作为子字符串的数组的子集。 <br>如果 include 为 False，则 Filter 返回不包含 match 作为子字符串的数组的子集。
|compare	|可选。 指示要使用的字符串比较类型的数值。 

Filter 函数返回的数组只包含足以构成匹配项数的元素。

### <a name="WorksheetFunctionTransposemethod">WorksheetFunction.Transpose method (Excel)<a/>

返回转置单元格区域，即将一行单元格区域转置成一列单元格区域，反之亦然。 **换位**必须以数组公式的形式输入到具有相同行数和列数的区域中, 因为数组具有列和行。 使用 "**转置**" 可移动工作表上数组的垂直和水平方向。

#### 语法
```vba
expression.Transpose (Arg1)
```

expression一个代表 `WorksheetFunction` 对象的变量。

#### Parameters

|Name	|Required/Optional	|Data type	|Description
|---|---|---|---
|Arg1	|Required	|Variant	|Array - 要进行转置的工作表中的单元格数组或区域。 <br>所谓数组的转置就是，将数组的第一行作为新数组的第一列，将数组的第二行作为新数组的第二列，依此类推。

#### 返回值

`Variant`

