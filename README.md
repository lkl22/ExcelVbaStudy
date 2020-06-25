# Excel VBA编程学习

在 Excel 众多的概念中，VBA 是最重要也是最难学的一部分。如果涉及到数据处理工作，VBA 几乎可以实现任何功能，从简单的数据处理，到批量数据分析，再到与 Office 其他软件交互，甚至与操作系统交互实现复杂的功能，VBA 几乎都可以胜任。


> Visual Basic for Applications（VBA）是 VisualBasic 的一种宏语言，是微软开发出来在其桌面应用程序中执行通用的自动化(OLE)任务的编程语言。主要能用来扩展 Windows 的应用程序功能，特别是Microsoft Office软件。

说简单点，VBA 是运行在 Microsoft Office 软件之上，可以用来编写非软件自带的功能的编程语言。Office 软件提供丰富的功能接口，VBA 可以调用它们，实现自定义的需求。基本上，能用鼠标和键盘能做的事情，VBA 也能做。

正如前文所述，VBA 可以运行在 Office 软件上，包括 Excel、Word、PPT、Outlook 等。VBA 语言在 Office 软件中是通用的，基本语法和用法都相同。但是每一个软件具有自己独有的对象，例如 Excel 有单元格对象，Word 有段落对象，PPT 有幻灯片对象。

#### **VBA 与宏有什么区别？**

在学习 VBA 过程中，经常会出现一个说法，「宏」。简单的说，宏是一段可以运行的 VBA 代码片段，也可以说是一个简称，并没有特别的不同之处。所以学习 VBA 时，不用纠结于两者到底有什么区别，只需要记住一点，宏是使用 VBA 编写的一段代码片段。

#### **学习 Excel VBA 有什么用处？**

前面我们说到，Excel中，VBA几乎可以实现任何功能，从简单的数据处理，到批量数据分析，再到与Office其他软件交互，甚至与操作系统交互实现复杂的功能，VBA几乎都可以胜任。以下是Excel VBA几个典型的用途。

* **节省时间**：只需一次点击就可以重复执行任意数量的操作。例如，现在要新建 20 个工作表，手动操作可能需要一分钟的时间。使用 VBA 只需一秒即可。
* **自动化任务**：只需一次点击就可以按预先设置好的步骤，自动完成操作。例如，插入一个图表并设调整格式，根据其复杂程度，可能需要多达几分钟时间。而使用VBA编写调整步骤，一次点击，几秒内即可完成所有的操作。
* **减少错误**：相比于手动操作出现的错误，只要正确编写 VBA 代码，执行过程中就不会出现错误。例如，从一区域中筛选指定数据，并复制到另外一个位置，手动操作可能会出现漏选的可能。但是使用 VBA，极短的时间内正确无误的完成操作。
* **与其他软件交互**：使用 VBA，可以在 Excel 里创建、更新 Word、PPT 等文件。还可以与系统交互，做到复制、移动、重命名其他文件等操作，无需打开其他文件。

## Excel VBA 基础

* [Excel VBA中的一些基本概念](./docs/basic/basicConcepts.md)（熟悉 VBA 中的基本概念）
* [启用 Excel 开发工具教程](./docs/basic/enableExcelDevTool.md)（准备 VBA 开发工具）
* [Excel VBA 设置宏安全性](./docs/basic/setMacroSecurity.md)（正确设置 VBA 开发安全选项）
* [Excel 保存包含 VBA 代码的工作簿](./docs/basic/saveWorkbookContainVBACode.md)（使用指定类型保存含 VBA 代码的工作簿）
* [使用 VBA 编辑器进行 Excel VBA 开发](./docs/basic/excelVBADevelopmentUsingVBAEditor.md)（熟悉 VBA 开发工具的用法）
* [编写你的第一个 VBA 宏](./docs/basic/writeFirstVBAMacro.md)（基础实战练习）
* [Excel 录制宏并查看宏代码](./docs/basic/recordMacroAndViewMacroCode.md)（写 VBA 代码的技巧）
* [Excel 中如何运行 VBA 代码？](./docs/basic/runVBACodeInExcel.md)（从工作表运行 VBA 代码）

## VBA 变量、类型、运算符

* [VBA 变量基础教程](./docs/variablesTypesOperators/variables.md)（VBA 核心概念）
* [VBA 常量基础教程](./docs/variablesTypesOperators/constant.md)（基础概念）
* [VBA 运算符基础教程](./docs/variablesTypesOperators/operators.md)（加减乘除+高级操作）
* [VBA 数据类型基础教程](./docs/variablesTypesOperators/types.md)（程序更高效、更精准）

## VBA 程序结构
* [VBA 程序结构入门](./docs/programStructure/introduction.md)（认识 VBA 程序骨架）
* [VBA 表达式和语句](docs/programStructure/expressionsAndStatements.md)（最基本的程序单元）
* [VBA 变量的声明和赋值]()（是程序动起来）
* [VBA 程序顺序结构]()（VBA 程序默认执行顺序）
* [VBA 程序选择结构]()（选择性的执行 VBA 代码）
* [VBA 程序循环结构]()（重复执行一段代码）
* [VBA With 结构]()（简化程序书写）
* [VBA GoTo 结构]()（程序之间跳转执行）
* [VBA 注释教程和实例]()（使程序更容易阅读和理解）

## 参考文献

[EXCEL VBA 入门到精通详细教程](https://www.lanrenexcel.com/excel-vba-tutorial/)