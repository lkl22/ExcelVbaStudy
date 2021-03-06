# Excel 中如何运行 VBA 代码？

编写 VBA 的最终目的是在 Excel 中运行它，得到特定的结果。所以，写完一段 VBA 代码后，下一步骤就是运行它。

需要指出的是，本篇所指的「运行」指的是，在 Excel 界面中运行，而非在 VBA 编辑器中运行。

一般来说，运行 VBA 有以下 3 种方法：
1. 从「开发工具」选项卡运行
2. 通过给「形状」指定宏的方式运行
3. 通过给「按钮」指定宏的方式运行

## 示例代码

为了演示运行 VBA 的多种方法，现准备了一个包含一个 VBA 宏的工作簿。包含的 VBA 宏是「[编写你的第一个 VBA 宏](./writeFirstVBAMacro.md)」教程中编写的代码，具体如下：
```vba
Sub MyCode()

    MsgBox "Hello World"
    
End Sub
```
此过程运行后弹出对话框，显示 “Hello World” 信息。

## 1 从「开发工具」选项卡运行

Excel 开发工具选项卡提供了一个查看工作簿包含的所有宏并且运行的功能。需要运行某一宏时，打开宏列表，选择想要运行的宏，点击「运行」即可。

具体步骤如下：

1. 找到「宏」按钮。
2. 点击宏按钮，会弹出工作簿包含的所有宏的列表。
3. 选择想要的宏，点击右侧「执行」按钮，运行宏。

## 2 通过给「形状」指定宏的方式运行

Excel 中的形状，可以为其指定宏，当鼠标点击时，宏自动运行。优点是，可以给任何形状指定宏。

具体方法如下：

1. 插入一个形状。
2. 右键形状，选择「指定宏」。
3. 在弹出的宏列表中，选择一个宏，后点击确定完成。

## 3 通过给「按钮」指定宏的方式运行

除了形状，Excel 提供一个内置的按钮，优点是可以自定义，并且有点击效果。

具体步骤如下：

1. 找到开发工具→插入命令。
2. 点击插入命令，从列表中选择「表单控件→按钮」。
3. 点击按钮，这时出现一个宏列表。其中选择想要指定的宏，点击确定，完成指定宏。

## 总结
以上几个方法适合在当前工作簿运行自己包含的代码。针对个人宏工作簿和 Excel 插件，还有几个更为快捷的运行 VBA 的方法，这在以后的教程介绍。