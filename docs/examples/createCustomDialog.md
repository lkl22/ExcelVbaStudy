# VBA 创建自定义对话框

请遵循以下过程创建自定义对话框：
* [1.创建用户窗体](#创建用户窗体)
* [2.操作方法：向用户窗体中添加控件](#向用户窗体中添加控件)
* [3.设置控件属性](#设置控件属性)
* [4.初始化控件属性](#初始化控件属性)
* [5.控件和对话框事件](#控件和对话框事件)
* [6.显示自定义对话框](#显示自定义对话框)
* [7.代码运行时使用控件值](#代码运行时使用控件值)

## 创建用户窗体

要创建自定义对话框，必须创建用户窗体。 要创建用户窗体，请单击“Visual Basic 编辑器”中 “**插入**” 菜单上的 “**用户窗体**”。

可使用 “**属性**” 窗口更改窗体的名称、行为和外观。 例如，要更改窗体的标题，可设置 `Caption` 属性。

## 向用户窗体中添加控件

在 “工具箱” 中找到要添加的控件并将它拖到窗体上。

若要向用户窗体中添加控件，请在 "**工具箱**" 中找到要添加的控件，将该控件拖到窗体上，然后拖动控件的调整控件，直到控件的轮廓大小和形状满足要求。

> 从窗体中将一个控件（或多个 "分组的" 控件）拖回到 "**工具箱**"，将创建一个可重复使用的控件模板。 This is a useful feature for implementing a standard interface for your applications.

将控件添加到窗体后，可使用 Visual Basic 编辑器中 "**格式**" 菜单上的命令调整控件的对齐方式和间距。

## 设置控件属性

在设计模式下右键单击控件，然后单击 “属性” 以显示 “属性” 窗口。

您可以在设计时（在运行任何宏之前）设置某些 **控件** 的属性。 在设计模式下，右键单击某一控件，然后单击“属性”**** 显示属性窗口。 属性名称显示在该窗口的左列中，属性的值显示在右列中。 在属性名称的右边输入新值可以设置属性的值。

## 初始化控件属性

可以在显示窗体之前的一个过程中初始化控件，或者为窗体的 `Initialize` 事件添加代码。

可以在运行时使用宏中的 Visual Basic 代码初始化 **控件**。 例如，可填充列表框、设置文本值或设置选项按钮。

以下示例使用 `AddItem` 方法向列表框中添加数据。 然后它设置文本框的值并显示窗体。
```vba
Private Sub GetUserName() 
    With UserForm1 
        .lstRegions.AddItem "North" 
        .lstRegions.AddItem "South" 
        .lstRegions.AddItem "East" 
        .lstRegions.AddItem "West" 
        .txtSalesPersonID.Text = "00000" 
        .Show 
        ' ... 
    End With 
End Sub
```
还可以在窗体的 `Initialize` 事件中用代码设置窗体上控件的初始值。 在 `Initialize` 事件中设置控件的初始值的好处在于：初始化代码与窗体存储在一起。 您可以将该窗体复制到其他项目中，这样，当运行 `Show` 方法来显示对话框时，将初始化其中的控件。
```vba
Private Sub UserForm_Initialize() 
    UserForm1.lstNames.AddItem "Test One" 
    UserForm1.lstNames.AddItem "Test Two" 
    UserForm1.txtUserName.Text = "Default Name" 
End Sub
```

## 控件和对话框事件

所有的控件都有一组预定义事件。 例如，命令按钮有一个当用户单击它时发生的 `Click` 事件。 您可以编写事件发生时运行的事件过程。

向对话框或文档中添加控件后, 添加事件过程以确定控件如何响应用户操作。

用户窗体和控件均拥有一组预定义事件。 例如，命令按钮具有 `Click` 事件，在用户单击命令按钮时，该事件发生，用户窗体具有 `Initialize` 事件，在加载窗体时，该事件运行。

若要编写控件或窗体事件过程, 请双击窗体或控件打开**模块**, 然后从 "**过程**" 列表框中选择事件。

事件过程包含控件的名称。 例如, `Command1_Click` 过程名称是控件名为`Command1`的Click事件。

如果为事件过程添加代码后更改该控件的名称，这些代码仍保留使用原名称的过程中。

例如, 假设您向`Command1` 控件的 **Click** 事件添加代码, 然后将该控件重命名为 `Command2`。 `Command2` 的 **click** 事件过程中将看不到任何代码。 您需要将 `Command1_Click` 代码移动到 `Command2_Click`。

> 若要简化开发过程，在编写代码之前命名控件是一个不错的做法。

## 显示自定义对话框

使用 `Show` 方法显示用户窗体。

若要在 Visual Basic 编辑器中测试对话框，请在 Visual Basic 编辑器的 “**运行**” 菜单上单击 “**运行子过程/用户窗体**”。

若要在 Visual Basic 中显示对话框，请使用 `Show` 方法。 下例显示了名为“`UserForm1`”的对话框。
```vba
Private Sub GetUserName() 
    UserForm1.Show 
End Sub
```

## 代码运行时使用控件值

**有些属性可以在运行时设置。 关闭对话框时，用户对对话框所做的更改将丢失。**

在运行 Visual Basic 代码时，可以设置和返回一些 **控件** 属性。 以下示例将一个文本框的 Text 属性设置为"Hello"。
```vba
TextBox1.Text = "Hello"
```
关闭窗体时，用户在窗体中输入的数据将丢失。 如果在卸载窗体后返回其中控件的值，则得到的是控件的初始值而非用户输入的值。

**如果要保存在窗体中输入的数据，可以在窗体仍在运行时将该信息保存到模块级变量中**。 以下示例显示一个窗体并保存窗体数据。
```vba
' Code in module to declare public variables. 
Public strRegion As String 
Public intSalesPersonID As Integer 
Public blnCancelled As Boolean 
 
' Code in form. 
Private Sub cmdCancel_Click() 
    Module1.blnCancelled = True 
    Unload Me 
End Sub 
 
Private Sub cmdOK_Click() 
    ' Save data. 
    intSalesPersonID = txtSalesPersonID.Text 
    strRegion = lstRegions.List(lstRegions.ListIndex) 
    Module1.blnCancelled = False 
    Unload Me 
End Sub 
 
Private Sub UserForm_Initialize() 
    Module1.blnCancelled = True 
End Sub 
 
' Code in module to display form. 
Sub LaunchSalesPersonForm() 
    frmSalesPeople.Show 
    If blnCancelled = True Then 
        MsgBox "Operation Cancelled!", vbExclamation 
    Else 
        MsgBox "The Salesperson's ID is: " & 
        intSalesPersonID & _ 
        "The Region is: " & strRegion 
    End If 
End Sub
```