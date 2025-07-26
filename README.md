# WPS/Excel VBA 智能粘贴工具 (Smart Paste Tool)

这是一个非常实用和灵活的VBA宏工具，旨在解决一个常见的表格操作难题：**如何将一列数据精确地粘贴到含有大量不规则合并单元格的另一列中**。

普通粘贴在这种情况下会失败，而这个工具通过交互式对话框，让用户通过鼠标选择源和目标，实现了完美的“智能”填充。

---

## ✨ 功能亮点

- **交互式操作**：无需修改任何代码，通过对话框引导用户选择数据源和目标位置。
- **通用性强**：不限制工作表名称、数据位置，适用于任何WPS或Excel文件。
- **智能识别合并单元格**：自动计算合并单元格的大小，确保数据被正确地、不重不漏地填入每个独立的单元格（或合并区域）中。
- **高效稳定**：代码经过优化，能够快速处理大量数据。

---

## 🚀 如何使用

整个过程非常简单，只需要三步：

1.  **导入代码**：
    - 在WPS或Excel中，按 `Alt + F11` 打开VBA编辑器。
    - 在菜单栏点击 `插入` -> `模块`。
    - 将下面的VBA代码完整复制并粘贴到新模块的空白区域中。
    - 关闭VBA编辑器。

2.  **运行宏**：
    - 回到表格界面，按 `Alt + F8` 打开宏列表。
    - 选择名为 `InteractiveSmartPaste` 的宏，点击 `运行`。

3.  **根据提示操作**：
    - **第1步**：会弹出一个提示框，要求您 **选择源数据**。此时，用鼠标框选您要复制的那一列数据，然后点击“确定”。
    - **第2步**：会弹出第二个提示框，要求您 **选择粘贴位置**。此时，用鼠标 **单击** 您希望数据开始填充的第一个单元格，然后点击“确定”。

完成！程序会自动将所有数据填充到目标位置。

---

## 📋 VBA 源代码

这是您可以直接使用的VBA代码。您也可以将此项目中的 `SmartPaste.bas` 文件通过VBA编辑器的 `文件` -> `导入文件` 功能直接导入到您的项目中。

```vba
Sub InteractiveSmartPaste()
    Dim sourceRange As Range
    Dim targetStartCell As Range
    Dim sourceData As Variant
    Dim targetSheet As Worksheet
    Dim currentTargetCell As Range
    Dim dataIndex As Long
    
    ' --- Step 1: 让用户用鼠标选择源数据 ---
    On Error Resume Next ' 如果用户点了“取消”，可以避免程序报错
    Set sourceRange = Application.InputBox("请用鼠标选择您要复制的源数据列。", "第1步：选择源数据", Type:=8)
    On Error GoTo 0 ' 恢复正常的错误处理
    
    ' 检查用户是否取消了选择
    If sourceRange Is Nothing Then
        MsgBox "操作已取消。", vbInformation
        Exit Sub
    End If
    
    ' --- Step 2: 让用户用鼠标选择粘贴的起始位置 ---
    On Error Resume Next
    Set targetStartCell = Application.InputBox("请用鼠标点击您要粘贴到的第一个单元格。", "第2步：选择粘贴位置", Type:=8)
    On Error GoTo 0
    
    ' 检查用户是否取消了选择
    If targetStartCell Is Nothing Then
        MsgBox "操作已取消。", vbInformation
        Exit Sub
    End If
    
    ' --- Step 3: 开始处理数据 ---
    
    ' 为了防止用户选择多个单元格，我们只取他选择区域的第一个单元格作为起点
    Set targetStartCell = targetStartCell.Cells(1, 1)
    Set targetSheet = targetStartCell.Worksheet
    
    ' 将源数据一次性读入数组，提高速度
    sourceData = sourceRange.Value
    
    ' 如果源数据是空的，则提示并退出
    If IsEmpty(sourceData) Then
        MsgBox "您选择的源数据区域是空的，没有内容可以粘贴。", vbExclamation
        Exit Sub
    End If

    dataIndex = 1 ' 数据索引，从1开始
    Set currentTargetCell = targetStartCell ' 当前要粘贴的目标单元格
    
    ' 循环处理，直到所有源数据都粘贴完毕
    Do While dataIndex <= UBound(sourceData, 1)
        
        ' 将数据粘贴到当前目标单元格
        currentTargetCell.Value = sourceData(dataIndex, 1)
        
        ' 准备处理下一条数据
        dataIndex = dataIndex + 1
        
        ' 如果数据已经用完，就跳出循环
        If dataIndex > UBound(sourceData, 1) Then Exit Do
        
        ' 关键一步：计算下一个可用的粘贴位置
        ' 它会跳过当前单元格所占用的所有合并区域
        Set currentTargetCell = targetSheet.Cells(currentTargetCell.MergeArea.Row + currentTargetCell.MergeArea.Rows.Count, currentTargetCell.Column)
        
    Loop
    
    MsgBox "太棒了！" & (dataIndex - 1) & " 条数据已成功填充！"
End Sub
```
