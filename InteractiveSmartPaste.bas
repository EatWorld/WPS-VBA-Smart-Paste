Attribute VB_Name = "Module1"
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
