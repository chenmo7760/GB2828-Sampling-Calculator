Option Explicit

' GB/T2828.1-2003/ISO2859-1 抽样计算模块
' 使用说明：
' 1. 在Excel中按Alt+F11打开VBA编辑器
' 2. 插入 > 模块，将本代码复制进去
' 3. 在工作表中设置输入单元格，调用函数GetSampleSize()

'=============================================================================
' 主函数：计算抽样样本量和Ac/Re值
'=============================================================================
' 参数：
'   批量PL: 批次总量
'   检验水平: S-1、S-2、S-3、S-4、Ⅰ、Ⅱ、Ⅲ
'   AQL值: 如0.01, 0.015, 0.025等
' 返回：数组 {样本量, Ac值, Re值, 是否全检}
'=============================================================================
Function 计算抽样(批量 As Long, 检验水平 As String, AQL As Double) As Variant
    Dim ws As Worksheet
    Dim 批量行 As Long
    Dim 检验水平列 As Long
    Dim 样本量字码 As String
    Dim 字码行 As Long
    Dim AQL列AC As Long
    Dim AQL列RE As Long
    Dim Ac值 As Variant
    Dim Re值 As Variant
    Dim 样本量 As Long
    Dim 结果(1 To 4) As Variant
    Dim i As Long
    Dim 实际行号 As Long
    
    ' 设置工作表（假设数据在当前活动工作表）
    On Error GoTo 错误处理
    Set ws = ActiveSheet
    
    ' 不在这里清除高亮，避免UDF调用时清除
    ' 高亮由公式计算完成后统一处理
    
    ' 步骤1：根据批量找到对应的行（下移2行后：第8-22行）
    批量行 = 0
    For i = 8 To 22
        If 批量 >= ws.Cells(i, 1).Value And 批量 <= ws.Cells(i, 2).Value Then
            批量行 = i
            Exit For
        End If
    Next i
    
    ' 处理500001以上的情况
    If 批量行 = 0 And 批量 > 500000 Then
        批量行 = 22
    End If
    
    If 批量行 = 0 Then
        结果(1) = "错误：批量超出范围"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤2：根据检验水平找到对应的列
    检验水平列 = 获取检验水平列(检验水平)
    If 检验水平列 = 0 Then
        结果(1) = "错误：无效的检验水平"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤3：获取样本量字码
    样本量字码 = ws.Cells(批量行, 检验水平列).Value
    
    ' 步骤4：在K列（第11列）中查找样本量字码对应的行（下移2行后：第8-22行）
    字码行 = 0
    For i = 8 To 22
        If ws.Cells(i, 11).Value = 样本量字码 Then
            字码行 = i
            Exit For
        End If
    Next i
    
    If 字码行 = 0 Then
        结果(1) = "错误：未找到样本量字码"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤5：获取初始样本量（可能会被后续步骤更新）
    样本量 = ws.Cells(字码行, 12).Value
    
    ' 步骤6：根据AQL值找到对应的AC和RE列
    AQL列AC = 获取AQL列(AQL, True)  ' AC列
    AQL列RE = 获取AQL列(AQL, False) ' RE列
    
    If AQL列AC = 0 Then
        结果(1) = "错误：无效的AQL值"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤7：获取Ac和Re值（处理"上"、"下"的情况）
    ' 注意：如果遇到"上"或"下"，会返回实际找到数值的行号
    实际行号 = 字码行  ' 初始化为字码行
    Ac值 = 获取单元格数值带行号(ws, 字码行, AQL列AC, 实际行号)
    
    ' 关键：如果实际行号与原字码行不同，说明发生了"上"或"下"查找
    ' 需要使用实际行号对应的样本量
    If 实际行号 <> 字码行 And 实际行号 > 0 Then
        样本量 = ws.Cells(实际行号, 12).Value
    End If
    
    Re值 = 获取单元格数值(ws, 实际行号, AQL列RE)
    
    ' 步骤8：判断是否全检
    ' 样本量：始终返回标准表中的样本量
    ' 检验类型：标注是全检还是抽检
    结果(1) = 样本量
    结果(2) = Ac值
    结果(3) = Re值
    
    If 样本量 >= 批量 Then
        结果(4) = "全检(需检" & 批量 & "件)"
    Else
        结果(4) = "抽检"
    End If
    
    ' 注意：Excel UDF不能可靠修改单元格格式
    ' 高亮功能由工作表事件（Worksheet_Calculate）触发
    
    计算抽样 = 结果
    Exit Function
    
错误处理:
    结果(1) = "错误: " & Err.Description
    结果(2) = "-"
    结果(3) = "-"
    结果(4) = "错误"
    计算抽样 = 结果
End Function

'=============================================================================
' 辅助函数：获取检验水平对应的列号
'=============================================================================
Private Function 获取检验水平列(检验水平 As String) As Long
    Select Case 检验水平
        Case "S-1": 获取检验水平列 = 4   ' D列
        Case "S-2": 获取检验水平列 = 5   ' E列
        Case "S-3": 获取检验水平列 = 6   ' F列
        Case "S-4": 获取检验水平列 = 7   ' G列
        Case "Ⅰ", "I": 获取检验水平列 = 8   ' H列
        Case "Ⅱ", "II": 获取检验水平列 = 9   ' I列
        Case "Ⅲ", "III": 获取检验水平列 = 10  ' J列
        Case Else: 获取检验水平列 = 0
    End Select
End Function

'=============================================================================
' 辅助函数：获取AQL对应的列号
'=============================================================================
Private Function 获取AQL列(AQL As Double, 是AC列 As Boolean) As Long
    Dim AQL数组 As Variant
    Dim i As Long
    Dim 基础列 As Long
    
    ' AQL值数组（从0.01开始，0不考虑）
    AQL数组 = Array(0.01, 0.015, 0.025, 0.04, 0.065, 0.1, 0.15, 0.25, 0.4, 0.65, _
                    1#, 1.5, 2.5, 4#, 6.5, 10#, 15#, 25#, 40#, 65#, 100#)
    
    ' 查找AQL值在数组中的位置
    For i = LBound(AQL数组) To UBound(AQL数组)
        If Abs(AQL数组(i) - AQL) < 0.0001 Then
            ' O列是第15列，对应AQL=0.01（第一个值）
            ' 每个AQL值占2列（AC和RE）
            基础列 = 15 + i * 2
            If 是AC列 Then
                获取AQL列 = 基础列
            Else
                获取AQL列 = 基础列 + 1
            End If
            Exit Function
        End If
    Next i
    
    获取AQL列 = 0
End Function

'=============================================================================
' 辅助函数：获取单元格数值（处理"上"、"下"的情况）
'=============================================================================
Private Function 获取单元格数值(ws As Worksheet, 行号 As Long, 列号 As Long) As Variant
    Dim 单元格值 As Variant
    Dim 当前行 As Long
    
    当前行 = 行号
    单元格值 = ws.Cells(当前行, 列号).Value
    
    ' 处理"箭"的情况（第22行可能有）
    If 单元格值 = "箭" Then
        单元格值 = "下"
    End If
    
    ' 循环查找直到找到数字（下移2行后：第8-22行）
    Do While 单元格值 = "上" Or 单元格值 = "下"
        If 单元格值 = "上" Then
            当前行 = 当前行 - 1
            If 当前行 < 8 Then Exit Do ' 防止越界
        ElseIf 单元格值 = "下" Then
            当前行 = 当前行 + 1
            If 当前行 > 22 Then Exit Do ' 防止越界
        End If
        单元格值 = ws.Cells(当前行, 列号).Value
    Loop
    
    ' 如果是空值，返回"-"
    If IsEmpty(单元格值) Or 单元格值 = "" Then
        获取单元格数值 = "-"
    Else
        获取单元格数值 = 单元格值
    End If
End Function

'=============================================================================
' 辅助函数：获取单元格数值并返回实际行号（处理"上"、"下"的情况）
' 参数：实际行号 - 通过引用返回，表示实际找到数值的行号
'=============================================================================
Private Function 获取单元格数值带行号(ws As Worksheet, 行号 As Long, 列号 As Long, ByRef 实际行号 As Long) As Variant
    Dim 单元格值 As Variant
    Dim 当前行 As Long
    
    当前行 = 行号
    单元格值 = ws.Cells(当前行, 列号).Value
    
    ' 处理"箭"的情况（第22行可能有）
    If 单元格值 = "箭" Then
        单元格值 = "下"
    End If
    
    ' 循环查找直到找到数字（下移2行后：第8-22行）
    Do While 单元格值 = "上" Or 单元格值 = "下"
        If 单元格值 = "上" Then
            当前行 = 当前行 - 1
            If 当前行 < 8 Then Exit Do ' 防止越界
        ElseIf 单元格值 = "下" Then
            当前行 = 当前行 + 1
            If 当前行 > 22 Then Exit Do ' 防止越界
        End If
        单元格值 = ws.Cells(当前行, 列号).Value
    Loop
    
    ' 返回实际找到数值的行号
    实际行号 = 当前行
    
    ' 如果是空值，返回"-"
    If IsEmpty(单元格值) Or 单元格值 = "" Then
        获取单元格数值带行号 = "-"
    Else
        获取单元格数值带行号 = 单元格值
    End If
End Function


'=============================================================================
' 宏：执行计算并高亮显示（可以绑定到按钮）
' 使用方法：插入 > 形状 > 按钮，右键 > 指定宏 > 执行抽样计算并高亮
'=============================================================================
Sub 执行抽样计算并高亮()
    Dim ws As Worksheet
    Dim 批量 As Long
    Dim 检验水平 As String
    Dim AQL As Double
    Dim 结果 As Variant
    Dim 批量行 As Long
    Dim 检验水平列 As Long
    Dim 样本量字码 As String
    Dim 字码行 As Long
    Dim 实际行号 As Long
    Dim AQL列AC As Long
    Dim AQL列RE As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    ' 从固定单元格读取输入
    On Error Resume Next
    批量 = ws.Range("B1").Value
    检验水平 = ws.Range("B2").Value
    AQL = ws.Range("B3").Value
    On Error GoTo 0
    
    If 批量 = 0 Or 检验水平 = "" Or AQL = 0 Then
        MsgBox "请在B1、B2、B3输入批量、检验水平和AQL值", vbExclamation
        Exit Sub
    End If
    
    ' 调用计算函数
    结果 = 计算抽样(批量, 检验水平, AQL)
    
    ' 输出结果到B4-B7
    ws.Range("B4").Value = 结果(1)
    ws.Range("B5").Value = 结果(2)
    ws.Range("B6").Value = 结果(3)
    ws.Range("B7").Value = 结果(4)
    
    ' 重新计算定位信息用于高亮
    ' 清除旧高亮
    ws.Range("A8:BH22").Interior.ColorIndex = xlNone
    
    ' 找到批量行
    批量行 = 0
    For i = 8 To 22
        If 批量 >= ws.Cells(i, 1).Value And 批量 <= ws.Cells(i, 2).Value Then
            批量行 = i
            Exit For
        End If
    Next i
    If 批量行 = 0 And 批量 > 500000 Then 批量行 = 22
    
    If 批量行 = 0 Then Exit Sub
    
    ' 找到检验水平列
    检验水平列 = 获取检验水平列(检验水平)
    If 检验水平列 = 0 Then Exit Sub
    
    ' 获取样本量字码
    样本量字码 = ws.Cells(批量行, 检验水平列).Value
    
    ' 找到字码行
    字码行 = 0
    For i = 8 To 22
        If ws.Cells(i, 11).Value = 样本量字码 Then
            字码行 = i
            Exit For
        End If
    Next i
    If 字码行 = 0 Then Exit Sub
    
    ' 获取AQL列
    AQL列AC = 获取AQL列(AQL, True)
    AQL列RE = 获取AQL列(AQL, False)
    If AQL列AC = 0 Then Exit Sub
    
    ' 获取实际行号（处理"上"/"下"）
    实际行号 = 字码行
    Call 获取单元格数值带行号(ws, 字码行, AQL列AC, 实际行号)
    
    ' 执行高亮
    If 实际行号 >= 8 And 实际行号 <= 22 Then
        ' 高亮样本量
        ws.Cells(实际行号, 12).Interior.Color = RGB(255, 255, 0) ' 黄色
        ws.Cells(实际行号, 12).Font.Bold = True
        
        ' 高亮AC
        If AQL列AC > 0 And AQL列AC <= 60 Then
            ws.Cells(实际行号, AQL列AC).Interior.Color = RGB(146, 208, 80) ' 浅绿色
            ws.Cells(实际行号, AQL列AC).Font.Bold = True
        End If
        
        ' 高亮RE
        If AQL列RE > 0 And AQL列RE <= 60 Then
            ws.Cells(实际行号, AQL列RE).Interior.Color = RGB(146, 208, 80) ' 浅绿色
            ws.Cells(实际行号, AQL列RE).Font.Bold = True
        End If
        
        ' 高亮样本量字码
        ws.Cells(批量行, 检验水平列).Interior.Color = RGB(255, 192, 0) ' 橙色
    End If
    
    MsgBox "计算完成！" & vbCrLf & _
           "样本量: " & 结果(1) & vbCrLf & _
           "AC: " & 结果(2) & vbCrLf & _
           "RE: " & 结果(3) & vbCrLf & _
           "类型: " & 结果(4), vbInformation, "抽样计算结果"
End Sub

'=============================================================================
' 自动执行版本（由工作表事件触发，不显示消息框）
'=============================================================================
Sub 自动执行抽样计算()
    Dim ws As Worksheet
    Dim 批量 As Long
    Dim 检验水平 As String
    Dim AQL As Double
    Dim 结果 As Variant
    Dim 批量行 As Long
    Dim 检验水平列 As Long
    Dim 样本量字码 As String
    Dim 字码行 As Long
    Dim 实际行号 As Long
    Dim AQL列AC As Long
    Dim AQL列RE As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    ' 从固定单元格读取输入
    On Error Resume Next
    批量 = ws.Range("B1").Value
    检验水平 = ws.Range("B2").Value
    AQL = ws.Range("B3").Value
    On Error GoTo 0
    
    ' 验证输入
    If 批量 = 0 Or 检验水平 = "" Or AQL = 0 Then
        Exit Sub ' 静默退出，不显示错误
    End If
    
    ' 调用计算函数（用于获取定位信息，不输出到单元格）
    结果 = 计算抽样(批量, 检验水平, AQL)
    
    ' 注意：不写入B4-B7，保留公式
    ' 如果B4-B7为空（没有公式），才写入值
    Application.EnableEvents = False
    If IsEmpty(ws.Range("B4").Formula) Or ws.Range("B4").Formula = "" Then
        ws.Range("B4").Value = 结果(1)
        ws.Range("B5").Value = 结果(2)
        ws.Range("B6").Value = 结果(3)
        ws.Range("B7").Value = 结果(4)
    End If
    Application.EnableEvents = True
    
    ' 重新计算定位信息用于高亮
    ' 清除旧高亮
    ws.Range("A8:BH22").Interior.ColorIndex = xlNone
    
    ' 找到批量行
    批量行 = 0
    For i = 8 To 22
        If 批量 >= ws.Cells(i, 1).Value And 批量 <= ws.Cells(i, 2).Value Then
            批量行 = i
            Exit For
        End If
    Next i
    If 批量行 = 0 And 批量 > 500000 Then 批量行 = 22
    
    If 批量行 = 0 Then Exit Sub
    
    ' 找到检验水平列
    检验水平列 = 获取检验水平列(检验水平)
    If 检验水平列 = 0 Then Exit Sub
    
    ' 获取样本量字码
    样本量字码 = ws.Cells(批量行, 检验水平列).Value
    
    ' 找到字码行
    字码行 = 0
    For i = 8 To 22
        If ws.Cells(i, 11).Value = 样本量字码 Then
            字码行 = i
            Exit For
        End If
    Next i
    If 字码行 = 0 Then Exit Sub
    
    ' 获取AQL列
    AQL列AC = 获取AQL列(AQL, True)
    AQL列RE = 获取AQL列(AQL, False)
    If AQL列AC = 0 Then Exit Sub
    
    ' 获取实际行号（处理"上"/"下"）
    实际行号 = 字码行
    Call 获取单元格数值带行号(ws, 字码行, AQL列AC, 实际行号)
    
    ' 执行高亮
    If 实际行号 >= 8 And 实际行号 <= 22 Then
        ' 高亮样本量
        ws.Cells(实际行号, 12).Interior.Color = RGB(255, 255, 0) ' 黄色
        ws.Cells(实际行号, 12).Font.Bold = True
        
        ' 高亮AC
        If AQL列AC > 0 And AQL列AC <= 60 Then
            ws.Cells(实际行号, AQL列AC).Interior.Color = RGB(146, 208, 80) ' 浅绿色
            ws.Cells(实际行号, AQL列AC).Font.Bold = True
        End If
        
        ' 高亮RE
        If AQL列RE > 0 And AQL列RE <= 60 Then
            ws.Cells(实际行号, AQL列RE).Interior.Color = RGB(146, 208, 80) ' 浅绿色
            ws.Cells(实际行号, AQL列RE).Font.Bold = True
        End If
        
        ' 高亮样本量字码
        ws.Cells(批量行, 检验水平列).Interior.Color = RGB(255, 192, 0) ' 橙色
    End If
End Sub

'=============================================================================
' 包装函数：计算样本量并触发高亮（用于B4单元格）
'=============================================================================
Function 获取样本量并高亮(批量 As Long, 检验水平 As String, AQL As Double) As Variant
    Dim 结果 As Variant
    结果 = 计算抽样(批量, 检验水平, AQL)
    
    ' 触发高亮（延迟执行，避免UDF限制）
    Application.OnTime Now + TimeValue("00:00:01"), "仅执行高亮"
    
    If IsArray(结果) Then
        获取样本量并高亮 = 结果(1)
    Else
        获取样本量并高亮 = 结果
    End If
End Function

'=============================================================================
' 仅执行高亮（不修改单元格值）
'=============================================================================
Sub 仅执行高亮()
    Dim ws As Worksheet
    Dim 批量 As Long
    Dim 检验水平 As String
    Dim AQL As Double
    Dim 批量行 As Long
    Dim 检验水平列 As Long
    Dim 样本量字码 As String
    Dim 字码行 As Long
    Dim 实际行号 As Long
    Dim AQL列AC As Long
    Dim AQL列RE As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    ' 从固定单元格读取输入
    On Error Resume Next
    批量 = ws.Range("B1").Value
    检验水平 = ws.Range("B2").Value
    AQL = ws.Range("B3").Value
    On Error GoTo 0
    
    If 批量 = 0 Or 检验水平 = "" Or AQL = 0 Then Exit Sub
    
    ' 清除旧高亮
    ws.Range("A8:BH22").Interior.ColorIndex = xlNone
    ws.Range("A8:BH22").Font.Bold = False
    
    ' 找到批量行
    批量行 = 0
    For i = 8 To 22
        If 批量 >= ws.Cells(i, 1).Value And 批量 <= ws.Cells(i, 2).Value Then
            批量行 = i
            Exit For
        End If
    Next i
    If 批量行 = 0 And 批量 > 500000 Then 批量行 = 22
    If 批量行 = 0 Then Exit Sub
    
    ' 找到检验水平列
    检验水平列 = 获取检验水平列(检验水平)
    If 检验水平列 = 0 Then Exit Sub
    
    ' 获取样本量字码
    样本量字码 = ws.Cells(批量行, 检验水平列).Value
    
    ' 找到字码行
    字码行 = 0
    For i = 8 To 22
        If ws.Cells(i, 11).Value = 样本量字码 Then
            字码行 = i
            Exit For
        End If
    Next i
    If 字码行 = 0 Then Exit Sub
    
    ' 获取AQL列
    AQL列AC = 获取AQL列(AQL, True)
    AQL列RE = 获取AQL列(AQL, False)
    If AQL列AC = 0 Then Exit Sub
    
    ' 获取实际行号
    实际行号 = 字码行
    Call 获取单元格数值带行号(ws, 字码行, AQL列AC, 实际行号)
    
    ' 执行高亮
    If 实际行号 >= 8 And 实际行号 <= 22 Then
        ' 高亮样本量
        ws.Cells(实际行号, 12).Interior.Color = RGB(255, 255, 0)
        ws.Cells(实际行号, 12).Font.Bold = True
        
        ' 高亮AC
        If AQL列AC > 0 And AQL列AC <= 60 Then
            ws.Cells(实际行号, AQL列AC).Interior.Color = RGB(146, 208, 80)
            ws.Cells(实际行号, AQL列AC).Font.Bold = True
        End If
        
        ' 高亮RE
        If AQL列RE > 0 And AQL列RE <= 60 Then
            ws.Cells(实际行号, AQL列RE).Interior.Color = RGB(146, 208, 80)
            ws.Cells(实际行号, AQL列RE).Font.Bold = True
        End If
        
        ' 高亮样本量字码
        ws.Cells(批量行, 检验水平列).Interior.Color = RGB(255, 192, 0)
    End If
End Sub

'=============================================================================
' 简化函数：只返回样本量
'=============================================================================
Function 获取样本量(批量 As Long, 检验水平 As String, AQL As Double) As Variant
    Dim 结果 As Variant
    结果 = 计算抽样(批量, 检验水平, AQL)
    If IsArray(结果) Then
        获取样本量 = 结果(1)
    Else
        获取样本量 = 结果
    End If
End Function

'=============================================================================
' 简化函数：返回Ac值
'=============================================================================
Function 获取Ac值(批量 As Long, 检验水平 As String, AQL As Double) As Variant
    Dim 结果 As Variant
    结果 = 计算抽样(批量, 检验水平, AQL)
    If IsArray(结果) Then
        获取Ac值 = 结果(2)
    Else
        获取Ac值 = "-"
    End If
End Function

'=============================================================================
' 简化函数：返回Re值
'=============================================================================
Function 获取Re值(批量 As Long, 检验水平 As String, AQL As Double) As Variant
    Dim 结果 As Variant
    结果 = 计算抽样(批量, 检验水平, AQL)
    If IsArray(结果) Then
        获取Re值 = 结果(3)
    Else
        获取Re值 = "-"
    End If
End Function

'=============================================================================
' 简化函数：返回检验类型（全检/抽检）
'=============================================================================
Function 获取检验类型(批量 As Long, 检验水平 As String, AQL As Double) As Variant
    Dim 结果 As Variant
    结果 = 计算抽样(批量, 检验水平, AQL)
    If IsArray(结果) Then
        获取检验类型 = 结果(4)
    Else
        获取检验类型 = "-"
    End If
End Function

