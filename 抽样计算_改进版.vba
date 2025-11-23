Option Explicit

' GB/T2828.1-2003/ISO2859-1 抽样计算模块（改进版）
' 改进：自动查找列位置，不依赖固定列号

'=============================================================================
' 主函数：计算抽样样本量和Ac/Re值
'=============================================================================
Function 计算抽样(批量 As Long, 检验水平 As String, AQL As Double) As Variant
    Dim ws As Worksheet
    Dim 批量行 As Long
    Dim 检验水平列 As Long
    Dim 样本量字码 As String
    Dim 字码行 As Long
    Dim 样本量字码列 As Long
    Dim 样本量列 As Long
    Dim AQL列AC As Long
    Dim AQL列RE As Long
    Dim Ac值 As Variant
    Dim Re值 As Variant
    Dim 样本量 As Long
    Dim 结果(1 To 4) As Variant
    Dim i As Long, j As Long
    
    ' 设置工作表
    Set ws = ActiveSheet
    
    ' 步骤0：查找关键列的位置（动态查找，不假设固定列号）
    样本量字码列 = 查找列标题(ws, "样本量字码")
    样本量列 = 查找列标题(ws, "样本量")
    
    If 样本量字码列 = 0 Or 样本量列 = 0 Then
        结果(1) = "错误：找不到样本量字码或样本量列"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤1：根据批量找到对应的行（第6-20行，数据行）
    批量行 = 0
    For i = 6 To 20
        Dim 最小批量 As Variant, 最大批量 As Variant
        最小批量 = ws.Cells(i, 1).Value
        最大批量 = ws.Cells(i, 2).Value
        
        ' 处理"500001以上"的情况
        If IsNumeric(最小批量) Then
            If IsNumeric(最大批量) Then
                If 批量 >= 最小批量 And 批量 <= 最大批量 Then
                    批量行 = i
                    Exit For
                End If
            Else
                ' 最大批量非数字，可能是"以上"
                If 批量 >= 最小批量 Then
                    批量行 = i
                    Exit For
                End If
            End If
        End If
    Next i
    
    If 批量行 = 0 Then
        结果(1) = "错误：批量超出范围"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤2：根据检验水平找到对应的列（动态查找）
    检验水平列 = 查找检验水平列(ws, 检验水平)
    
    If 检验水平列 = 0 Then
        结果(1) = "错误：无效的检验水平 [" & 检验水平 & "]"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤3：获取样本量字码
    样本量字码 = Trim(ws.Cells(批量行, 检验水平列).Value)
    
    If 样本量字码 = "" Then
        结果(1) = "错误：未找到样本量字码"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤4：在样本量字码列中查找对应的行
    字码行 = 0
    For i = 6 To 20
        If Trim(ws.Cells(i, 样本量字码列).Value) = 样本量字码 Then
            字码行 = i
            Exit For
        End If
    Next i
    
    If 字码行 = 0 Then
        结果(1) = "错误：未找到样本量字码 [" & 样本量字码 & "]"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤5：获取样本量
    样本量 = ws.Cells(字码行, 样本量列).Value
    
    ' 步骤6：根据AQL值找到对应的AC和RE列（动态查找）
    AQL列AC = 查找AQL列(ws, AQL, "AC")
    AQL列RE = 查找AQL列(ws, AQL, "RE")
    
    If AQL列AC = 0 Then
        结果(1) = "错误：无效的AQL值 [" & AQL & "]"
        计算抽样 = 结果
        Exit Function
    End If
    
    ' 步骤7：获取Ac和Re值（处理"上"、"下"的情况）
    Ac值 = 获取单元格数值(ws, 字码行, AQL列AC)
    Re值 = 获取单元格数值(ws, 字码行, AQL列RE)
    
    ' 步骤8：判断是否全检
    If 样本量 >= 批量 Then
        结果(1) = 批量
        结果(2) = Ac值
        结果(3) = Re值
        结果(4) = "全检"
    Else
        结果(1) = 样本量
        结果(2) = Ac值
        结果(3) = Re值
        结果(4) = "抽检"
    End If
    
    计算抽样 = 结果
End Function

'=============================================================================
' 辅助函数：查找列标题位置
'=============================================================================
Private Function 查找列标题(ws As Worksheet, 标题关键字 As String) As Long
    Dim i As Long, j As Long
    
    ' 在第3-5行中查找标题
    For i = 3 To 5
        For j = 1 To 50
            Dim 单元格内容 As String
            单元格内容 = Trim(ws.Cells(i, j).Value)
            If InStr(1, 单元格内容, 标题关键字, vbTextCompare) > 0 Then
                查找列标题 = j
                Exit Function
            End If
        Next j
    Next i
    
    查找列标题 = 0
End Function

'=============================================================================
' 辅助函数：查找检验水平列位置
'=============================================================================
Private Function 查找检验水平列(ws As Worksheet, 检验水平 As String) As Long
    Dim i As Long, j As Long
    Dim 标准化水平 As String
    
    ' 标准化检验水平输入
    Select Case UCase(检验水平)
        Case "S-1", "S1": 标准化水平 = "S-1"
        Case "S-2", "S2": 标准化水平 = "S-2"
        Case "S-3", "S3": 标准化水平 = "S-3"
        Case "S-4", "S4": 标准化水平 = "S-4"
        Case "Ⅰ", "I", "1": 标准化水平 = "Ⅰ"
        Case "Ⅱ", "II", "2": 标准化水平 = "Ⅱ"
        Case "Ⅲ", "III", "3": 标准化水平 = "Ⅲ"
        Case Else
            查找检验水平列 = 0
            Exit Function
    End Select
    
    ' 在第4-5行中查找检验水平标题
    For i = 4 To 5
        For j = 1 To 30
            Dim 单元格内容 As String
            单元格内容 = Trim(ws.Cells(i, j).Value)
            
            ' 精确匹配或包含匹配
            If 单元格内容 = 标准化水平 Then
                查找检验水平列 = j
                Exit Function
            End If
            
            ' 处理罗马数字的不同表示
            If (标准化水平 = "Ⅰ" And (单元格内容 = "I" Or 单元格内容 = "1")) Or _
               (标准化水平 = "Ⅱ" And (单元格内容 = "II" Or 单元格内容 = "2")) Or _
               (标准化水平 = "Ⅲ" And (单元格内容 = "III" Or 单元格内容 = "3")) Then
                查找检验水平列 = j
                Exit Function
            End If
        Next j
    Next i
    
    查找检验水平列 = 0
End Function

'=============================================================================
' 辅助函数：查找AQL对应的AC/RE列
'=============================================================================
Private Function 查找AQL列(ws As Worksheet, AQL As Double, 类型 As String) As Long
    Dim i As Long, j As Long
    Dim AQL行 As Long
    Dim AC_RE行 As Long
    
    ' 查找AQL值所在的行和列
    AQL行 = 0
    AC_RE行 = 0
    
    ' 在第3-5行中查找AQL标题行
    For i = 3 To 5
        For j = 10 To 60
            Dim 单元格内容 As Variant
            单元格内容 = ws.Cells(i, j).Value
            
            ' 找到AQL值
            If IsNumeric(单元格内容) Then
                If Abs(CDbl(单元格内容) - AQL) < 0.0001 Then
                    AQL行 = i
                    
                    ' 查找下一行是否有AC/RE标记
                    Dim 下一行内容 As String
                    下一行内容 = UCase(Trim(ws.Cells(i + 1, j).Value))
                    
                    If InStr(下一行内容, "AC") > 0 Then
                        If UCase(类型) = "AC" Then
                            查找AQL列 = j
                            Exit Function
                        Else ' RE
                            查找AQL列 = j + 1
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next j
    Next i
    
    查找AQL列 = 0
End Function

'=============================================================================
' 辅助函数：获取单元格数值（处理"上"、"下"的情况）
'=============================================================================
Private Function 获取单元格数值(ws As Worksheet, 行号 As Long, 列号 As Long) As Variant
    Dim 单元格值 As Variant
    Dim 当前行 As Long
    Dim 查找次数 As Integer
    
    当前行 = 行号
    单元格值 = ws.Cells(当前行, 列号).Value
    查找次数 = 0
    
    ' 处理"箭"的情况
    If 单元格值 = "箭" Or 单元格值 = "箭头" Then
        单元格值 = "下"
    End If
    
    ' 循环查找直到找到数字
    Do While (单元格值 = "上" Or 单元格值 = "下") And 查找次数 < 20
        查找次数 = 查找次数 + 1
        
        If 单元格值 = "上" Then
            当前行 = 当前行 - 1
            If 当前行 < 6 Then Exit Do
        ElseIf 单元格值 = "下" Then
            当前行 = 当前行 + 1
            If 当前行 > 20 Then Exit Do
        End If
        
        单元格值 = ws.Cells(当前行, 列号).Value
    Loop
    
    ' 如果是空值或仍然是"上"/"下"，返回"-"
    If IsEmpty(单元格值) Or 单元格值 = "" Or 单元格值 = "上" Or 单元格值 = "下" Then
        获取单元格数值 = "-"
    Else
        获取单元格数值 = 单元格值
    End If
End Function

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




