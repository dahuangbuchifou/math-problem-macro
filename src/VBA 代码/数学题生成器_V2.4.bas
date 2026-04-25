' ============================================================
' 幼升小数学题生成器 - VBA 代码
' 版本：V2.4.20260425.1115
' 文件名：数学题生成器_V2.4.bas
' 作者：工部尚书
' 创建日期：2026-04-24 19:00
' 最后更新：2026-04-25 11:15
' 说明：专为幼升小儿童设计的 Excel 数学题生成工具
' 支持：100以内加减法、两位数、三位数、连加连减、混合运算
' 特性：难度分级、专项练习、A4排版、多页生成、答案隐藏
' ============================================================
' V2.4 更新日志：
'   【修复】布局理解错误（严重）：改为 25行×2/3/4列=50/75/100题/页（竖版）
'     - ROWS_PER_PAGE = 25（横向 25 道题）
'     - 每页列数可选：2列/3列/4列（H8 单元格下拉选择）
'     - 列宽 20（2-4 列均匀分布在 A4 纸上）
'     - 行高自适应：25 行均匀分布（约 18-25）
'     - 字体大小 12（适应窄列）
'   【修复】参数刷新 Bug（最严重）：GenerateQuestions 完全不调用 InitParameterPanel
'     - 在 Cells.ClearContents 之前保存所有用户参数
'     - 清除后只重建标签（G列），不覆盖 H 列的值
'     - 恢复用户之前选择的所有值
'     - 参数面板新增 H8"每页列数"选项（2/3/4），H9 改为"打印模式"
' V2.2 更新日志：
'   【修复】参数被刷新 Bug：GenerateQuestions 先保存用户选择再清除面板
'   【修复】打印布局：改为 5列×5行=25题/页，更符合 A4 打印比例
'   【新增】打印模式选项：彩色模式（屏幕查看）/ 黑白模式（打印省墨）
'   【优化】颜色方案：改为极淡背景 + 细边框，打印友好不浪费墨水
'   【优化】版面美感：增大行高、调整字体、页边距合理
' V2.1 更新日志：
'   - 修复参数面板被清空 Bug（InitParameterPanel 移至 Cells.ClearContents 之后）
'   - 修复背景颜色太淡 Bug（使用更明显的柔和颜色方案）
' V2.0 更新日志：
'   - 修复 For Each 循环变量冲突 Bug
'   - 修复答案计算重复执行 Bug
'   - 修复 GenerateMixed 可能返回空字符串 Bug
'   - 移除无意义的负数判断逻辑
'   - 新增两位数加法/减法专项
'   - 新增三位数混合运算专项
'   - 新增 A4 版面自动填充（25列×N行=50/75/100题/页）
'   - 新增多页生成（最多 5 页）
'   - 新增页间分隔线和自动分页符
'   - 更新难度等级和练习模式下拉菜单
'   - 新增 G7/H7 页数参数
' ============================================================

' ==================== 全局变量 ====================
Dim wsQuestion As Worksheet      ' 题目页工作表
Dim wsAnswer As Worksheet        ' 答案页工作表
Dim wsRecord As Worksheet        ' 练习记录工作表

' ==================== 页面布局常量 ====================
Const ROWS_PER_PAGE As Integer = 25    ' 每页行数（纵向 25 道题）
Const MAX_PAGES As Integer = 5         ' 最大页数

' ==================== 打印友好颜色方案 ====================
' V2.3: 极淡背景色 + 细边框，打印省墨
' 每页使用不同的颜色方案，页内每 5 题轮换
' 页面 1：蓝色系（极淡）
' 页面 2：绿色系（极淡）
' 页面 3：暖色系（极淡）
' 页面 4：紫色系（极淡）
' 页面 5：黄色系（极淡）

Function GetPageColor(pageNum As Integer, questionIndex As Integer) As Long
    ' 根据页码和题目索引返回极淡但可区分的颜色（V2.3 打印友好）
    Dim colorIndex As Integer
    colorIndex = ((questionIndex - 1) \ 5) Mod 5 + 1  ' 每页内 5 种颜色轮换
    
    Select Case pageNum
        Case 1  ' 蓝色系（极淡）
            Select Case colorIndex
                Case 1: GetPageColor = RGB(232, 243, 251)  ' 极淡蓝
                Case 2: GetPageColor = RGB(225, 240, 252)  ' 更淡蓝
                Case 3: GetPageColor = RGB(218, 238, 248)  ' 浅蓝
                Case 4: GetPageColor = RGB(228, 242, 250)  ' 淡蓝
                Case 5: GetPageColor = RGB(235, 245, 252)  ' 最淡蓝
            End Select
        Case 2  ' 绿色系（极淡）
            Select Case colorIndex
                Case 1: GetPageColor = RGB(230, 248, 230)  ' 极淡绿
                Case 2: GetPageColor = RGB(225, 245, 225)  ' 更淡绿
                Case 3: GetPageColor = RGB(235, 250, 235)  ' 浅绿
                Case 4: GetPageColor = RGB(228, 247, 228)  ' 淡绿
                Case 5: GetPageColor = RGB(238, 252, 238)  ' 最淡绿
            End Select
        Case 3  ' 暖色系（极淡）
            Select Case colorIndex
                Case 1: GetPageColor = RGB(255, 244, 232)  ' 极淡橙
                Case 2: GetPageColor = RGB(255, 240, 225)  ' 更淡橙
                Case 3: GetPageColor = RGB(255, 248, 238)  ' 浅橙
                Case 4: GetPageColor = RGB(255, 242, 230)  ' 淡橙
                Case 5: GetPageColor = RGB(255, 250, 242)  ' 最淡橙
            End Select
        Case 4  ' 紫色系（极淡）
            Select Case colorIndex
                Case 1: GetPageColor = RGB(245, 235, 248)  ' 极淡紫
                Case 2: GetPageColor = RGB(240, 230, 245)  ' 更淡紫
                Case 3: GetPageColor = RGB(248, 240, 250)  ' 浅紫
                Case 4: GetPageColor = RGB(243, 233, 247)  ' 淡紫
                Case 5: GetPageColor = RGB(250, 243, 252)  ' 最淡紫
            End Select
        Case 5  ' 黄色系（极淡）
            Select Case colorIndex
                Case 1: GetPageColor = RGB(255, 255, 229)  ' 极淡黄
                Case 2: GetPageColor = RGB(255, 255, 220)  ' 更淡黄
                Case 3: GetPageColor = RGB(255, 255, 235)  ' 浅黄
                Case 4: GetPageColor = RGB(255, 255, 225)  ' 淡黄
                Case 5: GetPageColor = RGB(255, 255, 240)  ' 最淡黄
            End Select
    End Select
End Function

' ==================== 初始化工作表 ====================
Sub InitializeSheets()
    On Error GoTo InitError
    
    ' 设置题目页
    Set wsQuestion = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "题目页" Then
            Set wsQuestion = ws
            Exit For
        End If
    Next ws
    If wsQuestion Is Nothing Then
        Set wsQuestion = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsQuestion.Name = "题目页"
    End If
    
    ' 设置答案页
    Set wsAnswer = Nothing
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "答案页" Then
            Set wsAnswer = ws
            Exit For
        End If
    Next ws
    If wsAnswer Is Nothing Then
        Set wsAnswer = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsAnswer.Name = "答案页"
    End If
    
    ' 设置练习记录页
    Set wsRecord = Nothing
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "练习记录" Then
            Set wsRecord = ws
            Exit For
        End If
    Next ws
    If wsRecord Is Nothing Then
        Set wsRecord = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsRecord.Name = "练习记录"
        ' 初始化记录表头
        wsRecord.Range("A1").Value = "日期时间"
        wsRecord.Range("B1").Value = "难度等级"
        wsRecord.Range("C1").Value = "练习模式"
        wsRecord.Range("D1").Value = "题目数量"
        wsRecord.Range("A1:D1").Font.Bold = True
        wsRecord.Range("A1:D1").Interior.Color = RGB(220, 230, 240)
    End If
    
    ' 隐藏答案页（深度隐藏）
    wsAnswer.Visible = xlSheetVeryHidden
    
    Exit Sub
    
InitError:
    MsgBox "初始化工作表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ==================== 初始化参数面板 ====================
Sub InitParameterPanel()
    On Error Resume Next
    
    With wsQuestion
        ' 参数标签（G列）— V2.3: H8=每页列数, H9=打印模式
        .Range("G1").Value = "最大数字范围"
        .Range("G2").Value = "题目数量"
        .Range("G3").Value = "负数概率 (%)"
        .Range("G4").Value = "当前状态"
        .Range("G5").Value = "难度等级"
        .Range("G6").Value = "练习模式"
        .Range("G7").Value = "生成页数"
        .Range("G8").Value = "每页列数"
        .Range("G9").Value = "打印模式"
        
        ' 参数值（H列）- 默认值
        .Range("H1").Value = 100
        .Range("H2").Value = 50
        .Range("H3").Value = 0
        .Range("H4").Value = "就绪"
        .Range("H4").Interior.Color = RGB(200, 230, 255)
        .Range("H5").Value = "中级"
        .Range("H6").Value = "混合运算"
        .Range("H7").Value = 1
        .Range("H8").Value = 3
        .Range("H9").Value = "彩色"
        
        ' 设置标签样式
        .Range("G1:G9").Font.Bold = True
        .Range("G1:G9").Font.Size = 11
        .Range("G1:G9").HorizontalAlignment = xlRight
        .Range("G1:G9").VerticalAlignment = xlCenter
        
        ' 设置参数值样式
        .Range("H1:H9").Font.Size = 11
        .Range("H1:H9").HorizontalAlignment = xlCenter
        .Range("H1:H9").VerticalAlignment = xlCenter
        .Range("H1:H9").Borders.LineStyle = xlContinuous
        .Range("H1:H9").Borders.Color = RGB(200, 200, 200)
        
        ' 设置列宽
        .Columns("G").ColumnWidth = 14
        .Columns("H").ColumnWidth = 12
    End With
    
    ' 设置难度等级下拉菜单（H5）
    With wsQuestion.Range("H5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="初级,中级,高级,两位数,三位数"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置练习模式下拉菜单（H6）
    With wsQuestion.Range("H6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="混合运算,连加专项,连减专项,进位加法,退位减法,两位数加法,两位数减法,三位数混合"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置页数下拉菜单（H7）
    With wsQuestion.Range("H7").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="1,2,3,4,5"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置每页列数下拉菜单（H8）— V2.3 新增
    With wsQuestion.Range("H8").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="2,3,4"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置打印模式下拉菜单（H9）— V2.3 从 H8 移到 H9
    With wsQuestion.Range("H9").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="彩色,黑白"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    On Error GoTo 0
End Sub

' ==================== 获取难度参数 ====================
Sub GetDifficultyParams(ByRef maxNum As Integer, ByRef minNum As Integer, ByRef noCarry As Boolean, ByRef noBorrow As Boolean)
    Dim difficulty As String
    difficulty = Trim(wsQuestion.Range("H5").Value)
    
    ' 默认值
    maxNum = 100
    minNum = 1
    noCarry = False
    noBorrow = False
    
    Select Case difficulty
        Case "初级"
            maxNum = 20
            minNum = 1
            noCarry = True
            noBorrow = True
        Case "中级"
            maxNum = 50
            minNum = 1
            noCarry = False
            noBorrow = False
        Case "高级"
            maxNum = 100
            minNum = 1
            noCarry = False
            noBorrow = False
        Case "两位数"
            maxNum = 99
            minNum = 10
            noCarry = False
            noBorrow = False
        Case "三位数"
            maxNum = 999
            minNum = 100
            noCarry = False
            noBorrow = False
        Case Else
            maxNum = 100
            minNum = 1
            noCarry = False
            noBorrow = False
    End Select
End Sub

' ==================== 生成单个加法题目 ====================
Function GenerateAddition(maxNum As Integer, minNum As Integer, noCarry As Boolean) As String
    Dim a As Integer, b As Integer, result As Integer
    Dim attempts As Integer
    attempts = 0
    
    Do
        a = Int(Rnd() * (maxNum - minNum + 1)) + minNum
        b = Int(Rnd() * (maxNum - a + 1)) + 1
        result = a + b
        attempts = attempts + 1
        
        ' 防止死循环
        If attempts > 1000 Then
            GenerateAddition = ""
            Exit Function
        End If
        
        ' 不进位检查
        If noCarry Then
            If (a Mod 10) + (b Mod 10) >= 10 Then
                ' 需要进位，跳过
            Else
                Exit Do  ' 不进位，可以接受
            End If
        Else
            Exit Do  ' 允许进位，直接接受
        End If
    Loop While result > maxNum
    
    GenerateAddition = a & " + " & b & " = "
End Function

' ==================== 生成单个减法题目 ====================
Function GenerateSubtraction(maxNum As Integer, minNum As Integer, noBorrow As Boolean) As String
    Dim a As Integer, b As Integer
    Dim attempts As Integer
    attempts = 0
    
    Do
        a = Int(Rnd() * (maxNum - minNum + 1)) + minNum
        b = Int(Rnd() * (a - minNum)) + minNum
        attempts = attempts + 1
        
        ' 防止死循环
        If attempts > 1000 Then
            GenerateSubtraction = ""
            Exit Function
        End If
        
        ' 不退位检查
        If noBorrow Then
            If (a Mod 10) < (b Mod 10) Then
                ' 需要退位，跳过
            Else
                Exit Do  ' 不退位，可以接受
            End If
        Else
            Exit Do  ' 允许退位，直接接受
        End If
    Loop While b > a Or b < minNum
    
    GenerateSubtraction = a & " - " & b & " = "
End Function

' ==================== 生成连加题目 ====================
Function GenerateChainAdd(maxNum As Integer, minNum As Integer, noCarry As Boolean) As String
    Dim a As Integer, b As Integer, c As Integer, result As Integer
    Dim maxSingle As Integer
    Dim attempts As Integer
    attempts = 0
    
    ' 每个数最大约为 maxNum/3
    maxSingle = (maxNum - 3 * minNum) \ 3 + minNum
    If maxSingle < minNum Then maxSingle = minNum
    
    Do
        a = Int(Rnd() * (maxSingle - minNum + 1)) + minNum
        b = Int(Rnd() * (maxSingle - minNum + 1)) + minNum
        c = maxNum - a - b
        If c < minNum Then c = minNum
        result = a + b + c
        attempts = attempts + 1
        
        ' 防止死循环
        If attempts > 1000 Then
            GenerateChainAdd = ""
            Exit Function
        End If
        
        ' 不进位检查
        If noCarry Then
            If (a Mod 10) + (b Mod 10) + (c Mod 10) >= 10 Then
                ' 需要进位，跳过
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop While result > maxNum
    
    GenerateChainAdd = a & " + " & b & " + " & c & " = "
End Function

' ==================== 生成连减题目 ====================
Function GenerateChainSub(maxNum As Integer, minNum As Integer, noBorrow As Boolean) As String
    Dim a As Integer, b As Integer, c As Integer, remainder As Integer
    Dim attempts As Integer
    attempts = 0
    
    Do
        ' a 为被减数，范围 [minNum+2, maxNum]
        a = Int(Rnd() * (maxNum - minNum - 1)) + minNum + 2
        ' b 为第一个减数，范围 [minNum, a-minNum-1]
        b = Int(Rnd() * (a - minNum - 1)) + minNum
        ' remainder = a - b，c 从 [minNum, remainder-minNum] 中选
        remainder = a - b
        If remainder <= minNum + 1 Then
            c = minNum
        Else
            c = Int(Rnd() * (remainder - minNum - 1)) + minNum
        End If
        attempts = attempts + 1
        
        ' 防止死循环
        If attempts > 1000 Then
            GenerateChainSub = ""
            Exit Function
        End If
    Loop While a - b - c < 0 Or b >= a Or c >= (a - b)
    
    GenerateChainSub = a & " - " & b & " - " & c & " = "
End Function

' ==================== 生成混合运算题目 ====================
Function GenerateMixed(maxNum As Integer, minNum As Integer, noCarry As Boolean, noBorrow As Boolean) As String
    Dim opType As Integer
    Dim result As String
    Dim attempts As Integer
    attempts = 0
    
    Do
        opType = Int(Rnd() * 3)  ' 0:加法, 1:减法, 2:连加/连减
        result = ""
        
        Select Case opType
            Case 0
                result = GenerateAddition(maxNum, minNum, noCarry)
                If result = "" Then result = GenerateAddition(maxNum, minNum, False)
            Case 1
                result = GenerateSubtraction(maxNum, minNum, noBorrow)
                If result = "" Then result = GenerateSubtraction(maxNum, minNum, False)
            Case 2
                If Rnd() > 0.5 Then
                    result = GenerateChainAdd(maxNum, minNum, noCarry)
                    If result = "" Then result = GenerateChainAdd(maxNum, minNum, False)
                Else
                    result = GenerateChainSub(maxNum, minNum, noBorrow)
                End If
        End Select
        
        attempts = attempts + 1
        If attempts > 50 Then
            GenerateMixed = ""
            Exit Function
        End If
    Loop While result = ""
    
    GenerateMixed = result
End Function

' ==================== 生成两位数加法题目 ====================
Function GenerateTwoDigitAdd() As String
    Dim a As Integer, b As Integer, result As Integer
    Dim attempts As Integer
    attempts = 0
    
    Do
        a = Int(Rnd() * 90) + 10   ' 10-99
        b = Int(Rnd() * (99 - a + 1)) + 1  ' 确保 a+b <= 99
        result = a + b
        attempts = attempts + 1
        
        If attempts > 1000 Then
            GenerateTwoDigitAdd = ""
            Exit Function
        End If
    Loop While result > 99
    
    GenerateTwoDigitAdd = a & " + " & b & " = "
End Function

' ==================== 生成两位数减法题目 ====================
Function GenerateTwoDigitSub() As String
    Dim a As Integer, b As Integer
    Dim attempts As Integer
    attempts = 0
    
    Do
        a = Int(Rnd() * 90) + 10   ' 10-99
        b = Int(Rnd() * (a - 10)) + 10  ' 10 到 a-1
        attempts = attempts + 1
        
        If attempts > 1000 Then
            GenerateTwoDigitSub = ""
            Exit Function
        End If
    Loop While b >= a Or b < 10
    
    GenerateTwoDigitSub = a & " - " & b & " = "
End Function

' ==================== 生成三位数混合运算题目 ====================
Function GenerateThreeDigitMixed() As String
    Dim a As Integer, b As Integer, op As String
    Dim result As Integer
    Dim attempts As Integer
    attempts = 0
    
    Do
        a = Int(Rnd() * 900) + 100   ' 100-999
        
        If Rnd() > 0.5 Then
            ' 加法：a + b，确保结果 <= 999
            b = Int(Rnd() * (999 - a + 1)) + 1
            result = a + b
            op = "+"
        Else
            ' 减法：a - b，确保 b < a 且 b >= 100
            If a <= 200 Then
                b = Int(Rnd() * (a - 100)) + 100
            Else
                b = Int(Rnd() * (a - 100)) + 100
            End If
            result = a - b
            op = "-"
        End If
        
        attempts = attempts + 1
        If attempts > 1000 Then
            GenerateThreeDigitMixed = ""
            Exit Function
        End If
    Loop While result < 0 Or result > 999
    
    GenerateThreeDigitMixed = a & " " & op & " " & b & " = "
End Function

' ==================== 提取题目答案 ====================
Function ExtractAnswer(questionText As String) As Integer
    Dim parts() As String
    Dim expr As String
    Dim tokens() As String
    Dim result As Integer
    Dim t As Integer
    Dim i As Integer
    
    ' 分割题目和答案标记
    parts = Split(questionText, " = ")
    If UBound(parts) < 0 Then
        ExtractAnswer = 0
        Exit Function
    End If
    
    expr = parts(0)
    
    ' 解析表达式，从左到右计算
    tokens = Split(expr, " ")
    If UBound(tokens) < 0 Then
        ExtractAnswer = 0
        Exit Function
    End If
    
    result = Val(tokens(0))
    
    i = 1
    Do While i < UBound(tokens)
        If tokens(i) = "+" Then
            result = result + Val(tokens(i + 1))
        ElseIf tokens(i) = "-" Then
            result = result - Val(tokens(i + 1))
        End If
        i = i + 2
    Loop
    
    ExtractAnswer = result
End Function

' ==================== 核心功能：生成题目 ====================
Sub GenerateQuestions()
    On Error GoTo ErrorHandler
    
    Dim maxNum As Integer
    Dim minNum As Integer
    Dim noCarry As Boolean
    Dim noBorrow As Boolean
    Dim questionCount As Integer
    Dim practiceMode As String
    Dim difficulty As String
    Dim pageNum As Integer
    Dim totalPages As Integer
    Dim i As Integer
    Dim col As Integer
    Dim row As Integer
    Dim questionText As String
    Dim answerText As Integer
    Dim a As Integer, b As Integer
    Dim recordRow As Long
    Dim questionNum As Integer
    Dim currentPage As Integer
    Dim pageQuestionIndex As Integer
    Dim sepRow As Integer
    Dim totalRows As Integer
    Dim j As Integer
    Dim printMode As String
    Dim colsPerPage As Integer       ' V2.3: 每页列数（动态）
    Dim questionsPerPage As Integer  ' V2.3: 每页题目数（动态 = 25 × colsPerPage）
    Dim targetRowHeight As Integer   ' V2.3: 目标行高（根据 colsPerPage 自适应）
    Dim p As Integer
    Dim colInPage As Integer
    Dim isSep As Boolean
    
    ' ============================================================
    ' ★ V2.3 修复：在 Cells.ClearContents 之前保存所有用户参数
    ' ============================================================
    difficulty = Trim(wsQuestion.Range("H5").Value)
    practiceMode = Trim(wsQuestion.Range("H6").Value)
    questionCount = Val(wsQuestion.Range("H2").Value)
    totalPages = Val(wsQuestion.Range("H7").Value)
    colsPerPage = Val(wsQuestion.Range("H8").Value)
    printMode = Trim(wsQuestion.Range("H9").Value)
    
    ' 初始化工作表
    InitializeSheets
    
    ' ============================================================
    ' ★ V2.3 修复：完全不调用 InitParameterPanel！
    ' 只调用 InitParameterPanel_NoReset 重建标签和样式
    ' ============================================================
    
    ' 清除旧数据
    wsQuestion.Cells.ClearContents
    wsQuestion.Cells.ClearFormats
    wsQuestion.Cells.Interior.ColorIndex = xlNone
    wsAnswer.Cells.ClearContents
    wsAnswer.Cells.ClearFormats
    wsAnswer.Cells.Interior.ColorIndex = xlNone
    
    ' 只重建标签（不调用 InitParameterPanel）
    InitParameterPanel_NoReset
    
    ' ★ V2.3 修复：恢复用户之前选择的所有值
    wsQuestion.Range("H5").Value = difficulty
    wsQuestion.Range("H6").Value = practiceMode
    wsQuestion.Range("H2").Value = questionCount
    wsQuestion.Range("H7").Value = totalPages
    wsQuestion.Range("H8").Value = colsPerPage
    wsQuestion.Range("H9").Value = printMode
    
    ' ============================================================
    ' V2.3: 根据 colsPerPage 计算每页题目数
    ' ============================================================
    ' 校验 colsPerPage
    If colsPerPage < 2 Then colsPerPage = 2
    If colsPerPage > 4 Then colsPerPage = 4
    
    questionsPerPage = ROWS_PER_PAGE * colsPerPage  ' 25×2=50, 25×3=75, 25×4=100（竖版）
    
    ' 参数校验
    If questionCount <= 0 Then questionCount = questionsPerPage
    If questionCount > questionsPerPage * MAX_PAGES Then questionCount = questionsPerPage * MAX_PAGES
    If totalPages < 1 Then totalPages = 1
    If totalPages > MAX_PAGES Then totalPages = MAX_PAGES
    
    ' 根据页数调整题目数量
    If totalPages > 1 And questionCount < totalPages * questionsPerPage Then
        questionCount = totalPages * questionsPerPage
    End If
    
    ' 获取难度参数
    GetDifficultyParams maxNum, minNum, noCarry, noBorrow
    
    ' 更新状态
    wsQuestion.Range("H4").Value = "生成中..."
    wsQuestion.Range("H4").Interior.Color = RGB(255, 250, 200)
    
    ' 清除旧的分页符
    wsQuestion.ResetAllPageBreaks
    wsAnswer.ResetAllPageBreaks
    
    ' 计算总行数（每页 colsPerPage 行 + 页间分隔行）
    totalRows = totalPages * ROWS_PER_PAGE + (totalPages - 1)
    
    ' ==================== 设置题目页格式 ====================
    With wsQuestion
        .Cells.Font.Name = "微软雅黑"
        .Cells.Font.Size = 12   ' V2.3: 12号字体（适应 25 列窄布局）
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
        
        ' 设置列宽（V2.3: 25 列挤在 A4 纸上，列宽 13）
        For j = 1 To colsPerPage
            .Columns(j).ColumnWidth = 20
        Next j
        
        ' 设置行高（V2.3: 根据 colsPerPage 自适应）
        Select Case 25
            Case 25: targetRowHeight = 18
            Case Else: targetRowHeight = 18
            Case Else: targetRowHeight = 22
        End Select
        
        For j = 1 To totalRows
            ' 判断是否为分隔行
            isSep = False
            For p = 1 To totalPages - 1
                If j = p * (colsPerPage + 1) Then
                    isSep = True
                    Exit For
                End If
            Next p
            
            If isSep Then
                .Rows(j).RowHeight = 8  ' 分隔行更窄
            Else
                .Rows(j).RowHeight = targetRowHeight
            End If
        Next j
        
        ' 设置页间分隔线
        For p = 1 To totalPages - 1
            sepRow = p * (ROWS_PER_PAGE + 1)
            With .Rows(sepRow)
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
                .Interior.Color = RGB(250, 250, 250)
            End With
        Next p
        
        ' 设置分页符（在分隔行后插入水平分页符）
        For p = 2 To totalPages
            sepRow = (p - 1) * (ROWS_PER_PAGE + 1) + 1
            .HPageBreaks.Add Before:=.Rows(sepRow)
        Next p
        
        ' 设置打印页面
        .PageSetup.Orientation = xlPortrait
        .PageSetup.PaperSize = xlPaperA4
        .PageSetup.CenterHorizontally = True
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = totalPages
        .PageSetup.TopMargin = Application.CentimetersToPoints(1.2)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1.2)
        .PageSetup.LeftMargin = Application.CentimetersToPoints(0.8)
        .PageSetup.RightMargin = Application.CentimetersToPoints(0.8)
    End With
    
    ' ==================== 设置答案页格式 ====================
    With wsAnswer
        .Cells.Font.Name = "微软雅黑"
        .Cells.Font.Size = 12
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
        .Cells.Font.Color = RGB(200, 50, 50)  ' 答案用红色
        
        For j = 1 To colsPerPage
            .Columns(j).ColumnWidth = 20
        Next j
        
        For j = 1 To totalRows
            .Rows(j).RowHeight = targetRowHeight
        Next j
    End With
    
    ' ==================== 生成题目 ====================
    questionNum = 0
    
    For i = 1 To questionCount
        ' 计算当前题目所在的页码和页内索引
        currentPage = ((i - 1) \ questionsPerPage) + 1
        pageQuestionIndex = ((i - 1) Mod questionsPerPage) + 1
        
        ' 计算在工作表中的列和行
        row = ((pageQuestionIndex - 1) Mod ROWS_PER_PAGE) + 1
        colInPage = ((pageQuestionIndex - 1) \ ROWS_PER_PAGE) + 1
        
        ' 考虑分隔行偏移
        row = row + (currentPage - 1) * (ROWS_PER_PAGE + 1)
        col = colInPage
        
        ' 根据练习模式生成题目
        questionText = ""
        
        Select Case practiceMode
            Case "连加专项"
                questionText = GenerateChainAdd(maxNum, minNum, noCarry)
                If questionText = "" Then questionText = GenerateChainAdd(maxNum, minNum, False)
                
            Case "连减专项"
                questionText = GenerateChainSub(maxNum, minNum, noBorrow)
                
            Case "进位加法"
                ' 强制生成进位加法
                Do
                    a = Int(Rnd() * (maxNum - minNum + 1)) + minNum
                    b = Int(Rnd() * (maxNum - a)) + 1
                    If (a Mod 10) + (b Mod 10) >= 10 Then Exit Do
                Loop While a + b > maxNum
                questionText = a & " + " & b & " = "
                
            Case "退位减法"
                ' 强制生成退位减法
                Do
                    a = Int(Rnd() * (maxNum - minNum + 1)) + minNum
                    b = Int(Rnd() * (a - minNum)) + minNum
                    If (a Mod 10) < (b Mod 10) Then Exit Do
                Loop While b >= a
                questionText = a & " - " & b & " = "
                
            Case "两位数加法"
                questionText = GenerateTwoDigitAdd()
                
            Case "两位数减法"
                questionText = GenerateTwoDigitSub()
                
            Case "三位数混合"
                questionText = GenerateThreeDigitMixed()
                
            Case Else  ' 混合运算
                questionText = GenerateMixed(maxNum, minNum, noCarry, noBorrow)
        End Select
        
        ' 防护：如果题目生成失败，重试
        If questionText = "" Then
            questionText = GenerateMixed(maxNum, minNum, False, False)
            If questionText = "" Then
                ' 极端情况：生成简单题目
                a = Int(Rnd() * (maxNum - minNum)) + minNum
                b = Int(Rnd() * (maxNum - a)) + 1
                If a + b <= maxNum Then
                    questionText = a & " + " & b & " = "
                Else
                    questionText = maxNum & " - " & minNum & " = "
                End If
            End If
        End If
        
        ' 提取答案
        answerText = ExtractAnswer(questionText)
        
        ' ★ V2.3: 根据打印模式设置颜色
        If printMode = "黑白" Then
            ' 黑白模式：无背景色，只加细边框
            wsQuestion.Cells(row, col).Interior.ColorIndex = xlNone
            wsQuestion.Cells(row, col).Borders.LineStyle = xlContinuous
            wsQuestion.Cells(row, col).Borders.Color = RGB(180, 180, 180)
            wsQuestion.Cells(row, col).Borders.Weight = xlThin
            
            wsAnswer.Cells(row, col).Interior.ColorIndex = xlNone
            wsAnswer.Cells(row, col).Borders.LineStyle = xlContinuous
            wsAnswer.Cells(row, col).Borders.Color = RGB(180, 180, 180)
            wsAnswer.Cells(row, col).Borders.Weight = xlThin
        Else
            ' 彩色模式：极淡背景 + 细边框
            wsQuestion.Cells(row, col).Interior.Color = GetPageColor(currentPage, pageQuestionIndex)
            wsQuestion.Cells(row, col).Borders.LineStyle = xlContinuous
            wsQuestion.Cells(row, col).Borders.Color = RGB(200, 210, 220)
            wsQuestion.Cells(row, col).Borders.Weight = xlHairline
            
            wsAnswer.Cells(row, col).Interior.Color = GetPageColor(currentPage, pageQuestionIndex)
            wsAnswer.Cells(row, col).Borders.LineStyle = xlContinuous
            wsAnswer.Cells(row, col).Borders.Color = RGB(200, 210, 220)
            wsAnswer.Cells(row, col).Borders.Weight = xlHairline
        End If
        
        ' 写入题目
        wsQuestion.Cells(row, col).Value = questionText
        
        ' 写入答案
        wsAnswer.Cells(row, col).Value = answerText
        
        questionNum = questionNum + 1
    Next i
    
    ' ==================== 更新状态 ====================
    wsQuestion.Range("H4").Value = "已完成（" & questionNum & "题，" & totalPages & "页）"
    wsQuestion.Range("H4").Interior.Color = RGB(200, 240, 200)
    
    ' ==================== 记录练习信息 ====================
    recordRow = wsRecord.Cells(wsRecord.Rows.Count, 1).End(xlUp).Row + 1
    wsRecord.Cells(recordRow, 1).Value = Now()
    wsRecord.Cells(recordRow, 2).Value = difficulty
    wsRecord.Cells(recordRow, 3).Value = practiceMode
    wsRecord.Cells(recordRow, 4).Value = questionNum
    
    ' ==================== 完成提示 ====================
    MsgBox "✅ 题目生成完成！" & vbCrLf & _
           "难度：" & difficulty & vbCrLf & _
           "模式：" & practiceMode & vbCrLf & _
           "每页列数：" & colsPerPage & vbCrLf & _
           "题目数：" & questionNum & " 道" & vbCrLf & _
           "页数：" & totalPages & " 页" & vbCrLf & _
           "打印：" & printMode & "模式", _
           vbInformation, "生成成功"
    
    Exit Sub
    
ErrorHandler:
    wsQuestion.Range("H4").Value = "错误"
    wsQuestion.Range("H4").Interior.Color = RGB(255, 200, 200)
    MsgBox "❌ 生成题目时出错：" & vbCrLf & Err.Description, vbCritical, "错误"
End Sub

' ==================== 初始化参数面板（不覆盖用户值） ====================
' V2.3: 只设置标签、样式和下拉菜单，不覆盖 H 列已有值
' 参数面板布局：G8=每页列数, G9=打印模式
Sub InitParameterPanel_NoReset()
    On Error Resume Next
    
    With wsQuestion
        ' 只设置标签（不覆盖 H 列的值）
        .Range("G1").Value = "最大数字范围"
        .Range("G2").Value = "题目数量"
        .Range("G3").Value = "负数概率 (%)"
        .Range("G4").Value = "当前状态"
        .Range("G5").Value = "难度等级"
        .Range("G6").Value = "练习模式"
        .Range("G7").Value = "生成页数"
        .Range("G8").Value = "每页列数"
        .Range("G9").Value = "打印模式"
        
        ' 只设置标签样式
        .Range("G1:G9").Font.Bold = True
        .Range("G1:G9").Font.Size = 11
        .Range("G1:G9").HorizontalAlignment = xlRight
        .Range("G1:G9").VerticalAlignment = xlCenter
        
        ' 只设置参数值区域的样式（不覆盖值！）
        .Range("H1:H9").Font.Size = 11
        .Range("H1:H9").HorizontalAlignment = xlCenter
        .Range("H1:H9").VerticalAlignment = xlCenter
        .Range("H1:H9").Borders.LineStyle = xlContinuous
        .Range("H1:H9").Borders.Color = RGB(200, 200, 200)
        
        ' 设置列宽
        .Columns("G").ColumnWidth = 14
        .Columns("H").ColumnWidth = 12
    End With
    
    ' 设置难度等级下拉菜单（H5）
    With wsQuestion.Range("H5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="初级,中级,高级,两位数,三位数"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置练习模式下拉菜单（H6）
    With wsQuestion.Range("H6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="混合运算,连加专项,连减专项,进位加法,退位减法,两位数加法,两位数减法,三位数混合"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置页数下拉菜单（H7）
    With wsQuestion.Range("H7").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="1,2,3,4,5"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置每页列数下拉菜单（H8）— V2.3 新增
    With wsQuestion.Range("H8").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="2,3,4"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 设置打印模式下拉菜单（H9）— V2.3 从 H8 移到 H9
    With wsQuestion.Range("H9").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="彩色,黑白"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    On Error GoTo 0
End Sub

' ==================== 显示/隐藏答案 ====================
Sub ToggleAnswerSheet()
    On Error Resume Next
    
    InitializeSheets
    
    If wsAnswer.Visible = xlSheetVeryHidden Or wsAnswer.Visible = xlSheetHidden Then
        wsAnswer.Visible = xlSheetVisible
        wsAnswer.Activate
        MsgBox "📖 答案已显示", vbInformation, "提示"
    Else
        wsAnswer.Visible = xlSheetVeryHidden
        wsQuestion.Activate
        MsgBox "🔒 答案已隐藏", vbInformation, "提示"
    End If
    
    On Error GoTo 0
End Sub

' ==================== 打印预览 ====================
Sub PrintPreviewSheet()
    On Error GoTo ErrorHandler
    
    InitializeSheets
    
    If wsQuestion.Range("A1").Value = "" Then
        MsgBox "⚠️ 请先生成题目！", vbExclamation, "提示"
        Exit Sub
    End If
    
    wsQuestion.PrintPreview
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ 打印预览出错：" & vbCrLf & Err.Description, vbCritical, "错误"
End Sub

' ==================== 重置设置 ====================
Sub ResetSettings()
    On Error Resume Next
    
    InitializeSheets
    
    ' 清除题目
    wsQuestion.Cells.ClearContents
    wsQuestion.Cells.ClearFormats
    wsQuestion.Cells.Interior.ColorIndex = xlNone
    wsAnswer.Visible = xlSheetVeryHidden
    
    ' 重新初始化参数面板（重置时会覆盖为默认值，这是预期行为）
    InitParameterPanel
    
    ' 清除分页符
    wsQuestion.ResetAllPageBreaks
    
    MsgBox "✅ 设置已重置为默认值！" & vbCrLf & _
           "难度：中级 | 模式：混合运算 | 行数：3 | 题数：50 | 页数：1 | 打印：彩色", _
           vbInformation, "重置成功"
    
    On Error GoTo 0
End Sub

' ==================== 查看记录 ====================
Sub ViewRecords()
    On Error Resume Next
    
    InitializeSheets
    
    wsRecord.Visible = xlSheetVisible
    wsRecord.Activate
    
    ' 自动调整列宽
    wsRecord.Columns("A:D").AutoFit
    
    MsgBox "📊 已打开练习记录表", vbInformation, "提示"
    
    On Error GoTo 0
End Sub

' ==================== 导出PDF ====================
Sub ExportToPDF()
    On Error GoTo ErrorHandler
    
    InitializeSheets
    
    If wsQuestion.Range("A1").Value = "" Then
        MsgBox "⚠️ 请先生成题目！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Dim fileName As String
    Dim filePath As String
    Dim difficulty As String
    Dim currentDate As String
    
    difficulty = Trim(wsQuestion.Range("H5").Value)
    currentDate = Format(Now(), "yyyy-mm-dd")
    fileName = "数学题_" & difficulty & "_" & currentDate & ".pdf"
    filePath = Application.DefaultFilePath & "\" & fileName
    
    ' 导出PDF
    wsQuestion.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    MsgBox "✅ PDF 导出成功！" & vbCrLf & _
           "文件路径：" & filePath, _
           vbInformation, "导出成功"
    
    ' 打开所在文件夹
    Shell "explorer.exe /select," & filePath, vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ 导出PDF出错：" & vbCrLf & Err.Description, vbCritical, "错误"
End Sub

' ==================== 工作簿打开时初始化 ====================
Private Sub Workbook_Open()
    On Error Resume Next
    Randomize  ' 初始化随机数种子
    InitializeSheets
    InitParameterPanel
    On Error GoTo 0
End Sub

' ==================== 工作簿关闭时保存 ====================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ' 确保答案页隐藏
    If Not wsAnswer Is Nothing Then
        wsAnswer.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0
End Sub
