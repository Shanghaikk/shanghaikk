' ============================================================
' 理货表转化宏 - 替代按键精灵
' 用法: 打开原始理货表.xls → Alt+F11 打开VB编辑器
'       → 插入模块 → 粘贴本代码 → F5运行
' 要求: WPS Excel / Microsoft Excel
' ============================================================

Sub ConvertTallySheet()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim r As Long
    
    ' ─── 1. 读取当前数据 ──────────────────────
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    MsgBox "找到 " & (lastRow - 1) & " 行数据，开始转化...", vbInformation, "理货表转化"
    
    ' ─── 2. 创建新工作簿 ──────────────────────
    Set newWb = Workbooks.Add
    Set newWs = newWb.Sheets(1)
    newWs.Name = "理货表整理"
    
    ' ─── 3. 写入表头 ──────────────────────────
    Dim headers As Variant
    headers = Array("区域", "服务航线", "港区", "出船/共舱", "箱代理", _
                    "20E", "20F", "40E", "40F", "45E", "45F", _
                    "OtherE", "OtherF", "总自然箱量", "总TEU箱量", _
                    "持箱人", "服务航线中文", "船公司", "船代理")
    
    For i = 0 To UBound(headers)
        newWs.Cells(1, i + 1).Value = headers(i)
        ' 表头样式
        With newWs.Cells(1, i + 1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(21, 101, 192)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' ─── 4. 逐行转化 ──────────────────────────
    ' 原始列索引（基于 1）:
    ' A=区域(1) B=服务航线(2) C=船公司(3) D=船代理(4)
    ' E=持箱人(5) F=箱代理(6) G=港区(7) H=计划类型(8)
    ' I=内外贸(9) J=20E(10) K=20F(11) L=40E(12) M=40F(13)
    ' N=45E(14) O=45F(15) P=OtherE(16) Q=OtherF(17)
    ' R=总自然箱量(18) S=总TEU箱量(19)
    
    Dim src As Variant
    Dim arrArea, arrRoute
    Dim areaStr As String, routeStr As String
    Dim routeCode As String, routeName As String
    Dim portCode As String, holderStr As String
    Dim shipStr As String, shipType As String
    Dim agentStr As String, shipAgentStr As String
    
    r = 2 ' 新表从第2行开始
    
    For i = 2 To lastRow
        ' 读取原始行数据
        src = ws.Rows(i).Value
        
        ' 跳过全空行
        If IsEmpty(src(1, 1)) And IsEmpty(src(1, 2)) Then GoTo NextRow
        
        ' ═══ 区域转中文 ═══
        areaStr = CStr(src(1, 1))
        If InStr(areaStr, "/") > 0 Then
            arrArea = Split(areaStr, "/")
            areaStr = arrArea(1) ' 取中文名
        End If
        
        ' ═══ 服务航线拆分 ═══
        routeStr = CStr(src(1, 2))
        routeCode = routeStr
        routeName = ""
        If InStr(routeStr, "/") > 0 Then
            arrRoute = Split(routeStr, "/")
            routeCode = arrRoute(0)
            If UBound(arrRoute) >= 1 Then routeName = arrRoute(1)
        End If
        
        ' ═══ 港区映射 ═══
        portCode = MapPort(CStr(src(1, 7)))
        
        ' ═══ 出船/共舱 ═══
        shipStr = SimplifyCompany(CStr(src(1, 3)))
        holderStr = SimplifyCompany(CStr(src(1, 5)))
        shipType = "共舱"
        
        Dim shipCode As String, holderCode As String
        shipCode = CStr(src(1, 3))
        holderCode = CStr(src(1, 5))
        If InStr(shipCode, "/") > 0 Then shipCode = Split(shipCode, "/")(0)
        If InStr(holderCode, "/") > 0 Then holderCode = Split(holderCode, "/")(0)
        If shipCode = holderCode Or shipStr = holderStr Then
            shipType = "出船"
        End If
        
        ' ═══ 箱代理简化 ═══
        agentStr = SimplifyAgent(CStr(src(1, 6)))
        
        ' ═══ 持箱人简化 ═══
        holderStr = SimplifyCompany(CStr(src(1, 5)))
        
        ' ═══ 船公司简化 ═══
        shipStr = SimplifyCompany(CStr(src(1, 3)))
        
        ' ═══ 船代理简化 ═══
        shipAgentStr = SimplifyAgent(CStr(src(1, 4)))
        
        ' ═══ 写入新行 ═══
        newWs.Cells(r, 1).Value = areaStr
        newWs.Cells(r, 2).Value = routeCode
        newWs.Cells(r, 3).Value = portCode
        newWs.Cells(r, 4).Value = shipType
        newWs.Cells(r, 5).Value = agentStr
        newWs.Cells(r, 6).Value = NumericVal(src(1, 10)) ' 20E
        newWs.Cells(r, 7).Value = NumericVal(src(1, 11)) ' 20F
        newWs.Cells(r, 8).Value = NumericVal(src(1, 12)) ' 40E
        newWs.Cells(r, 9).Value = NumericVal(src(1, 13)) ' 40F
        newWs.Cells(r, 10).Value = NumericVal(src(1, 14)) ' 45E
        newWs.Cells(r, 11).Value = NumericVal(src(1, 15)) ' 45F
        newWs.Cells(r, 12).Value = NumericVal(src(1, 16)) ' OtherE
        newWs.Cells(r, 13).Value = NumericVal(src(1, 17)) ' OtherF
        newWs.Cells(r, 14).Value = NumericVal(src(1, 18)) ' 总自然箱量
        newWs.Cells(r, 15).Value = NumericVal(src(1, 19)) ' 总TEU箱量
        newWs.Cells(r, 16).Value = holderStr
        newWs.Cells(r, 17).Value = routeName
        newWs.Cells(r, 18).Value = shipStr
        newWs.Cells(r, 19).Value = shipAgentStr
        
        ' 数字右对齐
        Dim c As Integer
        For c = 6 To 15
            newWs.Cells(r, c).HorizontalAlignment = xlRight
        Next c
        
        r = r + 1
        
NextRow:
    Next i
    
    ' ─── 5. 调整列宽 ──────────────────────────
    newWs.Columns("A").ColumnWidth = 10
    newWs.Columns("B").ColumnWidth = 10
    newWs.Columns("C").ColumnWidth = 8
    newWs.Columns("D").ColumnWidth = 10
    newWs.Columns("E").ColumnWidth = 12
    newWs.Columns("F:K").ColumnWidth = 8
    newWs.Columns("L:M").ColumnWidth = 8
    newWs.Columns("N:O").ColumnWidth = 12
    newWs.Columns("P:R").ColumnWidth = 22
    newWs.Columns("S").ColumnWidth = 16
    
    ' ─── 6. 冻结首行 + 筛选 ───────────────────
    newWs.Activate
    newWs.Range("A2").Select
    ActiveWindow.FreezePanes = True
    newWs.Range("A1:S1").AutoFilter
    
    ' ─── 7. 保存 ──────────────────────────────
    Dim savePath As String
    savePath = Left(ws.Parent.FullName, InStrRev(ws.Parent.FullName, ".") - 1) & "-理货表整理.xlsx"
    newWb.SaveAs savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "✅ 转化完成！" & vbCrLf & _
           "行数: " & (r - 2) & " 行" & vbCrLf & _
           "保存至: " & savePath, vbInformation, "成功"
End Sub

' ═══ 辅助函数 ═══════════════════════════════════════════

Function MapPort(val As String) As String
    Select Case val
        Case "W1/浦东", "W1浦": MapPort = "WGQ1"
        Case "W2/振东", "W2振": MapPort = "WGQ2"
        Case "W4/沪东", "W4沪": MapPort = "WGQ4"
        Case "W5/明东", "W5明": MapPort = "WGQ5"
        Case "Y1/盛东", "Y1盛", "YS": MapPort = "YS"
        Case "Y3/冠东", "Y3冠": MapPort = "YS3"
        Case "Y4/尚东", "Y4尚": MapPort = "YS4"
        Case Else
            If InStr(val, "/") > 0 Then
                Dim parts As Variant
                parts = Split(val, "/")
                MapPort = parts(UBound(parts))
            Else
                MapPort = val
            End If
    End Select
End Function

Function SimplifyCompany(val As String) As String
    Dim prefixes As Variant
    prefixes = Array("ANL/澳航/", "APL/美国总统/", "CNC/正利航业/", "CMA/达飞/")
    Dim i As Integer
    For i = 0 To UBound(prefixes)
        If val = prefixes(i) Or Left(val, Len(prefixes(i))) = prefixes(i) Then
            SimplifyCompany = "CMA/达飞/"
            Exit Function
        End If
    Next i
    SimplifyCompany = val
End Function

Function SimplifyAgent(val As String) As String
    If InStr(val, "上港联合") > 0 Then
        SimplifyAgent = "联代"
        Exit Function
    End If
    If InStr(val, "联东") > 0 Then
        SimplifyAgent = "联代"
        Exit Function
    End If
    If InStr(val, "中联船代") > 0 Or InStr(val, "中联") > 0 Then
        SimplifyAgent = "中联"
        Exit Function
    End If
    
    ' 格式: MA/外运/021XHD002633 → 取中间段
    Dim parts As Variant
    If InStr(val, "/") > 0 Then
        parts = Split(val, "/")
        If UBound(parts) >= 2 Then
            ' 第三段含数字/字母 → 取第二段
            Dim thirdPart As String
            thirdPart = parts(2)
            Dim j As Integer
            Dim hasDigit As Boolean
            hasDigit = False
            For j = 1 To Len(thirdPart)
                If Mid(thirdPart, j, 1) Like "[0-9A-Za-z]" Then
                    hasDigit = True
                    Exit For
                End If
            Next j
            If hasDigit Then
                SimplifyAgent = parts(1)
            Else
                SimplifyAgent = parts(2)
            End If
        ElseIf UBound(parts) = 1 Then
            SimplifyAgent = parts(1)
        End If
    Else
        SimplifyAgent = val
    End If
End Function

Function NumericVal(val) As Double
    If IsEmpty(val) Or Not IsNumeric(val) Then
        NumericVal = 0
    Else
        NumericVal = CDbl(val)
    End If
End Function
