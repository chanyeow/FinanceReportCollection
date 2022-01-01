Attribute VB_Name = "模块6"
Sub fixPrice()
Attribute fixPrice.VB_ProcData.VB_Invoke_Func = " \n14"
    Worksheets(4).Select
    total_row = ActiveSheet.[a65536].End(xlUp).Row
    
    If total_row <> 1 Then
        ActiveSheet.Range("g2:g" & total_row).ClearContents
    End If
    
    Dim obj
    Set obj = CreateObject("WinHttp.WinHttpRequest.5.1")
       
    cellIndex = 7
    For i = 2 To total_row
        closePrice = ActiveSheet.Cells(i, 7)
        If closePrice = "" Then
            code = ActiveSheet.Cells(i, 2)
            If code <> lastCode Then
                marketCode = getMarketCode(code)
                
                Url = "http://push2.eastmoney.com/api/qt/stock/get?cb=jQuery1123015927363688471385_1618668042046&fltt=2&invt=2&secid=" & marketCode & "." & code & "&fields=f57%2Cf58%2Cf43%2Cf47%2Cf48%2Cf168%2Cf169%2Cf170%2Cf152&ut=b2884a393a59ad64002292a3e90d46a5&_=1618668042049"
                obj.Open "GET", Url, True
                obj.send
                obj.WaitForResponse
                t1 = BytesToBstr(obj.ResponseBody, "UTF-8")
                t1 = Split(t1, "(")(1)
                Set x = CreateObject("ScriptControl"): x.Language = "JScript"
                x.AddCode ("var query = (" & t1)
                
                rData = x.Eval("query.data")
                If rData <> "Null" Then
                    ActiveSheet.Cells(i, cellIndex) = format_int(x.Eval("query.data.f43"))
                End If
                lastCode = code
            Else
                ActiveSheet.Cells(i, cellIndex) = ActiveSheet.Cells(i - 1, cellIndex)
            End If
        End If
    Next
        
    MsgBox "收盘价补全完毕"
End Sub

Sub calc()
Attribute calc.VB_ProcData.VB_Invoke_Func = " \n14"

    total_row = Sheet5.[a65536].End(xlUp).Row
    If total_row <> 1 Then
        Sheet5.Range("a2:i" & total_row).ClearContents
    End If


    total_row = Sheet4.[a65536].End(xlUp).Row
    Worksheets(4).Select
    lowPer = ActiveSheet.Cells(2, 11)
    upPer = ActiveSheet.Cells(3, 11)
    Index = 2
    For i = 2 To total_row
        excutivePrice = Val(ActiveSheet.Cells(i, 6))
        curPrice = Val(ActiveSheet.Cells(i, 7))
        If curPrice <> 0 Then
            newLowPrice = excutivePrice + excutivePrice * lowPer / 100
            newUpPrice = excutivePrice + excutivePrice * upPer / 100
            If (newLowPrice < curPrice) And (newUpPrice > curPrice) Then
                Sheet5.Cells(Index, 1) = ActiveSheet.Cells(i, 1)
                Sheet5.Cells(Index, 2) = ActiveSheet.Cells(i, 2)
                Sheet5.Cells(Index, 3) = ActiveSheet.Cells(i, 3)
                Sheet5.Cells(Index, 4) = ActiveSheet.Cells(i, 4)
                Sheet5.Cells(Index, 5) = ActiveSheet.Cells(i, 5)
                Sheet5.Cells(Index, 6) = ActiveSheet.Cells(i, 6)
                Sheet5.Cells(Index, 7) = ActiveSheet.Cells(i, 7)
                Sheet5.Cells(Index, 8) = (curPrice - excutivePrice) / curPrice
                Sheet5.Cells(Index, 9) = newLowPrice
                
                Index = Index + 1
            End If
        End If
    Next
    Worksheets(5).Select
    MsgBox "计算完成"
    
End Sub
