Attribute VB_Name = "模块9"
Sub getSlumpedStocks()
Attribute getSlumpedStocks.VB_ProcData.VB_Invoke_Func = " \n14"
'
' getSlumpedStocks 宏
'

'
    If MsgBox("点击确认将开始抓取数据，时间较长（请看左下角进度条，抓取中请勿操作）", vbYesNo) <> vbYes Then Exit Sub
    
    iOriginIndex = 2    '开始index
    iIndex = iOriginIndex
    iSelfIndex = 5
    iMaxRow = Sheet9.[a65536].End(xlUp).Row + iIndex
    
    'Worksheets(1).UsedRange.Columns.Count   'Sheet9.Range("IV1").End(xlToLeft).Column
    'iMaxLine = Sheet11.Cells.SpecialCells(xlCellTypeLastCell).Column
    'Sheet11.Range("a5" & total_row).ClearContents
    
    'iMaxRow = 5
    
    Dim obj
    Set obj = CreateObject("WinHttp.WinHttpRequest.5.1")
    Do
        Sheet11.Cells(iSelfIndex, "a") = Sheet9.Cells(iIndex, 1)
        Sheet11.Cells(iSelfIndex, "b") = Sheet9.Cells(iIndex, 2)
        
        code = Sheet9.Cells(iIndex, 1)
        marketCode = getMarketCode(code)
        
        ' 周K
        ' Url = "http://74.push2his.eastmoney.com/api/qt/stock/kline/get?cb=jQuery112407697307125612478_1638805346501&secid=" & marketCode & "." & code & "&ut=fa5fd1943c7b386f172d6893dbfba10b&fields1=f1%2Cf2%2Cf3%2Cf4%2Cf5%2Cf6&fields2=f51%2Cf52%2Cf53%2Cf54%2Cf55%2Cf56%2Cf57%2Cf58%2Cf59%2Cf60%2Cf61&klt=101&fqt=0&beg=20170101&end=20500101&smplmt=460&lmt=1000000&_=1638805346528"
        
        ' 月k
        Url = "http://90.push2his.eastmoney.com/api/qt/stock/kline/get?cb=jQuery112402816109881259974_1638893884237&secid=" & marketCode & "." & code & "&ut=fa5fd1943c7b386f172d6893dbfba10b&fields1=f1%2Cf2%2Cf3%2Cf4%2Cf5%2Cf6&fields2=f51%2Cf52%2Cf53%2Cf54%2Cf55%2Cf56%2Cf57%2Cf58%2Cf59%2Cf60%2Cf61&klt=103&fqt=0&beg=20170101&end=20500101&smplmt=460&lmt=1000000&_=1638893884280"
        
        obj.Open "GET", Url, True
        obj.send
        obj.WaitForResponse
        t1 = BytesToBstr(obj.ResponseBody, "UTF-8")
        t1 = Split(t1, "(")(1)
        Set x = CreateObject("ScriptControl"): x.Language = "JScript"
        x.AddCode ("var query = (" & t1)
        
        rData = x.Eval("query.data")
        If rData <> "Null" Then
            ikLineLen = x.Eval("query.data.klines.length")
            
            iLineBegin = 3
            For i = 1 To ikLineLen
                sStr = x.Eval("query.data.klines[" & i - 1 & "]")
                Sheet11.Cells(iSelfIndex, iLineBegin) = Split(sStr, ",")(2)
                
                iLineBegin = iLineBegin + 1
            Next
        Else
            Sheet11.Cells(iIndex, "a").Interior.ColorIndex = 3
            Sheet11.Cells(iIndex, "b").Interior.ColorIndex = 3
        End If
        
        Application.StatusBar = GetProgress(iIndex - iOriginIndex, iMaxRow - iOriginIndex)
        
        iIndex = iIndex + 1
        iSelfIndex = iSelfIndex + 1
    Loop Until iIndex > iMaxRow
    MsgBox "A股所有数据抓取完毕", , "提示"
End Sub
