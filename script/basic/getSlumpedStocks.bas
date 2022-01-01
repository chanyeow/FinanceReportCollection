Attribute VB_Name = "ģ��10"
Sub calcSlumpedStocks()
Attribute calcSlumpedStocks.VB_ProcData.VB_Invoke_Func = " \n14"
'
' calcSlumpedStocks ��
'

'
    iSlumpedPercent = Sheet11.Range("H5")
    iPricePercent = Sheet11.Range("H6")
    iMaxRow = Sheet11.[a65536].End(xlUp).Row
    
    sNotic = "�Ƿ���" & iSlumpedPercent & "%����ɸѡ�����������Ľ����"
    
    If MsgBox(sNotic, vbYesNo) <> vbYes Then Exit Sub
    
    iBeginLine = 2
    iWriteIndex = 2
    Dim bPriceLess  As Boolean
    Dim iPriceDiv As Single
    
    For i = iBeginLine To iMaxRow
        iTotalKLine = Sheet11.Range("E" & i)
        iSlumpedKLine = Sheet11.Range("F" & i)
        iFirstKlinePrice = Sheet11.Cells(i, "c")
        iLastKlinePrice = Sheet11.Cells(i, "d")
        
        If iTotalKLine <> 0 And iFirstKlinePrice <> 0 Then
            iDiv = iSlumpedKLine / iTotalKLine * 100
            If iLastKlinePrice > iFirstKlinePrice Then
                bPriceLess = False
            Else
                iPriceDiv = (iFirstKlinePrice - iLastKlinePrice) / iFirstKlinePrice * 100
                If iPriceDiv > iPricePercent Then
                bPriceLess = True
            Else
                bPriceLess = False
            End If
        End If
            If (iDiv > iSlumpedPercent) And bPriceLess Then
                Sheet11.Cells(i, "a").Interior.ColorIndex = 3
                
                Sheet11.Cells(iWriteIndex, "j") = Sheet11.Cells(i, "b")
                Sheet11.Cells(iWriteIndex, "k") = iPriceDiv
                iWriteIndex = iWriteIndex + 1
            End If
        End If
    Next
    
    MsgBox "����ɸѡ���", , "��ʾ"

End Sub


Sub clearSlumped()
    iMaxRow = Sheet11.[a65536].End(xlUp).Row
    iMaxLine = Sheet11.Cells.SpecialCells(xlCellTypeLastCell).Column
    Sheet11.Range(Cells(2, 1), Cells(iMaxRow, iMaxLine)).Interior.ColorIndex = 0
End Sub