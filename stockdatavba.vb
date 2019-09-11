Sub vbastock():
    
    'set dimensions
    Dim vol As Double
    Dim i As Long
    Dim yrlychange As Single
    Dim yrlypercent As Single
    Dim begin As Long
    Dim rc As Long
    Dim j As Integer
    Dim book As Variant
    Set book = ActiveWorkbook.Worksheets
    
    
For Each ws In book
    j = 0
    vol = 0
    yrlychange = 0
    begin = 2

    'set names and stuff
     ws.Range("i1").Value = "ticker name"
     ws.Range("j1").Value = "yrly change"
     ws.Range("k1").Value = "yrly percent"
     ws.Range("l1").Value = "total vol"
     
    'set initial val
    j = 0
    begin = 2
    yrlychange = 0
    vol = 0
    
    'rc and stuff
    rc = ws.Cells(ws.Rows.Count, "a").End(xlUp).Row
    
    
    'loop and stuff
    
    For i = 2 To rc
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            vol = vol + ws.Cells(i, 7).Value
            
            If vol = 0 Then
                ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("j" & 2 + j).Value = 0
                ws.Range("k" & 2 + j).Value = "%" & 0
                ws.Range("l" & 2 + j).Value = 0
            
            Else
                If ws.Cells(begin, 3) = 0 Then
                    For fund = begin To i
                        If ws.Cells(fund, 3).Value <> 0 Then
                            begin = fund
                        Exit For
                    End If
                Next fund
            End If
            
            'next loop stuff
        
            yrlychange = ws.Cells(i, 6) - ws.Cells(begin, 3)
            yrlypercent = (yrlychange / ws.Cells(begin, 3) * 100)
         
            
            'NEXT
             begin = i + 1
         
            ws.Range("i" & j + 2).Value = ws.Cells(i, 1).Value
            ws.Range("j" & j + 2).Value = yrlychange
            ws.Range("k" & j + 2).Value = "%" & yrlypercent
            ws.Range("l" & j + 2).Value = vol
            
        'condi rice
        Select Case yrlychange
            Case Is > 0
                Range("j" & j + 2).Interior.ColorIndex = 4
            Case Is < 0
                Range("j" & j + 2).Interior.ColorIndex = 3
            Case Else
                Range("j" & j + 2).Interior.ColorIndex = 0
        End Select
         
    End If
        
    'other stuff
    
    vol = 0
    yrlychange = 0
    j = j + 1
    
    Else
        vol = vol + ws.Cells(i, 7).Value
    End If
Next i
Next ws
    
End Sub
