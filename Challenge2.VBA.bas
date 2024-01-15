Attribute VB_Name = "Module1"

Sub calculatestock()

    Dim lastrow As Long
    Dim ws As Worksheet
    Dim year_change As Double
    Dim percent_change As Double
    Dim volume As Double
    Dim summaryrow As Long
    Dim openprice As Double
    
    For Each ws In Worksheets
    
    
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
       
        openprice = ws.Cells(2, 3).Value
        volume = 0
        summaryrow = 2
        
        For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            
                volume = volume + ws.Cells(i, 7).Value
            
           
            Else
            
                yearly_change = ws.Cells(i, 6).Value - openprice
                percent_change = yearly_change / openprice
                
                'display the results
                ws.Cells(summaryrow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summaryrow, 10).Value = yearly_change
                ws.Cells(summaryrow, 11).Value = percent_change
                 ws.Cells(summaryrow, 11).NumberFormat = "0.00%"
                ws.Cells(summaryrow, 12).Value = volume + ws.Cells(i, 7).Value
                
                'reset or reinitialize your variables
                summaryrow = summaryrow + 1
                openprice = ws.Cells(i + 1, 3).Value
                volume = 0
                
            
            End If
        
        Next i
      
 Next ws

End Sub

