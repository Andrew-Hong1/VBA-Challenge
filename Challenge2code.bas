Attribute VB_Name = "Module1"
Dim vol As Double
Dim closing As Double
Dim opening As Double
Dim yearchange As Variant
Dim percentchange As Variant
Dim greatestpercentincrease As Variant
Dim greatestpercentdecrease As Variant
Dim greatestvolume As Variant
Dim greatestticker As String
Dim lowestticker As String
Dim greatestvolumeticker As String
Dim rownumber As Integer

Sub CreditCharge()

'loops for each worksheet
For Each ws In Worksheets

'sets the last row for each worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

    'Adds the following headers to every worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'sets default row number for the summary table to 2
    rownumber = 2
    
    'sets default greatest percent changes to 0
    greatestpercentincrease = 0
    greatestpercentdecrease = 0
    greatestvolume = 0
    
    'sets the current opening price for AAB
    opening = ws.Cells(2, 3).Value
    
'-----------------------------------------------------------------------------------
'Adds the tickers, the volume, and the yearly change to the summary table
'------------------------------------------------------------------------------------
        For c = 2 To Lastrow
        
            Ticker = ws.Cells(c, 1).Value
            
            'If the next ticker is inequal to the current ticker then
            If ws.Cells(c + 1, 1).Value <> ws.Cells(c, 1).Value Then
                
                'Adds the ticker name to the summary table
                Ticker = ws.Cells(c, 1).Value
                ws.Cells(rownumber, "I").Value = Ticker
                
                'Adds the total volume to the summary table
                vol = ws.Cells(c, 7).Value + vol
                ws.Cells(rownumber, "L").Value = vol
                
                'Sets the current ticker's closing price
                closing = ws.Cells(c, 6).Value
                
                'Adds the current ticker's yearly change to the summary table
                ws.Cells(rownumber, "J").Value = closing - opening
                
                'Sets the variable for year change as the value in the summary table
                yearchange = ws.Cells(rownumber, "J").Value
                
'---------------------------------------------------------------------------------------
'Sets the yearly change color based on the value in the cells
'---------------------------------------------------------------------------------------
                If yearchange > 0 Then
                    
                    ws.Cells(rownumber, "J").Interior.ColorIndex = 4
                    
                Else
                    
                    ws.Cells(rownumber, "J").Interior.ColorIndex = 3
                    
                End If
'---------------------------------------------------------------------------------------
'calculates percent change to the summary table and formats it to percentage
'---------------------------------------------------------------------------------------
                
                ws.Cells(rownumber, "K").Value = yearchange / opening
                
                percentchange = ws.Cells(rownumber, "K").Value
                
                ws.Cells(rownumber, "K") = Format(ws.Cells(rownumber, "K"), "0.00%")
                
'-------------------------------------------------------------------------------
'Calculates the greatest percent increase, greatest percent decrease, and greatest total
'volume and puts them in the chart
'-------------------------------------------------------------------------------
                
                If percentchange > greatestpercentincrease Then
                
                    greatestpercentincrease = percentchange
                
                    ws.Cells(2, "Q").Value = greatestpercentincrease
                    
                    ws.Cells(2, "Q") = Format(ws.Cells(2, "Q"), "0.00%")
                    
                    greatestticker = ws.Cells(rownumber, "I").Value
                    
                    ws.Cells(2, "P").Value = greatestticker
                    
                End If
                
                If percentchange < greatestpercentdecrease Then
                    
                    greatestpercentdecrease = percentchange
                    
                    ws.Cells(3, "Q").Value = greatestpercentdecrease
                    
                    ws.Cells(3, "Q") = Format(ws.Cells(3, "Q"), "0.00%")
                    
                    lowestticker = ws.Cells(rownumber, "I").Value
                    
                    ws.Cells(3, "P").Value = lowestticker
                    
                End If
                
                If vol > greatestvolume Then
                
                    greatestvolume = vol
                    
                    ws.Cells(4, "Q").Value = greatestvolume
                    
                    greatestvolume = ws.Cells(4, "Q").Value
                    
                    greatestvolumeticker = ws.Cells(rownumber, "I").Value
                    
                    ws.Cells(4, "P").Value = greatestvolumeticker
                    
                End If
                
'-----------------------------------------------------------------------------------
'Resets the values
'-----------------------------------------------------------------------------------
                
                'Sets the opening price to the next ticker's column value
                opening = ws.Cells(c + 1, 3).Value
                
                'Resets closing price
                closing = 0
                
                'resets vol
                vol = 0
                
                'Adds 1 to the rownumber
                rownumber = rownumber + 1
                
            Else
            
                'Adds up the volume whenever the tickers are the same
                vol = ws.Cells(c, 7).Value + vol
                
            End If
        
        Next c
        
Next ws
    
    
End Sub



