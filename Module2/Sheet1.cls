VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ModuleProject()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputCell As Range
    Dim i As Long
    
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Set outputCell = ws.Cells(2, 9) ' Start outputting in cell I2
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                outputCell.Value = ws.Cells(i, 1).Value
                Set outputCell = outputCell.Offset(1, 0) ' Move to the next row in column I
            End If
        Next i
    Next ws
End Sub
Sub Prices()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim ticker As String
    Dim firstRow As Long
    Dim currentRow As Long
    Dim outputRow As Long
    
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        firstRow = 2 ' starts from row 2
        currentRow = 2 ' starts from row 2
        outputRow = 2  ' starts from row 2
        
        Do While currentRow <= lastRow
            ticker = ws.Cells(currentRow, 1).Value
            openingPrice = Val(ws.Cells(currentRow, 3).Value)
            totalVolume = 0 ' Initialize total volume for the current ticker, this will later be used to calculate total stock vol
            
            ' Find the last row for the current ticker to be  able to output it
            Do While ws.Cells(currentRow, 1).Value = ticker
                totalVolume = totalVolume + Val(ws.Cells(currentRow, 7).Value)
                currentRow = currentRow + 1
                If currentRow > lastRow Then Exit Do
            Loop
            
            ' Calculate the math equations
            closingPrice = Val(ws.Cells(currentRow - 1, 6).Value)
            yearlyChange = closingPrice - openingPrice
            percentChange = (yearlyChange / openingPrice) * 100
            
            ' Output the results for the current and each ticker
            ws.Cells(outputRow, 10).Value = yearlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = totalVolume ' Output total stock volume in column L, this will need to be formatted to fit
            
            ' Move to the next output row for the next ticker
            outputRow = outputRow + 1
        Loop
    Next ws
End Sub

Sub AnalysisWork()
    Dim ws As Worksheet
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        For currentRow = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            percentChange = ws.Cells(currentRow, 11).Value
            totalVolume = ws.Cells(currentRow, 12).Value

       
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                greatestIncreaseTicker = ws.Cells(currentRow, 9).Value ' Update the ticker symbol for increase (Column I)
            End If

        
            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                greatestDecreaseTicker = ws.Cells(currentRow, 9).Value ' Update the ticker symbol for decrease (Column I)
            End If

      
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                greatestVolumeTicker = ws.Cells(currentRow, 9).Value ' Update the ticker symbol for volume (Column I)
            End If
        Next currentRow

        ' Output the results on a specific row in each worksheet
        Dim outputRow As Long
        outputRow = 3 ' results will be displayed in row 3

        ' Write the results to each worksheet
        ws.Cells(outputRow, 15).Value = greatestIncreaseTicker
        ws.Cells(outputRow, 16).Value = maxIncrease

        ws.Cells(outputRow, 15).Offset(0, 1).Value = maxIncrease

        ws.Cells(outputRow + 1, 15).Value = greatestDecreaseTicker
        ws.Cells(outputRow + 16).Value = maxDecrease

        ws.Cells(outputRow + 1, 15).Offset(0, 1).Value = maxDecrease

        ws.Cells(outputRow + 2, 15).Value = greatestVolumeTicker
        ws.Cells(outputRow + 16).Value = maxVolume

        ws.Cells(outputRow + 2, 15).Offset(0, 1).Value = maxVolume
    Next ws
End Sub
