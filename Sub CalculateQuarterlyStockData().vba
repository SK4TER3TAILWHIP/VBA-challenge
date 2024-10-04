Sub CalculateQuarterlyStockData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim currentQuarter As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim i As Long, j As Long
    Dim startRow As Long
    Dim sheetsToProcess As Variant
    Dim quarterlyChange As Double
    Dim percentChange As Double

    ' Variables to track greatest changes and volumes
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' List of sheets to process
    sheetsToProcess = Array("Q1", "Q2", "Q3", "Q4") ' Your specified sheet names

    On Error GoTo ErrorHandler

    For Each sheetName In sheetsToProcess
        ' Check if the sheet exists
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        On Error GoTo ErrorHandler
        
        ' If the sheet is not found, skip to the next
        If ws Is Nothing Then
            MsgBox "Sheet '" & sheetName & "' does not exist.", vbExclamation
            GoTo NextSheet
        End If

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        startRow = 2 ' Assuming headers are in row 1
        j = 2 ' Output starting row for tickers in column I

        ' Reset greatest variables for the current sheet
        greatestIncrease = -1E+308 ' Very small number
        greatestDecrease = 1E+308 ' Very large number
        greatestVolume = 0

        ' Loop through each row of data
        For i = startRow To lastRow
            If IsEmpty(ws.Cells(i, 1)) Then Exit For ' Exit if there's no ticker

            ticker = ws.Cells(i, 1).Value
            currentQuarter = Format(ws.Cells(i, 2).Value, "yyyy") & " Q" & DatePart("q", ws.Cells(i, 2).Value)

            If i = startRow Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' New ticker found, reset values
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If

            ' Accumulate total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' Check if we are at the last entry for this quarter
            If i = lastRow Or (ws.Cells(i + 1, 1).Value <> ticker) Or (Format(ws.Cells(i + 1, 2).Value, "yyyy") & " Q" & DatePart("q", ws.Cells(i + 1, 2).Value) <> currentQuarter) Then
                ' This is the last row of the quarter for this ticker
                closePrice = ws.Cells(i, 6).Value

                ' Calculate quarterly change and percentage change
                quarterlyChange = closePrice - openPrice
                percentChange = (quarterlyChange / openPrice) * 100

                ' Output the results
                ws.Cells(j, 9).Value = ticker ' Column I for Ticker
                ws.Cells(j, 10).Value = quarterlyChange ' Column J for Quarterly Change
                ws.Cells(j, 11).Value = percentChange ' Column K for Percent Change
                ws.Cells(j, 12).Value = totalVolume ' Column L for Total Volume

                ' Apply color formatting to Quarterly Change
                With ws.Cells(j, 10) ' Column J for Quarterly Change
                    If quarterlyChange < 0 Then
                        .Interior.Color = RGB(255, 0, 0) ' Red for negative
                    ElseIf quarterlyChange = 0 Then
                        .Interior.Color = RGB(255, 255, 255) ' White for no change
                    Else
                        .Interior.Color = RGB(0, 255, 0) ' Green for positive
                    End If
                End With

                ' Track greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If

                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If

                j = j + 1 ' Move to the next output row
            End If
        Next i

        ' Output greatest values to columns O, P, and Q
        Dim resultRow As Long
        resultRow = ws.Cells(ws.Rows.Count, 15).End(xlUp).Row + 1 ' Find next available row in column O

        ' Greatest Percentage Increase
        ws.Cells(resultRow, 15).Value = "Greatest % Increase" ' Column O
        ws.Cells(resultRow, 16).Value = greatestIncreaseTicker ' Column P
        ws.Cells(resultRow, 17).Value = greatestIncrease ' Column Q

        resultRow = resultRow + 1 ' Move to the next row

        ' Greatest Percentage Decrease
        ws.Cells(resultRow, 15).Value = "Greatest % Decrease" ' Column O
        ws.Cells(resultRow, 16).Value = greatestDecreaseTicker ' Column P
        ws.Cells(resultRow, 17).Value = greatestDecrease ' Column Q

        resultRow = resultRow + 1 ' Move to the next row

        ' Greatest Total Volume
        ws.Cells(resultRow, 15).Value = "Greatest Total Volume" ' Column O
        ws.Cells(resultRow, 16).Value = greatestVolumeTicker ' Column P
        ws.Cells(resultRow, 17).Value = greatestVolume ' Column Q

NextSheet:
    Next sheetName

    MsgBox "Quarterly stock data calculation completed for all sheets!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub