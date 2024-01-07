Attribute VB_Name = "Module1"
Sub stocks()

'The code for looping through all worksheets came from Stack Overflow
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
Application.ScreenUpdating = False
'End first part of worksheet loop

Dim TickerSymbol As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As LongLong
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim RowNum As Double
Dim Increase As Double
Dim Decrease As Double
Dim Volume As LongLong

'The code below came from class unless specified otherwise
finalrow = Cells(Rows.Count, 1).End(xlUp).Row
finalrow2 = Cells(Rows.Count, 9).End(xlUp).Row

ActiveSheet.Cells(1, 9).Value = "Ticker"
ActiveSheet.Cells(1, 10).Value = "Yearly Change"
ActiveSheet.Cells(1, 11).Value = "Percent Change"
ActiveSheet.Cells(1, 12).Value = "Total Stock Volume"
ActiveSheet.Cells(2, 15).Value = "Greatest % Increase"
ActiveSheet.Cells(3, 15).Value = "Greatest % Decrease"
ActiveSheet.Cells(4, 15).Value = "Greatest Total Volume"
ActiveSheet.Cells(1, 16).Value = "Ticker"
ActiveSheet.Cells(1, 17).Value = "Value"

YearlyChange = 0
PercentChange = 0
ClosePrice = 0
OpenPrice = 0
TotalStockVolume = 0
RowNum = 2
Increase = 0
Decrease = 0
Volume = 0

For i = 2 To finalrow

If Cells(i + 1, 1) <> Cells(i, 1) Then
        ClosePrice = ClosePrice + Cells(i, 6)
        YearlyChange = ClosePrice - OpenPrice
        Cells(RowNum, 10) = YearlyChange
        TotalStockVolume = TotalStockVolume + Cells(i, 7)
        Cells(RowNum, 12) = TotalStockVolume
        TotalStockVolume = 0
            If YearlyChange < 0 Then
                Cells(RowNum, 10).Interior.ColorIndex = 3
                    Else: Cells(RowNum, 10).Interior.ColorIndex = 4
            End If
        PercentChange = YearlyChange / OpenPrice
        Cells(RowNum, 11) = PercentChange
        'The NumberFormat code came from recording a macro and seeing the VBA results
        Cells(RowNum, 11).NumberFormat = "0.00%"
        TickerSymbol = Cells(i, 1)
        Cells(RowNum, 9) = TickerSymbol
        RowNum = RowNum + 1
        ClosePrice = 0
        
    ElseIf Cells(i - 1, 1) <> Cells(i, 1) Then
    OpenPrice = Cells(i, 3)
    TotalStockVolume = TotalStockVolume + Cells(i, 7)
    
    ElseIf Cells(i + 1, 1) = Cells(i, 1) Then
    TotalStockVolume = TotalStockVolume + Cells(i, 7)
        
End If
    
Next i

For i = 2 To finalrow2

If Cells(i, 11) > Increase Then
    Increase = Cells(i, 11)
    Cells(2, 16) = Cells(i, 9)
    Cells(2, 17) = Increase
    Cells(2, 17).NumberFormat = "0.00%"
    
    ElseIf Cells(i, 11) < Decrease Then
    Decrease = Cells(i, 11)
    Cells(3, 16) = Cells(i, 9)
    Cells(3, 17) = Decrease
    Cells(3, 17).NumberFormat = "0.00%"
    
    ElseIf Cells(i, 12) > Volume Then
    Volume = Cells(i, 12)
    Cells(4, 16) = Cells(i, 9)
    Cells(4, 17) = Volume
    End If

Next i

'The autofit code came from recording a macro and seeing the VBA results
Cells.Select
Selection.Columns.AutoFit
Cells(1, 1).Select

'The final part of the worksheet loop macro from Stack Overflow
Application.ScreenUpdating = True

Next ws

End Sub

