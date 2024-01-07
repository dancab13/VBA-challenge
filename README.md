# VBA-challenge
Solutions and code for Challenge 2

This repo contains VBA scripts and other required materials for Challenge 2 of the Columbia BCS Data Analytics bootcamp.

The VBA code:
1) Retrieves the ticker symbol, annual open price, annual closing price, and total volume of stock for each discrete stock in the worksheets; 
2) Retrieves the stock symbol and value of the stock with the greatest percentage increase in price, greatest percentage decrease in price, and greatest total volume in price (for the year); and,
3) Prints this data for each worksheet (that is, for each year of data) onto the respective worksheets (e.g., 2018 results on the worksheet with 2018 data).  

The solved Excel file is available upon request, though the instructions didn't specify we add it.

Some notes:
*The code for looping through the worksheets is from Stack Overflow:
  Dim ws As Worksheet

  For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
  Application.ScreenUpdating = False
  Application.ScreenUpdating = True

  Next ws

*The code for autofitting all cell column widths is from recording a macro and seeing the VBA code:

  Cells.Select
  Selection.Columns.AutoFit
  Cells(1, 1).Select

All other code is my own (from class and testing) unless otherwise specified.
