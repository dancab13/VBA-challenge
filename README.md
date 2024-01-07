# VBA-challenge
Solutions and code for Challenge 2

This repo contains VBA scripts and other required materials for Challenge 2 of the Columbia BCS Data Analytics bootcamp.

The solved Excel file is available upon request, though the instructions didn't specify we add it.

Some notes:
-The code for looping through the worksheets is from Stack Overflow:
  Dim ws As Worksheet

  For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
  Application.ScreenUpdating = False
  Application.ScreenUpdating = True

  Next ws

-The code for autofitting all cell column widths is from recording a macro and seeing the VBA code:

  Cells.Select
  Selection.Columns.AutoFit
  Cells(1, 1).Select

All other code is my own (from class and testing) unless otherwise specified.
