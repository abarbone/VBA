'' Convert cells to numbers, if appropriate
''
'' Note: The function will try to convert dates to General and may not be
''   appropriate for all types.  This is mostly useful for converting numbers
''   that have been (accidentally) saved as text.  Or, if data were imported
''   from another program, it may contain mixed types that Excel can handle. 
''
'' From: https://www.exceltip.com/cells-ranges-rows-and-columns-in-vba/change-text-to-number-using-vba.html'
'' Highlight cells that need to be update and run macro'
'' Quick enough to be run on entire worksheet?  So far'

Sub ConvertTextToNumbers()
  On Error Resume Next
  Dim rSelection As Range
  Set rSelection = rSelection
  rSelection.Select

  With Selection
    Selection.NumberFormat = "General"
    .Value = .Value
  End With

  rSelection.Select
  Set rSelection = Nothing
End Sub
