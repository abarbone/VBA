'From: https://www.excelcampus.com/keyboard-shortcuts/move-selected-range/'

Sub Move_Selection_Right_Column()
'Moves selection one column to the right

  On Error Resume Next
    Selection.Offset(0, 1).Select
  On Error GoTo 0

End Sub

Sub Move_Selection_Left_Column()
'Moves selection one column to the left

  On Error Resume Next
    Selection.Offset(0, -1).Select
  On Error GoTo 0

End Sub

Sub Move_Selection_Down_Row()
'Moves selection down one row

  On Error Resume Next
    Selection.Offset(1, 0).Select
  On Error GoTo 0

End Sub

Sub Move_Selection_Up_Row()
'Moves selection up one row

  On Error Resume Next
    Selection.Offset(-1, 0).Select
  On Error GoTo 0

End Sub

Sub Add_Move_Selection_Keyboard_Shortcuts()
'Create keyboard shortcuts to call Move Selection macros
  '% - Alt
  '+ - Shift
  'Help for OnKey Method: https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/application-onkey-method-excel

  'Set Alt+Right Arrow to call the macro
  Application.OnKey "%{RIGHT}", "Move_Selection_Right_Column"

  'Set Alt+Left Arrow to call the macro
  Application.OnKey "%{LEFT}", "Move_Selection_Left_Column"

  'Set Alt+Shift+Down Arrow to call the macro
  Application.OnKey "+%{Down}", "Move_Selection_Down_Row"

  'Set Alt+Shift+Up Arrow to call the macro
  Application.OnKey "+%{Up}", "Move_Selection_Up_Row"

End Sub
