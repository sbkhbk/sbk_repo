Sub Range_End_Method()
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
            "Last Column: " & lCol
  
End Sub
