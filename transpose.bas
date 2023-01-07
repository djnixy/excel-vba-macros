Sub Transpose_PasteSpecial()
    'x = 1
    'i = 2
    
    'For i = 2 To 5
    '    Sheets("source").Activate
    '    Sheets("source").Range(Cells(i, 1), Cells(i, 1)).Copy
       
    '    Sheets("destination").Activate
    '    Sheets("destination").Range(Cells(x, 3), Cells(x + 9, 3)).PasteSpecial Transpose:=True

    '    x = x + 10
    'Next i
    
    x = 1
    i = 2
    For i = 2 To 5

        Sheets("source").Activate
        Sheets("source").Range(Cells(i, 2), Cells(i, 11)).Copy
        'Range(Cells(6, 3), Cells(6, 12)).PasteSpecial Paste:=xlPasteValues
       
        Sheets("destination").Activate
        Sheets("destination").Range(Cells(x, 4), Cells(x + 9, 4)).PasteSpecial Transpose:=True

        x = x + 10
    Next i
End Sub
