Dim ultimo
Cells(100000, 1).Select
Selection.End(xlUp).Select
ultimo = ActiveCell.Row

Range(cells(10, 22), cells(10, 37))).Select
selection.copy

Range(cells(11, 22), cells(ultimo, 37))).Select
ActiveSheet.paste