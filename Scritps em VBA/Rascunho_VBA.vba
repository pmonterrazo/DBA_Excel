Sub Macro11()
'
' Macro11 Macro

    Dim ultimo

    Cells(100000, 2).Select
    Selection.End(xlUp).select
    ultimo = ActiveCell.Row


    Application.ScreenUpdating = False
 
'
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "COD CLI"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "600ML"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "300ML"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "1L"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "REFR. PEQ"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "REFR. GRND"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "MESA PLAS"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VALUE(RC[-8])"
    Range("J2").Select
    Selection.Copy
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("J6241").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("J2").Select
    ActiveSheet.Range("$J$1:$O$6241").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.End(xlUp).Select
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("J3").Select
    Selection.End(xlUp).Select
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Selection.End(xlUp).Select
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("J2").Select
    ActiveSheet.Paste
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents

    Range("J2").Select
    ActiveSheet.Range("$J$1:$O$2238").RemoveDuplicates Columns:=1, Header:= _
        xlYes

    Cells(100000, 9).Select
    Selection.End(xlUp).select
    ultimo = ActiveCell.Row

    Cells(2, 11).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C[-9],RC[-1],C[-8])"
    Cells(2, 12).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C2,RC10,C[-8])"
    Cells(2, 13).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C2,RC[-3],C[-8])"
    Cells(2, 14).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C2,RC[-4],C[-8])"
    Cells(2, 15).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C2,RC[-5],C[-8])"
    Cells(2, 16).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C2,RC[-6],C[-8])"

    Range(Cells(2, 11), Cells(2, 16)).Select
    Selection.Copy
    Range(Cells(3, 11), Cells(ultimo, 18)).Select
    ActiveSheet.Paste
    
End Sub


=SE(AG10>=1;SEERRO(AG10*3-$AE10*3-SOMASE(HIST_CONSUMO!C:C;D10;HIST_CONSUMO!F:F););0)

=SE(AG10>=1;(AG10*3)-SOMASE(HIST_CONSUMO!N:N;D10;HIST_CONSUMO!O:O);"")