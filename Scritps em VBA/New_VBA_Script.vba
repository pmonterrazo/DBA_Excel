Sub Atualizar()
'
' Atualizar Macro
'
    Application.ScreenUpdating = False

'
    Range("F11").Select
    Selection.ListObject.QueryTable.Refresh 
    Range("F19").Select

        
        
  'Atualiza Faturamento Medio
  '
  '  Range("B4").Select
  ' Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
  'Atualiza Historico de Consumo
  
   ' Range("E14").Select
   ' Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    'Formatando_Tabela
    
    Range("Consulta5[[#Headers],[COD. CLIENTE]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Font
        .Color = -16744448
        .TintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 32768
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Range("L10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("L10:N10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("D10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("AB11").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Replace What:="   ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("Consulta5[[#Headers],[COD. CLIENTE]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Size = 10
    Columns("P:P").EntireColumn.AutoFit
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    'Ajustando as colunas de Retornave, CX Plastica, Refri...
    
    Columns("Q:Q").ColumnWidth = 4.43
    Columns("R:R").ColumnWidth = 4
    Columns("S:S").ColumnWidth = 3.29
    Columns("T:T").ColumnWidth = 3.14
    Columns("U:U").ColumnWidth = 2.86
    Columns("V:V").ColumnWidth = 3.86
    Columns("W:W").ColumnWidth = 3.57
    Columns("X:X").ColumnWidth = 3.71
    Columns("Y:Y").ColumnWidth = 3
    Columns("Z:Z").ColumnWidth = 4.57
    Columns("AA:AA").ColumnWidth = 4
    Columns("AB:AB").ColumnWidth = 4.86
    Columns("AC:AC").ColumnWidth = 3.57
    
    'Se 600ml < 30 Então....
    
    Range("AC10").Select
    ActiveCell.FormulaR1C1 = "=IF([@[600ML.]]<30,""GIRO 600ML"","" "")"
    
    'Condição SE TIVER GIRO 600ML FICAR VERMELHO
    
    Range("AC10").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="GIRO 600ML", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
   ' Selection.AutoFill Destination:=Range("Consulta5[MOTIVO]")
    Range("Consulta5[MOTIVO]").Select
    
    'Inserindo e Formatando Coluna de FAT MED
    
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("Consulta5[[#Headers],[Coluna1]]").Select
  ' ' ActiveCell.FormulaR1C1 = "FAT MED."
    Range("P10").Select
    
    'Formatando coluna de LIBERAR
    
    Range("Consulta5[[#Headers],[LIBERAR]]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Range("Q10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("Q10:Q46").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="NÃO", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'Formatando e ajustando todas as colunas
    
    ActiveCell.FormulaR1C1 = _
    "=SUMIF('HIST. CONSUMO'!C[-28],[@[COD. CLIENTE]],'HIST. CONSUMO'!C[-25])"
    Range("AE11").Select
    ActiveWindow.SmallScroll Down:=-33
    Range("AE10").Select
    Selection.AutoFill Destination:=Range("Consulta5[600ML.]"), Type:= _
        xlFillDefault
    Range("Consulta5[600ML.]").Select
    Range("AF10").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF('HIST. CONSUMO'!C[-29],[@[COD. CLIENTE]],'HIST. CONSUMO'!C[-27])+SUMIF('HIST. CONSUMO'!C[-29],[@[COD. CLIENTE]],'HIST. CONSUMO'!C[-25])"
    Range("AF10").Select
    Selection.AutoFill Destination:=Range("Consulta5[300ML.]"), Type:= _
        xlFillDefault
    Range("Consulta5[300ML.]").Select
    Range("AG10").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF('HIST. CONSUMO'!C[-30],[@[COD. CLIENTE]],'HIST. CONSUMO'!C[-29])+SUMIF('HIST. CONSUMO'!C[-30],[@[COD. CLIENTE]],'HIST. CONSUMO'!C[-24])"
    Range("AG10").Select
    Selection.AutoFill Destination:=Range("Consulta5[1L.]"), Type:= _
        xlFillDefault
    Range("Consulta5[1L.]").Select
    Range("AD10").Select
    ActiveCell.FormulaR1C1 = "=IF([@[600ML.]]<=30,""GIRO 600ML"","" "")"
    Range("AD10").Select
    Selection.AutoFill Destination:=Range("Consulta5[MOTIVO]"), Type:= _
        xlFillDefault
    Range("Consulta5[MOTIVO]").Select
    Range("AJ10").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(CEV!C[-33],[@[COD. CLIENTE]],CEV!C[-33],[@[COD. CLIENTE]])"
    Range("AJ10").Select
    Selection.AutoFill Destination:=Range("Consulta5[QNTD. COMODATO]"), Type:= _
        xlFillDefault
    Range("Consulta5[QNTD. COMODATO]").Select
    Range("AK10").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C[-34],[@[COD. CLIENTE]],CEV!C[-33])"
    Range("AK10").Select
    Selection.AutoFill Destination:=Range("Consulta5[[600ML  ]]"), Type:= _
        xlFillDefault
    Range("Consulta5[[600ML  ]]").Select
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    Range("AL10").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C[-35],[@[COD. CLIENTE]],CEV!C[-33])"
    Range("AL10").Select
    Selection.AutoFill Destination:=Range("Consulta5[[300ML  ]]"), Type:= _
        xlFillDefault
    Range("Consulta5[[300ML  ]]").Select
    Range("AM10").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C[-35],[@PEDIDO],CEV!C[-33])"
    Range("AM10").Select
    Selection.AutoFill Destination:=Range("Consulta5[[1L  ]]"), Type:= _
        xlFillDefault
    Range("Consulta5[[1L  ]]").Select
    Range("AN10").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C[-35],[@[RAZÃO SOCIAL]],CEV!C[-33])"
    Range("AN10").Select
    Selection.AutoFill Destination:=Range("Consulta5[REFRI. PEQ]"), Type:= _
        xlFillDefault
    Range("Consulta5[REFRI. PEQ]").Select
    Range("AO10").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C[-35],[@VD],CEV!C[-33])"
    Range("AO10").Select
    Selection.AutoFill Destination:=Range("Consulta5[REFRI. GRAND]"), Type:= _
        xlFillDefault
    Range("Consulta5[REFRI. GRAND]").Select
    Range("AP10").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C[-35],[@SUP],CEV!C[-33])"
    Range("AP10").Select
    Selection.AutoFill Destination:=Range("Consulta5[MESA PLAST.]"), Type:= _
        xlFillDefault
    Range("Consulta5[MESA PLAST.]").Select
    Range("AQ10").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C[-35],[@PRAZO],CEV!C[-33])"
    Range("AQ10").Select
    Selection.AutoFill Destination:=Range("Consulta5[MESA MAD.]"), Type:= _
        xlFillDefault
    Range("Consulta5[MESA MAD.]").Select
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    Range("P10").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF('Fat. Medio'!C[-15],[@[COD. CLIENTE]],'Fat. Medio'!C[-13])"
    Range("P10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("P10").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF('Fat. Medio'!C[-15],[@[COD. CLIENTE]],'Fat. Medio'!C[-13])/3"
    Range("P10").Select
    Selection.AutoFill Destination:=Range("Consulta5[FAT MED.]"), Type:= _
        xlFillDefault
    Range("Consulta5[FAT MED.]").Select
        Columns("I:I").ColumnWidth = 8.14
    Columns("I:I").ColumnWidth = 9.57
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Columns("Q:Q").ColumnWidth = 8.29
    Columns("Q:Q").ColumnWidth = 13.43
    ActiveWindow.SmallScroll Down:=-6
    Range("Q10").Select
    Range("D10").Select
    ActiveWindow.SmallScroll Down:=-3
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="   ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    Range("Consulta5[[#Headers],[LIBERAR]]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 32768
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    Range("Q10").Select
    ActiveCell.FormulaR1C1 = ""
    Range("Q11").Select
    Sheets("Fat. Medio").Select
    Range("B8").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("HIST. CONSUMO").Select
    Range("B6").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    Application.ScreenUpdating = True
    
    MsgBox "Atualização Finalizada com Sucesso!"

End Sub

Sub cev()
'
' cev Macro
'

'
    Range("D13").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
Sub Macro12()
'
' Macro12 Macro
'

'

End Sub
Sub Macro13()
'
' Macro13 Macro
'

'

End Sub
Sub Macro14()
'
' Macro14 Macro
'

'

End Sub
