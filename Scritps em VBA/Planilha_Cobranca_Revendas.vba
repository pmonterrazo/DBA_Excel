Dim ultimo
Dim ultimo_cadcli
Dim ultimo_tituloaberto
Dim Total As Long
Dim x As Long
Dim Largura As Long
Dim PERCENTUAL As Double

Sub BarraEvolucao()

Total = 10000
usfBarraEvolucao.Show
Largura = usfBarraEvolucao.lblBarraEvolucao.Width

End Sub

Sub Limpar_Filtros()
'Updated by Extendoffice 20210625
    Dim xAF As AutoFilter
    Dim xFs As Filters
    Dim xLos As ListObjects
    Dim xLo As ListObject
    Dim xRg As Range
    Dim xWs As Worksheet
    Dim xIntC, xF1, xF2, xCount As Integer
    Application.ScreenUpdating = False
    On Error Resume Next
    For Each xWs In Application.Worksheets
        xWs.ShowAllData
        Set xLos = xWs.ListObjects
        xCount = xLos.Count
        For xF1 = 1 To xCount
         Set xLo = xLos.Item(xF1)
         Set xRg = xLo.Range
         xIntC = xRg.Columns.Count
         For xF2 = 1 To xIntC
            xLo.Range.AutoFilter Field:=xF2
         Next
        Next
    Next
    Application.ScreenUpdating = True

End Sub

Sub Atualizacao_Planilha()

    BarraEvolucao

    'ATUALIZANDO OS DADOS DO BANCO DE DADOS
    
    DoEvents
    PERCENTUAL = 1000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"

    Application.ScreenUpdating = False
    
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_STATUS").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_V_AV").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_GA").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_TP._COBR.").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_VEN").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_ANO").ClearManualFilter
    
    Limpar_Filtros
    
    Sheets("Analises").Select
    Sheets("Titulo Aberto").Visible = True
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 1500 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    Sheets("Analises").Select
    Sheets("Cadastro de Cliente").Visible = True
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 1700 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    Sheets("Titulo Aberto").Select
    Cells(1, 1).Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 1800 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    Sheets("Titulo Aberto").Select
    Cells(1, 22).Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 2000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    
    Sheets("Cadastro de Cliente").Select
    Cells(1, 1).Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 2500 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    '#####FIM DA ATUALIZACAO OS DADOS DO BANCO DE DADOS
    
    
    
    
    'COLHENDO O ULTIMO NA PLANILHA DE TITULOS ABERTOS
    
    Sheets("Titulo Aberto").Select
    Cells(1000000, 2).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
    
    'FIM DA RECOLHA DO ULTIMO NA PLANILHA DE TITULOS ABERTOS
    
    
    
    
    'COPIANDO AS COLUNAS TITULOS E CODIGO DO CLIENTE DA PLANILHA DE TITULOS ABERTOS
    
    Sheets("Analises").Select
    Range(Cells(8, 5), Cells(1000000, 24)).Select
    Selection.EntireRow.Delete
    
    Sheets("Titulo Aberto").Select
    Range(Cells(2, 1), Cells(ultimo, 2)).Select
    Selection.Copy
    
    Sheets("Analises").Select
    Cells(8, 5).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'FIM DA COPIA DAS COLUNAS TITULOS E CODIGO DO CLIENTE DA PLANILHA DE TITULOS ABERTOS
    
    
    
    'ADICIONANDO AS FORMULAS NAS CELULAS PRINCIPAIS DA PLANILHA ANALISES
    
    Sheets("Analises").Select
    
    Cells(8, 7).FormulaR1C1 = "=VLOOKUP(RC[-1],'Cadastro de Cliente'!C[-6]:C[-4],2,0)"
    Cells(8, 8).FormulaR1C1 = "=VLOOKUP(RC[-2],'Cadastro de Cliente'!C[-7]:C[-5],3,0)"
    Cells(8, 9).FormulaR1C1 = "=VLOOKUP(RC[-4],'Titulo Aberto'!C[-8]:C[-5],4,0)"
    Cells(8, 10).FormulaR1C1 = "=VLOOKUP(RC[-5],'Titulo Aberto'!C[-9]:C[-5],5,0)"
    
    Cells(8, 11).FormulaR1C1 = "=VLOOKUP(RC[-5],'Titulo Aberto'!C[-9]:C[-5],5,0)"
    Cells(8, 12).FormulaR1C1 = "=IF(RC[-2]<TODAY(),""VENCIDO"",""A VENCER"")"
    Cells(8, 13).FormulaR1C1 = "=VLOOKUP(RC[-7],'Titulo Aberto'!C[-11]:C[-5],7,0)"
    Cells(8, 14).FormulaR1C1 = "=IF(SUMIF('Titulo Aberto'!C[8],RC[-9],'Titulo Aberto'!C[10])>1,SUMIF('Titulo Aberto'!C[8],RC[-9],'Titulo Aberto'!C[10]),RC[1])"
    
    Cells(8, 15).FormulaR1C1 = "=SUMIF('Titulo Aberto'!C[-14],RC[-10],'Titulo Aberto'!C[-5])"
    Cells(8, 16).FormulaR1C1 = "=SUMIF('Titulo Aberto'!C[10],RC[-11],'Titulo Aberto'!C[12])"
    Cells(8, 17).FormulaR1C1 = "=RC[-2]+RC[-1]"
    
    Cells(8, 18).FormulaR1C1 = "=VLOOKUP([@[COD.CLI]],'Cadastro de Cliente'!C[-17]:C[-13],4,0)"
    Cells(8, 19).FormulaR1C1 = "=VLOOKUP([@[COD.CLI]],'Cadastro de Cliente'!C[-18]:C[-14],5,0)"
    Cells(8, 20).FormulaR1C1 = "=VLOOKUP([@[COD.CLI]],'Cadastro de Cliente'!C[-19]:C[-11],9,0)"
    Cells(8, 21).FormulaR1C1 = "=VLOOKUP([@[COD.CLI]],'Cadastro de Cliente'!C[-20]:C[-11],10,0)"
    Cells(8, 22).FormulaR1C1 = "=VLOOKUP([@[COD.CLI]],'Cadastro de Cliente'!C[-21]:C[-11],11,0)"
    
    
    Cells(8, 23).FormulaR1C1 = "=YEAR(RC[-14])"
    Cells(8, 24).FormulaR1C1 = "=VLOOKUP(RC[-18],'Cadastro de Cliente'!C[-23]:C[-17],7,0)"
    
    Sheets("Titulo Aberto").Select
    
    Range(Cells(2, 26), Cells(ultimo, 28)).Select
    Selection.Clear
    
    Cells(8, 26).FormulaR1C1 = "=RC[-25]"
    Cells(8, 27).FormulaR1C1 = "=DAYS360(RC[-22],RC[-23])"
    Cells(8, 28).FormulaR1C1 = "=RC[-1]*0.003"
    
    
    'FIM DA ADICAO AS FORMULAS NAS CELULAS PRINCIPAIS DA PLANILHA ANALISES
    
    
    'BARRA DE LOADING DE 30% ATÉ 40% (NAO AFETA O CODIGO)
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 3000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    DoEvents
    PERCENTUAL = 4000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    'FIM --> BARRA DE LOADING DE 30% ATÉ 40% (NAO AFETA O CODIGO)
    
    
    
    'COPIANDO AS LINHAS PRINCIPAIS DA COLUNA G ATÉ X ( OU SEJA DA 7 A 24 )
    
    Sheets("Analises").Select
    
    Range(Cells(8, 7), Cells(8, 24)).Select
    Selection.Copy
    
    Range(Cells(9, 7), Cells(ultimo, 24)).Select
    ActiveSheet.Paste
     
    'FIM  --- > COPIANDO AS LINHAS PRINCIPAIS DA COLUNA G ATÉ X ( OU SEJA DA 7 A 24 )
    
    
    
    'FORMATANDO AS COLUNAS NO CADASTRO DE CLIENTE
    
    Sheets("Cadastro de Cliente").Select
    
    Cells(1000000, 1).Select
    Selection.End(xlUp).Select
    ultimo_cadcli = ActiveCell.Row
    
    Range(Cells(2, 1), Cells(ultimo_cadcli, 11)).Select
    
    Selection.Replace What:="   ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 5000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    
    'FIM  --- > FORMATANDO AS COLUNAS NO CADASTRO DE CLIENTE
    
    
    'FORMATANDO AS COLUNAS NO TITULOS ABERTOS
    
    Sheets("Titulo Aberto").Select
    
    Cells(1000000, 1).Select
    Selection.End(xlUp).Select
    ultimo_tituloaberto = ActiveCell.Row
    
    Range(Cells(2, 1), Cells(ultimo_tituloaberto, 20)).Select
    Selection.Replace What:="  ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 7000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    Range(Cells(2, 26), Cells(2, 28)).Select
    Selection.Copy
    
    Range(Cells(3, 26), Cells(ultimo, 28)).Select
    ActiveSheet.Paste
    
    Range(Cells(3, 26), Cells(ultimo, 28)).Select
    Selection.Copy
    
    Cells(2, 26).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'FIM --- > FORMATANDO AS COLUNAS NO TITULOS ABERTOS
    
       
    'AJUSTE DAS COLUNAS NA PLANILHA ANALISES
    
    Sheets("Analises").Select
    
    Range(Cells(9, 5), Cells(ultimo, 24)).Select

    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("J:J").ColumnWidth = 13
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").EntireColumn.AutoFit
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
    Columns("S:S").EntireColumn.AutoFit
    Columns("T:T").EntireColumn.AutoFit
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit
    Columns("W:W").EntireColumn.AutoFit
    Columns("X:X").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 8000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
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
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 9000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    Application.ScreenUpdating = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
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
    
    
    'FIM -----> AJUSTE DAS COLUNAS NA PLANILHA ANALISES


    Sheets("Titulo Aberto").Select
    Range(Cells(2, 1), Cells(ultimo, 2)).Select
    Selection.Copy
    
    Sheets("Analises").Select
    Cells(8, 5).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        
    Range(Cells(8, 5), Cells(ultimo, 24)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Cells(6, 5).Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Selection.Font.Size = 11
    Selection.Font.Size = 12
    
    Cells(6, 15).Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Selection.Font.Size = 14
    Selection.Font.Size = 12

    
    Sheets("Titulo Aberto").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Cadastro de Cliente").Select
    ActiveWindow.SelectedSheets.Visible = False
    
    
    Application.ScreenUpdating = True
    
    DoEvents
    PERCENTUAL = 10000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = PERCENTUAL * Largura
    usfBarraEvolucao.lblValor = Round(PERCENTUAL * 100, 1) & "%"
    
    
    usfBarraEvolucao.lblStatus = "        Concluido!"

    Application.Wait (Now + TimeValue("0:00:03"))

    Unload usfBarraEvolucao
    
    Sheets("Analises").Select
    Cells(8, 5).Select
    
    
    
End Sub


