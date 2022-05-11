Sub Atualizando_PLP()


'Atualizando_Planilha

    Application.ScreenUpdating = False
      
      
    Dim ultimo
    Cells(100000, 12).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
    
    progresso.Show
    
    progresso.Analise.Visible = True
    Sheets("ANALISE").Select
    Cells(9, 4).Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    progresso.AnaliseOK.Visible = True
    
    progresso.Historico.Visible = True
    Sheets("HIST_CONSUMO").Select
    Range("HIST_CONSUMO[[#Headers],[NUM.PEDIDO]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    progresso.HistoricoOK.Visible = True
    
    progresso.Itens.Visible = True
    Sheets("ITENS_PEDIDOS").Select
    Range("ITENS_DO_PEDIDO[[#Headers],[COD. CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    progresso.ItensOK.Visible = True
    
    progresso.Fat.Visible = True
    Sheets("FAT_MEDIO").Select
    Range("FATURAMENTO_MEDIO[[#Headers],[COD CLI]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    progresso.FatOK.Visible = True
    
    progresso.Titulo.Visible = True
    Sheets("TITL_CLIENTE").Select
    Range("TITULO_DO_CLIENTE[[#Headers],[COD. CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    progresso.TituloOK.Visible = True
    
    progresso.Cev.Visible = True
    Sheets("CEV").Select
    Range("CEV_PROD[[#Headers],[NUM. CONTRATO]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("CEV_QTD_CONTR[[#Headers],[NUM CONTR.]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    progresso.cevOk.Visible = True
    
    
    'Formatando a Planilha
    
    progresso.Formulas.Visible = True
    
    Cells(11, 14).Select
    Range(Cells(11, 37), Cells(10000, 37)).Activate
    Selection.Clear
    
    ' Formatando_Dados_Info_Cliente
    
    Cells(9, 4).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
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
    Selection.Replace What:="   ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Columns("D:D").EntireColumn.AutoFit
    Selection.Font.Size = 10
    Cells(9, 4).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    
    
    'Configurando as Formulas
    
    Cells(10, 14).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(TODAY()<TITL_CLIENTE!R[-8]C[-11],SUMIF(TITL_CLIENTE!C[-13],RC[-10],TITL_CLIENTE!C[-10]),0)"
    Cells(10, 15).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(TODAY()>RC[1],SUMIF(TITL_CLIENTE!C[-14],RC[-11],TITL_CLIENTE!C[-11]),)"
    Cells(10, 16).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(VALUE(RC[-12]),TITL_CLIENTE!C[-15]:C[-13],3,0),"""")"
    Cells(10, 17).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(FAT_MEDIO!C[-16],RC[-13],FAT_MEDIO!C[-14])/3"
    Cells(10, 18).Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,""NÃO"",""LIBERAR"")"
    Cells(10, 19).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C[-12],RC[-15],ITENS_PEDIDOS!C[-10])"
    Cells(10, 20).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C[-12],RC[-16],ITENS_PEDIDOS!C[-10])"
    Cells(10, 21).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C[-12],RC[-17],ITENS_PEDIDOS!C[-10])"
    Cells(10, 22).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C[-12],RC[-18],ITENS_PEDIDOS!C[-10])"
    Cells(10, 23).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C[-12],RC[-19],ITENS_PEDIDOS!C[-10])"
    Cells(10, 24).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C[-12],RC[-20],ITENS_PEDIDOS!C[-10])"
    Cells(10, 25).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[6]=0,"" "",IF(AND(RC[7]>=1,RC[1]=0),""GIRO ZERO 600ML"",IF(AND(RC[7]>=1,RC[1]<RC[7]*3),""BAIXO GIRO 600ML"",IF(AND(RC[8]>=1,RC[2]=0),""GIRO ZERO 300ML"",IF(AND(RC[8]>=1,RC[2]<RC[8]*3),""BAIXO GIRO 300ML"",IF(AND(RC[9]>=1,RC[3]=0),""GIRO ZERO 1L"",IF(AND(RC[9]>=1,RC[3]<RC[9]*3),""BAIXO GIRO 1L"","""")))))))"
    Cells(10, 26).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[6]>=1,(RC[6]*3)-SUMIF(HIST_CONSUMO!C[-23],RC[-22],HIST_CONSUMO!C[-22])/3,"""")"
    Cells(10, 27).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[6]>=1,(RC[6]*3)-SUMIF(HIST_CONSUMO!C[-24],RC[-23],HIST_CONSUMO!C[-21])/3,"""")"
    Cells(10, 28).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[6]>=1,(RC[6]*3)-SUMIF(HIST_CONSUMO!C[-25],RC[-24],HIST_CONSUMO!C[-21])/3,"""")"
    Cells(10, 29).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[6]>=1,SUMIF(FAT_MEDIO!C[-28],RC[-25],FAT_MEDIO!C[-26])/3<1000),1000-SUMIF(FAT_MEDIO!C[-28],RC[-25],FAT_MEDIO!C[-26])/3,"""")"
    Cells(10, 30).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[6]>=1,SUMIF(FAT_MEDIO!C[-29],RC[-26],FAT_MEDIO!C[-27])/3<1200),1200-SUMIF(FAT_MEDIO!C[-29],RC[-26],FAT_MEDIO!C[-27])/3,"""")"
    Cells(10, 31).Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(CEV!C[-19],RC[-27])"
    Cells(10, 32).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C2,RC4,CEV!C[-29])"
    Cells(10, 33).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C2,RC4,CEV!C[-29])"
    Cells(10, 34).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C2,RC4,CEV!C[-29])"
    Cells(10, 35).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C2,RC4,CEV!C[-29])"
    Cells(10, 36).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C2,RC4,CEV!C[-29])"
    Cells(10, 37).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(CEV!C2,RC4,CEV!C[-29])"
  
    'Sincronizando_Planilha_Manual

    Sheets("ANALISE").Select
    Range("D10").Select
    
    Cells(100000, 12).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
        
    Range(Cells(10, 14), Cells(10, 37)).Select
    Selection.Copy

    Range(Cells(11, 14), Cells(ultimo, 37)).Select
    ActiveSheet.Paste
    
    Application.ScreenUpdating = True
    
    progresso.FormulasOK.Visible = True
    progresso.atualizando = "Atualizado!!!"
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    Unload progresso
    MsgBox "Atualização Finalizada com Sucesso"
    
End Sub2