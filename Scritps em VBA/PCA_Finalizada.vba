Sub Atualizar()

    frmLoading.Show
    
    Application.ScreenUpdating = False
    
    
    Dim ultimo
    Cells(100000, 11).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
    
    
    frmLoading.lblAna.Visible = True
    Sheets("ANALISE").Select
    Range("INFO_CLIENTE[[#Headers],[COD. CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okAna.Visible = True
    
    
    frmLoading.lblCurva.Visible = True
    Sheets("ABC_QNT").Select
    Range("ABC_QNTD[[#Headers],[COD.CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    
    Sheets("CURVA_ABC").Select
    Range("ABC_BANCO[[#Headers],[COD.CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okCurv.Visible = True
    
    
    frmLoading.lblTitulo.Visible = True
    Sheets("TITULO_CLIENTE_ABERTO").Select
    Range("TITULO_CLIENTE_ABERTO[[#Headers],[COD.CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okTitAb.Visible = True
    
    frmLoading.lblFatu.Visible = True
    Sheets("FATURAMENTO_MEDIO").Select
    Range("FATURAMENTO_MEDIO[[#Headers],[COD CLI]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okFatMe.Visible = True
    
    frmLoading.lblTituloBai.Visible = True
    Sheets("TITULO_CLIENTE_BAIXADO").Select
    Range("TITULO_CLIENTE_BAIXADO[[#Headers],[COD. CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okTitBai.Visible = True
    
    frmLoading.lblLimite.Visible = True
    Sheets("LIMITE_CREDITO").Select
    Range("LIMITE_DE_CREDITO_CLIENTE[[#Headers],[COD. CLIENTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okLimCre.Visible = True
    
    frmLoading.lblHist.Visible = True
    Sheets("HISTORICO_CONSUMO").Select
    Range("HISTORICO_DE_CONSUMO[[#Headers],[NUM.PEDIDO]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okHistCo.Visible = True
   
    frmLoading.lblCev.Visible = True
    Sheets("CEV").Select
    Range("CEV[[#Headers],[NUM. CONTRATO]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("QTD_CEV[[#Headers],[CONTRATO]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    frmLoading.okCev.Visible = True
    
    
'Limpando_Conteudo

    Cells(14, 13).Select
    Range(Cells(14, 13), Cells(100000, 39)).Activate
    Selection.Clear

'Fim_Limpando_Conteudo


'Loading_Formatando
    
    frmFormat.lblFormat.Visible = True
    
'Iniciando_Formatação


     Cells(13, 13).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(OR(RC[-9]="""",RC[-9]=0),0,IF(AND(OR(RC[-2]=""6 DIAS ESP"",RC[-2]=""12 DIAS ESP"",RC[-2]=""A VISTA DINH.""),RC[3]=0),"""",IF(AND(RC[4]="""",OR(RC[5]<10),RC[5]=""0""),"""",IF(SUMIFS(TITULO_CLIENTE_ABERTO!C[-11],TITULO_CLIENTE_ABERTO!C[-12],RC[-9],TITULO_CLIENTE_ABERTO!C[-8],""914"")>0,""ACORDO"",IF(RC[4]<>"""",IF(DAYS360(RC[4],TODAY(),)>90,""DÉBITO > DO Q" & _
        "UE 90 DIAS"",IF(AND(RC[5]>=8,RC[5]<>""0""),""ATRASO MÉDIO > DO QUE 8 DIAS"","""")),IF(RC[5]<>""0"",IF(RC[5]>=8,""ATRASO MÉDIO > DO QUE 8 DIAS"",""""),"""")))))),"" "")" & _
        ""
    Cells(13, 14).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-10]="""","""",IF(OR(RC[-3]=""6 DIAS ESP"",RC[-3]=""A VISTA DINH."",RC[-3]=""12 DIAS ESP""),"""",IF(AND(RC[3]="""",RC[2]=0,RC[2]=""0""),"""",IF(OR(RC[-1]=""DÉBITO > DO QUE 90 DIAS"",RC[-1]=""ACORDO""),""A VISTA DINH."",IF(AND(RIGHT(RC[-3],2)<=""06"",RC[4]>=10),""A VISTA DINH."",IF(AND(RIGHT(RC[-3],2)>=""07"",RIGHT(RC[-3],2)<=""12"",RC[4]>=10),""06 DIAS"",IF(A" & _
        "ND(RIGHT(RC[-3],2)>""12"",RC[4]>=10),"""","""")))))))" & _
        ""
    Cells(13, 15).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[2]>=TODAY(),SUMIF(TITULO_CLIENTE_ABERTO[COD.CLIENTE],RC[-11],TITULO_CLIENTE_ABERTO[VALOR]),)"
    Cells(13, 16).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[1]<TODAY(),SUMIF(TITULO_CLIENTE_ABERTO!C[-15],RC[-12],TITULO_CLIENTE_ABERTO!C[-14]),)"
    Cells(13, 17).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX(TITULO_CLIENTE_ABERTO!C4,MATCH(RC[-13],TITULO_CLIENTE_ABERTO!C1,0)),"" "")"
    Cells(13, 18).Select
    ActiveCell.FormulaR1C1 = _
        "=AVERAGE(DAYS(TITULO_CLIENTE_BAIXADO!R[-11]C[-14],TITULO_CLIENTE_BAIXADO!R[-11]C[-12]))"
    Cells(13, 19).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-15],LIMITE_CREDITO!C[-18]:C[-16],2,0)"
    Cells(13, 20).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(FATURAMENTO_MEDIO!C[-19],INFO_CLIENTE[@[COD. CLIENTE]],FATURAMENTO_MEDIO!C[-17])/3"
    Cells(13, 21).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-17],LIMITE_CREDITO!C[-20]:C[-18],3,0)"
    Cells(13, 22).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(CURVA_ABC!C[-21],RC[-18],CURVA_ABC!C[-16])+SUMIF(CURVA_ABC!C[-21],RC[-18],CURVA_ABC!C[-15])"
    Cells(13, 23).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R13C[-19]="""","""",IF(RC11=""A VISTA DINH."",""A VISTA"",IF(AND(RC21>=0,RC21<REGRAS!R6C23),REGRAS!R6C23,IF(AND(RC21>=REGRAS!R6C23,RC21<REGRAS!R7C23),REGRAS!R7C23,IF(AND(RC21>=REGRAS!R7C23,RC21<REGRAS!R8C23),REGRAS!R8C23,IF(AND(RC21>=REGRAS!R8C23,RC21<REGRAS!R9C23),REGRAS!R9C23,IF(AND(RC21>=REGRAS!R9C23,RC21<REGRAS!R10C23),REGRAS!R10C23,IF(AND(RC21>=REGRAS!R10C2" & _
        "3,RC21<REGRAS!R11C23),REGRAS!R11C23,IF(AND(RC21>=REGRAS!R10C23,RC21<REGRAS!R11C23),REGRAS!R11C23,REGRAS!R11C23)))))))))" & _
        ""
    Cells(13, 24).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-20]="""",RC[-20]=0),"""",CONCAT(""Nivel "",IF(RC[-1]=""A VISTA"",0,IF(RC[-1]="""",0,IF(RC[-1]=1000,1,IF(RC[-1]=2000,2,IF(RC[-1]=4000,3,IF(RC[-1]=8000,4,IF(RC[-1]=16000,5,IF(RC[-1]=32000,6,IF(RC[-1]=64000,7,"""")))))))))))"
    Cells(13, 25).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[10]>=1,RC[1]=0),""GIRO ZERO 600ML"",IF(AND(RC[10]>=1,RC[1]<=RC[10]*3),""BAIXO GIRO 600ML"",IF(AND(RC[11]>=1,RC[2]=0),""GIRO ZERO 300ML"",IF(AND(RC[11]>=1,RC[2]<=RC[11]*3),""BAIXO GIRO 300ML"",IF(AND(RC[12]>=1,RC[3]=0),""GIRO ZERO 1L"",IF(AND(RC[12]>=1,RC[3]<=RC[12]*3),""BAIXO GIRO 1L"",""""))))))"
    Cells(13, 26).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[9]>=1,(RC[9]*3)-SUMIF(HISTORICO_CONSUMO!C3,R[-3]C4,HISTORICO_CONSUMO!C4)/3,"""")"
    Cells(13, 27).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[9]>=1,(RC[9]*3)-SUMIF(HISTORICO_CONSUMO!C3,R[-3]C4,HISTORICO_CONSUMO!C6)/3,"""")"
    Cells(13, 28).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[9]>=1,(RC[9]*3)-SUMIF(HISTORICO_CONSUMO!C3,R[-3]C4,HISTORICO_CONSUMO!C7)/3,"""")"
    Cells(13, 29).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[9]>=1,(RC[9]*3)-SUMIF(HISTORICO_CONSUMO!C3,R[-3]C4,HISTORICO_CONSUMO!C8)/3,"""")"
    Cells(13, 30).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(ABC_QNT!C[-29],RC[-26],ABC_QNT!C[-28])"
    Cells(13, 31).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(ABC_QNT!C[-30],RC[-27],ABC_QNT!C[-27])"
    Cells(13, 32).Select
    ActiveCell.FormulaR1C1 = "=SUMIF(ABC_QNT!C[-31],RC[-28],ABC_QNT!C[-27])"
    Cells(13, 33).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF('CEV'!C2,RC[-29])>=1,COUNTIF('CEV'!C2,RC[-29]),"""")"
    Cells(13, 35).Select
    ActiveCell.FormulaR1C1 = "=SUMIF('CEV'!C[-33],RC4,'CEV'!C[-32])"
    Cells(13, 36).Select
    ActiveCell.FormulaR1C1 = "=SUMIF('CEV'!C[-32],RC[-32],'CEV'!C[-31])"
    Cells(13, 37).Select
    ActiveCell.FormulaR1C1 = "=SUMIF('CEV'!C[-32],RC[-33],'CEV'!C[-31])"
    Cells(13, 38).Select
    ActiveCell.FormulaR1C1 = "=SUMIF('CEV'!C[-32],RC[-34],'CEV'!C[-31])"
    Cells(13, 39).Select
    ActiveCell.FormulaR1C1 = "=SUMIF('CEV'!C[-32],RC[-35],'CEV'!C[-31])"
 
  
'Finalizando_Formatação



'Sincronizando_Planilha_Manual


    Sheets("ANALISE").Select

    Cells(100000, 11).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
    
            
    Range(Cells(13, 13), Cells(13, 39)).Select
    Selection.Copy

    Range(Cells(14, 13), Cells(ultimo, 39)).Select
    ActiveSheet.Paste


'Fim_da_Sincronização
    
    Application.ScreenUpdating = True
    
    frmFormat.okFormat.Visible = True
    
    Unload frmFormat
           
    MsgBox "Atualização Finalizada!"

End Sub