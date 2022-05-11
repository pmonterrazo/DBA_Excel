    Sheets("ANALISE").Select
    Cells(10, 14).FormulaR1C1 = _
        "=IFERROR(IF(VLOOKUP(RC[-10],TITL_CLIENTE!C[-13]:C[-11],3,0)>=TODAY(),SUMIFS(TITL_CLIENTE!C[-10],TITL_CLIENTE!C[-13],RC[-10],TITL_CLIENTE!C[-11],"">=""&TODAY()),""""),"""")"
    Cells(10, 15).FormulaR1C1 = _
        "=IFERROR(IF(VLOOKUP(RC[-11],TITL_CLIENTE!C[-14]:C[-12],3,0)<TODAY(),SUMIFS(TITL_CLIENTE!C[-11],TITL_CLIENTE!C[-14],RC[-11],TITL_CLIENTE!C[-12],""<""&TODAY()),""""),"""")"
    Cells(10, 16).FormulaR1C1 = _
        "=IFERROR(IF(VLOOKUP(RC[-12],TITL_CLIENTE!C[-15]:C[-13],3,0)<TODAY(),VLOOKUP(RC[-12],TITL_CLIENTE!C[-15]:C[-13],3,0),""""),"""")"
    Cells(10, 17).FormulaR1C1 = "=SUMIF(FAT_MEDIO!C[-16],RC[-13],FAT_MEDIO!C[-14])/3"
    Cells(10, 18).FormulaR1C1 = "=IF(RC[-2]<TODAY(),""NÃO"",""LIBERAR"")"
    Cells(10, 19).FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C1,RC4,ITENS_PEDIDOS!C[-16])"
    Cells(10, 20).FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C1,RC4,ITENS_PEDIDOS!C[-16])"
    Cells(10, 21).FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C1,RC4,ITENS_PEDIDOS!C[-16])"
    Cells(10, 22).FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C1,RC4,ITENS_PEDIDOS!C[-16])"
    Cells(10, 23).FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C1,RC4,ITENS_PEDIDOS!C[-16])"
    Cells(10, 24).FormulaR1C1 = _
        "=SUMIF(ITENS_PEDIDOS!C1,RC4,ITENS_PEDIDOS!C[-16])"
    Cells(10, 25).FormulaR1C1 = _
       "=IF(RC[6]=0,"" "",IF(AND(RC[7]>=1,RC[1]=0),""GIRO ZERO 600ML"",IF(AND(RC[7]>=1,RC[1]<RC[7]*3),""BAIXO GIRO 600ML"",IF(AND(RC[8]>=1,RC[2]=0),""GIRO ZERO 300ML"",IF(AND(RC[8]>=1,RC[2]<RC[8]*3),""BAIXO GIRO 300ML"",IF(AND(RC[9]>=1,RC[3]=0),""GIRO ZERO 1L"",IF(AND(RC[9]>=1,RC[3]<RC[9]*3),""BAIXO GIRO 1L"","""")))))))"
    Cells(10, 26).FormulaR1C1 = _
        "=IF(RC[6]>=1,(RC[6]*3)-SUMIF(HIST_CONSUMO!C[-23],RC[-22],HIST_CONSUMO!C[-22])/3,"""")"
    Cells(10, 27).FormulaR1C1 = _
        "=IF(RC[6]>=1,(RC[6]*3)-SUMIF(HIST_CONSUMO!C[-24],RC[-23],HIST_CONSUMO!C[-21])/3,"""")"
    Cells(10, 28).FormulaR1C1 = _
        "=IF(RC[6]>=1,(RC[6]*3)-SUMIF(HIST_CONSUMO!C[-25],RC[-24],HIST_CONSUMO!C[-21])/3,"""")"
    Cells(10, 29).FormulaR1C1 = _
         "=IF(AND(RC[6]>=1,SUMIF(FAT_MEDIO!C[-28],RC[-25],FAT_MEDIO!C[-26])/3<1000),1000-SUMIF(FAT_MEDIO!C[-28],RC[-25],FAT_MEDIO!C[-26])/3,"""")"
    Cells(10, 30).FormulaR1C1 = _
       "=IF(AND(RC[6]>=1,SUMIF(FAT_MEDIO!C[-29],RC[-26],FAT_MEDIO!C[-27])/3<1200),1200-SUMIF(FAT_MEDIO!C[-29],RC[-26],FAT_MEDIO!C[-27])/3,"""")"
    Cells(10, 31).FormulaR1C1 = "=COUNTIF(CEV!C[-18],RC[-27])"
    Cells(10, 32).FormulaR1C1 = "=SUMIF(CEV!C1,RC4,CEV!C[-29])"
    Cells(10, 33).FormulaR1C1 = "=SUMIF(CEV!C1,RC4,CEV!C[-29])"
    Cells(10, 34).FormulaR1C1 = "=SUMIF(CEV!C1,RC4,CEV!C[-29])"
    Cells(10, 35).FormulaR1C1 = "=SUMIF(CEV!C1,RC4,CEV!C[-29])"
    Cells(10, 36).FormulaR1C1 = "=SUMIF(CEV!C1,RC4,CEV!C[-29])"
    Cells(10, 37).FormulaR1C1 = "=SUMIF(CEV!C1,RC4,CEV!C[-29])"
    Cells(8, 5).FormulaR1C1 = "=SUBTOTAL(3,R[2]C:R[684]C)"