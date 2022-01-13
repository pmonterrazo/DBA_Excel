Dim Total As Long
Dim x As Long
Dim Largura As Long
Dim Percentual As Double

Sub deletarConsultas()
    Dim pq As Object
    For Each pq In ThisWorkbook.Queries
        pq.Delete
    Next
 End Sub
Sub deletarQuerys()
    Dim qt As QueryTable
    Dim WSh As Worksheet
    For Each WSh In ThisWorkbook.Worksheets
        For Each qt In WSh.QueryTables
              qt.Delete
         Next qt
    Next WSh
End Sub
Sub deletarConexoes()
    Dim Conn As WorkbookConnection
    For Each Conn In ThisWorkbook.Connections
        Conn.Delete
    Next Conn
End Sub

Sub Barra_Evolucao()


Total = 10000
usfBarraEvolucao.Show
Largura = usfBarraEvolucao.lblBarraEvolucao.Width

End Sub



Sub Atualizar()
    
    Barra_Evolucao
    
    Application.Wait (Now + TimeValue("0:00:05"))
    
    DoEvents
    Percentual = 2000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = Percentual * Largura
    usfBarraEvolucao.lblValor = Round(Percentual * 100, 1) & "%"
    
 
    Application.ScreenUpdating = False

    deletarQuerys
    deletarConsultas
    deletarConexoes
    
    'Achar_Ultima_Celula
    
    Dim ultimo
    Cells(100000, 1).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
    
    
    'Deletar_Infor_Anterior
    
    Rows("7:10000").Select
    Selection.Delete Shift:=xlUp
    If ultimo > 7 Then
        Range(Cells(7, 1), Cells(ultimo, 20)).Select
        Selection.ClearContents
    End If
    Application.Wait (Now + TimeValue("0:00:05"))
    Application.ScreenUpdating = True
    
    DoEvents
    Percentual = 4000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = Percentual * Largura
    usfBarraEvolucao.lblValor = Round(Percentual * 100, 1) & "%"
    
        ' Atualizar Macro
   'fim
    Application.ScreenUpdating = False
    
        Application.CutCopyMode = False
        ActiveWorkbook.Queries.Add Name:="Consulta1", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Fonte = Odbc.Query(""dsn=DBCONTROL_3336_001"", ""SELECT DISTINCT A.CRMOVMOV_NDUPL AS 'NUM.TITULO', A.CRMOVMOV_CCLI AS 'COD. CLIENTE', B.VDCLICLI_RAZAO50 AS 'RAZÃO SOCIAL', CAST(RIGHT(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),4) AS VARCHAR(300)) AS" & _
        " 'DATA EMISSAO', CAST(RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) AS VARCHAR(300)) AS 'DATA VENCIMENTO', DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_D" & _
        "TV AS VARCHAR(300)),2) AS DATE),CURDATE()) AS DIAS, (CASE WHEN DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) > 0 THEN 'V' ELSE 'AV' END) AS 'V/AV', A.CRMOVMOV_MOD AS 'TP. COBRANÇA', C.VDPEDPPT_VLR_FCEM AS 'VLR. ORIGINAL',A.CRMO" & _
        "VMOV_VALOR AS 'VLR. TITULO', (DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) * 0.002 * A.CRMOVMOV_VALOR) AS JUROS, ((DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)" & _
        "),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) * 0.002 * A.CRMOVMOV_VALOR) + A.CRMOVMOV_VALOR) AS 'VLR. TOTAL', A.CRMOVMOV_VEN, A.CRMOVMOV_SUPVD AS 'GA', B.VDCLICLI_CONTATO AS 'CONTATO', B.VDCLICLI_FONE AS 'TELEFONE 1', B.VDCLICLI_FONE2 AS 'TELEFONE 2', B.VDCLICLI_CEL1 AS 'TELEFONE 3', B.VDCLICLI_CEL2 AS 'TELEFONE 4', LEFT(CAST(A.CRMOVMO" & _
        "V_DTE AS VARCHAR(300)),4) AS ANO#(lf)#(lf)FROM CADMOV01 AS A #(lf)INNER JOIN CADCLI01 AS B #(lf)ON CAST(A.CRMOVMOV_CCLI AS VARCHAR(300))=CAST(VDCLICLI_REGI AS VARCHAR(300)) || REPEAT('0',4-LENGTH(CAST(VDCLICLI_NUM AS VARCHAR(300)))) || CAST(VDCLICLI_NUM AS VARCHAR(300))#(lf)INNER JOIN NPEDPT01 AS C#(lf)ON A.CRMOVMOV_NPED=VDPEDPPT_NPEDPT"")," & Chr(13) & "" & Chr(10) & "    #""Tipo Alterado"" =" & _
        " Table.TransformColumnTypes(Fonte,{{""DATA EMISSAO"", type date}, {""DATA VENCIMENTO"", type date}, {""VLR. ORIGINAL"", Currency.Type}, {""VLR. TITULO"", Currency.Type}, {""JUROS"", Currency.Type}, {""VLR. TOTAL"", Currency.Type}})," & Chr(13) & "" & Chr(10) & "    #""Colunas Renomeadas"" = Table.RenameColumns(#""Tipo Alterado"",{{""CRMOVMOV_VEN"", ""VEN""}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Colunas Renomeadas"""
    Application.CutCopyMode = False
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Consulta1;Extended Properties=""""" _
        , Destination:=Range("$A$7")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Consulta1]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Consulta1"
        .Refresh BackgroundQuery:=False
    End With
        
    
    Application.ScreenUpdating = True
    'inicio
    
    DoEvents
    Percentual = 6000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = Percentual * Largura
    usfBarraEvolucao.lblValor = Round(Percentual * 100, 1) & "%"
    
    'FORMATANDO_PLANILHA
    Application.Wait (Now + TimeValue("0:00:05"))
    
    Application.ScreenUpdating = False
    
        Range("Consulta1[#Headers]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 25600
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
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
        .VerticalAlignment = xlBottom
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
    ActiveWindow.SmallScroll Down:=-27
    ActiveWindow.FreezePanes = False
    Range("Consulta1[[#Headers],[NUM.TITULO]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("H8").Select
    ActiveWindow.FreezePanes = True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("I:L").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    Columns("P:S").Select
    Selection.NumberFormat = "(00) 0 0000-0000"
    
    Application.ScreenUpdating = True
 
        
    DoEvents
    Percentual = 10000 / Total
    usfBarraEvolucao.lblBarraEvolucao.Width = Percentual * Largura
    usfBarraEvolucao.lblValor = Round(Percentual * 100, 1) & "%"
    
    usfBarraEvolucao.Caption = "Concluido"
    Unload usfBarraEvolucao
    
    MsgBox "Atualização Finalizada com Sucesso!"
End Sub
Sub Macro2()
'
' Macro2 Macro
'

'
    
End Sub
Sub teste()
'
' teste Macro
'

'
    Range("I2:J5").Select
    Selection.EntireColumn.Hidden = True
End Sub

Sub Macro11()
'
' Macro11 Macro
'

'

End Sub
