Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = vbFormControlMenu Then
      MsgBox "Use o Botão de FECHAR Abaixo!"
      Cancel = True
   End If
   
End Sub

Private Sub btnConsulta_Click()
    If cbOcoStatus = "" Then
        MsgBox "Selecione a Ocorrencia"
        Exit Sub
    End If
    
    Dim linha
    
    Dim ultimo
    Sheets("Analise").Select
    Cells(1000000, 1).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
    
    For linha = 6 To ultimo
        If Cells(linha, 1) = cbOcoStatus Then
            If Cells(linha, 3) = "ATIVO" Then
                cxAtivo.BackColor = RGB(0, 255, 0)
            Else
                cxAtivo.BackColor = RGB(220, 20, 60)
            End If
        End If
     Next
     
    For linha = 6 To ultimo
        If Cells(linha, 1) = cbOcoStatus Then
            If Cells(linha, 4) = "DISPON. PALM" Then
                cxInativo.BackColor = RGB(0, 255, 0)
            Else
                cxInativo.BackColor = RGB(220, 20, 60)
            End If
        End If
     Next
    
     

    Sheets("Analise").Select
    Cells(1000000, 6).Select
    Selection.End(xlUp).Select
    ultimo = ActiveCell.Row
    
    For linha = 6 To ultimo
        If Cells(linha, 6) = cbOcoStatus Then
            edtValor = FormatCurrency(Cells(linha, 7), 2)
            Exit For
        Else
            edtValor.Text = "Sem Movimento"
            
        End If
     Next
     
        
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()
Workbooks("Automacao_ocorrencia.xlsm").Close
End Sub

