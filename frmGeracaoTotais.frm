Option Explicit

Private Sub UserForm_Initialize()
    Dim plan As Worksheet
    
    Me.Caption = "Gerar Aba de Totais"
    
    lstAgrupadoras.MultiSelect = fmMultiSelectMulti
    lstTotalizadoras.MultiSelect = fmMultiSelectMulti
    
    cboAbaOrigem.Clear
    
    For Each plan In ThisWorkbook.Worksheets
        cboAbaOrigem.AddItem plan.Name
    Next plan
End Sub

Private Sub cboAbaOrigem_Change()
    Dim plan As Worksheet
    Dim coluna As Long
    Dim titulo As String
    
    If cboAbaOrigem.Value = vbNullString Then Exit Sub
    
    Set plan = ThisWorkbook.Worksheets(cboAbaOrigem.Value)
    
    lstAgrupadoras.Clear
    lstTotalizadoras.Clear
    
    coluna = 1
    
    Do While plan.Cells(1, coluna).Text <> vbNullString
        titulo = plan.Cells(1, coluna).Text
        
        lstAgrupadoras.AddItem titulo
        lstTotalizadoras.AddItem titulo
        
        coluna = coluna + 1
    Loop
    
    txtAbaSaida.Text = "Totais" & cboAbaOrigem.Value
End Sub

Private Sub cmdGerar_Click()
    On Error GoTo Erro
    
    If cboAbaOrigem.Value = vbNullString Then
        MsgBox "Informe a aba de origem.", vbExclamation, "Wordex"
        Exit Sub
    End If
    
    If txtAbaSaida.Text = vbNullString Then
        MsgBox "Informe a aba de saída.", vbExclamation, "Wordex"
        Exit Sub
    End If
    
    GerarAbaTotais _
        cboAbaOrigem.Value, _
        txtAbaSaida.Text, _
        ObterItensSelecionados(lstAgrupadoras), _
        ObterItensSelecionados(lstTotalizadoras)
    
    MsgBox "Aba de totais gerada.", vbInformation, "Wordex"
    Unload Me
    Exit Sub

Erro:
    MsgBox Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Function ObterItensSelecionados(ByRef lista As MSForms.ListBox) As Variant
    Dim itens() As String
    Dim i As Long
    Dim qtd As Long
    
    qtd = 0
    
    For i = 0 To lista.ListCount - 1
        If lista.Selected(i) Then
            ReDim Preserve itens(0 To qtd)
            itens(qtd) = lista.List(i)
            qtd = qtd + 1
        End If
    Next i
    
    If qtd = 0 Then
        ObterItensSelecionados = Array()
    Else
        ObterItensSelecionados = itens
    End If
End Function