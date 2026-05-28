Option Explicit

Private Const FORMATO_COUNT As String = "###,###,###,##0"

Public Sub GerarAbaTotais( _
    ByVal nomeAbaOrigem As String, _
    ByVal nomeAbaSaida As String, _
    ByVal colunasAgrupadoras As Variant, _
    ByVal colunasTotalizadoras As Variant, _
    ByVal valoresCalcular As Variant, _
    ByVal gerarParaGrafico As Boolean _
)
    Dim origem As Worksheet
    Dim destino As Worksheet
    Dim grupos As Object
    Dim valoresGrupos As Object
    Dim linha As Long
    Dim chave As String
    Dim stats As Object
    
    If Totais_ArrayVazio(colunasTotalizadoras) Then
        Err.Raise vbObjectError + 3000, "GerarAbaTotais", _
            "Selecione pelo menos uma coluna para totalizar."
    End If

    If Totais_ArrayVazio(valoresCalcular) Then
        Err.Raise vbObjectError + 3003, "GerarAbaTotais", _
            "Selecione pelo menos um valor para calcular: Sum, Min, Max, Avg ou Count."
    End If
    
    Totais_ValidarValoresCalcular valoresCalcular
    
    Set origem = ThisWorkbook.Worksheets(nomeAbaOrigem)
    Set destino = Totais_ObterOuCriarAba(nomeAbaSaida)
    
    Set grupos = CreateObject("Scripting.Dictionary")
    Set valoresGrupos = CreateObject("Scripting.Dictionary")
    
    linha = 2
    
    Do While origem.Cells(linha, K_REQUIRED_COL).Text <> vbNullString
        chave = Totais_ObterChaveGrupo(origem, linha, colunasAgrupadoras)
        
        If Not grupos.Exists(chave) Then
            Set stats = CreateObject("Scripting.Dictionary")
            grupos.Add chave, stats
            valoresGrupos.Add chave, Totais_ObterValoresGrupo(origem, linha, colunasAgrupadoras)
        Else
            Set stats = grupos(chave)
        End If
        
        Totais_AcumularLinha origem, linha, colunasTotalizadoras, stats
        
        linha = linha + 1
    Loop
    
    Totais_EscreverResultado destino, origem, grupos, valoresGrupos, colunasAgrupadoras, colunasTotalizadoras, valoresCalcular, gerarParaGrafico
End Sub

Private Sub Totais_AcumularLinha( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal colunasTotalizadoras As Variant, _
    ByRef stats As Object _
)
    Dim i As Long
    Dim campo As String
    Dim coluna As Long
    Dim cell As Range
    Dim valor As Variant
    Dim numero As Double
    Dim countKey As String
    Dim sumKey As String
    Dim minKey As String
    Dim maxKey As String
    
    For i = LBound(colunasTotalizadoras) To UBound(colunasTotalizadoras)
        campo = CStr(colunasTotalizadoras(i))
        coluna = Totais_ObterNumeroColuna(plan, campo)
        Set cell = plan.Cells(linha, coluna)
        
        valor = cell.Value2
        
        If Not IsEmpty(valor) _
           And Trim$(cell.Text) <> vbNullString _
           And IsNumeric(valor) Then
           
            numero = CDbl(valor)
            
            countKey = campo & "Count"
            sumKey = campo & "Sum"
            minKey = campo & "Min"
            maxKey = campo & "Max"
            
            If Not stats.Exists(countKey) Then
                stats.Add countKey, 0&
                stats.Add sumKey, 0#
                stats.Add minKey, numero
                stats.Add maxKey, numero
            End If
            
            stats(countKey) = CLng(stats(countKey)) + 1
            stats(sumKey) = CDbl(stats(sumKey)) + numero
            
            If numero < CDbl(stats(minKey)) Then stats(minKey) = numero
            If numero > CDbl(stats(maxKey)) Then stats(maxKey) = numero
        End If
    Next i
End Sub

Private Sub Totais_EscreverResultado( _
    ByRef destino As Worksheet, _
    ByRef origem As Worksheet, _
    ByRef grupos As Object, _
    ByRef valoresGrupos As Object, _
    ByVal colunasAgrupadoras As Variant, _
    ByVal colunasTotalizadoras As Variant, _
    ByVal valoresCalcular As Variant, _
    ByVal gerarParaGrafico As Boolean _
)
    Dim linhaDestino As Long
    Dim colunaDestino As Long
    Dim colunaOrigem As Long
    Dim i As Long
    Dim j As Long
    Dim chave As Variant
    Dim valores As Variant
    Dim campo As String
    Dim stats As Object
    Dim tipoTotal As String
    Dim nomeTotal As String
    Dim valorTotal As Variant

    destino.Cells.Clear

    colunaDestino = 1

    If Not Totais_ArrayVazio(colunasAgrupadoras) Then
        For i = LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras)
            campo = CStr(colunasAgrupadoras(i))
            colunaOrigem = Totais_ObterNumeroColuna(origem, campo)

            destino.Cells(K_HEADERS_ROW, colunaDestino).Value = campo
            Totais_CopiarFormatoColuna origem, colunaOrigem, destino, colunaDestino

            colunaDestino = colunaDestino + 1
        Next i
    End If

    For i = LBound(colunasTotalizadoras) To UBound(colunasTotalizadoras)
        campo = CStr(colunasTotalizadoras(i))
        colunaOrigem = Totais_ObterNumeroColuna(origem, campo)

        For j = LBound(valoresCalcular) To UBound(valoresCalcular)
            tipoTotal = Totais_NormalizarTipoTotal(CStr(valoresCalcular(j)))
            nomeTotal = campo & tipoTotal
            Totais_EscreverCabecalhoTotal destino, origem, colunaOrigem, colunaDestino, nomeTotal, tipoTotal, gerarParaGrafico, Totais_ColunaPareceData(origem, colunaOrigem)
        Next j
    Next i

    destino.Rows(K_HEADERS_ROW).Font.Bold = True

    linhaDestino = K_HEADERS_ROW + 1

    For Each chave In grupos.Keys
        Set stats = grupos(chave)
        valores = valoresGrupos(chave)
        colunaDestino = 1

        If Not Totais_ArrayVazio(colunasAgrupadoras) Then
            For i = LBound(valores) To UBound(valores)
                destino.Cells(linhaDestino, colunaDestino).Value = valores(i)
                colunaDestino = colunaDestino + 1
            Next i
        End If

        For i = LBound(colunasTotalizadoras) To UBound(colunasTotalizadoras)
            campo = CStr(colunasTotalizadoras(i))

            For j = LBound(valoresCalcular) To UBound(valoresCalcular)
                tipoTotal = Totais_NormalizarTipoTotal(CStr(valoresCalcular(j)))
                nomeTotal = campo & tipoTotal
                valorTotal = Totais_CalcularValorTotal(stats, campo, tipoTotal)
                Totais_EscreverValorTotal destino, linhaDestino, colunaDestino, nomeTotal, tipoTotal, valorTotal, gerarParaGrafico, Totais_ColunaPareceData(origem, Totais_ObterNumeroColuna(origem, campo))
            Next j
        Next i

        linhaDestino = linhaDestino + 1
    Next chave
End Sub

Private Sub Totais_EscreverCabecalhoTotal( _
    ByRef destino As Worksheet, _
    ByRef origem As Worksheet, _
    ByVal colunaOrigem As Long, _
    ByRef colunaDestino As Long, _
    ByVal nomeTotal As String, _
    ByVal tipoTotal As String, _
    ByVal gerarParaGrafico As Boolean, _
    ByVal fonteEhData As Boolean _
)
    destino.Cells(K_HEADERS_ROW, colunaDestino).Value = nomeTotal

    Select Case tipoTotal
        Case "Count"
            Totais_AplicarFormatoCount destino, colunaDestino

        Case "Avg"
            Totais_CopiarFormatoColuna origem, colunaOrigem, destino, colunaDestino
            destino.Columns(colunaDestino).NumberFormat = _
                Totais_FormatoAvg(origem.Columns(colunaOrigem).NumberFormat)

        Case Else
            Totais_CopiarFormatoColuna origem, colunaOrigem, destino, colunaDestino
    End Select

    colunaDestino = colunaDestino + 1

    If gerarParaGrafico Then
        destino.Cells(K_HEADERS_ROW, colunaDestino).Value = nomeTotal & "_Value"
        Totais_AplicarFormatoValorGrafico destino, colunaDestino, fonteEhData And (tipoTotal = "Min" Or tipoTotal = "Max")
        colunaDestino = colunaDestino + 1

        destino.Cells(K_HEADERS_ROW, colunaDestino).Value = nomeTotal & "_Label"
        Totais_AplicarFormatoLegendaGrafico destino, colunaDestino
        colunaDestino = colunaDestino + 1
    End If
End Sub

Private Sub Totais_EscreverValorTotal( _
    ByRef destino As Worksheet, _
    ByVal linhaDestino As Long, _
    ByRef colunaDestino As Long, _
    ByVal nomeTotal As String, _
    ByVal tipoTotal As String, _
    ByVal valorTotal As Variant, _
    ByVal gerarParaGrafico As Boolean, _
    ByVal fonteEhData As Boolean _
)
    destino.Cells(linhaDestino, colunaDestino).Value = valorTotal
    colunaDestino = colunaDestino + 1

    If gerarParaGrafico Then
        destino.Cells(linhaDestino, colunaDestino).Value = Totais_ValorGrafico(valorTotal, fonteEhData, tipoTotal)
        colunaDestino = colunaDestino + 1

        destino.Cells(linhaDestino, colunaDestino).Value = Totais_LegendaGrafico(nomeTotal)
        colunaDestino = colunaDestino + 1
    End If
End Sub

Private Sub Totais_ValidarValoresCalcular(ByVal valoresCalcular As Variant)
    Dim i As Long
    Dim tipoTotal As String

    For i = LBound(valoresCalcular) To UBound(valoresCalcular)
        tipoTotal = Totais_NormalizarTipoTotal(CStr(valoresCalcular(i)))
    Next i
End Sub

Private Function Totais_NormalizarTipoTotal(ByVal tipoTotal As String) As String
    Select Case LCase$(Trim$(tipoTotal))
        Case "sum"
            Totais_NormalizarTipoTotal = "Sum"
        Case "min"
            Totais_NormalizarTipoTotal = "Min"
        Case "max"
            Totais_NormalizarTipoTotal = "Max"
        Case "avg"
            Totais_NormalizarTipoTotal = "Avg"
        Case "count"
            Totais_NormalizarTipoTotal = "Count"
        Case Else
            Err.Raise vbObjectError + 3004, "GerarAbaTotais", _
                "Valor a calcular inválido: '" & tipoTotal & "'. Use Sum, Min, Max, Avg ou Count."
    End Select
End Function

Private Function Totais_CalcularValorTotal( _
    ByRef stats As Object, _
    ByVal campo As String, _
    ByVal tipoTotal As String _
) As Variant
    Dim countValue As Long
    Dim sumValue As Double

    Select Case tipoTotal
        Case "Count"
            Totais_CalcularValorTotal = Totais_ValorLong(stats, campo & "Count")

        Case "Sum"
            Totais_CalcularValorTotal = Totais_ValorDouble(stats, campo & "Sum")

        Case "Min"
            Totais_CalcularValorTotal = Totais_ValorOuVazio(stats, campo & "Min")

        Case "Max"
            Totais_CalcularValorTotal = Totais_ValorOuVazio(stats, campo & "Max")

        Case "Avg"
            countValue = Totais_ValorLong(stats, campo & "Count")
            sumValue = Totais_ValorDouble(stats, campo & "Sum")

            If countValue > 0 Then
                Totais_CalcularValorTotal = sumValue / countValue
            Else
                Totais_CalcularValorTotal = vbNullString
            End If
    End Select
End Function

Private Function Totais_LegendaGrafico(ByVal nomeTotal As String) As String
    Dim sufixos As Variant
    Dim i As Long
    Dim sufixo As String

    sufixos = Array("Count", "Sum", "Min", "Max", "Avg")

    For i = LBound(sufixos) To UBound(sufixos)
        sufixo = CStr(sufixos(i))

        If Len(nomeTotal) > Len(sufixo) Then
            If Right$(nomeTotal, Len(sufixo)) = sufixo Then
                Totais_LegendaGrafico = Left$(nomeTotal, Len(nomeTotal) - Len(sufixo))
                Exit Function
            End If
        End If
    Next i

    Totais_LegendaGrafico = nomeTotal
End Function

Private Sub Totais_AplicarFormatoValorGrafico( _
    ByRef destino As Worksheet, _
    ByVal colunaDestino As Long, _
    ByVal valorComoDataTexto As Boolean _
)
    With destino.Columns(colunaDestino)
        If valorComoDataTexto Then
            .NumberFormat = "@"
            .HorizontalAlignment = xlLeft
        Else
            .NumberFormat = "General"
            .HorizontalAlignment = xlRight
        End If

        .VerticalAlignment = xlCenter
        .ColumnWidth = 18
    End With
End Sub

Private Function Totais_ValorGrafico( _
    ByVal valorTotal As Variant, _
    ByVal fonteEhData As Boolean, _
    ByVal tipoTotal As String _
) As Variant
    If IsEmpty(valorTotal) Or CStr(valorTotal) = vbNullString Then
        Totais_ValorGrafico = vbNullString
        Exit Function
    End If

    If fonteEhData And (tipoTotal = "Min" Or tipoTotal = "Max") Then
        If IsNumeric(valorTotal) Then
            Totais_ValorGrafico = Format$(CDate(CDbl(valorTotal)), "yyyy-mm-dd")
            Exit Function
        End If

        If IsDate(valorTotal) Then
            Totais_ValorGrafico = Format$(CDate(valorTotal), "yyyy-mm-dd")
            Exit Function
        End If
    End If

    Totais_ValorGrafico = valorTotal
End Function

Private Function Totais_ColunaPareceData(ByRef plan As Worksheet, ByVal coluna As Long) As Boolean
    Dim formato As String
    Dim linha As Long
    Dim v As Variant

    formato = LCase$(CStr(plan.Columns(coluna).NumberFormat))

    If Totais_FormatoPareceData(formato) Then
        Totais_ColunaPareceData = True
        Exit Function
    End If

    linha = K_DETAILS_ROW

    Do While plan.Cells(linha, K_REQUIRED_COL).Text <> vbNullString
        v = plan.Cells(linha, coluna).Value

        If IsDate(v) And Not IsNumeric(plan.Cells(linha, coluna).Text) Then
            Totais_ColunaPareceData = True
            Exit Function
        End If

        linha = linha + 1
    Loop
End Function

Private Function Totais_FormatoPareceData(ByVal formato As String) As Boolean
    Dim limpo As String

    limpo = Totais_RemoverLiteraisFormato(LCase$(formato))

    If InStr(1, limpo, "[$-f800]", vbTextCompare) > 0 Then
        Totais_FormatoPareceData = True
        Exit Function
    End If

    Totais_FormatoPareceData = _
        (InStr(1, limpo, "d", vbTextCompare) > 0 Or InStr(1, limpo, "y", vbTextCompare) > 0 Or InStr(1, limpo, "a", vbTextCompare) > 0) _
        And InStr(1, limpo, "m", vbTextCompare) > 0
End Function

Private Function Totais_RemoverLiteraisFormato(ByVal formato As String) As String
    Dim i As Long
    Dim ch As String
    Dim dentroAspas As Boolean
    Dim resultado As String

    For i = 1 To Len(formato)
        ch = Mid$(formato, i, 1)

        If ch = """" Then
            dentroAspas = Not dentroAspas
        ElseIf Not dentroAspas Then
            If ch = "\" Then
                i = i + 1
            Else
                resultado = resultado & ch
            End If
        End If
    Next i

    Totais_RemoverLiteraisFormato = resultado
End Function

Private Sub Totais_AplicarFormatoLegendaGrafico(ByRef destino As Worksheet, ByVal colunaDestino As Long)
    With destino.Columns(colunaDestino)
        .NumberFormat = "@"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .ColumnWidth = 22
    End With
End Sub

Private Function Totais_ObterChaveGrupo( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal colunasAgrupadoras As Variant _
) As String

    Dim i As Long
    Dim chave As String
    Dim coluna As Long
    
    If Totais_ArrayVazio(colunasAgrupadoras) Then
        Totais_ObterChaveGrupo = "#TOTAL_GERAL#"
        Exit Function
    End If
    
    chave = vbNullString
    
    For i = LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras)
        coluna = Totais_ObterNumeroColuna(plan, CStr(colunasAgrupadoras(i)))
        chave = chave & ChrW(30) & CStr(plan.Cells(linha, coluna).Value2)
    Next i
    
    Totais_ObterChaveGrupo = chave
End Function

Private Function Totais_ObterValoresGrupo( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal colunasAgrupadoras As Variant _
) As Variant

    Dim valores() As Variant
    Dim i As Long
    Dim coluna As Long
    
    If Totais_ArrayVazio(colunasAgrupadoras) Then
        Totais_ObterValoresGrupo = Array()
        Exit Function
    End If
    
    ReDim valores(LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras))
    
    For i = LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras)
        coluna = Totais_ObterNumeroColuna(plan, CStr(colunasAgrupadoras(i)))
        valores(i) = plan.Cells(linha, coluna).Value2
    Next i
    
    Totais_ObterValoresGrupo = valores
End Function

Private Function Totais_ObterOuCriarAba(ByVal nomeAba As String) As Worksheet
    On Error Resume Next
    Set Totais_ObterOuCriarAba = ThisWorkbook.Worksheets(nomeAba)
    On Error GoTo 0
    
    If Not Totais_ObterOuCriarAba Is Nothing Then
        If MsgBox("A aba '" & nomeAba & "' já existe. Deseja substituí-la?", vbYesNo + vbQuestion, "Wordex") = vbNo Then
            Err.Raise vbObjectError + 3001, "Totais_ObterOuCriarAba", "Operação cancelada."
        End If
        
        Totais_ObterOuCriarAba.Cells.Clear
        Exit Function
    End If
    
    Set Totais_ObterOuCriarAba = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Totais_ObterOuCriarAba.Name = nomeAba
End Function

Private Function Totais_ObterNumeroColuna(ByRef plan As Worksheet, ByVal tituloColuna As String) As Long
    Dim coluna As Long
    
    coluna = 1
    
    Do While plan.Cells(1, coluna).Text <> vbNullString
        If plan.Cells(1, coluna).Text = tituloColuna Then
            Totais_ObterNumeroColuna = coluna
            Exit Function
        End If
        
        coluna = coluna + 1
    Loop
    
    Err.Raise vbObjectError + 3002, "Totais_ObterNumeroColuna", _
        "Coluna '" & tituloColuna & "' não encontrada na aba '" & plan.Name & "'."
End Function

Private Sub Totais_CopiarFormatoColuna( _
    ByRef origem As Worksheet, _
    ByVal colunaOrigem As Long, _
    ByRef destino As Worksheet, _
    ByVal colunaDestino As Long _
)
    With destino.Columns(colunaDestino)
        .NumberFormat = origem.Columns(colunaOrigem).NumberFormat
        .HorizontalAlignment = origem.Columns(colunaOrigem).HorizontalAlignment
        .VerticalAlignment = origem.Columns(colunaOrigem).VerticalAlignment
        .ColumnWidth = origem.Columns(colunaOrigem).ColumnWidth
        .Font.Name = origem.Columns(colunaOrigem).Font.Name
        .Font.Size = origem.Columns(colunaOrigem).Font.Size
    End With
End Sub

Private Sub Totais_AplicarFormatoCount(ByRef destino As Worksheet, ByVal colunaDestino As Long)
    With destino.Columns(colunaDestino)
        .NumberFormat = FORMATO_COUNT
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ColumnWidth = 16
    End With
End Sub

Private Function Totais_FormatoAvg(ByVal formatoOriginal As String) As String
    Dim secoes() As String
    Dim i As Long
    
    formatoOriginal = Trim$(formatoOriginal)
    
    If formatoOriginal = vbNullString Or LCase$(formatoOriginal) = "general" Then
        Totais_FormatoAvg = "0.0"
        Exit Function
    End If
    
    secoes = Split(formatoOriginal, ";")
    
    For i = LBound(secoes) To UBound(secoes)
        secoes(i) = Totais_FormatoAvgSecao(secoes(i))
    Next i
    
    Totais_FormatoAvg = Join(secoes, ";")
End Function

Private Function Totais_FormatoAvgSecao(ByVal secao As String) As String
    Dim i As Long
    Dim ultimoPlaceholder As Long
    Dim posDecimal As Long
    Dim ch As String
    
    ultimoPlaceholder = 0
    
    For i = Len(secao) To 1 Step -1
        ch = Mid$(secao, i, 1)
        
        If ch = "0" Or ch = "#" Or ch = "?" Then
            ultimoPlaceholder = i
            Exit For
        End If
    Next i
    
    If ultimoPlaceholder = 0 Then
        Totais_FormatoAvgSecao = secao
        Exit Function
    End If
    
    posDecimal = 0
    
    For i = ultimoPlaceholder To 1 Step -1
        ch = Mid$(secao, i, 1)
        
        If ch = "." Then
            posDecimal = i
            Exit For
        End If
    Next i
    
    If posDecimal > 0 Then
        ' Já possui casas decimais. Mantém como está.
        Totais_FormatoAvgSecao = secao
    Else
        ' Não possui casas decimais. Acrescenta uma.
        Totais_FormatoAvgSecao = _
            Left$(secao, ultimoPlaceholder) & ".0" & Mid$(secao, ultimoPlaceholder + 1)
    End If
End Function

Private Function Totais_ArrayVazio(ByVal arr As Variant) As Boolean
    On Error GoTo Vazio
    
    Totais_ArrayVazio = (UBound(arr) < LBound(arr))
    Exit Function

Vazio:
    Totais_ArrayVazio = True
End Function

Private Function Totais_ValorLong(ByRef dic As Object, ByVal chave As String) As Long
    If dic.Exists(chave) Then
        Totais_ValorLong = CLng(dic(chave))
    Else
        Totais_ValorLong = 0
    End If
End Function

Private Function Totais_ValorDouble(ByRef dic As Object, ByVal chave As String) As Double
    If dic.Exists(chave) Then
        Totais_ValorDouble = CDbl(dic(chave))
    Else
        Totais_ValorDouble = 0#
    End If
End Function

Private Function Totais_ValorOuVazio(ByRef dic As Object, ByVal chave As String) As Variant
    If dic.Exists(chave) Then
        Totais_ValorOuVazio = dic(chave)
    Else
        Totais_ValorOuVazio = vbNullString
    End If
End Function

' Retorna um literal JSON sem formatação visual da célula.
' Use esta função no gerador de JSON para colunas *_Value:
' - número: sem aspas e com ponto decimal
' - data: com aspas em ISO yyyy-mm-dd
' - texto: com aspas e escape JSON
' - vazio: null
Public Function Wordex_JsonValorSemFormatacao(ByVal cell As Range) As String
    Dim v As Variant
    Dim valorTexto As String

    v = cell.Value

    If IsEmpty(v) Or Trim$(cell.Text) = vbNullString Then
        Wordex_JsonValorSemFormatacao = "null"
        Exit Function
    End If

    ' Data precisa ser testada antes de número, porque data no Excel
    ' pode ser armazenada internamente como número serial.
    If IsDate(v) And Not IsNumeric(cell.Text) Then
        Wordex_JsonValorSemFormatacao = """" & Format$(CDate(v), "yyyy-mm-dd") & """"
        Exit Function
    End If

    If IsNumeric(v) Then
        ' JSON/JavaScript exige ponto como separador decimal e não aceita
        ' separador de milhar em literais numéricos.
        valorTexto = Trim$(CStr(cell.Value2))
        valorTexto = Replace$(valorTexto, Application.ThousandsSeparator, vbNullString)
        valorTexto = Replace$(valorTexto, Application.DecimalSeparator, ".")
        valorTexto = Replace$(valorTexto, " ", vbNullString)

        Wordex_JsonValorSemFormatacao = valorTexto
        Exit Function
    End If

    Wordex_JsonValorSemFormatacao = """" & Wordex_JsonEscapeTexto(CStr(v)) & """"
End Function

Private Function Wordex_JsonEscapeTexto(ByVal texto As String) As String
    texto = Replace$(texto, "\", "\\")
    texto = Replace$(texto, Chr$(34), Chr$(92) & Chr$(34))
    texto = Replace$(texto, vbCrLf, "\n")
    texto = Replace$(texto, vbCr, "\n")
    texto = Replace$(texto, vbLf, "\n")

    Wordex_JsonEscapeTexto = texto
End Function

