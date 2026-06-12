Attribute VB_Name = "WordexConsulta"
Option Explicit

' =====================================================================
' WORDEX - Totais e gráficos calculados em memória (sem abas Totais*)
'
'   nomeAbaOrigem        - aba com os registros (ex.: "Produtos")
'   colunasAgrupadoras   - colunas de agrupamento separadas por vírgula;
'                          vazio = total geral (uma linha)
'   colunasTotalizadoras - colunas numéricas a totalizar (ex.: "Preço, Quantidade")
'   valoresCalcular      - operações (ex.: "Sum" ou "Count, Sum, Min, Max, Avg")
'   criterios (ParamArray) - pares opcionais NomeDaColuna, Valor
'
' Exemplos:
'   =ObterTotal("Produtos"; "ClienteId"; "Preço, Quantidade"; "Sum"; "ClienteId"; A2)
'   =ObterTotais("Produtos"; ""; "Preço, Quantidade"; "Count, Sum, Min, Max, Avg")
'   =ObterGrafico("Produtos"; "CategoriaId"; "Preço"; "Sum")
'   =ObterRegistroGrafico("Produtos"; "CategoriaId"; "Preço"; "Sum")
' =====================================================================

Private Const CONSULTA_FORMATO_COUNT As String = "###,###,###,##0"

Public Function ObterTotal( _
    ByVal nomeAbaOrigem As String, _
    ByVal colunasAgrupadoras As String, _
    ByVal colunasTotalizadoras As String, _
    ByVal valoresCalcular As String, _
    ParamArray criterios() As Variant _
) As String

    Dim arrCriterios As Variant

    On Error GoTo Erro

    arrCriterios = CVar(criterios)

    ObterTotal = Consulta_GerarJson( _
        nomeAbaOrigem, _
        colunasAgrupadoras, _
        colunasTotalizadoras, _
        valoresCalcular, _
        False, _
        arrCriterios, _
        True _
    )
    Exit Function

Erro:
    ObterTotal = "[{""Erro"": " & Consulta_JsonString(Err.Description) & "}]"
End Function

Public Function ObterGrafico( _
    ByVal nomeAbaOrigem As String, _
    ByVal colunasAgrupadoras As String, _
    ByVal colunasTotalizadoras As String, _
    ByVal valoresCalcular As String, _
    ParamArray criterios() As Variant _
) As String

    Dim arrCriterios As Variant

    On Error GoTo Erro

    arrCriterios = CVar(criterios)

    ObterGrafico = Consulta_GerarJson( _
        nomeAbaOrigem, _
        colunasAgrupadoras, _
        colunasTotalizadoras, _
        valoresCalcular, _
        True, _
        arrCriterios, _
        True _
    )
    Exit Function

Erro:
    ObterGrafico = "[{""Erro"": " & Consulta_JsonString(Err.Description) & "}]"
End Function

Public Function ObterTotais( _
    ByVal nomeAbaOrigem As String, _
    ByVal colunasAgrupadoras As String, _
    ByVal colunasTotalizadoras As String, _
    ByVal valoresCalcular As String, _
    ParamArray criterios() As Variant _
) As String

    Dim arrCriterios As Variant

    On Error GoTo Erro

    arrCriterios = CVar(criterios)

    ObterTotais = Consulta_RegistroParaArray(Consulta_GerarJson( _
        nomeAbaOrigem, _
        colunasAgrupadoras, _
        colunasTotalizadoras, _
        valoresCalcular, _
        False, _
        arrCriterios, _
        False _
    ))
    Exit Function

Erro:
    ObterTotais = "[{""Erro"": " & Consulta_JsonString(Err.Description) & "}]"
End Function

' Alias legado (=ObterRegistroTotal nas fórmulas antigas da planilha).
Public Function ObterRegistroTotal( _
    ByVal nomeAbaOrigem As String, _
    ByVal colunasAgrupadoras As String, _
    ByVal colunasTotalizadoras As String, _
    ByVal valoresCalcular As String, _
    ParamArray criterios() As Variant _
) As String

    ObterRegistroTotal = ObterTotais( _
        nomeAbaOrigem, _
        colunasAgrupadoras, _
        colunasTotalizadoras, _
        valoresCalcular, _
        criterios _
    )
End Function

Public Function ObterRegistroGrafico( _
    ByVal nomeAbaOrigem As String, _
    ByVal colunasAgrupadoras As String, _
    ByVal colunasTotalizadoras As String, _
    ByVal valoresCalcular As String, _
    ParamArray criterios() As Variant _
) As String

    Dim arrCriterios As Variant

    On Error GoTo Erro

    arrCriterios = CVar(criterios)

    ObterRegistroGrafico = Consulta_RegistroParaArray(Consulta_GerarJson( _
        nomeAbaOrigem, _
        colunasAgrupadoras, _
        colunasTotalizadoras, _
        valoresCalcular, _
        True, _
        arrCriterios, _
        False _
    ))
    Exit Function

Erro:
    ObterRegistroGrafico = "{""Erro"": " & Consulta_JsonString(Err.Description) & "}"
End Function

Private Function Consulta_GerarJson( _
    ByVal nomeAbaOrigem As String, _
    ByVal colunasAgrupadoras As String, _
    ByVal colunasTotalizadoras As String, _
    ByVal valoresCalcular As String, _
    ByVal gerarParaGrafico As Boolean, _
    ByVal criterios As Variant, _
    ByVal retornarArray As Boolean _
) As String

    Dim origem As Worksheet
    Dim arrAgrupadoras As Variant
    Dim arrTotalizadoras As Variant
    Dim arrCalcular As Variant
    Dim grupos As Object
    Dim valoresGrupos As Object

    Set origem = ThisWorkbook.Worksheets(nomeAbaOrigem)

    arrAgrupadoras = Consulta_QuebrarLista(colunasAgrupadoras)
    arrTotalizadoras = Consulta_QuebrarLista(colunasTotalizadoras)
    arrCalcular = Consulta_QuebrarLista(valoresCalcular)

    criterios = Consulta_NormalizarCriterios(criterios)
    Consulta_ValidarCriterios criterios

    Set grupos = CreateObject("Scripting.Dictionary")
    Set valoresGrupos = CreateObject("Scripting.Dictionary")

    Consulta_CalcularGrupos origem, arrAgrupadoras, arrTotalizadoras, criterios, grupos, valoresGrupos

    Consulta_GerarJson = Consulta_GruposParaJson( _
        origem, _
        grupos, _
        valoresGrupos, _
        arrAgrupadoras, _
        arrTotalizadoras, _
        arrCalcular, _
        gerarParaGrafico, _
        criterios, _
        retornarArray _
    )
End Function

Private Sub Consulta_CalcularGrupos( _
    ByRef origem As Worksheet, _
    ByVal colunasAgrupadoras As Variant, _
    ByVal colunasTotalizadoras As Variant, _
    ByVal criterios As Variant, _
    ByRef grupos As Object, _
    ByRef valoresGrupos As Object _
)
    Dim linha As Long
    Dim chave As String
    Dim stats As Object

    linha = K_DETAILS_ROW

    Do While origem.Cells(linha, K_REQUIRED_COL).Text <> vbNullString
        If Consulta_RegistroAtendeCriterios(origem, linha, criterios, True) Then
            chave = Consulta_ObterChaveGrupo(origem, linha, colunasAgrupadoras)

            If Not grupos.Exists(chave) Then
                Set stats = CreateObject("Scripting.Dictionary")
                grupos.Add chave, stats
                valoresGrupos.Add chave, Consulta_ObterValoresGrupo(origem, linha, colunasAgrupadoras)
            Else
                Set stats = grupos(chave)
            End If

            Consulta_AcumularLinha origem, linha, colunasTotalizadoras, stats
        End If

        linha = linha + 1
    Loop
End Sub

Private Function Consulta_GruposParaJson( _
    ByRef origem As Worksheet, _
    ByRef grupos As Object, _
    ByRef valoresGrupos As Object, _
    ByVal colunasAgrupadoras As Variant, _
    ByVal colunasTotalizadoras As Variant, _
    ByVal valoresCalcular As Variant, _
    ByVal gerarParaGrafico As Boolean, _
    ByVal criterios As Variant, _
    ByVal retornarArray As Boolean _
) As String

    Dim json As String
    Dim virgula As String
    Dim chave As Variant
    Dim stats As Object
    Dim valores As Variant
    Dim objJson As String

    If Consulta_ArrayVazio(colunasTotalizadoras) Then
        Err.Raise vbObjectError + 5100, "ObterTotal", _
            "Selecione pelo menos uma coluna para totalizar."
    End If

    If Consulta_ArrayVazio(valoresCalcular) Then
        Err.Raise vbObjectError + 5101, "ObterTotal", _
            "Selecione pelo menos um valor para calcular: Sum, Min, Max, Avg ou Count."
    End If

    Consulta_ValidarValoresCalcular valoresCalcular

    json = vbNullString
    virgula = vbNullString

    For Each chave In grupos.Keys
        Set stats = grupos(chave)
        valores = valoresGrupos(chave)

        If Consulta_GrupoAtendeCriterios( _
            origem, colunasAgrupadoras, valores, colunasTotalizadoras, valoresCalcular, stats, gerarParaGrafico, criterios _
        ) Then
            objJson = Consulta_GrupoParaJsonObject( _
                origem, colunasAgrupadoras, valores, colunasTotalizadoras, valoresCalcular, stats, gerarParaGrafico _
            )

            If Not retornarArray Then
                Consulta_GruposParaJson = objJson
                Exit Function
            End If

            json = json & virgula & objJson
            virgula = ","
        End If
    Next chave

    If retornarArray Then
        Consulta_GruposParaJson = "[" & json & "]"
    Else
        Consulta_GruposParaJson = "{}"
    End If
End Function

Private Function Consulta_GrupoParaJsonObject( _
    ByRef origem As Worksheet, _
    ByVal colunasAgrupadoras As Variant, _
    ByVal valoresGrupo As Variant, _
    ByVal colunasTotalizadoras As Variant, _
    ByVal valoresCalcular As Variant, _
    ByRef stats As Object, _
    ByVal gerarParaGrafico As Boolean _
) As String

    Dim json As String
    Dim virgula As String
    Dim i As Long
    Dim j As Long
    Dim campo As String
    Dim tipoTotal As String
    Dim nomeTotal As String
    Dim valorTotal As Variant

    json = "{"
    virgula = vbNullString

    If Not Consulta_ArrayVazio(colunasAgrupadoras) Then
        For i = LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras)
            campo = CStr(colunasAgrupadoras(i))

            json = json & virgula & Consulta_JsonString(campo) & ": "
            json = json & Wordex_JsonCampoValor(origem, campo, valoresGrupo(i))

            virgula = ","
        Next i
    End If

    For i = LBound(colunasTotalizadoras) To UBound(colunasTotalizadoras)
        campo = CStr(colunasTotalizadoras(i))

        For j = LBound(valoresCalcular) To UBound(valoresCalcular)
            tipoTotal = Consulta_NormalizarTipoTotal(CStr(valoresCalcular(j)))
            nomeTotal = campo & tipoTotal
            valorTotal = Consulta_CalcularValorTotal(stats, campo, tipoTotal)

            json = json & virgula & Consulta_JsonString(nomeTotal) & ": "
            json = json & Wordex_JsonCampoValor(origem, campo, valorTotal)
            virgula = ","

            If gerarParaGrafico Then
                json = json & virgula & Consulta_JsonString(nomeTotal & "_Label") & ": "
                json = json & Wordex_JsonCampoValor(origem, nomeTotal & "_Label", Consulta_LegendaGrafico(nomeTotal))
                virgula = ","
            End If
        Next j
    Next i

    json = json & "}"
    Consulta_GrupoParaJsonObject = json
End Function

Private Function Consulta_RegistroParaArray(ByVal conteudo As String) As String
    conteudo = Trim$(conteudo)

    If conteudo = vbNullString Or conteudo = "{}" Then
        Consulta_RegistroParaArray = "[]"
    ElseIf Left$(conteudo, 1) = "[" Then
        Consulta_RegistroParaArray = conteudo
    Else
        Consulta_RegistroParaArray = "[" & conteudo & "]"
    End If
End Function

Private Function Consulta_GrupoAtendeCriterios( _
    ByRef origem As Worksheet, _
    ByVal colunasAgrupadoras As Variant, _
    ByVal valoresGrupo As Variant, _
    ByVal colunasTotalizadoras As Variant, _
    ByVal valoresCalcular As Variant, _
    ByRef stats As Object, _
    ByVal gerarParaGrafico As Boolean, _
    ByVal criterios As Variant _
) As Boolean

    Dim i As Long
    Dim qtd As Long
    Dim nomeCampo As String
    Dim valorEsperado As Variant
    Dim valorAtual As Variant
    Dim colunaEncontrada As Boolean

    qtd = Consulta_QuantidadeCriterios(criterios)

    If qtd = 0 Then
        Consulta_GrupoAtendeCriterios = True
        Exit Function
    End If

    For i = LBound(criterios) To UBound(criterios) Step 2
        nomeCampo = CStr(criterios(i))
        valorEsperado = criterios(i + 1)
        valorAtual = Consulta_ObterValorCampoGrupo( _
            origem, colunasAgrupadoras, valoresGrupo, colunasTotalizadoras, valoresCalcular, stats, gerarParaGrafico, nomeCampo, colunaEncontrada _
        )

        If colunaEncontrada Then
            If valorAtual <> valorEsperado Then
                Consulta_GrupoAtendeCriterios = False
                Exit Function
            End If
        End If
    Next i

    Consulta_GrupoAtendeCriterios = True
End Function

Private Function Consulta_ObterValorCampoGrupo( _
    ByRef origem As Worksheet, _
    ByVal colunasAgrupadoras As Variant, _
    ByVal valoresGrupo As Variant, _
    ByVal colunasTotalizadoras As Variant, _
    ByVal valoresCalcular As Variant, _
    ByRef stats As Object, _
    ByVal gerarParaGrafico As Boolean, _
    ByVal nomeCampo As String, _
    ByRef colunaEncontrada As Boolean _
) As Variant

    Dim i As Long
    Dim j As Long
    Dim campo As String
    Dim tipoTotal As String
    Dim nomeTotal As String
    Dim valorGrafico As Variant
    Dim fonteEhData As Boolean
    Dim colunaOrigem As Long

    colunaEncontrada = False

    If Not Consulta_ArrayVazio(colunasAgrupadoras) Then
        For i = LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras)
            If CStr(colunasAgrupadoras(i)) = nomeCampo Then
                Consulta_ObterValorCampoGrupo = valoresGrupo(i)
                colunaEncontrada = True
                Exit Function
            End If
        Next i
    End If

    For i = LBound(colunasTotalizadoras) To UBound(colunasTotalizadoras)
        campo = CStr(colunasTotalizadoras(i))
        colunaOrigem = Consulta_ObterNumeroColuna(origem, campo)
        fonteEhData = Consulta_ColunaPareceData(origem, colunaOrigem)

        For j = LBound(valoresCalcular) To UBound(valoresCalcular)
            tipoTotal = Consulta_NormalizarTipoTotal(CStr(valoresCalcular(j)))
            nomeTotal = campo & tipoTotal

            If nomeCampo = nomeTotal Then
                Consulta_ObterValorCampoGrupo = Consulta_CalcularValorTotal(stats, campo, tipoTotal)
                colunaEncontrada = True
                Exit Function
            End If

            If gerarParaGrafico Then
                If nomeCampo = nomeTotal & "_Label" Then
                    Consulta_ObterValorCampoGrupo = Consulta_LegendaGrafico(nomeTotal)
                    colunaEncontrada = True
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Function Consulta_ObterFormatoTotal( _
    ByRef origem As Worksheet, _
    ByVal colunaOrigem As Long, _
    ByVal tipoTotal As String _
) As String

    Select Case tipoTotal
        Case "Count"
            Consulta_ObterFormatoTotal = CONSULTA_FORMATO_COUNT

        Case "Avg"
            Consulta_ObterFormatoTotal = Consulta_FormatoAvg(origem.Columns(colunaOrigem).NumberFormat)

        Case Else
            Consulta_ObterFormatoTotal = origem.Columns(colunaOrigem).NumberFormat
    End Select
End Function

Private Function Consulta_ValorComoTexto(ByVal valor As Variant, ByVal numberFormat As String) As String
    If IsEmpty(valor) Or CStr(valor) = vbNullString Then
        Consulta_ValorComoTexto = vbNullString
    Else
        Consulta_ValorComoTexto = Format$(valor, numberFormat)
    End If
End Function

Private Sub Consulta_AcumularLinha( _
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
        coluna = Consulta_ObterNumeroColuna(plan, campo)
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

Private Sub Consulta_ValidarValoresCalcular(ByVal valoresCalcular As Variant)
    Dim i As Long

    For i = LBound(valoresCalcular) To UBound(valoresCalcular)
        Consulta_NormalizarTipoTotal CStr(valoresCalcular(i))
    Next i
End Sub

Private Function Consulta_NormalizarTipoTotal(ByVal tipoTotal As String) As String
    Select Case LCase$(Trim$(tipoTotal))
        Case "sum"
            Consulta_NormalizarTipoTotal = "Sum"
        Case "min"
            Consulta_NormalizarTipoTotal = "Min"
        Case "max"
            Consulta_NormalizarTipoTotal = "Max"
        Case "avg"
            Consulta_NormalizarTipoTotal = "Avg"
        Case "count"
            Consulta_NormalizarTipoTotal = "Count"
        Case Else
            Err.Raise vbObjectError + 5102, "ObterTotal", _
                "Valor a calcular inválido: '" & tipoTotal & "'. Use Sum, Min, Max, Avg ou Count."
    End Select
End Function

Private Function Consulta_CalcularValorTotal( _
    ByRef stats As Object, _
    ByVal campo As String, _
    ByVal tipoTotal As String _
) As Variant

    Dim countValue As Long
    Dim sumValue As Double

    Select Case tipoTotal
        Case "Count"
            Consulta_CalcularValorTotal = Consulta_ValorLong(stats, campo & "Count")

        Case "Sum"
            Consulta_CalcularValorTotal = Consulta_ValorDouble(stats, campo & "Sum")

        Case "Min"
            Consulta_CalcularValorTotal = Consulta_ValorOuVazio(stats, campo & "Min")

        Case "Max"
            Consulta_CalcularValorTotal = Consulta_ValorOuVazio(stats, campo & "Max")

        Case "Avg"
            countValue = Consulta_ValorLong(stats, campo & "Count")
            sumValue = Consulta_ValorDouble(stats, campo & "Sum")

            If countValue > 0 Then
                Consulta_CalcularValorTotal = sumValue / countValue
            Else
                Consulta_CalcularValorTotal = vbNullString
            End If
    End Select
End Function

Private Function Consulta_LegendaGrafico(ByVal nomeTotal As String) As String
    Dim sufixos As Variant
    Dim i As Long
    Dim sufixo As String

    sufixos = Array("Count", "Sum", "Min", "Max", "Avg")

    For i = LBound(sufixos) To UBound(sufixos)
        sufixo = CStr(sufixos(i))

        If Len(nomeTotal) > Len(sufixo) Then
            If Right$(nomeTotal, Len(sufixo)) = sufixo Then
                Consulta_LegendaGrafico = Left$(nomeTotal, Len(nomeTotal) - Len(sufixo))
                Exit Function
            End If
        End If
    Next i

    Consulta_LegendaGrafico = nomeTotal
End Function

Private Function Consulta_ValorGrafico( _
    ByVal valorTotal As Variant, _
    ByVal fonteEhData As Boolean, _
    ByVal tipoTotal As String _
) As Variant

    If IsEmpty(valorTotal) Or CStr(valorTotal) = vbNullString Then
        Consulta_ValorGrafico = vbNullString
        Exit Function
    End If

    If fonteEhData And (tipoTotal = "Min" Or tipoTotal = "Max") Then
        If IsNumeric(valorTotal) Then
            Consulta_ValorGrafico = Format$(CDate(CDbl(valorTotal)), "yyyy-mm-dd")
            Exit Function
        End If

        If IsDate(valorTotal) Then
            Consulta_ValorGrafico = Format$(CDate(valorTotal), "yyyy-mm-dd")
            Exit Function
        End If
    End If

    Consulta_ValorGrafico = valorTotal
End Function

Private Function Consulta_ColunaPareceData(ByRef plan As Worksheet, ByVal coluna As Long) As Boolean
    Dim formato As String
    Dim linha As Long
    Dim v As Variant

    formato = LCase$(CStr(plan.Columns(coluna).NumberFormat))

    If Consulta_FormatoPareceData(formato) Then
        Consulta_ColunaPareceData = True
        Exit Function
    End If

    linha = K_DETAILS_ROW

    Do While plan.Cells(linha, K_REQUIRED_COL).Text <> vbNullString
        v = plan.Cells(linha, coluna).Value

        If IsDate(v) And Not IsNumeric(plan.Cells(linha, coluna).Text) Then
            Consulta_ColunaPareceData = True
            Exit Function
        End If

        linha = linha + 1
    Loop
End Function

Private Function Consulta_FormatoPareceData(ByVal formato As String) As Boolean
    Dim limpo As String

    limpo = Consulta_RemoverLiteraisFormato(LCase$(formato))

    If InStr(1, limpo, "[$-f800]", vbTextCompare) > 0 Then
        Consulta_FormatoPareceData = True
        Exit Function
    End If

    Consulta_FormatoPareceData = _
        (InStr(1, limpo, "d", vbTextCompare) > 0 Or InStr(1, limpo, "y", vbTextCompare) > 0 Or InStr(1, limpo, "a", vbTextCompare) > 0) _
        And InStr(1, limpo, "m", vbTextCompare) > 0
End Function

Private Function Consulta_RemoverLiteraisFormato(ByVal formato As String) As String
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

    Consulta_RemoverLiteraisFormato = resultado
End Function

Private Function Consulta_ObterChaveGrupo( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal colunasAgrupadoras As Variant _
) As String

    Dim i As Long
    Dim chave As String
    Dim coluna As Long

    If Consulta_ArrayVazio(colunasAgrupadoras) Then
        Consulta_ObterChaveGrupo = "#TOTAL_GERAL#"
        Exit Function
    End If

    chave = vbNullString

    For i = LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras)
        coluna = Consulta_ObterNumeroColuna(plan, CStr(colunasAgrupadoras(i)))
        chave = chave & ChrW(30) & CStr(plan.Cells(linha, coluna).Value2)
    Next i

    Consulta_ObterChaveGrupo = chave
End Function

Private Function Consulta_ObterValoresGrupo( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal colunasAgrupadoras As Variant _
) As Variant

    Dim valores() As Variant
    Dim i As Long
    Dim coluna As Long

    If Consulta_ArrayVazio(colunasAgrupadoras) Then
        Consulta_ObterValoresGrupo = Array()
        Exit Function
    End If

    ReDim valores(LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras))

    For i = LBound(colunasAgrupadoras) To UBound(colunasAgrupadoras)
        coluna = Consulta_ObterNumeroColuna(plan, CStr(colunasAgrupadoras(i)))
        valores(i) = plan.Cells(linha, coluna).Value2
    Next i

    Consulta_ObterValoresGrupo = valores
End Function

Private Function Consulta_ObterNumeroColuna(ByRef plan As Worksheet, ByVal tituloColuna As String) As Long
    Dim coluna As Long

    coluna = 1

    Do While plan.Cells(K_HEADERS_ROW, coluna).Text <> vbNullString
        If plan.Cells(K_HEADERS_ROW, coluna).Text = tituloColuna Then
            Consulta_ObterNumeroColuna = coluna
            Exit Function
        End If

        coluna = coluna + 1
    Loop

    Err.Raise vbObjectError + 5103, "Consulta_ObterNumeroColuna", _
        "Coluna '" & tituloColuna & "' não encontrada na aba '" & plan.Name & "'."
End Function

Private Function Consulta_FormatoAvg(ByVal formatoOriginal As String) As String
    Dim secoes() As String
    Dim i As Long

    formatoOriginal = Trim$(formatoOriginal)

    If formatoOriginal = vbNullString Or LCase$(formatoOriginal) = "general" Then
        Consulta_FormatoAvg = "0.0"
        Exit Function
    End If

    secoes = Split(formatoOriginal, ";")

    For i = LBound(secoes) To UBound(secoes)
        secoes(i) = Consulta_FormatoAvgSecao(secoes(i))
    Next i

    Consulta_FormatoAvg = Join(secoes, ";")
End Function

Private Function Consulta_FormatoAvgSecao(ByVal secao As String) As String
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
        Consulta_FormatoAvgSecao = secao
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
        Consulta_FormatoAvgSecao = secao
    Else
        Consulta_FormatoAvgSecao = _
            Left$(secao, ultimoPlaceholder) & ".0" & Mid$(secao, ultimoPlaceholder + 1)
    End If
End Function

Private Function Consulta_ArrayVazio(ByVal arr As Variant) As Boolean
    On Error GoTo Vazio

    Consulta_ArrayVazio = (UBound(arr) < LBound(arr))
    Exit Function

Vazio:
    Consulta_ArrayVazio = True
End Function

Private Function Consulta_ValorLong(ByRef dic As Object, ByVal chave As String) As Long
    If dic.Exists(chave) Then
        Consulta_ValorLong = CLng(dic(chave))
    Else
        Consulta_ValorLong = 0
    End If
End Function

Private Function Consulta_ValorDouble(ByRef dic As Object, ByVal chave As String) As Double
    If dic.Exists(chave) Then
        Consulta_ValorDouble = CDbl(dic(chave))
    Else
        Consulta_ValorDouble = 0#
    End If
End Function

Private Function Consulta_ValorOuVazio(ByRef dic As Object, ByVal chave As String) As Variant
    If dic.Exists(chave) Then
        Consulta_ValorOuVazio = dic(chave)
    Else
        Consulta_ValorOuVazio = vbNullString
    End If
End Function

Private Function Consulta_QuebrarLista(ByVal texto As Variant) As Variant
    Dim bruto As Variant
    Dim itens() As String
    Dim i As Long
    Dim qtd As Long
    Dim item As String
    Dim textoLimpo As String

    If IsMissing(texto) Or IsEmpty(texto) Or IsNull(texto) Then
        Consulta_QuebrarLista = Array()
        Exit Function
    End If

    textoLimpo = Trim$(CStr(texto))

    If textoLimpo = vbNullString Then
        Consulta_QuebrarLista = Array()
        Exit Function
    End If

    bruto = Split(textoLimpo, ",")
    ReDim itens(0 To UBound(bruto))

    For i = LBound(bruto) To UBound(bruto)
        item = Trim$(CStr(bruto(i)))

        If item <> vbNullString Then
            itens(qtd) = item
            qtd = qtd + 1
        End If
    Next i

    If qtd = 0 Then
        Consulta_QuebrarLista = Array()
    Else
        ReDim Preserve itens(0 To qtd - 1)
        Consulta_QuebrarLista = itens
    End If
End Function

Private Function Consulta_NormalizarCriterios(ByVal criterios As Variant) As Variant
    On Error GoTo Vazio

    If Not IsArray(criterios) Then GoTo Vazio

    If (UBound(criterios) - LBound(criterios) + 1) = 1 Then
        If IsArray(criterios(LBound(criterios))) Then
            Consulta_NormalizarCriterios = Consulta_NormalizarCriterios(criterios(LBound(criterios)))
            Exit Function
        End If

        If IsMissing(criterios(LBound(criterios))) _
           Or IsEmpty(criterios(LBound(criterios))) Then
            GoTo Vazio
        End If
    End If

    If UBound(criterios) >= LBound(criterios) Then
        Consulta_NormalizarCriterios = criterios
        Exit Function
    End If

Vazio:
    Consulta_NormalizarCriterios = Array()
End Function

Private Function Consulta_RegistroAtendeCriterios( _
    ByRef plan As Worksheet, _
    ByVal numeroLinha As Long, _
    ByVal criterios As Variant, _
    Optional ByVal ignorarColunaAusente As Boolean = False _
) As Boolean

    Dim i As Long
    Dim qtd As Long
    Dim nomeCampo As String
    Dim valorEsperado As Variant
    Dim valorAtual As Variant
    Dim numeroColuna As Long

    qtd = Consulta_QuantidadeCriterios(criterios)

    If qtd = 0 Then
        Consulta_RegistroAtendeCriterios = True
        Exit Function
    End If

    For i = LBound(criterios) To UBound(criterios) Step 2
        nomeCampo = CStr(criterios(i))
        valorEsperado = criterios(i + 1)

        If ignorarColunaAusente Then
            numeroColuna = Consulta_TentarObterNumeroColuna(plan, nomeCampo)

            If numeroColuna = 0 Then GoTo ProximoPar

            valorAtual = plan.Cells(numeroLinha, numeroColuna).Value2
        Else
            valorAtual = Consulta_ObterValorColuna(plan, numeroLinha, nomeCampo)
        End If

        If valorAtual <> valorEsperado Then
            Consulta_RegistroAtendeCriterios = False
            Exit Function
        End If

ProximoPar:
    Next i

    Consulta_RegistroAtendeCriterios = True
End Function

Private Function Consulta_TentarObterNumeroColuna( _
    ByRef plan As Worksheet, _
    ByVal tituloColuna As String _
) As Long

    Dim coluna As Long

    coluna = 1

    Do While plan.Cells(K_HEADERS_ROW, coluna).Text <> vbNullString
        If plan.Cells(K_HEADERS_ROW, coluna).Text = tituloColuna Then
            Consulta_TentarObterNumeroColuna = coluna
            Exit Function
        End If

        coluna = coluna + 1
    Loop
End Function

Private Sub Consulta_ValidarCriterios(ByVal criterios As Variant)
    Dim qtd As Long

    qtd = Consulta_QuantidadeCriterios(criterios)

    If qtd = 0 Then Exit Sub

    If qtd Mod 2 <> 0 Then
        Err.Raise vbObjectError + 5104, "Consulta_ValidarCriterios", _
            "Critérios devem estar em pares: NomeDaColuna, Valor."
    End If
End Sub

Private Function Consulta_QuantidadeCriterios(ByVal criterios As Variant) As Long
    On Error GoTo SemCriterios

    If Not IsArray(criterios) Then
        Consulta_QuantidadeCriterios = 0
        Exit Function
    End If

    If UBound(criterios) < LBound(criterios) Then
        Consulta_QuantidadeCriterios = 0
        Exit Function
    End If

    Consulta_QuantidadeCriterios = UBound(criterios) - LBound(criterios) + 1
    Exit Function

SemCriterios:
    Consulta_QuantidadeCriterios = 0
End Function

Private Function Consulta_ObterValorColuna( _
    ByRef plan As Worksheet, _
    ByVal numeroLinha As Long, _
    ByVal tituloColuna As String _
) As Variant

    Dim numeroColuna As Long

    numeroColuna = Consulta_ObterNumeroColuna(plan, tituloColuna)
    Consulta_ObterValorColuna = plan.Cells(numeroLinha, numeroColuna).Value2
End Function

Private Function Consulta_JsonString(ByVal valor As String) As String
    Consulta_JsonString = """" & Consulta_JsonEscape(valor) & """"
End Function

Private Function Consulta_JsonEscape(ByVal valor As String) As String
    valor = Replace(valor, "\", "\\")
    valor = Replace(valor, """", "\""")
    valor = Replace(valor, vbCrLf, "\n")
    valor = Replace(valor, vbCr, "\n")
    valor = Replace(valor, vbLf, "\n")
    valor = Replace(valor, vbTab, "\t")

    Consulta_JsonEscape = valor
End Function
