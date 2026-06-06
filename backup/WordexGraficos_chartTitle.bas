Attribute VB_Name = "WordexGraficos"
Option Explicit

Private Const WX_HEADERS_ROW As Long = 1
Private Const WX_DETAILS_ROW As Long = 2
Private Const WX_REQUIRED_COL As Long = 1
Private Const WX_SEP As String = "|~WX~|"

' =====================================================================
' WORDEX - Funções de gráfico para uso direto em célula da planilha
'
' Exemplo 1D:
'   =ObterGrafico("Produtos"; "ClienteId, CategoriaId"; "Preço"; "Sum"; "GRÁFICO DE VENDAS"; A4; "ClienteId"; A3)
'
' Exemplo 2D:
'   =ObterGrafico2D("Vendas"; "Cliente, Mes"; "Preço"; "Sum"; "VENDAS POR CLIENTE/MÊS"; "Cliente")
'
' Regras:
'   - camposGrupo aceita lista separada por vírgula.
'   - camposValor e agregacoes devem ter apenas um item para gráfico.
'   - criterios vêm em pares: "Campo"; Valor.
'   - chartTitle pode ser uma célula ou texto literal.
'   - tituloItem pode ser uma célula, texto literal ou nome de coluna.
' =====================================================================

Public Function ObterGrafico( _
    ByVal nomeAba As String, _
    ByVal camposGrupo As String, _
    ByVal camposValor As String, _
    ByVal agregacoes As String, _
    ByVal chartTitle As Variant, _
    ByVal tituloItem As Variant, _
    ParamArray criterios() As Variant _
) As String
    On Error GoTo Erro

    Dim plan As Worksheet
    Dim grupos As Variant
    Dim campoValor As String
    Dim operacao As String
    Dim colunaValor As Long
    Dim dic As Object
    Dim ordem As Collection
    Dim linha As Long
    Dim chave As String
    Dim stats As Object
    Dim arrCriterios As Variant
    Dim tituloCalculado As String

    Set plan = ThisWorkbook.Worksheets(nomeAba)

    grupos = WX_QuebrarLista(camposGrupo)
    campoValor = WX_UnicoItem(camposValor, "camposValor")
    operacao = WX_NormalizarOperacao(WX_UnicoItem(agregacoes, "agregacoes"))
    colunaValor = WX_ObterNumeroColuna(plan, campoValor)

    arrCriterios = CVar(criterios)
    WX_ValidarCriterios arrCriterios

    Set dic = CreateObject("Scripting.Dictionary")
    Set ordem = New Collection

    linha = WX_DETAILS_ROW

    Do While plan.Cells(linha, WX_REQUIRED_COL).Text <> vbNullString
        If WX_RegistroAtendeCriterios(plan, linha, arrCriterios) Then
            chave = WX_ObterChaveGrupo(plan, linha, grupos)

            If Not dic.Exists(chave) Then
                Set stats = WX_NovoStats()
                tituloCalculado = WX_ResolverTituloItem(plan, linha, tituloItem, grupos)
                stats("Title") = tituloCalculado
                dic.Add chave, stats
                ordem.Add chave
            Else
                Set stats = dic(chave)
            End If

            WX_Acumular plan.Cells(linha, colunaValor), stats
        End If

        linha = linha + 1
    Loop

    ObterGrafico = WX_MontarJsonGrafico(plan, colunaValor, campoValor, operacao, chartTitle, dic, ordem)
    Exit Function

Erro:
    ObterGrafico = "{""Erro"":" & WX_JsonString(Err.Description) & "}"
End Function

Public Function ObterGrafico2D( _
    ByVal nomeAba As String, _
    ByVal camposGrupo As String, _
    ByVal camposValor As String, _
    ByVal agregacoes As String, _
    ByVal chartTitle As Variant, _
    ByVal tituloItem As Variant, _
    ParamArray criterios() As Variant _
) As String
    On Error GoTo Erro

    Dim plan As Worksheet
    Dim grupos As Variant
    Dim campoValor As String
    Dim operacao As String
    Dim colunaValor As Long
    Dim colunaGrupoLinha As Long
    Dim colunaGrupoColuna As Long
    Dim dicLinhas As Object
    Dim dicColunas As Object
    Dim dicCelulas As Object
    Dim ordemLinhas As Collection
    Dim ordemColunas As Collection
    Dim linha As Long
    Dim chaveLinha As String
    Dim chaveColuna As String
    Dim chaveCelula As String
    Dim stats As Object
    Dim arrCriterios As Variant
    Dim tituloLinha As String
    Dim tituloColuna As String

    Set plan = ThisWorkbook.Worksheets(nomeAba)

    grupos = WX_QuebrarLista(camposGrupo)

    If WX_ArrayQuantidade(grupos) <> 2 Then
        Err.Raise vbObjectError + 4100, "ObterGrafico2D", _
            "ObterGrafico2D precisa de exatamente dois campos em camposGrupo. Exemplo: ""Cliente, Mes""."
    End If

    campoValor = WX_UnicoItem(camposValor, "camposValor")
    operacao = WX_NormalizarOperacao(WX_UnicoItem(agregacoes, "agregacoes"))
    colunaValor = WX_ObterNumeroColuna(plan, campoValor)
    colunaGrupoLinha = WX_ObterNumeroColuna(plan, CStr(grupos(LBound(grupos))))
    colunaGrupoColuna = WX_ObterNumeroColuna(plan, CStr(grupos(LBound(grupos) + 1)))

    arrCriterios = CVar(criterios)
    WX_ValidarCriterios arrCriterios

    Set dicLinhas = CreateObject("Scripting.Dictionary")
    Set dicColunas = CreateObject("Scripting.Dictionary")
    Set dicCelulas = CreateObject("Scripting.Dictionary")
    Set ordemLinhas = New Collection
    Set ordemColunas = New Collection

    linha = WX_DETAILS_ROW

    Do While plan.Cells(linha, WX_REQUIRED_COL).Text <> vbNullString
        If WX_RegistroAtendeCriterios(plan, linha, arrCriterios) Then
            chaveLinha = CStr(plan.Cells(linha, colunaGrupoLinha).Value2)
            chaveColuna = CStr(plan.Cells(linha, colunaGrupoColuna).Value2)

            tituloLinha = WX_ResolverTituloItem(plan, linha, tituloItem, grupos)
            tituloColuna = plan.Cells(linha, colunaGrupoColuna).Text

            If Not dicLinhas.Exists(chaveLinha) Then
                dicLinhas.Add chaveLinha, tituloLinha
                ordemLinhas.Add chaveLinha
            End If

            If Not dicColunas.Exists(chaveColuna) Then
                dicColunas.Add chaveColuna, tituloColuna
                ordemColunas.Add chaveColuna
            End If

            chaveCelula = chaveLinha & WX_SEP & chaveColuna

            If Not dicCelulas.Exists(chaveCelula) Then
                Set stats = WX_NovoStats()
                dicCelulas.Add chaveCelula, stats
            Else
                Set stats = dicCelulas(chaveCelula)
            End If

            WX_Acumular plan.Cells(linha, colunaValor), stats
        End If

        linha = linha + 1
    Loop

    ObterGrafico2D = WX_MontarJsonGrafico2D(plan, colunaValor, campoValor, operacao, chartTitle, _
                                           dicLinhas, dicColunas, dicCelulas, _
                                           ordemLinhas, ordemColunas)
    Exit Function

Erro:
    ObterGrafico2D = "{""Erro"":" & WX_JsonString(Err.Description) & "}"
End Function

Private Function WX_MontarJsonGrafico( _
    ByRef plan As Worksheet, _
    ByVal colunaValor As Long, _
    ByVal campoValor As String, _
    ByVal operacao As String, _
    ByVal chartTitle As Variant, _
    ByRef dic As Object, _
    ByRef ordem As Collection _
) As String
    Dim json As String
    Dim virgula As String
    Dim i As Long
    Dim chave As String
    Dim stats As Object
    Dim valor As Variant
    Dim label As String

    json = "{" & _
           """ChartTitle"":" & WX_JsonString(WX_TextoParametro(chartTitle)) & "," & _
           """ItemTitle"":" & WX_JsonString(campoValor) & "," & _
           """Items"":["

    virgula = vbNullString

    For i = 1 To ordem.Count
        chave = CStr(ordem(i))
        Set stats = dic(chave)
        valor = WX_ObterValorAgregado(stats, operacao)
        label = WX_FormatarLabel(plan, colunaValor, valor, operacao)

        json = json & virgula & "{" & _
               """title"":" & WX_JsonString(CStr(stats("Title"))) & "," & _
               """value"":" & WX_JsonValorNumerico(valor) & "," & _
               """label"":" & WX_JsonString(label) & _
               "}"

        virgula = ","
    Next i

    json = json & "]}"
    WX_MontarJsonGrafico = json
End Function

Private Function WX_MontarJsonGrafico2D( _
    ByRef plan As Worksheet, _
    ByVal colunaValor As Long, _
    ByVal campoValor As String, _
    ByVal operacao As String, _
    ByVal chartTitle As Variant, _
    ByRef dicLinhas As Object, _
    ByRef dicColunas As Object, _
    ByRef dicCelulas As Object, _
    ByRef ordemLinhas As Collection, _
    ByRef ordemColunas As Collection _
) As String
    Dim json As String
    Dim virgula As String
    Dim virgulaValores As String
    Dim i As Long
    Dim j As Long
    Dim chaveLinha As String
    Dim chaveColuna As String
    Dim chaveCelula As String
    Dim stats As Object
    Dim valor As Variant
    Dim label As String

    json = "{" & _
           """ChartTitle"":" & WX_JsonString(WX_TextoParametro(chartTitle)) & "," & _
           """ItemTitle"":" & WX_JsonString(campoValor) & "," & _
           """Columns"":["

    virgula = vbNullString

    For j = 1 To ordemColunas.Count
        chaveColuna = CStr(ordemColunas(j))
        json = json & virgula & "{""title"":" & WX_JsonString(CStr(dicColunas(chaveColuna))) & "}"
        virgula = ","
    Next j

    json = json & "],""Items"":["

    virgula = vbNullString

    For i = 1 To ordemLinhas.Count
        chaveLinha = CStr(ordemLinhas(i))

        json = json & virgula & "{" & _
               """title"":" & WX_JsonString(CStr(dicLinhas(chaveLinha))) & "," & _
               """values"":["

        virgulaValores = vbNullString

        For j = 1 To ordemColunas.Count
            chaveColuna = CStr(ordemColunas(j))
            chaveCelula = chaveLinha & WX_SEP & chaveColuna

            If dicCelulas.Exists(chaveCelula) Then
                Set stats = dicCelulas(chaveCelula)
                valor = WX_ObterValorAgregado(stats, operacao)
                label = WX_FormatarLabel(plan, colunaValor, valor, operacao)
            Else
                valor = Empty
                label = vbNullString
            End If

            json = json & virgulaValores & "{" & _
                   """title"":" & WX_JsonString(CStr(dicColunas(chaveColuna))) & "," & _
                   """value"":" & WX_JsonValorNumerico(valor) & "," & _
                   """label"":" & WX_JsonString(label) & _
                   "}"

            virgulaValores = ","
        Next j

        json = json & "]}"
        virgula = ","
    Next i

    json = json & "]}"
    WX_MontarJsonGrafico2D = json
End Function

Private Function WX_NovoStats() As Object
    Dim stats As Object

    Set stats = CreateObject("Scripting.Dictionary")
    stats.Add "Count", 0&
    stats.Add "Sum", 0#
    stats.Add "Min", Empty
    stats.Add "Max", Empty

    Set WX_NovoStats = stats
End Function

Private Sub WX_Acumular(ByRef cel As Range, ByRef stats As Object)
    Dim valor As Variant
    Dim numero As Double

    valor = cel.Value2

    If Trim$(cel.Text) = vbNullString Then Exit Sub

    If IsNumeric(valor) Then
        numero = CDbl(valor)

        stats("Count") = CLng(stats("Count")) + 1
        stats("Sum") = CDbl(stats("Sum")) + numero

        If IsEmpty(stats("Min")) Then
            stats("Min") = numero
            stats("Max") = numero
        Else
            If numero < CDbl(stats("Min")) Then stats("Min") = numero
            If numero > CDbl(stats("Max")) Then stats("Max") = numero
        End If
    End If
End Sub

Private Function WX_ObterValorAgregado(ByRef stats As Object, ByVal operacao As String) As Variant
    Select Case operacao
        Case "Count"
            WX_ObterValorAgregado = CLng(stats("Count"))
        Case "Sum"
            WX_ObterValorAgregado = CDbl(stats("Sum"))
        Case "Min"
            WX_ObterValorAgregado = stats("Min")
        Case "Max"
            WX_ObterValorAgregado = stats("Max")
        Case "Avg"
            If CLng(stats("Count")) > 0 Then
                WX_ObterValorAgregado = CDbl(stats("Sum")) / CLng(stats("Count"))
            Else
                WX_ObterValorAgregado = Empty
            End If
    End Select
End Function


Private Function WX_TextoParametro(ByVal valor As Variant) As String
    On Error GoTo ValorInvalido

    If IsObject(valor) Then
        If TypeName(valor) = "Range" Then
            WX_TextoParametro = CStr(valor.Cells(1, 1).Text)
        Else
            WX_TextoParametro = CStr(valor)
        End If
    ElseIf IsError(valor) Or IsEmpty(valor) Then
        WX_TextoParametro = vbNullString
    Else
        WX_TextoParametro = CStr(valor)
    End If

    Exit Function

ValorInvalido:
    WX_TextoParametro = vbNullString
End Function

Private Function WX_ResolverTituloItem( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal tituloItem As Variant, _
    ByVal grupos As Variant _
) As String
    Dim texto As String
    Dim coluna As Long

    texto = CStr(tituloItem)

    If texto <> vbNullString Then
        coluna = WX_TentarObterNumeroColuna(plan, texto)

        If coluna > 0 Then
            WX_ResolverTituloItem = CStr(plan.Cells(linha, coluna).Text)
        Else
            WX_ResolverTituloItem = texto
        End If

        Exit Function
    End If

    If Not WX_ArrayVazio(grupos) Then
        coluna = WX_ObterNumeroColuna(plan, CStr(grupos(LBound(grupos))))
        WX_ResolverTituloItem = CStr(plan.Cells(linha, coluna).Text)
    Else
        WX_ResolverTituloItem = vbNullString
    End If
End Function

Private Function WX_ObterChaveGrupo( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal grupos As Variant _
) As String
    Dim i As Long
    Dim coluna As Long
    Dim chave As String

    If WX_ArrayVazio(grupos) Then
        WX_ObterChaveGrupo = "#TOTAL_GERAL#"
        Exit Function
    End If

    For i = LBound(grupos) To UBound(grupos)
        coluna = WX_ObterNumeroColuna(plan, CStr(grupos(i)))
        chave = chave & WX_SEP & CStr(plan.Cells(linha, coluna).Value2)
    Next i

    WX_ObterChaveGrupo = chave
End Function

Private Function WX_RegistroAtendeCriterios( _
    ByRef plan As Worksheet, _
    ByVal linha As Long, _
    ByVal criterios As Variant _
) As Boolean
    Dim i As Long
    Dim qtd As Long
    Dim nomeCampo As String
    Dim valorEsperado As Variant
    Dim valorAtual As Variant
    Dim coluna As Long

    qtd = WX_QuantidadeCriterios(criterios)

    If qtd = 0 Then
        WX_RegistroAtendeCriterios = True
        Exit Function
    End If

    For i = LBound(criterios) To UBound(criterios) Step 2
        nomeCampo = CStr(criterios(i))
        valorEsperado = criterios(i + 1)
        coluna = WX_ObterNumeroColuna(plan, nomeCampo)
        valorAtual = plan.Cells(linha, coluna).Value2

        If CStr(valorAtual) <> CStr(valorEsperado) Then
            WX_RegistroAtendeCriterios = False
            Exit Function
        End If
    Next i

    WX_RegistroAtendeCriterios = True
End Function

Private Sub WX_ValidarCriterios(ByVal criterios As Variant)
    Dim qtd As Long

    qtd = WX_QuantidadeCriterios(criterios)

    If qtd = 0 Then Exit Sub

    If qtd Mod 2 <> 0 Then
        Err.Raise vbObjectError + 4000, "WX_ValidarCriterios", _
            "Critérios devem estar em pares: NomeDaColuna, Valor."
    End If
End Sub

Private Function WX_QuantidadeCriterios(ByVal criterios As Variant) As Long
    On Error GoTo SemCriterios

    WX_QuantidadeCriterios = UBound(criterios) - LBound(criterios) + 1
    Exit Function

SemCriterios:
    WX_QuantidadeCriterios = 0
End Function

Private Function WX_QuebrarLista(ByVal texto As String) As Variant
    Dim bruto As Variant
    Dim itens() As String
    Dim i As Long
    Dim qtd As Long
    Dim item As String

    texto = Trim$(texto)

    If texto = vbNullString Then
        WX_QuebrarLista = Array()
        Exit Function
    End If

    bruto = Split(texto, ",")
    ReDim itens(0 To UBound(bruto))

    For i = LBound(bruto) To UBound(bruto)
        item = Trim$(CStr(bruto(i)))

        If item <> vbNullString Then
            itens(qtd) = item
            qtd = qtd + 1
        End If
    Next i

    If qtd = 0 Then
        WX_QuebrarLista = Array()
    Else
        ReDim Preserve itens(0 To qtd - 1)
        WX_QuebrarLista = itens
    End If
End Function

Private Function WX_UnicoItem(ByVal texto As String, ByVal nomeParametro As String) As String
    Dim itens As Variant

    itens = WX_QuebrarLista(texto)

    If WX_ArrayQuantidade(itens) = 0 Then
        Err.Raise vbObjectError + 4010, "WX_UnicoItem", _
            "Informe um valor para " & nomeParametro & "."
    End If

    If WX_ArrayQuantidade(itens) > 1 Then
        Err.Raise vbObjectError + 4011, "WX_UnicoItem", _
            "Para gráfico, " & nomeParametro & " deve ter apenas um item."
    End If

    WX_UnicoItem = CStr(itens(LBound(itens)))
End Function

Private Function WX_NormalizarOperacao(ByVal operacao As String) As String
    Select Case LCase$(Trim$(operacao))
        Case "sum"
            WX_NormalizarOperacao = "Sum"
        Case "min"
            WX_NormalizarOperacao = "Min"
        Case "max"
            WX_NormalizarOperacao = "Max"
        Case "avg"
            WX_NormalizarOperacao = "Avg"
        Case "count"
            WX_NormalizarOperacao = "Count"
        Case Else
            Err.Raise vbObjectError + 4020, "WX_NormalizarOperacao", _
                "Operação inválida: '" & operacao & "'. Use Sum, Min, Max, Avg ou Count."
    End Select
End Function

Private Function WX_ArrayQuantidade(ByVal arr As Variant) As Long
    On Error GoTo Vazio

    WX_ArrayQuantidade = UBound(arr) - LBound(arr) + 1
    Exit Function

Vazio:
    WX_ArrayQuantidade = 0
End Function

Private Function WX_ArrayVazio(ByVal arr As Variant) As Boolean
    WX_ArrayVazio = (WX_ArrayQuantidade(arr) = 0)
End Function

Private Function WX_ObterNumeroColuna(ByRef plan As Worksheet, ByVal tituloColuna As String) As Long
    Dim coluna As Long

    coluna = 1

    Do While plan.Cells(WX_HEADERS_ROW, coluna).Text <> vbNullString
        If plan.Cells(WX_HEADERS_ROW, coluna).Text = tituloColuna Then
            WX_ObterNumeroColuna = coluna
            Exit Function
        End If

        coluna = coluna + 1
    Loop

    Err.Raise vbObjectError + 4030, "WX_ObterNumeroColuna", _
        "Coluna '" & tituloColuna & "' não encontrada na aba '" & plan.Name & "'."
End Function

Private Function WX_TentarObterNumeroColuna(ByRef plan As Worksheet, ByVal tituloColuna As String) As Long
    On Error GoTo NaoEncontrada

    WX_TentarObterNumeroColuna = WX_ObterNumeroColuna(plan, tituloColuna)
    Exit Function

NaoEncontrada:
    WX_TentarObterNumeroColuna = 0
End Function

Private Function WX_FormatarLabel( _
    ByRef plan As Worksheet, _
    ByVal colunaValor As Long, _
    ByVal valor As Variant, _
    ByVal operacao As String _
) As String
    Dim formato As String
    Dim texto As String

    If IsEmpty(valor) Then
        WX_FormatarLabel = vbNullString
        Exit Function
    End If

    If operacao = "Count" Then
        WX_FormatarLabel = Format$(CLng(valor), "#,##0")
        Exit Function
    End If

    formato = CStr(plan.Columns(colunaValor).NumberFormat)

    If Trim$(formato) <> vbNullString And LCase$(Trim$(formato)) <> "general" Then
        On Error Resume Next
        texto = Application.WorksheetFunction.Text(valor, formato)
        If Err.Number = 0 And texto <> vbNullString Then
            WX_FormatarLabel = texto
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    WX_FormatarLabel = CStr(valor)
End Function

Private Function WX_JsonValorNumerico(ByVal valor As Variant) As String
    If IsEmpty(valor) Or CStr(valor) = vbNullString Then
        WX_JsonValorNumerico = "null"
    ElseIf IsNumeric(valor) Then
        WX_JsonValorNumerico = WX_NumeroInvariant(CDbl(valor))
    Else
        WX_JsonValorNumerico = "null"
    End If
End Function

Private Function WX_NumeroInvariant(ByVal valor As Double) As String
    Dim texto As String

    texto = Trim$(CStr(valor))
    texto = Replace$(texto, ChrW$(160), vbNullString)
    texto = Replace$(texto, " ", vbNullString)

    If Application.ThousandsSeparator <> vbNullString Then
        texto = Replace$(texto, Application.ThousandsSeparator, vbNullString)
    End If

    If Application.DecimalSeparator <> "." Then
        texto = Replace$(texto, Application.DecimalSeparator, ".")
    End If

    WX_NumeroInvariant = texto
End Function

Private Function WX_JsonString(ByVal valor As String) As String
    WX_JsonString = """" & WX_JsonEscape(valor) & """"
End Function

Private Function WX_JsonEscape(ByVal valor As String) As String
    valor = Replace$(valor, Chr$(92), Chr$(92) & Chr$(92))
    valor = Replace$(valor, Chr$(34), Chr$(92) & Chr$(34))
    valor = Replace$(valor, vbCrLf, "\n")
    valor = Replace$(valor, vbCr, "\n")
    valor = Replace$(valor, vbLf, "\n")
    valor = Replace$(valor, vbTab, "\t")

    WX_JsonEscape = valor
End Function
