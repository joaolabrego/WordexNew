Attribute VB_Name = "WordexTotaisGraficos"
Option Explicit

' ============================================================
' Wordex - Totais para gráficos
' ------------------------------------------------------------
' Ideia central:
'   Gráfico = visualização de uma coleção de totais.
'
' Uso típico em célula do ROOT:
'   =ObterTotaisGrafico("Produtos";"Categoria.Nome";"Preço";"Sum";"Preço";"Quantidade";"Sum";"Quantidade")
'
' Parâmetros da função:
'   nomeAbaDados      = aba fonte, por exemplo "Produtos"
'   colunaCategoria   = coluna agrupadora. Aceita campo simples ou campo JSON simples:
'                       "CategoriaId", "ClienteId", "Categoria.Nome", etc.
'   definicoes        = trios repetidos:
'                       colunaValor, operacao, legenda
'
' Operações aceitas:
'   Count, Sum, Min, Max, Avg
'
' Resultado:
'   JSON array com uma linha por grupo e colunas numéricas agregadas.
'
' Exemplo de saída:
'   [
'     {"Categoria":"Eletrônico","PreçoSum":100,"PreçoSum_Label":"Preço"},
'     {"Categoria":"Livros","PreçoSum":50,"PreçoSum_Label":"Preço"}
'   ]
'
' Para gráfico de barras agrupadas no Wordex:
'   ColumnCategory = "Categoria"
'   ColumnValues   = "PreçoSum;QuantidadeSum"
' ============================================================

Public Function ObterTotaisGrafico(ByVal nomeAbaDados As String, _
                                   ByVal colunaCategoria As String, _
                                   ParamArray definicoes() As Variant) As String
    On Error GoTo TrataErro

    Dim qtdArgs As Long
    qtdArgs = QuantidadeParamArray(definicoes)

    If qtdArgs = 0 Or (qtdArgs Mod 3) <> 0 Then
        Err.Raise vbObjectError + 3101, "ObterTotaisGrafico", _
            "Informe trios: ColunaValor, Operacao, Legenda."
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomeAbaDados)

    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If ultimaLinha < 2 Then
        ObterTotaisGrafico = "[]"
        Exit Function
    End If

    Dim nomeCampoCategoriaSaida As String
    nomeCampoCategoriaSaida = NomeCampoCategoriaSaida(colunaCategoria)

    Dim grupos As Object
    Set grupos = CreateObject("Scripting.Dictionary")

    Dim ordemGrupos As Collection
    Set ordemGrupos = New Collection

    Dim linha As Long
    For linha = 2 To ultimaLinha
        Dim categoriaValor As String
        categoriaValor = CStr(ObterValorCampo(ws, linha, colunaCategoria))

        Dim chaveGrupo As String
        chaveGrupo = categoriaValor

        If Not grupos.Exists(chaveGrupo) Then
            Dim registro As Object
            Set registro = CreateObject("Scripting.Dictionary")
            registro(nomeCampoCategoriaSaida) = categoriaValor
            grupos.Add chaveGrupo, registro
            ordemGrupos.Add chaveGrupo
        End If

        Dim reg As Object
        Set reg = grupos(chaveGrupo)

        Dim i As Long
        For i = LBound(definicoes) To UBound(definicoes) Step 3
            Dim colunaValor As String
            Dim operacao As String
            Dim legenda As String
            Dim nomeSaida As String

            colunaValor = CStr(definicoes(i))
            operacao = NormalizarOperacao(CStr(definicoes(i + 1)))
            legenda = CStr(definicoes(i + 2))
            nomeSaida = NomeCampoValorSaida(colunaValor, operacao)

            Dim valorBruto As Variant
            valorBruto = ObterValorCampo(ws, linha, colunaValor)

            Dim valorNumerico As Double
            Dim temNumero As Boolean
            temNumero = TentarConverterNumero(valorBruto, valorNumerico)

            AtualizarAgregado reg, nomeSaida, operacao, valorBruto, valorNumerico, temNumero
            reg(nomeSaida & "_Label") = legenda
        Next i
    Next linha

    ObterTotaisGrafico = MontarJsonTotais(grupos, ordemGrupos, nomeCampoCategoriaSaida, definicoes)
    Exit Function

TrataErro:
    ObterTotaisGrafico = "{""Erro"":" & JsonString(Err.Description) & "}"
End Function

' Exemplo pronto para a planilha atual.
' Cria/atualiza a propriedade ROOT.TotaisProdutosPorCategoria em ROOT!linha 2.
Public Sub GerarTotaisProdutosPorCategoria()
    Dim json As String

    json = ObterTotaisGrafico( _
        "Produtos", _
        "Categoria.Nome", _
        "Preço", "Sum", "Preço", _
        "Quantidade", "Sum", "Quantidade")

    GravarValorNoRoot "TotaisProdutosPorCategoria", json

    MsgBox "TotaisProdutosPorCategoria atualizado no ROOT.", vbInformation
End Sub

' ----------------------------------------------------------------
' Montagem do JSON final
' ----------------------------------------------------------------

Private Function MontarJsonTotais(ByVal grupos As Object, _
                                  ByVal ordemGrupos As Collection, _
                                  ByVal nomeCampoCategoriaSaida As String, _
                                  ByRef definicoes() As Variant) As String
    Dim partes() As String
    ReDim partes(1 To ordemGrupos.Count)

    Dim g As Long
    For g = 1 To ordemGrupos.Count
        Dim chaveGrupo As String
        chaveGrupo = CStr(ordemGrupos(g))

        Dim reg As Object
        Set reg = grupos(chaveGrupo)

        Dim campos As Collection
        Set campos = New Collection

        campos.Add JsonString(nomeCampoCategoriaSaida) & ":" & JsonString(CStr(reg(nomeCampoCategoriaSaida)))

        Dim i As Long
        For i = LBound(definicoes) To UBound(definicoes) Step 3
            Dim colunaValor As String
            Dim operacao As String
            Dim legenda As String
            Dim nomeSaida As String
            Dim valorFinal As Variant

            colunaValor = CStr(definicoes(i))
            operacao = NormalizarOperacao(CStr(definicoes(i + 1)))
            legenda = CStr(definicoes(i + 2))
            nomeSaida = NomeCampoValorSaida(colunaValor, operacao)
            valorFinal = ValorAgregadoFinal(reg, nomeSaida, operacao)

            campos.Add JsonString(nomeSaida) & ":" & JsonNumeroOuNull(valorFinal)
            campos.Add JsonString(nomeSaida & "_Label") & ":" & JsonString(legenda)
        Next i

        partes(g) = "{" & JoinCollection(campos, ",") & "}"
    Next g

    MontarJsonTotais = "[" & Join(partes, ",") & "]"
End Function

Private Sub AtualizarAgregado(ByVal reg As Object, _
                              ByVal nomeSaida As String, _
                              ByVal operacao As String, _
                              ByVal valorBruto As Variant, _
                              ByVal valorNumerico As Double, _
                              ByVal temNumero As Boolean)
    Dim kCount As String
    Dim kSum As String
    Dim kMin As String
    Dim kMax As String

    kCount = "__" & nomeSaida & "_Count"
    kSum = "__" & nomeSaida & "_Sum"
    kMin = "__" & nomeSaida & "_Min"
    kMax = "__" & nomeSaida & "_Max"

    If Not reg.Exists(kCount) Then reg(kCount) = 0#
    If Not reg.Exists(kSum) Then reg(kSum) = 0#

    Select Case operacao
        Case "Count"
            If Len(Trim$(CStr(valorBruto))) > 0 Then
                reg(kCount) = CDbl(reg(kCount)) + 1#
            End If

        Case "Sum", "Avg", "Min", "Max"
            If temNumero Then
                reg(kCount) = CDbl(reg(kCount)) + 1#
                reg(kSum) = CDbl(reg(kSum)) + valorNumerico

                If Not reg.Exists(kMin) Then
                    reg(kMin) = valorNumerico
                ElseIf valorNumerico < CDbl(reg(kMin)) Then
                    reg(kMin) = valorNumerico
                End If

                If Not reg.Exists(kMax) Then
                    reg(kMax) = valorNumerico
                ElseIf valorNumerico > CDbl(reg(kMax)) Then
                    reg(kMax) = valorNumerico
                End If
            End If
    End Select
End Sub

Private Function ValorAgregadoFinal(ByVal reg As Object, _
                                    ByVal nomeSaida As String, _
                                    ByVal operacao As String) As Variant
    Dim kCount As String
    Dim kSum As String
    Dim kMin As String
    Dim kMax As String

    kCount = "__" & nomeSaida & "_Count"
    kSum = "__" & nomeSaida & "_Sum"
    kMin = "__" & nomeSaida & "_Min"
    kMax = "__" & nomeSaida & "_Max"

    Dim qtd As Double
    If reg.Exists(kCount) Then qtd = CDbl(reg(kCount)) Else qtd = 0#

    Select Case operacao
        Case "Count"
            ValorAgregadoFinal = qtd

        Case "Sum"
            If reg.Exists(kSum) Then ValorAgregadoFinal = CDbl(reg(kSum)) Else ValorAgregadoFinal = 0#

        Case "Min"
            If qtd = 0 Or Not reg.Exists(kMin) Then ValorAgregadoFinal = Null Else ValorAgregadoFinal = CDbl(reg(kMin))

        Case "Max"
            If qtd = 0 Or Not reg.Exists(kMax) Then ValorAgregadoFinal = Null Else ValorAgregadoFinal = CDbl(reg(kMax))

        Case "Avg"
            If qtd = 0 Or Not reg.Exists(kSum) Then
                ValorAgregadoFinal = Null
            Else
                ValorAgregadoFinal = CDbl(reg(kSum)) / qtd
            End If
    End Select
End Function

' ----------------------------------------------------------------
' Leitura de campos simples e campos JSON simples: Categoria.Nome
' ----------------------------------------------------------------

Private Function ObterValorCampo(ByVal ws As Worksheet, ByVal linha As Long, ByVal campo As String) As Variant
    Dim partes() As String
    partes = Split(CStr(campo), ".")

    Dim colunaBase As Long
    colunaBase = ObterNumeroColunaLocal(ws, partes(0))

    Dim valor As Variant
    valor = ws.Cells(linha, colunaBase).Value2

    If UBound(partes) >= 1 Then
        Dim i As Long
        Dim jsonAtual As String
        jsonAtual = CStr(valor)

        For i = 1 To UBound(partes)
            jsonAtual = ExtrairValorJsonSimples(jsonAtual, partes(i))
        Next i

        ObterValorCampo = jsonAtual
    Else
        ObterValorCampo = valor
    End If
End Function

Private Function ObterNumeroColunaLocal(ByVal ws As Worksheet, ByVal nomeColuna As String) As Long
    Dim ultimaColuna As Long
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To ultimaColuna
        If Trim$(CStr(ws.Cells(1, c).Value2)) = nomeColuna Then
            ObterNumeroColunaLocal = c
            Exit Function
        End If
    Next c

    Err.Raise vbObjectError + 3102, "ObterNumeroColunaLocal", _
        "Título de coluna '" & nomeColuna & "' não existe na aba '" & ws.Name & "'."
End Function

Private Function ExtrairValorJsonSimples(ByVal json As String, ByVal nomePropriedade As String) As String
    Dim busca As String
    busca = """" & nomePropriedade & """"

    Dim p As Long
    p = InStr(1, json, busca, vbTextCompare)
    If p = 0 Then
        ExtrairValorJsonSimples = ""
        Exit Function
    End If

    p = InStr(p + Len(busca), json, ":")
    If p = 0 Then
        ExtrairValorJsonSimples = ""
        Exit Function
    End If

    p = p + 1
    Do While p <= Len(json) And Mid$(json, p, 1) Like "[ " & vbTab & vbCr & vbLf & "]"
        p = p + 1
    Loop

    If p > Len(json) Then
        ExtrairValorJsonSimples = ""
        Exit Function
    End If

    Dim ch As String
    ch = Mid$(json, p, 1)

    If ch = """" Then
        ExtrairValorJsonSimples = LerStringJson(json, p)
    ElseIf ch = "{" Then
        ExtrairValorJsonSimples = LerBlocoJson(json, p, "{", "}")
    ElseIf ch = "[" Then
        ExtrairValorJsonSimples = LerBlocoJson(json, p, "[", "]")
    Else
        ExtrairValorJsonSimples = LerValorJsonNaoString(json, p)
    End If
End Function

Private Function LerStringJson(ByVal json As String, ByVal posAspasInicial As Long) As String
    Dim i As Long
    Dim resultado As String
    Dim ch As String
    Dim escapando As Boolean

    For i = posAspasInicial + 1 To Len(json)
        ch = Mid$(json, i, 1)

        If escapando Then
            resultado = resultado & ch
            escapando = False
        ElseIf ch = "\" Then
            escapando = True
        ElseIf ch = """" Then
            LerStringJson = resultado
            Exit Function
        Else
            resultado = resultado & ch
        End If
    Next i

    LerStringJson = resultado
End Function

Private Function LerValorJsonNaoString(ByVal json As String, ByVal posInicial As Long) As String
    Dim i As Long
    For i = posInicial To Len(json)
        Dim ch As String
        ch = Mid$(json, i, 1)

        If ch = "," Or ch = "}" Or ch = "]" Then
            LerValorJsonNaoString = Trim$(Mid$(json, posInicial, i - posInicial))
            Exit Function
        End If
    Next i

    LerValorJsonNaoString = Trim$(Mid$(json, posInicial))
End Function

Private Function LerBlocoJson(ByVal json As String, ByVal posInicial As Long, _
                              ByVal abre As String, ByVal fecha As String) As String
    Dim nivel As Long
    Dim i As Long
    Dim dentroString As Boolean
    Dim escapando As Boolean
    Dim ch As String

    For i = posInicial To Len(json)
        ch = Mid$(json, i, 1)

        If dentroString Then
            If escapando Then
                escapando = False
            ElseIf ch = "\" Then
                escapando = True
            ElseIf ch = """" Then
                dentroString = False
            End If
        Else
            If ch = """" Then
                dentroString = True
            ElseIf ch = abre Then
                nivel = nivel + 1
            ElseIf ch = fecha Then
                nivel = nivel - 1
                If nivel = 0 Then
                    LerBlocoJson = Mid$(json, posInicial, i - posInicial + 1)
                    Exit Function
                End If
            End If
        End If
    Next i

    LerBlocoJson = Mid$(json, posInicial)
End Function

' ----------------------------------------------------------------
' Escrita auxiliar no ROOT
' ----------------------------------------------------------------

Private Sub GravarValorNoRoot(ByVal nomeCampo As String, ByVal valor As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ROOT")

    Dim coluna As Long
    coluna = ObterOuCriarColuna(ws, nomeCampo)

    ws.Cells(2, coluna).Value = valor
End Sub

Private Function ObterOuCriarColuna(ByVal ws As Worksheet, ByVal nomeColuna As String) As Long
    Dim ultimaColuna As Long
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To ultimaColuna
        If Trim$(CStr(ws.Cells(1, c).Value2)) = nomeColuna Then
            ObterOuCriarColuna = c
            Exit Function
        End If
    Next c

    ObterOuCriarColuna = ultimaColuna + 1
    ws.Cells(1, ObterOuCriarColuna).Value = nomeColuna
End Function

' ----------------------------------------------------------------
' Utilitários
' ----------------------------------------------------------------

Private Function QuantidadeParamArray(ByRef arr() As Variant) As Long
    On Error GoTo SemItens
    QuantidadeParamArray = UBound(arr) - LBound(arr) + 1
    Exit Function
SemItens:
    QuantidadeParamArray = 0
End Function

Private Function NormalizarOperacao(ByVal operacao As String) As String
    Dim op As String
    op = LCase$(Trim$(operacao))

    Select Case op
        Case "count", "cont", "contagem"
            NormalizarOperacao = "Count"
        Case "sum", "soma"
            NormalizarOperacao = "Sum"
        Case "min", "mín", "minimo", "mínimo"
            NormalizarOperacao = "Min"
        Case "max", "máx", "maximo", "máximo"
            NormalizarOperacao = "Max"
        Case "avg", "media", "média"
            NormalizarOperacao = "Avg"
        Case Else
            Err.Raise vbObjectError + 3103, "NormalizarOperacao", _
                "Operação inválida: " & operacao & ". Use Count, Sum, Min, Max ou Avg."
    End Select
End Function

Private Function NomeCampoCategoriaSaida(ByVal campoCategoria As String) As String
    Dim p As Long
    p = InStr(1, campoCategoria, ".")

    If p > 0 Then
        NomeCampoCategoriaSaida = LimparNomeCampo(Left$(campoCategoria, p - 1))
    Else
        NomeCampoCategoriaSaida = LimparNomeCampo(campoCategoria)
    End If
End Function

Private Function NomeCampoValorSaida(ByVal campoValor As String, ByVal operacao As String) As String
    NomeCampoValorSaida = LimparNomeCampo(UltimaParteCampo(campoValor)) & operacao
End Function

Private Function UltimaParteCampo(ByVal campo As String) As String
    Dim partes() As String
    partes = Split(campo, ".")
    UltimaParteCampo = partes(UBound(partes))
End Function

Private Function LimparNomeCampo(ByVal texto As String) As String
    LimparNomeCampo = Replace(Trim$(texto), " ", "")
End Function

Private Function TentarConverterNumero(ByVal valor As Variant, ByRef numero As Double) As Boolean
    On Error GoTo Falhou

    If IsError(valor) Or IsEmpty(valor) Then
        TentarConverterNumero = False
        Exit Function
    End If

    If VarType(valor) <> vbString Then
        If IsNumeric(valor) Then
            numero = CDbl(valor)
            TentarConverterNumero = True
            Exit Function
        End If
    End If

    Dim s As String
    s = Trim$(CStr(valor))

    If Len(s) = 0 Then
        TentarConverterNumero = False
        Exit Function
    End If

    s = Replace(s, "R$", "")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")

    Dim posVirgula As Long
    Dim posPonto As Long
    posVirgula = InStrRev(s, ",")
    posPonto = InStrRev(s, ".")

    If posVirgula > 0 And posPonto > 0 Then
        If posVirgula > posPonto Then
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        Else
            s = Replace(s, ",", "")
        End If
    ElseIf posVirgula > 0 Then
        s = Replace(s, ".", "")
        s = Replace(s, ",", ".")
    End If

    numero = CDbl(s)
    TentarConverterNumero = True
    Exit Function

Falhou:
    TentarConverterNumero = False
End Function

Private Function JsonNumeroOuNull(ByVal valor As Variant) As String
    If IsNull(valor) Then
        JsonNumeroOuNull = "null"
    Else
        JsonNumeroOuNull = NumeroJson(CDbl(valor))
    End If
End Function

Private Function NumeroJson(ByVal numero As Double) As String
    Dim s As String
    s = CStr(numero)

    If Application.ThousandsSeparator <> "" Then
        s = Replace(s, Application.ThousandsSeparator, "")
    End If

    If Application.DecimalSeparator <> "." Then
        s = Replace(s, Application.DecimalSeparator, ".")
    End If

    NumeroJson = s
End Function

Private Function JsonString(ByVal texto As String) As String
    Dim s As String
    s = CStr(texto)

    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")

    JsonString = """" & s & """"
End Function

Private Function JoinCollection(ByVal col As Collection, ByVal separador As String) As String
    Dim arr() As String
    ReDim arr(1 To col.Count)

    Dim i As Long
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i

    JoinCollection = Join(arr, separador)
End Function
