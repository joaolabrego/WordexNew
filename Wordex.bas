Attribute VB_Name = "Wordex"
Option Explicit

' Crudex primitivo: importar Wordex.bas + WordexConsulta.bas no crudex.xlsm.
' Remover do VBA do workbook: WordexTotais, WordexGraficos, frmTotais.

Public Const K_HEADERS_ROW As Long = 1
Public Const K_DETAILS_ROW As Long = 2
Public Const K_REQUIRED_COL As Long = 1

Public Sub GerarJson()
    Dim json As String
    Application.CalculateFull
    json = ObterRegistros("ROOT")
    SalvarJsonEmDisco json
End Sub

Private Sub SalvarJsonEmDisco(ByVal json As String)
    Dim caminhoJson As String
    
    If ThisWorkbook.Path = vbNullString Then
        Err.Raise vbObjectError + 2000, "SalvarJsonEmDisco", _
            "Salve a planilha antes de gerar o JSON."
    End If
    
    caminhoJson = ThisWorkbook.Path & Application.PathSeparator & _
                  NomeArquivoSemExtensao(ThisWorkbook.Name) & ".json"
    
    SalvarTextoUtf8 caminhoJson, json
    
    MsgBox "JSON salvo em:" & vbCrLf & caminhoJson, vbInformation, "Wordex"
End Sub

Private Function NomeArquivoSemExtensao(ByVal nomeArquivo As String) As String
    Dim pos As Long
    
    pos = InStrRev(nomeArquivo, ".")
    
    If pos > 0 Then
        NomeArquivoSemExtensao = Left$(nomeArquivo, pos - 1)
    Else
        NomeArquivoSemExtensao = nomeArquivo
    End If
End Function

Private Sub SalvarTextoUtf8(ByVal caminhoArquivo As String, ByVal texto As String)
    Dim stream As New ADODB.stream
    
    With stream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText texto
        .SaveToFile caminhoArquivo, 2
        .Close
    End With
End Sub

Public Function ObterRegistro(ByVal nomeAba As String, ParamArray criterios() As Variant) As String
    On Error GoTo Erro

    Dim plan As Worksheet
    Dim numeroLinha As Long
    Dim arrCriterios As Variant

    Set plan = Application.Worksheets(nomeAba)

    arrCriterios = CVar(criterios)
    numeroLinha = ObterNumeroLinha(plan, K_DETAILS_ROW, arrCriterios)

    If numeroLinha > 0 Then
        ObterRegistro = ObterRegistroJSON(plan, numeroLinha)
    Else
        ObterRegistro = ObterRegistroJSON(plan, K_DETAILS_ROW, True)
    End If

    Exit Function

Erro:
    ObterRegistro = "{""Erro"": " & JsonString(Err.Description) & "}"
End Function

Public Function ObterRegistros(ByVal nomeAba As String, ParamArray criterios() As Variant) As String
    On Error GoTo Erro

    Dim plan As Worksheet
    Dim numeroLinha As Long
    Dim json As String
    Dim virgula As String
    Dim arrCriterios As Variant

    Set plan = Application.Worksheets(nomeAba)

    arrCriterios = CVar(criterios)
    ValidarCriterios arrCriterios

    json = "["
    virgula = vbNullString
    numeroLinha = K_DETAILS_ROW

    Do While plan.Cells(numeroLinha, K_REQUIRED_COL).Text <> vbNullString
        If RegistroAtendeCriterios(plan, numeroLinha, arrCriterios) Then
            json = json & virgula & ObterRegistroJSON(plan, numeroLinha)
            virgula = ","
        End If

        numeroLinha = numeroLinha + 1
    Loop

    json = json & "]"

    ObterRegistros = json
    Exit Function

Erro:
    ObterRegistros = "[{""Erro"": " & JsonString(Err.Description) & "}]"
End Function

Private Function ObterNumeroLinha(ByRef plan As Worksheet, ByVal numeroLinhaInicial As Long, ByVal criterios As Variant) As Long
    Dim numeroLinha As Long

    ValidarCriterios criterios

    numeroLinha = numeroLinhaInicial

    Do While plan.Cells(numeroLinha, K_REQUIRED_COL).Text <> vbNullString
        If RegistroAtendeCriterios(plan, numeroLinha, criterios) Then
            ObterNumeroLinha = numeroLinha
            Exit Function
        End If

        numeroLinha = numeroLinha + 1
    Loop

    ObterNumeroLinha = 0
End Function

Private Function RegistroAtendeCriterios(ByRef plan As Worksheet, ByVal numeroLinha As Long, ByVal criterios As Variant) As Boolean
    Dim i As Long
    Dim qtd As Long
    Dim nomeCampo As String
    Dim valorEsperado As Variant
    Dim valorAtual As Variant

    qtd = QuantidadeCriterios(criterios)

    If qtd = 0 Then
        RegistroAtendeCriterios = True
        Exit Function
    End If

    For i = LBound(criterios) To UBound(criterios) Step 2
        nomeCampo = CStr(criterios(i))
        valorEsperado = criterios(i + 1)
        valorAtual = ObterValorColuna(plan, numeroLinha, nomeCampo)

        If valorAtual <> valorEsperado Then
            RegistroAtendeCriterios = False
            Exit Function
        End If
    Next i

    RegistroAtendeCriterios = True
End Function

Private Sub ValidarCriterios(ByVal criterios As Variant)
    Dim qtd As Long

    qtd = QuantidadeCriterios(criterios)

    If qtd = 0 Then Exit Sub

    If qtd Mod 2 <> 0 Then
        Err.Raise vbObjectError + 1000, "ValidarCriterios", _
            "Critérios devem estar em pares: NomeDaColuna, Valor."
    End If
End Sub

Private Function QuantidadeCriterios(ByVal criterios As Variant) As Long
    On Error GoTo SemCriterios

    QuantidadeCriterios = UBound(criterios) - LBound(criterios) + 1
    Exit Function

SemCriterios:
    QuantidadeCriterios = 0
End Function

Private Function ObterNumeroColuna(ByRef plan As Worksheet, ByVal tituloColuna As String) As Long
    Dim numeroColuna As Long
    Dim tituloAtual As String

    numeroColuna = 1

    Do While plan.Cells(K_HEADERS_ROW, numeroColuna).Text <> vbNullString
        tituloAtual = plan.Cells(K_HEADERS_ROW, numeroColuna).Text

        If tituloAtual = tituloColuna Then
            ObterNumeroColuna = numeroColuna
            Exit Function
        End If

        numeroColuna = numeroColuna + 1
    Loop

    Err.Raise vbObjectError + 1001, "ObterNumeroColuna", _
        "Título de coluna '" & tituloColuna & "' não existe na aba '" & plan.Name & "'."
End Function

Private Function ObterValorColuna(ByRef plan As Worksheet, ByVal numeroLinha As Long, ByVal tituloColuna As String, Optional ByVal textual As Boolean = False) As Variant
    Dim numeroColuna As Long
    Dim cell As Range

    numeroColuna = ObterNumeroColuna(plan, tituloColuna)
    Set cell = plan.Cells(numeroLinha, numeroColuna)

    If textual Then
        ObterValorColuna = cell.Text
    Else
        ObterValorColuna = cell.Value2
    End If
End Function

Private Function ObterRegistroJSON(ByRef plan As Worksheet, ByVal numeroLinha As Long, Optional ByVal Vazio As Boolean = False) As String
    Dim numeroColuna As Long
    Dim json As String
    Dim virgula As String
    Dim tituloColuna As String
    Dim valorTexto As String
    Dim cell As Range

    numeroColuna = 1
    json = "{"
    virgula = vbNullString

    Do While plan.Cells(K_HEADERS_ROW, numeroColuna).Text <> vbNullString
        tituloColuna = plan.Cells(K_HEADERS_ROW, numeroColuna).Text
        Set cell = plan.Cells(numeroLinha, numeroColuna)

        If LCase$(Right$(Trim$(tituloColuna), 5)) = "_kind" Then
            GoTo ProximaColuna
        End If

        If Vazio Then
            valorTexto = vbNullString
        Else
            valorTexto = cell.Text
        End If

        If Not Vazio And InStr(valorTexto, """ChartTitle""") > 0 Then
            GoTo ProximaColuna
        End If

        If Not Vazio And InStr(valorTexto, """Erro""") > 0 And Left$(Trim$(valorTexto), 1) = "{" Then
            GoTo ProximaColuna
        End If

        json = json & virgula & JsonString(tituloColuna) & ": "

        If Not Vazio And IsJsonRaw(valorTexto) Then
            If Wordex_DeveEmbalarDatasource(tituloColuna) Then
                json = json & Wordex_EmbalarJsonColuna(tituloColuna, valorTexto, Wordex_ObterKindExplicito(plan, numeroLinha, tituloColuna))
            Else
                json = json & valorTexto
            End If
        Else
            json = json & Wordex_JsonCampoTipado(plan, numeroLinha, tituloColuna, cell, Vazio)
        End If

        virgula = ","
ProximaColuna:
        numeroColuna = numeroColuna + 1
    Loop

    json = json & "}"

    ObterRegistroJSON = json
End Function

Public Function Wordex_JsonCampoTipado( _
    ByRef plan As Worksheet, _
    ByVal numeroLinha As Long, _
    ByVal tituloColuna As String, _
    ByVal cel As Range, _
    ByVal vazio As Boolean _
) As String
    Dim kind As String
    Dim valueJson As String

    kind = Wordex_InferirKind(plan, numeroLinha, tituloColuna, cel)

    If vazio Then
        valueJson = "null"
    Else
        Select Case kind
            Case "number"
                valueJson = Wordex_JsonValorSemFormatacao(cel)
            Case "boolean"
                valueJson = LCase$(CStr(CBool(cel.Value)))
            Case "datetime"
                valueJson = """" & Format$(CDate(cel.Value), "yyyy-mm-dd") & """"
            Case Else
                valueJson = JsonString(Trim$(cel.Text))
        End Select
    End If

    Wordex_JsonCampoTipado = "{""Kind"": " & JsonString(kind) & ", ""Value"": " & valueJson & "}"
End Function

Public Function Wordex_InferirKind( _
    ByRef plan As Worksheet, _
    ByVal numeroLinha As Long, _
    ByVal tituloColuna As String, _
    ByVal cel As Range _
) As String
    Dim nome As String
    Dim kindExplicito As String
    Dim valor As Variant
    Dim valorTexto As String

    kindExplicito = Wordex_ObterKindExplicito(plan, numeroLinha, tituloColuna)

    If kindExplicito <> vbNullString Then
        If LCase$(kindExplicito) = "text" Then kindExplicito = "string"
        Wordex_InferirKind = LCase$(kindExplicito)
        Exit Function
    End If

    nome = LCase$(Trim$(tituloColuna))
    valor = cel.Value2
    valorTexto = Trim$(cel.Text)

    If Right$(nome, 6) = "_value" Then
        Wordex_InferirKind = "number"
        Exit Function
    End If

    If Right$(nome, 6) = "_label" Then
        Wordex_InferirKind = "string"
        Exit Function
    End If

    If Wordex_NomeCampoPareceImagem(nome) And Wordex_PareceUrlImagem(valorTexto) Then
        Wordex_InferirKind = "image"
        Exit Function
    End If

    If Wordex_PareceUrlImagem(valorTexto) Then
        Wordex_InferirKind = "image"
        Exit Function
    End If

    Select Case VarType(valor)
        Case vbBoolean
            Wordex_InferirKind = "boolean"

        Case vbDate
            Wordex_InferirKind = "datetime"

        Case Else
            Wordex_InferirKind = "string"
    End Select
End Function

Private Function Wordex_ObterKindExplicito( _
    ByRef plan As Worksheet, _
    ByVal numeroLinha As Long, _
    ByVal tituloColuna As String _
) As String
    Dim numeroColuna As Long
    Dim tituloKind As String

    On Error GoTo SemKind

    tituloKind = Trim$(tituloColuna) & "_Kind"
    numeroColuna = ObterNumeroColuna(plan, tituloKind)
    Wordex_ObterKindExplicito = Trim$(CStr(plan.Cells(numeroLinha, numeroColuna).Value2))
    Exit Function

SemKind:
    Wordex_ObterKindExplicito = vbNullString
End Function

Private Function Wordex_NomeCampoPareceImagem(ByVal nomeCampo As String) As Boolean
    Wordex_NomeCampoPareceImagem = _
        InStr(nomeCampo, "logo") > 0 Or _
        InStr(nomeCampo, "imagem") > 0 Or _
        InStr(nomeCampo, "image") > 0 Or _
        InStr(nomeCampo, "foto") > 0 Or _
        InStr(nomeCampo, "picture") > 0 Or _
        InStr(nomeCampo, "avatar") > 0 Or _
        InStr(nomeCampo, "url") > 0
End Function

Private Function Wordex_PareceUrlImagem(ByVal valorTexto As String) As Boolean
    Dim texto As String

    texto = LCase$(Trim$(valorTexto))

    If texto = vbNullString Then Exit Function

    Wordex_PareceUrlImagem = _
        Left$(texto, 7) = "http://" Or _
        Left$(texto, 8) = "https://" Or _
        Left$(texto, 5) = "data:" Or _
        InStr(texto, ".png") > 0 Or _
        InStr(texto, ".jpg") > 0 Or _
        InStr(texto, ".jpeg") > 0 Or _
        InStr(texto, ".gif") > 0 Or _
        InStr(texto, ".webp") > 0 Or _
        InStr(texto, ".svg") > 0 Or _
        InStr(texto, ".bmp") > 0
End Function

Private Function Wordex_DeveEmbalarDatasource(ByVal tituloColuna As String) As Boolean
    Dim nome As String

    nome = LCase$(Trim$(tituloColuna))

    Wordex_DeveEmbalarDatasource = _
        (nome = "clientes") Or _
        (nome = "produtos") Or _
        (Left$(nome, 6) = "totais") Or _
        (Right$(nome, 7) = "grafico")
End Function

Private Function Wordex_InferirKindDatasource(ByVal tituloColuna As String) As String
    Dim nome As String

    nome = LCase$(Trim$(tituloColuna))

    ' Kind do wrapper: collection | total | histogram.
    ' FKs não passam por aqui — ObterRegistroJSON só embala quando Wordex_DeveEmbalarDatasource = True.
    If Right$(nome, 7) = "grafico" Then
        Wordex_InferirKindDatasource = "histogram"
    ElseIf Left$(nome, 6) = "totais" Then
        Wordex_InferirKindDatasource = "total"
    Else
        Wordex_InferirKindDatasource = "collection"
    End If
End Function

Private Function Wordex_EmbalarJsonColuna( _
    ByVal tituloColuna As String, _
    ByVal valorTexto As String, _
    ByVal kindExplicito As String _
) As String
    Dim kind As String
    Dim conteudo As String

    kind = kindExplicito
    If kind = vbNullString Then kind = Wordex_InferirKindDatasource(tituloColuna)

    conteudo = Trim$(valorTexto)

    If Wordex_JaEmbaladoComoDatasource(conteudo) Then
        Wordex_EmbalarJsonColuna = conteudo
        Exit Function
    End If

    If Left$(conteudo, 1) = "[" Then
        Wordex_EmbalarJsonColuna = "{""Kind"": """ & kind & """, ""Items"": " & conteudo & "}"
    ElseIf Left$(conteudo, 1) = "{" Then
        Wordex_EmbalarJsonColuna = "{""Kind"": """ & kind & """, ""Items"": [" & conteudo & "]}"
    Else
        Wordex_EmbalarJsonColuna = conteudo
    End If
End Function

Private Function Wordex_JaEmbaladoComoDatasource(ByVal conteudo As String) As Boolean
    Dim texto As String

    texto = LCase$(Trim$(conteudo))

    If Left$(texto, 1) <> "{" Then Exit Function

    Wordex_JaEmbaladoComoDatasource = _
        InStr(texto, """kind""") > 0 And _
        InStr(texto, """items""") > 0
End Function

Public Function Wordex_JsonValorSemFormatacao(ByVal cel As Range) As String
    Dim valor As Variant
    Dim valorTexto As String
    Dim numeroJson As String

    valor = cel.Value

    Select Case VarType(valor)
        Case vbEmpty, vbNull
            Wordex_JsonValorSemFormatacao = "null"

        Case vbBoolean
            Wordex_JsonValorSemFormatacao = LCase$(CStr(CBool(valor)))

        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            Wordex_JsonValorSemFormatacao = Wordex_JsonNumeroInvariant(valor)

        Case vbDate
            Wordex_JsonValorSemFormatacao = """" & Format$(CDate(valor), "yyyy-mm-dd") & """"

        Case vbString
            valorTexto = Trim$(CStr(valor))

            If valorTexto = vbNullString Then
                Wordex_JsonValorSemFormatacao = "null"
            ElseIf Wordex_JsonNormalizarNumeroTexto(valorTexto, numeroJson) Then
                Wordex_JsonValorSemFormatacao = numeroJson
            Else
                Wordex_JsonValorSemFormatacao = """" & Wordex_JsonEscapeTexto(valorTexto) & """"
            End If

        Case vbError
            Wordex_JsonValorSemFormatacao = "null"

        Case Else
            valorTexto = Trim$(CStr(valor))

            If valorTexto = vbNullString Then
                Wordex_JsonValorSemFormatacao = "null"
            ElseIf Wordex_JsonNormalizarNumeroTexto(valorTexto, numeroJson) Then
                Wordex_JsonValorSemFormatacao = numeroJson
            Else
                Wordex_JsonValorSemFormatacao = """" & Wordex_JsonEscapeTexto(valorTexto) & """"
            End If
    End Select
End Function

Private Function Wordex_JsonNumeroInvariant(ByVal valor As Variant) As String
    Dim texto As String
    Dim sepMilhar As String
    Dim sepDecimal As String

    texto = CStr(valor)

    sepMilhar = Application.International(xlThousandsSeparator)
    sepDecimal = Application.International(xlDecimalSeparator)

    If sepMilhar <> vbNullString Then
        texto = Replace$(texto, sepMilhar, vbNullString)
    End If

    If sepDecimal <> "." Then
        texto = Replace$(texto, sepDecimal, ".")
    End If

    Wordex_JsonNumeroInvariant = texto
End Function

Private Function Wordex_JsonNormalizarNumeroTexto(ByVal texto As String, ByRef numeroJson As String) As Boolean
    Dim s As String
    Dim posVirgula As Long
    Dim posPonto As Long

    s = Trim$(texto)

    If s = vbNullString Then Exit Function

    If Left$(s, 1) = "+" Then
        s = Mid$(s, 2)
    End If

    posVirgula = InStrRev(s, ",")
    posPonto = InStrRev(s, ".")

    If posVirgula > 0 And posPonto > 0 Then
        If posVirgula > posPonto Then
            s = Replace$(s, ".", vbNullString)
            s = Replace$(s, ",", ".")
        Else
            s = Replace$(s, ",", vbNullString)
        End If
    ElseIf posVirgula > 0 Then
        s = Replace$(s, ".", vbNullString)
        s = Replace$(s, ",", ".")
    ElseIf posPonto > 0 Then
        s = Replace$(s, ",", vbNullString)
    End If

    s = Wordex_JsonAjustarZerosEsquerdaNumero(s)

    If Wordex_JsonNumeroValido(s) Then
        numeroJson = s
        Wordex_JsonNormalizarNumeroTexto = True
    End If
End Function

Private Function Wordex_JsonAjustarZerosEsquerdaNumero(ByVal numero As String) As String
    Dim sinal As String
    Dim parteInteira As String
    Dim parteDecimal As String
    Dim posPonto As Long

    If Left$(numero, 1) = "-" Then
        sinal = "-"
        numero = Mid$(numero, 2)
    End If

    posPonto = InStr(1, numero, ".", vbBinaryCompare)

    If posPonto > 0 Then
        parteInteira = Left$(numero, posPonto - 1)
        parteDecimal = Mid$(numero, posPonto)
    Else
        parteInteira = numero
        parteDecimal = vbNullString
    End If

    Do While Len(parteInteira) > 1 And Left$(parteInteira, 1) = "0"
        parteInteira = Mid$(parteInteira, 2)
    Loop

    If parteInteira = vbNullString Then
        parteInteira = "0"
    End If

    Wordex_JsonAjustarZerosEsquerdaNumero = sinal & parteInteira & parteDecimal
End Function

Private Function Wordex_JsonNumeroValido(ByVal numero As String) As Boolean
    Dim i As Long
    Dim inicio As Long
    Dim ch As String
    Dim temDigito As Boolean
    Dim temPonto As Boolean

    If numero = vbNullString Then Exit Function

    If Left$(numero, 1) = "-" Then
        If Len(numero) = 1 Then Exit Function
        inicio = 2
    Else
        inicio = 1
    End If

    For i = inicio To Len(numero)
        ch = Mid$(numero, i, 1)

        Select Case ch
            Case "0" To "9"
                temDigito = True
            Case "."
                If temPonto Then Exit Function
                temPonto = True
            Case Else
                Exit Function
        End Select
    Next i

    If Not temDigito Then Exit Function
    If Right$(numero, 1) = "." Then Exit Function
    If Mid$(numero, inicio, 1) = "." Then Exit Function

    Wordex_JsonNumeroValido = True
End Function

Private Function Wordex_JsonEscapeTexto(ByVal texto As String) As String
    texto = Replace$(texto, "\", "\\")
    texto = Replace$(texto, Chr$(34), Chr$(92) & Chr$(34))
    texto = Replace$(texto, vbCrLf, "")
    texto = Replace$(texto, vbCr, "")
    texto = Replace$(texto, vbLf, "")

    Wordex_JsonEscapeTexto = texto
End Function

Private Function IsJsonRaw(ByVal valor As String) As Boolean
    valor = Trim$(valor)

    If Len(valor) < 2 Then
        IsJsonRaw = False
        Exit Function
    End If

    IsJsonRaw = _
        (Left$(valor, 1) = "{" And Right$(valor, 1) = "}") Or _
        (Left$(valor, 1) = "[" And Right$(valor, 1) = "]")
End Function

Private Function JsonString(ByVal valor As String) As String
    JsonString = """" & JsonEscape(valor) & """"
End Function

Private Function JsonEscape(ByVal valor As String) As String
    valor = Replace(valor, "\", "\\")
    valor = Replace(valor, """", "\""")
    valor = Replace(valor, vbCrLf, "\n")
    valor = Replace(valor, vbCr, "\n")
    valor = Replace(valor, vbLf, "\n")
    valor = Replace(valor, vbTab, "\t")

    JsonEscape = valor
End Function
