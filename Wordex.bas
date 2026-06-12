Attribute VB_Name = "Wordex"
Option Explicit

' Wordex ponte: importar Wordex.bas + WordexConsulta.bas no wordex.xlsm.
' Layout de cada aba de dados:
'   Linha 1 (K_HEADERS_ROW) = nomes dos campos
'   Linha 2 (K_TYPES_ROW)   = tipos: string, number, date, datetime, boolean,
'                               collection, object, totals, graph, image, ...
'   Linha 3+ (K_DETAILS_ROW) = registros

Public Const K_HEADERS_ROW As Long = 1
Public Const K_TYPES_ROW As Long = 2
Public Const K_DETAILS_ROW As Long = 3
Public Const K_REQUIRED_COL As Long = 1

Public Sub GerarJson()
    Dim json As String
    Application.CalculateFull
    json = ObterRegistros("ROOT")
    SalvarJsonEmDisco json
End Sub

Public Sub LerImagem()
    Dim dlg As FileDialog
    Dim caminho As String

    Set dlg = Application.FileDialog(msoFileDialogFilePicker)

    With dlg
        .Title = "Selecionar imagem"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Imagens", "*.png;*.jpg;*.jpeg;*.gif;*.webp;*.bmp;*.tif;*.tiff"
        .Filters.Add "Todos os arquivos", "*.*"

        If .Show <> -1 Then Exit Sub

        caminho = .SelectedItems(1)
    End With

    ActiveCell.Value = caminho

    MsgBox "Caminho da imagem colocado na célula ativa:" & vbCrLf & caminho, vbInformation, "Wordex"
End Sub

Private Sub SalvarJsonEmDisco(ByVal json As String)
    Dim caminhoJson As Variant
    Dim nomePadrao As String
    Dim caminhoInicial As String

    nomePadrao = NomeArquivoSemExtensao(ThisWorkbook.Name) & ".json"

    If ThisWorkbook.Path <> vbNullString Then
        caminhoInicial = ThisWorkbook.Path & Application.PathSeparator & nomePadrao
    Else
        caminhoInicial = nomePadrao
    End If

    caminhoJson = Application.GetSaveAsFilename( _
        InitialFileName:=caminhoInicial, _
        FileFilter:="JSON (*.json), *.json", _
        Title:="Salvar JSON")

    If caminhoJson = False Then Exit Sub

    caminhoJson = Wordex_CaminhoComExtensaoJson(CStr(caminhoJson))

    On Error GoTo Erro
    SalvarTextoUtf8 CStr(caminhoJson), json

    MsgBox "JSON salvo em:" & vbCrLf & caminhoJson, vbInformation, "Wordex"
    Exit Sub

Erro:
    MsgBox "Não foi possível salvar o JSON:" & vbCrLf & Err.Description, vbCritical, "Wordex"
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

Private Function Wordex_CaminhoComExtensaoJson(ByVal caminhoCompleto As String) As String
    Dim pasta As String
    Dim nomeArquivo As String
    Dim pos As Long

    pos = InStrRev(caminhoCompleto, Application.PathSeparator)

    If pos > 0 Then
        pasta = Left$(caminhoCompleto, pos)
        nomeArquivo = Mid$(caminhoCompleto, pos + 1)
    Else
        pasta = vbNullString
        nomeArquivo = caminhoCompleto
    End If

    Wordex_CaminhoComExtensaoJson = pasta & NomeArquivoSemExtensao(nomeArquivo) & ".json"
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
    Dim kindJson As String

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
            kindJson = Wordex_ObterKindColuna(plan, tituloColuna)

            If Wordex_DeveEmbalarJsonRaw(kindJson) Then
                json = json & Wordex_EmbalarJsonColuna(tituloColuna, valorTexto, kindJson)
            ElseIf kindJson = "object" Then
                json = json & Wordex_EmbalarJsonObject(valorTexto)
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

    If vazio Then
        kind = Wordex_ObterKindColuna(plan, tituloColuna)
        Wordex_JsonCampoTipado = "{""Kind"": " & JsonString(kind) & ", ""Value"": null}"
    Else
        Wordex_JsonCampoTipado = Wordex_JsonCampoValor(plan, tituloColuna, cel.Value2, cel)
    End If
End Function

' Campo {Kind, Value} usando o tipo da linha 2 da aba informada.
Public Function Wordex_JsonCampoValor( _
    ByRef plan As Worksheet, _
    ByVal tituloColuna As String, _
    ByVal valor As Variant, _
    Optional ByVal cel As Range = Nothing _
) As String
    Dim kind As String
    Dim valueJson As String
    Dim valorTexto As String

    If Right$(LCase$(Trim$(tituloColuna)), 6) = "_label" Then
        kind = "string"
    Else
        kind = Wordex_ObterKindColuna(plan, tituloColuna)
    End If

    If cel Is Nothing Then
        valorTexto = Trim$(CStr(valor))
    Else
        valorTexto = Trim$(cel.Text)
    End If

    If kind = "image" Or Wordex_PareceReferenciaImagem(valorTexto) Then
        kind = "image"
        valueJson = Wordex_JsonValorImagem(valorTexto)
    ElseIf cel Is Nothing Then
        valueJson = Wordex_JsonValorPorKind(kind, valor, Nothing)
    Else
        valueJson = Wordex_JsonValorPorKind(kind, valor, cel)
    End If

    Wordex_JsonCampoValor = "{""Kind"": " & JsonString(kind) & ", ""Value"": " & valueJson & "}"
End Function

Public Function Wordex_ObterKindColuna( _
    ByRef plan As Worksheet, _
    ByVal tituloColuna As String _
) As String
    Dim numeroColuna As Long
    Dim kind As String

    numeroColuna = ObterNumeroColuna(plan, tituloColuna)
    kind = Trim$(plan.Cells(K_TYPES_ROW, numeroColuna).Text)

    If kind = vbNullString Then
        Err.Raise vbObjectError + 1002, "Wordex_ObterKindColuna", _
            "Defina o tipo na linha " & K_TYPES_ROW & " da aba '" & plan.Name & _
            "' para a coluna '" & tituloColuna & "'."
    End If

    Wordex_ObterKindColuna = Wordex_NormalizarKind(kind)
End Function

Private Function Wordex_NormalizarKind(ByVal kind As String) As String
    kind = LCase$(Trim$(kind))

    ' Vocabulário da planilha (linha 2) → Kind do JSON consumido pelo Wordex.
    Select Case kind
        Case "text"
            Wordex_NormalizarKind = "string"
        Case "totals", "total"
            Wordex_NormalizarKind = "totals"
        Case "graph", "histogram"
            Wordex_NormalizarKind = "graph"
        Case Else
            Wordex_NormalizarKind = kind
    End Select
End Function

Private Function Wordex_ValorVazio(ByVal valor As Variant, Optional ByVal cel As Range = Nothing) As Boolean
    If IsEmpty(valor) Or IsNull(valor) Then
        Wordex_ValorVazio = True
        Exit Function
    End If

    If Not cel Is Nothing Then
        Wordex_ValorVazio = Trim$(cel.Text) = vbNullString
        Exit Function
    End If

    If VarType(valor) = vbString Then
        Wordex_ValorVazio = Trim$(CStr(valor)) = vbNullString
    End If
End Function

Private Function Wordex_JsonValorPorKind( _
    ByVal kind As String, _
    ByVal valor As Variant, _
    Optional ByVal cel As Range = Nothing _
) As String
    Dim valorTexto As String

    If Wordex_ValorVazio(valor, cel) Then
        Wordex_JsonValorPorKind = "null"
        Exit Function
    End If

    Select Case kind
        Case "number"
            If Not cel Is Nothing Then
                Wordex_JsonValorPorKind = Wordex_JsonValorSemFormatacao(cel)
            ElseIf IsNumeric(valor) Then
                Wordex_JsonValorPorKind = Wordex_JsonNumeroInvariant(valor)
            Else
                Wordex_JsonValorPorKind = "null"
            End If

        Case "boolean"
            If Not cel Is Nothing Then
                Wordex_JsonValorPorKind = LCase$(CStr(CBool(cel.Value)))
            Else
                Wordex_JsonValorPorKind = LCase$(CStr(CBool(valor)))
            End If

        Case "datetime", "date"
            If Not cel Is Nothing Then
                Wordex_JsonValorPorKind = """" & Format$(CDate(cel.Value), "yyyy-mm-dd") & """"
            Else
                Wordex_JsonValorPorKind = """" & Format$(CDate(valor), "yyyy-mm-dd") & """"
            End If

        Case "object"
            If Not cel Is Nothing Then
                valorTexto = Trim$(cel.Text)

                If IsJsonRaw(valorTexto) Then
                    Wordex_JsonValorPorKind = valorTexto
                Else
                    Wordex_JsonValorPorKind = JsonString(valorTexto)
                End If
            ElseIf VarType(valor) = vbString And IsJsonRaw(CStr(valor)) Then
                Wordex_JsonValorPorKind = Trim$(CStr(valor))
            Else
                Wordex_JsonValorPorKind = JsonString(CStr(valor))
            End If

        Case "image"
            If Not cel Is Nothing Then
                Wordex_JsonValorPorKind = Wordex_JsonValorImagem(Trim$(cel.Text))
            Else
                Wordex_JsonValorPorKind = Wordex_JsonValorImagem(Trim$(CStr(valor)))
            End If

        Case Else
            If Not cel Is Nothing Then
                Wordex_JsonValorPorKind = JsonString(Trim$(cel.Text))
            Else
                Wordex_JsonValorPorKind = JsonString(CStr(valor))
            End If
    End Select
End Function

Private Function Wordex_DeveEmbalarJsonRaw(ByVal kind As String) As Boolean
    Select Case kind
        Case "collection", "totals", "graph"
            Wordex_DeveEmbalarJsonRaw = True
        Case Else
            Wordex_DeveEmbalarJsonRaw = False
    End Select
End Function

Private Function Wordex_EmbalarJsonObject(ByVal valorTexto As String) As String
    Dim conteudo As String

    conteudo = Trim$(valorTexto)

    If Wordex_JaEmbaladoComoCampoTipado(conteudo) Then
        Wordex_EmbalarJsonObject = conteudo
        Exit Function
    End If

    Wordex_EmbalarJsonObject = "{""Kind"": ""object"", ""Value"": " & conteudo & "}"
End Function

Private Function Wordex_JaEmbaladoComoCampoTipado(ByVal conteudo As String) As Boolean
    Dim texto As String
    Dim posValue As Long

    texto = LCase$(Trim$(conteudo))

    If Left$(texto, 1) <> "{" Then Exit Function

    ' Wrapper raiz {Kind, Value}. Objetos aninhados (ObterRegistro) também
    ' contêm "Kind"/"Value" nos campos — não podem ser confundidos com wrapper.
    If Left$(texto, 8) <> "{""kind"":" Then Exit Function

    posValue = InStr(1, texto, """value"":", vbBinaryCompare)

    Wordex_JaEmbaladoComoCampoTipado = (posValue > 8)
End Function

Private Function Wordex_EmbalarJsonColuna( _
    ByVal tituloColuna As String, _
    ByVal valorTexto As String, _
    ByVal kind As String _
) As String
    Dim conteudo As String

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

Private Function Wordex_JsonValorImagem(ByVal valorTexto As String) As String
    Dim caminho As String
    Dim dataUri As String

    valorTexto = Trim$(valorTexto)

    If valorTexto = vbNullString Then
        Wordex_JsonValorImagem = "null"
        Exit Function
    End If

    If Left$(LCase$(valorTexto), 5) = "data:" Then
        Wordex_JsonValorImagem = JsonString(valorTexto)
        Exit Function
    End If

    If Wordex_PareceUrlHttp(valorTexto) Then
        Wordex_JsonValorImagem = JsonString(valorTexto)
        Exit Function
    End If

    caminho = Wordex_ResolverCaminhoImagem(valorTexto)

    On Error GoTo Falha
    dataUri = Wordex_ImageFileToDataUri(caminho)
    Wordex_JsonValorImagem = JsonString(dataUri)
    Exit Function

Falha:
    Wordex_JsonValorImagem = JsonString(valorTexto)
End Function

Private Function Wordex_PareceReferenciaImagem(ByVal valorTexto As String) As Boolean
    Dim texto As String
    Dim ext As String
    Dim pos As Long

    texto = Trim$(valorTexto)

    If texto = vbNullString Then Exit Function

    If Left$(LCase$(texto), 5) = "data:" Then
        Wordex_PareceReferenciaImagem = True
        Exit Function
    End If

    If Wordex_PareceUrlHttp(texto) Then
        Wordex_PareceReferenciaImagem = Wordex_PareceExtensaoImagem(texto)
        Exit Function
    End If

    pos = InStrRev(texto, ".")

    If pos = 0 Then Exit Function

    ext = LCase$(Mid$(texto, pos + 1))
    Wordex_PareceReferenciaImagem = Wordex_ExtensaoImagemValida(ext)
End Function

Private Function Wordex_PareceExtensaoImagem(ByVal caminho As String) As Boolean
    Dim pos As Long
    Dim ext As String

    pos = InStrRev(caminho, ".")

    If pos = 0 Then Exit Function

    ext = LCase$(Mid$(caminho, pos + 1))
    ext = Left$(ext, InStr(ext & "?", "?") - 1)
    ext = Left$(ext, InStr(ext & "#", "#") - 1)

    Wordex_PareceExtensaoImagem = Wordex_ExtensaoImagemValida(ext)
End Function

Private Function Wordex_ExtensaoImagemValida(ByVal ext As String) As Boolean
    Select Case LCase$(ext)
        Case "png", "jpg", "jpeg", "gif", "webp", "bmp", "tif", "tiff", "ico", "svg"
            Wordex_ExtensaoImagemValida = True
        Case Else
            Wordex_ExtensaoImagemValida = False
    End Select
End Function

Private Function Wordex_PareceUrlHttp(ByVal valorTexto As String) As Boolean
    Dim texto As String

    texto = LCase$(Trim$(valorTexto))
    Wordex_PareceUrlHttp = (Left$(texto, 7) = "http://" Or Left$(texto, 8) = "https://")
End Function

Private Function Wordex_ResolverCaminhoImagem(ByVal caminho As String) As String
    Dim relativo As String

    caminho = Trim$(caminho)

    If caminho = vbNullString Then
        Wordex_ResolverCaminhoImagem = vbNullString
        Exit Function
    End If

    If Len(Dir$(caminho, vbNormal)) > 0 Then
        Wordex_ResolverCaminhoImagem = caminho
        Exit Function
    End If

    If ThisWorkbook.Path <> vbNullString Then
        relativo = ThisWorkbook.Path & Application.PathSeparator & caminho

        If Len(Dir$(relativo, vbNormal)) > 0 Then
            Wordex_ResolverCaminhoImagem = relativo
            Exit Function
        End If
    End If

    Wordex_ResolverCaminhoImagem = caminho
End Function

Private Function Wordex_EncodeFileBase64(ByVal filePath As String) As String
    Dim stream As Object
    Dim xmlDoc As Object
    Dim node As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filePath

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
    Set node = xmlDoc.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = stream.Read

    Wordex_EncodeFileBase64 = node.Text
    stream.Close
End Function

Private Function Wordex_ImageMimePorExtensao(ByVal filePath As String) As String
    Dim ext As String

    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "png": Wordex_ImageMimePorExtensao = "image/png"
        Case "jpg", "jpeg": Wordex_ImageMimePorExtensao = "image/jpeg"
        Case "gif": Wordex_ImageMimePorExtensao = "image/gif"
        Case "webp": Wordex_ImageMimePorExtensao = "image/webp"
        Case "bmp": Wordex_ImageMimePorExtensao = "image/bmp"
        Case "tif", "tiff": Wordex_ImageMimePorExtensao = "image/tiff"
        Case Else: Wordex_ImageMimePorExtensao = "application/octet-stream"
    End Select
End Function

Private Function Wordex_ImageFileToDataUri(ByVal filePath As String) As String
    Wordex_ImageFileToDataUri = "data:" & Wordex_ImageMimePorExtensao(filePath) & _
        ";base64," & Wordex_EncodeFileBase64(filePath)
End Function
