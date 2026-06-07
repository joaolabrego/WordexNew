Attribute VB_Name = "Wordex"
Option Explicit

Public Const K_HEADERS_ROW As Long = 1
Public Const K_DETAILS_ROW As Long = 2
Public Const K_REQUIRED_COL As Long = 1

Public Sub AbrirFormularioTotais()
    frmTotais.Show
End Sub

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

        json = json & virgula & JsonString(tituloColuna) & ": "

        If Not Vazio And IsJsonRaw(valorTexto) Then
            json = json & valorTexto
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
        valueJson = Wordex_JsonValorSemFormatacao(cel)
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
        Wordex_InferirKind = "text"
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

        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble
            Wordex_InferirKind = "number"

        Case vbCurrency, vbDecimal
            Wordex_InferirKind = "currency"

        Case vbString
            Wordex_InferirKind = "text"

        Case Else
            Wordex_InferirKind = "text"
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
