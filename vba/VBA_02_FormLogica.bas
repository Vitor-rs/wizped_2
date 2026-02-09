' ===========================================================
' FORM CODE: frmAlunos v4 (code-behind)
' Cole TODO este codigo no code-behind do frmAlunos (F7)
'
' RECURSOS:
'   - Autocomplete com realce [match] para Alunos e Livros
'   - Navegacao por seta no dropdown de sugestoes
'   - Hover simulado (MouseMove)
'   - Botao dropdown no Livro (mostra todos)
'   - DblClick na Agenda para editar (sem apagar)
'   - Valores padrao ao criar novo aluno
'   - Hora formatada HH:MM (corrige decimal)
'   - Busca accent-insensitive e case-insensitive
' ===========================================================

Option Explicit

' --- Estado do formulario ---
Private mEditando As Boolean
Private mLinhaAtual As Long

' --- Livro selecionado (TextBox, nao ComboBox) ---
Private mIDLivroSelecionado As Variant

' --- Agenda: ID sendo editado (DblClick) ---
Private mAgendaEditandoID As Long   ' 0 = modo adicionar, >0 = modo editar

' --- Flags anti-recursao ---
Private mSuprimirBusca As Boolean
Private mSuprimirLivro As Boolean

' ===========================================================
' INICIALIZACAO
' ===========================================================

Private Sub UserForm_Initialize()
    mEditando = False
    mLinhaAtual = 0
    mIDLivroSelecionado = Empty
    mAgendaEditandoID = 0
    mSuprimirBusca = False
    mSuprimirLivro = False
    CarregarLookups
    LimparForm
End Sub

' ===========================================================
' LOOKUPS
' ===========================================================

Private Sub CarregarLookups()
    Dim ws As Worksheet, r As Long
    
    Set ws = ThisWorkbook.Sheets("BD_Experiencia")
    cmbExperiencia.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbExperiencia.AddItem
        cmbExperiencia.List(cmbExperiencia.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbExperiencia.List(cmbExperiencia.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    Set ws = ThisWorkbook.Sheets("BD_Modalidades")
    cmbModalidade.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbModalidade.AddItem
        cmbModalidade.List(cmbModalidade.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbModalidade.List(cmbModalidade.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    Set ws = ThisWorkbook.Sheets("BD_Status")
    cmbStatus.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbStatus.AddItem
        cmbStatus.List(cmbStatus.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbStatus.List(cmbStatus.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    Set ws = ThisWorkbook.Sheets("BD_Contrato")
    cmbContrato.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbContrato.AddItem
        cmbContrato.List(cmbContrato.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbContrato.List(cmbContrato.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    Set ws = ThisWorkbook.Sheets("BD_Professores")
    cmbProfessor.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbProfessor.AddItem
        cmbProfessor.List(cmbProfessor.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbProfessor.List(cmbProfessor.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    cmbDia.Clear
    Dim dias As Variant: dias = Array("2a", "3a", "4a", "5a", "6a", "Sab")
    Dim d As Variant
    For Each d In dias: cmbDia.AddItem d: Next d
    
    Set ws = ThisWorkbook.Sheets("BD_Horarios")
    cmbHora.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Formatar hora como texto HH:MM caso venha como decimal
        Dim horaCell As Variant: horaCell = ws.Cells(r, 2).Value
        cmbHora.AddItem FormatarHora(horaCell)
    Next r
End Sub

' ===========================================================
' FORMATAR HORA (converte decimal do Excel para HH:MM)
'
' O Excel armazena horas como fracoes do dia:
'   07:00 -> 0,291666...
'   08:00 -> 0,333333...
' Esta funcao detecta isso e converte para "07:00", "08:00" etc.
' Se ja for texto ("07:00"), retorna como esta.
' ===========================================================

Private Function FormatarHora(valor As Variant) As String
    If IsEmpty(valor) Then
        FormatarHora = ""
    ElseIf IsNumeric(valor) Then
        ' Se e um numero entre 0 e 1, e um horario do Excel
        If CDbl(valor) >= 0 And CDbl(valor) < 1 Then
            FormatarHora = Format(CDate(valor), "HH:MM")
        Else
            FormatarHora = CStr(valor)
        End If
    ElseIf IsDate(valor) Then
        FormatarHora = Format(valor, "HH:MM")
    Else
        FormatarHora = CStr(valor)
    End If
End Function

' ===========================================================
'  AUTOCOMPLETE DE ALUNOS
' ===========================================================

Private Sub txtBusca_Change()
    If mSuprimirBusca Then Exit Sub
    Dim termo As String: termo = Trim(txtBusca.Value)
    If Len(termo) = 0 Then
        lstSugestoes.Visible = False: Exit Sub
    End If
    FiltrarSugestoesAlunos termo
End Sub

' Setas no campo de busca -> navegar lista
Private Sub txtBusca_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 40 Then  ' Seta pra baixo
        If lstSugestoes.Visible And lstSugestoes.ListCount > 0 Then
            If lstSugestoes.ListIndex < 0 Then lstSugestoes.ListIndex = 0
            lstSugestoes.SetFocus
            KeyCode = 0
        End If
    ElseIf KeyCode = 13 Then  ' Enter
        If lstSugestoes.Visible And lstSugestoes.ListIndex >= 0 Then
            SelecionarSugestaoAluno: KeyCode = 0
        End If
    ElseIf KeyCode = 27 Then  ' Escape
        lstSugestoes.Visible = False: KeyCode = 0
    End If
End Sub

' Navegacao dentro do lstSugestoes
Private Sub lstSugestoes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then  ' Enter = selecionar
        If lstSugestoes.ListIndex >= 0 Then SelecionarSugestaoAluno
        KeyCode = 0
    ElseIf KeyCode = 27 Then  ' Escape = voltar ao campo
        lstSugestoes.Visible = False
        txtBusca.SetFocus: KeyCode = 0
    ElseIf KeyCode = 38 Then  ' Seta pra cima no primeiro item = voltar
        If lstSugestoes.ListIndex = 0 Then
            lstSugestoes.Visible = False
            txtBusca.SetFocus: KeyCode = 0
        End If
    End If
End Sub

' Duplo-clique seleciona
Private Sub lstSugestoes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SelecionarSugestaoAluno
End Sub

' Hover simulado: seleciona visualmente o item sob o cursor
Private Sub lstSugestoes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                    ByVal X As Single, ByVal Y As Single)
    If lstSugestoes.ListCount = 0 Then Exit Sub
    Dim itens As Long: itens = lstSugestoes.ListCount
    If itens > 20 Then itens = 20
    Dim altItem As Single: altItem = lstSugestoes.Height / itens
    If altItem < 1 Then Exit Sub
    Dim idx As Long: idx = Int(Y / altItem)
    If idx < 0 Then idx = 0
    If idx >= lstSugestoes.ListCount Then idx = lstSugestoes.ListCount - 1
    If lstSugestoes.ListIndex <> idx Then lstSugestoes.ListIndex = idx
End Sub

Private Sub FiltrarSugestoesAlunos(termo As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Alunos")
    Dim lastRow As Long: lastRow = UltimaLinhaAlunos(ws)
    Dim termoNorm As String: termoNorm = RemoverAcentos(LCase(termo))
    
    lstSugestoes.Clear
    Dim r As Long, cont As Long: cont = 0
    Dim nomeAluno As String, nomeNorm As String, idAluno As Variant
    
    For r = 2 To lastRow
        nomeAluno = CStr(ws.Cells(r, 2).Value)
        If Len(nomeAluno) = 0 Then GoTo ProxAluno
        
        nomeNorm = RemoverAcentos(LCase(nomeAluno))
        Dim encontrou As Boolean: encontrou = False
        
        If InStr(1, nomeNorm, termoNorm, vbBinaryCompare) > 0 Then encontrou = True
        
        If Not encontrou And IsNumeric(termo) Then
            idAluno = ws.Cells(r, 1).Value
            If Not IsEmpty(idAluno) Then
                If InStr(1, CStr(idAluno), termo, vbBinaryCompare) > 0 Then encontrou = True
            End If
        End If
        
        If encontrou Then
            lstSugestoes.AddItem
            lstSugestoes.List(cont, 0) = r
            idAluno = ws.Cells(r, 1).Value
            lstSugestoes.List(cont, 1) = IIf(IsEmpty(idAluno), "-", CStr(idAluno))
            lstSugestoes.List(cont, 2) = RealcarTexto(nomeAluno, termo)
            cont = cont + 1
            If cont >= 20 Then Exit For
        End If
ProxAluno:
    Next r
    
    MostrarListBox lstSugestoes, cont
End Sub

Private Sub SelecionarSugestaoAluno()
    If lstSugestoes.ListIndex < 0 Then Exit Sub
    Dim linhaAluno As Long
    linhaAluno = CLng(lstSugestoes.List(lstSugestoes.ListIndex, 0))
    lstSugestoes.Visible = False
    CarregarAluno linhaAluno
    Feedback "Aluno carregado: " & txtNome.Value, False
End Sub

' ===========================================================
'  AUTOCOMPLETE DE LIVROS
' ===========================================================

Private Sub txtLivro_Change()
    If mSuprimirLivro Then Exit Sub
    mIDLivroSelecionado = Empty   ' Limpar selecao ao digitar
    
    Dim termo As String: termo = Trim(txtLivro.Value)
    If Len(termo) = 0 Then
        lstLivroSugestoes.Visible = False: Exit Sub
    End If
    FiltrarSugestoesLivros termo
End Sub

Private Sub txtLivro_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 40 Then
        If lstLivroSugestoes.Visible And lstLivroSugestoes.ListCount > 0 Then
            If lstLivroSugestoes.ListIndex < 0 Then lstLivroSugestoes.ListIndex = 0
            lstLivroSugestoes.SetFocus: KeyCode = 0
        Else
            ' Se lista nao esta visivel, abrir com todos os livros
            MostrarTodosLivros: KeyCode = 0
        End If
    ElseIf KeyCode = 13 Then
        If lstLivroSugestoes.Visible And lstLivroSugestoes.ListIndex >= 0 Then
            SelecionarSugestaoLivro: KeyCode = 0
        End If
    ElseIf KeyCode = 27 Then
        lstLivroSugestoes.Visible = False: KeyCode = 0
    End If
End Sub

Private Sub lstLivroSugestoes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If lstLivroSugestoes.ListIndex >= 0 Then SelecionarSugestaoLivro
        KeyCode = 0
    ElseIf KeyCode = 27 Then
        lstLivroSugestoes.Visible = False
        txtLivro.SetFocus: KeyCode = 0
    ElseIf KeyCode = 38 Then
        If lstLivroSugestoes.ListIndex = 0 Then
            lstLivroSugestoes.Visible = False
            txtLivro.SetFocus: KeyCode = 0
        End If
    End If
End Sub

Private Sub lstLivroSugestoes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SelecionarSugestaoLivro
End Sub

Private Sub lstLivroSugestoes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                         ByVal X As Single, ByVal Y As Single)
    If lstLivroSugestoes.ListCount = 0 Then Exit Sub
    Dim itens As Long: itens = lstLivroSugestoes.ListCount
    If itens > 20 Then itens = 20
    Dim altItem As Single: altItem = lstLivroSugestoes.Height / itens
    If altItem < 1 Then Exit Sub
    Dim idx As Long: idx = Int(Y / altItem)
    If idx < 0 Then idx = 0
    If idx >= lstLivroSugestoes.ListCount Then idx = lstLivroSugestoes.ListCount - 1
    If lstLivroSugestoes.ListIndex <> idx Then lstLivroSugestoes.ListIndex = idx
End Sub

' Botao dropdown: mostra TODOS os livros ativos
Private Sub btnLivroDD_Click()
    If lstLivroSugestoes.Visible Then
        lstLivroSugestoes.Visible = False
    Else
        MostrarTodosLivros
    End If
End Sub

Private Sub MostrarTodosLivros()
    FiltrarSugestoesLivros ""   ' String vazia = sem filtro = mostra todos
End Sub

Private Sub FiltrarSugestoesLivros(termo As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Livros")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim termoNorm As String: termoNorm = RemoverAcentos(LCase(termo))
    Dim semFiltro As Boolean: semFiltro = (Len(termoNorm) = 0)
    
    lstLivroSugestoes.Clear
    
    ' Coletar livros ativos que batem
    Dim nRes As Long: nRes = 0
    Dim r As Long, nomeLivro As String
    For r = 2 To lastRow
        If ws.Cells(r, 6).Value = True Then
            If semFiltro Then
                nRes = nRes + 1
            Else
                nomeLivro = CStr(ws.Cells(r, 2).Value)
                If InStr(1, RemoverAcentos(LCase(nomeLivro)), termoNorm, vbBinaryCompare) > 0 Then
                    nRes = nRes + 1
                End If
            End If
        End If
    Next r
    
    If nRes = 0 Then
        lstLivroSugestoes.Visible = False: Exit Sub
    End If
    
    ' Armazenar e ordenar por Ordem
    ReDim resultados(1 To nRes, 1 To 3) As Variant
    Dim idx As Long: idx = 0
    For r = 2 To lastRow
        If ws.Cells(r, 6).Value = True Then
            nomeLivro = CStr(ws.Cells(r, 2).Value)
            Dim incluir As Boolean
            If semFiltro Then
                incluir = True
            Else
                incluir = (InStr(1, RemoverAcentos(LCase(nomeLivro)), termoNorm, vbBinaryCompare) > 0)
            End If
            If incluir Then
                idx = idx + 1
                resultados(idx, 1) = ws.Cells(r, 1).Value
                resultados(idx, 2) = ws.Cells(r, 2).Value
                resultados(idx, 3) = ws.Cells(r, 4).Value
            End If
        End If
    Next r
    
    ' Bubble sort por Ordem
    Dim i As Long, j As Long
    Dim t1 As Variant, t2 As Variant, t3 As Variant
    For i = 1 To nRes - 1
        For j = 1 To nRes - i
            If resultados(j, 3) > resultados(j + 1, 3) Then
                t1 = resultados(j, 1): resultados(j, 1) = resultados(j + 1, 1): resultados(j + 1, 1) = t1
                t2 = resultados(j, 2): resultados(j, 2) = resultados(j + 1, 2): resultados(j + 1, 2) = t2
                t3 = resultados(j, 3): resultados(j, 3) = resultados(j + 1, 3): resultados(j + 1, 3) = t3
            End If
        Next j
    Next i
    
    ' Popular ListBox (max 20)
    Dim cont As Long: cont = 0
    For i = 1 To nRes
        lstLivroSugestoes.AddItem
        lstLivroSugestoes.List(cont, 0) = resultados(i, 1)
        If semFiltro Then
            lstLivroSugestoes.List(cont, 1) = CStr(resultados(i, 2))
        Else
            lstLivroSugestoes.List(cont, 1) = RealcarTexto(CStr(resultados(i, 2)), termo)
        End If
        cont = cont + 1
        If cont >= 20 Then Exit For
    Next i
    
    MostrarListBox lstLivroSugestoes, cont
End Sub

Private Sub SelecionarSugestaoLivro()
    If lstLivroSugestoes.ListIndex < 0 Then Exit Sub
    
    Dim idLivro As Long
    idLivro = CLng(lstLivroSugestoes.List(lstLivroSugestoes.ListIndex, 0))
    mIDLivroSelecionado = idLivro
    
    ' Buscar nome limpo do livro
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Livros")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CLng(ws.Cells(r, 1).Value) = idLivro Then
            mSuprimirLivro = True
            txtLivro.Value = ws.Cells(r, 2).Value
            mSuprimirLivro = False
            ' Auto-preencher experiencia padrao
            SelecionarCombo cmbExperiencia, ws.Cells(r, 5).Value
            Exit For
        End If
    Next r
    
    lstLivroSugestoes.Visible = False
    AtualizarTipoPreview
End Sub

' ===========================================================
'  REALCE DE TEXTO E NORMALIZACAO
' ===========================================================

' Insere [colchetes] ao redor de cada ocorrencia do termo.
' Busca accent-insensitive, resultado mantem acentos originais.
' Ex: RealcarTexto("Arthur Alonso Flores", "lo") -> "Arthur A[lo]nso F[lo]res"

Private Function RealcarTexto(original As String, termo As String) As String
    If Len(termo) = 0 Then RealcarTexto = original: Exit Function
    
    Dim origNorm As String: origNorm = RemoverAcentos(LCase(original))
    Dim termoNorm As String: termoNorm = RemoverAcentos(LCase(termo))
    Dim lenT As Long: lenT = Len(termoNorm)
    
    ' Encontrar posicoes de match
    Dim posicoes() As Long, nPos As Long: nPos = 0
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, origNorm, termoNorm, vbBinaryCompare)
        If pos = 0 Then Exit Do
        nPos = nPos + 1
        ReDim Preserve posicoes(1 To nPos)
        posicoes(nPos) = pos
        pos = pos + lenT
    Loop
    
    If nPos = 0 Then RealcarTexto = original: Exit Function
    
    ' Inserir [ ] de tras pra frente (nao desloca posicoes)
    Dim resultado As String: resultado = original
    Dim k As Long
    For k = nPos To 1 Step -1
        Dim p As Long: p = posicoes(k)
        resultado = Left(resultado, p - 1) & "[" & Mid(resultado, p, lenT) & "]" & Mid(resultado, p + lenT)
    Next k
    RealcarTexto = resultado
End Function

Private Function RemoverAcentos(texto As String) As String
    Dim i As Long, resultado As String, c As String, codigo As Long
    resultado = ""
    For i = 1 To Len(texto)
        c = Mid(texto, i, 1): codigo = AscW(c)
        Select Case codigo
            Case 192 To 197: c = "A"
            Case 224 To 229: c = "a"
            Case 200 To 203: c = "E"
            Case 232 To 235: c = "e"
            Case 204 To 207: c = "I"
            Case 236 To 239: c = "i"
            Case 210 To 214: c = "O"
            Case 242 To 246: c = "o"
            Case 217 To 220: c = "U"
            Case 249 To 252: c = "u"
            Case 199: c = "C": Case 231: c = "c"
            Case 209: c = "N": Case 241: c = "n"
        End Select
        resultado = resultado & c
    Next i
    RemoverAcentos = resultado
End Function

' ===========================================================
'  BUSCA EXATA (botao Buscar)
' ===========================================================

Private Sub btnBuscar_Click()
    lstSugestoes.Visible = False
    
    Dim termo As String: termo = Trim(txtBusca.Value)
    If Len(termo) = 0 Then
        Feedback "Digite um ID ou nome para buscar.", True
        txtBusca.SetFocus: Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Alunos")
    Dim lastRow As Long: lastRow = UltimaLinhaAlunos(ws)
    Dim r As Long
    Dim termoNorm As String: termoNorm = RemoverAcentos(LCase(termo))
    
    If IsNumeric(termo) Then
        For r = 2 To lastRow
            If CStr(ws.Cells(r, 1).Value) = CStr(CLng(termo)) Then
                CarregarAluno r
                Feedback "Encontrado: " & ws.Cells(r, 2).Value, False: Exit Sub
            End If
        Next r
    End If
    
    For r = 2 To lastRow
        If InStr(1, RemoverAcentos(LCase(CStr(ws.Cells(r, 2).Value))), termoNorm, vbBinaryCompare) > 0 Then
            CarregarAluno r
            Feedback "Encontrado: " & ws.Cells(r, 2).Value, False: Exit Sub
        End If
    Next r
    
    Feedback "Nenhum aluno encontrado para: " & termo, True
End Sub

Private Sub btnLimpar_Click()
    FecharOverlays
    LimparForm
    Feedback "", False
End Sub

' ===========================================================
'  CARREGAR ALUNO
' ===========================================================

Private Sub CarregarAluno(linhaAluno As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Alunos")
    
    mEditando = True
    mLinhaAtual = linhaAluno
    mAgendaEditandoID = 0
    
    ' ID (trava em modo edicao)
    If Not IsEmpty(ws.Cells(linhaAluno, 1).Value) Then
        txtID.Value = CStr(ws.Cells(linhaAluno, 1).Value)
    Else
        txtID.Value = ""
    End If
    txtID.Enabled = False
    
    ' Nome
    txtNome.Value = CStr(ws.Cells(linhaAluno, 2).Value)
    
    ' Livro (TextBox + mIDLivroSelecionado)
    Dim idLivro As Variant: idLivro = ws.Cells(linhaAluno, 5).Value
    If Not IsEmpty(idLivro) Then
        mIDLivroSelecionado = CLng(idLivro)
        Dim wsL As Worksheet: Set wsL = ThisWorkbook.Sheets("BD_Livros")
        Dim rr As Long
        mSuprimirLivro = True
        For rr = 2 To wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row
            If CLng(wsL.Cells(rr, 1).Value) = CLng(idLivro) Then
                txtLivro.Value = wsL.Cells(rr, 2).Value: Exit For
            End If
        Next rr
        mSuprimirLivro = False
    Else
        mIDLivroSelecionado = Empty
        mSuprimirLivro = True: txtLivro.Value = "": mSuprimirLivro = False
    End If
    
    SelecionarCombo cmbExperiencia, ws.Cells(linhaAluno, 6).Value
    SelecionarCombo cmbModalidade, ws.Cells(linhaAluno, 7).Value
    
    ' VIP
    Dim vipVal As Variant: vipVal = ws.Cells(linhaAluno, 8).Value
    chkVIP.Value = IIf(IsEmpty(vipVal) Or IsNull(vipVal), False, CBool(vipVal))
    
    SelecionarCombo cmbStatus, ws.Cells(linhaAluno, 3).Value
    SelecionarCombo cmbContrato, ws.Cells(linhaAluno, 4).Value
    SelecionarCombo cmbProfessor, ws.Cells(linhaAluno, 9).Value
    
    ' Data
    If IsDate(ws.Cells(linhaAluno, 10).Value) Then
        txtData.Value = Format(ws.Cells(linhaAluno, 10).Value, "dd/mm/yyyy")
    Else
        txtData.Value = ""
    End If
    
    ' Obs
    txtObs.Value = IIf(IsEmpty(ws.Cells(linhaAluno, 11).Value), "", CStr(ws.Cells(linhaAluno, 11).Value))
    
    ' Agenda + Preview
    CarregarAgenda ws.Cells(linhaAluno, 1).Value
    AtualizarTipoPreview
End Sub

' Carrega a agenda do aluno, formatando hora como HH:MM
Private Sub CarregarAgenda(idAluno As Variant)
    lstAgenda.Clear
    If IsEmpty(idAluno) Then Exit Sub
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Agenda")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).Value) = CStr(idAluno) Then
            lstAgenda.AddItem
            lstAgenda.List(lstAgenda.ListCount - 1, 0) = ws.Cells(r, 1).Value
            ' Hora: formatar como HH:MM (corrige o bug do decimal)
            lstAgenda.List(lstAgenda.ListCount - 1, 1) = FormatarHora(ws.Cells(r, 3).Value)
            lstAgenda.List(lstAgenda.ListCount - 1, 2) = CStr(ws.Cells(r, 4).Value)
        End If
    Next r
End Sub

' ===========================================================
'  NOVO ALUNO (com valores padrao)
' ===========================================================

Public Sub btnNovo_Click()
    FecharOverlays
    LimparForm
    
    txtID.Enabled = True
    
    ' Valores padrao
    SelecionarCombo cmbStatus, 1       ' Ativo
    SelecionarCombo cmbContrato, 1     ' Matricula
    SelecionarCombo cmbModalidade, 1   ' Presencial
    txtData.Value = Format(Date, "dd/mm/yyyy")   ' Data de hoje
    
    txtID.SetFocus
    Feedback "Novo Aluno. Status=Ativo, Contrato=Matricula, Data=hoje.", False
End Sub

' ===========================================================
'  SALVAR
' ===========================================================

Private Sub btnSalvar_Click()
    FecharOverlays
    
    Dim idStr As String: idStr = Trim(txtID.Value)
    
    If Len(idStr) = 0 Then
        Feedback "ID (SponteWeb) e obrigatorio.", True: txtID.SetFocus: Exit Sub
    End If
    If Not IsNumeric(idStr) Then
        Feedback "ID deve ser numerico.", True: txtID.SetFocus: Exit Sub
    End If
    If Len(Trim(txtNome.Value)) = 0 Then
        Feedback "Nome e obrigatorio.", True: txtNome.SetFocus: Exit Sub
    End If
    
    ' Presencial requer Experiencia
    If ValorCombo(cmbModalidade) = 1 Then
        If IsEmpty(ValorCombo(cmbExperiencia)) Then
            Feedback "Modalidade Presencial requer Experiencia.", True
            cmbExperiencia.SetFocus: Exit Sub
        End If
    End If
    
    ' Livro: se tem texto mas nao selecionou da lista
    If Len(Trim(txtLivro.Value)) > 0 And IsEmpty(mIDLivroSelecionado) Then
        Feedback "Selecione um livro da lista de sugestoes.", True
        txtLivro.SetFocus: Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Alunos")
    Dim lastRow As Long: lastRow = UltimaLinhaAlunos(ws)
    
    ' Verificar ID duplicado (modo novo)
    If Not mEditando Then
        Dim r As Long
        For r = 2 To lastRow
            If CStr(ws.Cells(r, 1).Value) = CStr(CLng(idStr)) Then
                Feedback "ID " & idStr & " ja existe! Use Buscar para editar.", True: Exit Sub
            End If
        Next r
    End If
    
    ' Linha de gravacao
    Dim linhaGravar As Long
    If mEditando Then
        linhaGravar = mLinhaAtual
    Else
        Dim nomeAluno As String: nomeAluno = Trim(txtNome.Value)
        Dim encontrou As Boolean: encontrou = False
        For r = 2 To lastRow
            If IsEmpty(ws.Cells(r, 1).Value) Or ws.Cells(r, 1).Value = "" Then
                If StrComp(Trim(CStr(ws.Cells(r, 2).Value)), nomeAluno, vbTextCompare) = 0 Then
                    linhaGravar = r: encontrou = True: Exit For
                End If
            End If
        Next r
        If Not encontrou Then linhaGravar = lastRow + 1
    End If
    
    ' Gravar campos
    ws.Cells(linhaGravar, 1).Value = CLng(idStr)
    ws.Cells(linhaGravar, 2).Value = Trim(txtNome.Value)
    
    If Not IsEmpty(mIDLivroSelecionado) Then
        ws.Cells(linhaGravar, 5).Value = CLng(mIDLivroSelecionado)
    Else: ws.Cells(linhaGravar, 5).Value = Empty
    End If
    
    GravarCombo ws, linhaGravar, 6, cmbExperiencia
    GravarCombo ws, linhaGravar, 7, cmbModalidade
    ws.Cells(linhaGravar, 8).Value = chkVIP.Value
    GravarCombo ws, linhaGravar, 3, cmbStatus
    GravarCombo ws, linhaGravar, 4, cmbContrato
    GravarCombo ws, linhaGravar, 9, cmbProfessor
    
    If Len(Trim(txtData.Value)) > 0 Then
        If IsDate(txtData.Value) Then
            ws.Cells(linhaGravar, 10).Value = CDate(txtData.Value)
            ws.Cells(linhaGravar, 10).NumberFormat = "dd/mm/yyyy"
        Else
            ws.Cells(linhaGravar, 10).Value = txtData.Value
        End If
    Else: ws.Cells(linhaGravar, 10).Value = Empty
    End If
    
    ws.Cells(linhaGravar, 11).Value = Trim(txtObs.Value)
    
    mEditando = True: mLinhaAtual = linhaGravar: txtID.Enabled = False
    Feedback "Aluno salvo! (linha " & linhaGravar & ")", False
End Sub

' ===========================================================
'  EXCLUIR
' ===========================================================

Private Sub btnExcluir_Click()
    FecharOverlays
    If Not mEditando Then
        Feedback "Nenhum aluno carregado para excluir.", True: Exit Sub
    End If
    
    If MsgBox("Excluir " & txtNome.Value & "?" & vbCrLf & _
              "Esta acao nao pode ser desfeita.", _
              vbQuestion + vbYesNo, "Confirmar") = vbNo Then Exit Sub
    
    ' Remover horarios (reverso)
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("BD_Agenda")
    Dim idAluno As String: idAluno = CStr(txtID.Value)
    Dim rr As Long
    For rr = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If CStr(wsA.Cells(rr, 2).Value) = idAluno Then wsA.Rows(rr).Delete
    Next rr
    
    ThisWorkbook.Sheets("BD_Alunos").Rows(mLinhaAtual).Delete
    LimparForm
    Feedback "Aluno excluido.", False
End Sub

' ===========================================================
'  AGENDA: Adicionar / Remover / DblClick para Editar
' ===========================================================

Private Sub btnAddHora_Click()
    FecharOverlays
    
    If Len(Trim(txtID.Value)) = 0 Then
        Feedback "Salve o aluno antes de adicionar horarios.", True: Exit Sub
    End If
    If cmbDia.ListIndex = -1 Then
        Feedback "Selecione um dia.", True: cmbDia.SetFocus: Exit Sub
    End If
    If cmbHora.ListIndex = -1 Then
        Feedback "Selecione uma hora.", True: cmbHora.SetFocus: Exit Sub
    End If
    
    Dim idAluno As Long: idAluno = CLng(txtID.Value)
    Dim dia As String: dia = cmbDia.Value
    Dim hora As String: hora = cmbHora.Value
    
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("BD_Agenda")
    Dim lastRow As Long: lastRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    
    ' === MODO EDITAR: atualizar registro existente ===
    If mAgendaEditandoID > 0 Then
        Dim rr As Long
        For rr = 2 To lastRow
            If CLng(wsA.Cells(rr, 1).Value) = mAgendaEditandoID Then
                ' Verificar duplicata (excluindo o proprio registro)
                Dim rrChk As Long
                For rrChk = 2 To lastRow
                    If rrChk <> rr And CLng(wsA.Cells(rrChk, 2).Value) = idAluno Then
                        If CStr(wsA.Cells(rrChk, 3).Value) = hora And CStr(wsA.Cells(rrChk, 4).Value) = dia Then
                            Feedback "Horario " & hora & " " & dia & " ja existe.", True
                            Exit Sub
                        End If
                    End If
                Next rrChk
                
                ' Atualizar
                wsA.Cells(rr, 3).Value = hora
                wsA.Cells(rr, 3).NumberFormat = "@"
                wsA.Cells(rr, 4).Value = dia
                
                mAgendaEditandoID = 0
                CarregarAgenda idAluno
                Feedback "Horario atualizado: " & hora & " " & dia, False
                Exit Sub
            End If
        Next rr
        
        ' Se nao encontrou o registro (foi deletado?), resetar modo
        mAgendaEditandoID = 0
    End If
    
    ' === MODO ADICIONAR: novo registro ===
    ' Verificar duplicata
    For rr = 2 To lastRow
        If CLng(wsA.Cells(rr, 2).Value) = idAluno Then
            If CStr(wsA.Cells(rr, 3).Value) = hora And CStr(wsA.Cells(rr, 4).Value) = dia Then
                Feedback "Horario " & hora & " " & dia & " ja existe.", True: Exit Sub
            End If
        End If
    Next rr
    
    ' Proximo ID
    Dim maxID As Long: maxID = 0
    For rr = 2 To lastRow
        If wsA.Cells(rr, 1).Value > maxID Then maxID = wsA.Cells(rr, 1).Value
    Next rr
    
    Dim nl As Long: nl = lastRow + 1
    wsA.Cells(nl, 1).Value = maxID + 1
    wsA.Cells(nl, 2).Value = idAluno
    wsA.Cells(nl, 3).NumberFormat = "@"
    wsA.Cells(nl, 3).Value = hora
    wsA.Cells(nl, 4).Value = dia
    
    CarregarAgenda idAluno
    Feedback "Horario adicionado: " & hora & " " & dia, False
End Sub

Private Sub btnRemHora_Click()
    FecharOverlays
    If lstAgenda.ListIndex = -1 Then
        Feedback "Selecione um horario para remover.", True: Exit Sub
    End If
    
    Dim idAgenda As Long
    idAgenda = CLng(lstAgenda.List(lstAgenda.ListIndex, 0))
    
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("BD_Agenda")
    Dim rr As Long
    For rr = 2 To wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
        If CLng(wsA.Cells(rr, 1).Value) = idAgenda Then
            wsA.Rows(rr).Delete: Exit For
        End If
    Next rr
    
    mAgendaEditandoID = 0
    CarregarAgenda CLng(txtID.Value)
    Feedback "Horario removido.", False
End Sub

' DblClick na agenda: preenche combos para edicao SEM apagar
Private Sub lstAgenda_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstAgenda.ListIndex = -1 Then Exit Sub
    
    ' Guardar o ID_Agenda sendo editado
    mAgendaEditandoID = CLng(lstAgenda.List(lstAgenda.ListIndex, 0))
    
    ' Preencher combos com valores atuais
    Dim horaVal As String: horaVal = lstAgenda.List(lstAgenda.ListIndex, 1)
    Dim diaVal As String: diaVal = lstAgenda.List(lstAgenda.ListIndex, 2)
    
    Dim i As Long
    For i = 0 To cmbHora.ListCount - 1
        If CStr(cmbHora.List(i)) = horaVal Then cmbHora.ListIndex = i: Exit For
    Next i
    For i = 0 To cmbDia.ListCount - 1
        If CStr(cmbDia.List(i)) = diaVal Then cmbDia.ListIndex = i: Exit For
    Next i
    
    Feedback "Editando horario. Altere Dia/Hora e clique + Adicionar.", False
End Sub

' ===========================================================
'  PREVIEW DO TIPO (tempo real)
' ===========================================================

Private Sub cmbExperiencia_Change(): AtualizarTipoPreview: End Sub
Private Sub cmbModalidade_Change(): AtualizarTipoPreview: End Sub
Private Sub chkVIP_Change(): AtualizarTipoPreview: End Sub

Private Sub AtualizarTipoPreview()
    Dim tipo As String: tipo = ""
    Dim r As Long
    
    Dim idExp As Variant: idExp = ValorCombo(cmbExperiencia)
    If Not IsEmpty(idExp) Then
        Dim wsExp As Worksheet: Set wsExp = ThisWorkbook.Sheets("BD_Experiencia")
        For r = 2 To wsExp.Cells(wsExp.Rows.Count, 1).End(xlUp).Row
            If CLng(wsExp.Cells(r, 1).Value) = CLng(idExp) Then
                tipo = wsExp.Cells(r, 3).Value: Exit For
            End If
        Next r
    End If
    
    If chkVIP.Value Then tipo = IIf(Len(tipo) > 0, tipo & " VIP", "VIP")
    
    Dim idMod As Variant: idMod = ValorCombo(cmbModalidade)
    If Not IsEmpty(idMod) Then
        If CLng(idMod) <> 1 Then
            Dim wsMod As Worksheet: Set wsMod = ThisWorkbook.Sheets("BD_Modalidades")
            For r = 2 To wsMod.Cells(wsMod.Rows.Count, 1).End(xlUp).Row
                If CLng(wsMod.Cells(r, 1).Value) = CLng(idMod) Then
                    Dim apelido As String: apelido = wsMod.Cells(r, 3).Value
                    tipo = IIf(Len(tipo) > 0, tipo & " " & apelido, apelido)
                    Exit For
                End If
            Next r
        End If
    End If
    
    If Len(tipo) = 0 Then tipo = "..."
    lblTipoPreview.Caption = tipo
End Sub

' ===========================================================
'  FECHAR
' ===========================================================

Private Sub btnFechar_Click()
    Unload Me
End Sub

' ===========================================================
'  UTILITARIOS
' ===========================================================

Private Sub LimparForm()
    mEditando = False: mLinhaAtual = 0
    mIDLivroSelecionado = Empty: mAgendaEditandoID = 0
    
    txtID.Value = "": txtID.Enabled = True
    txtNome.Value = "": txtData.Value = "": txtObs.Value = ""
    
    mSuprimirBusca = True: txtBusca.Value = "": mSuprimirBusca = False
    mSuprimirLivro = True: txtLivro.Value = "": mSuprimirLivro = False
    
    cmbExperiencia.ListIndex = -1: cmbModalidade.ListIndex = -1
    cmbStatus.ListIndex = -1: cmbContrato.ListIndex = -1
    cmbProfessor.ListIndex = -1: cmbDia.ListIndex = -1: cmbHora.ListIndex = -1
    
    chkVIP.Value = False
    lstAgenda.Clear
    lstSugestoes.Clear: lstSugestoes.Visible = False
    lstLivroSugestoes.Clear: lstLivroSugestoes.Visible = False
    lblTipoPreview.Caption = "": lblFeedback.Caption = ""
End Sub

Private Sub FecharOverlays()
    lstSugestoes.Visible = False
    lstLivroSugestoes.Visible = False
End Sub

Private Sub SelecionarCombo(cmb As MSForms.ComboBox, valor As Variant)
    If IsEmpty(valor) Or IsNull(valor) Then cmb.ListIndex = -1: Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If CStr(cmb.List(i, 0)) = CStr(valor) Then cmb.ListIndex = i: Exit Sub
    Next i
    cmb.ListIndex = -1
End Sub

Private Function ValorCombo(cmb As MSForms.ComboBox) As Variant
    If cmb.ListIndex = -1 Then ValorCombo = Empty Else ValorCombo = cmb.List(cmb.ListIndex, 0)
End Function

Private Sub GravarCombo(ws As Worksheet, linha As Long, col As Long, cmb As MSForms.ComboBox)
    If Not IsEmpty(ValorCombo(cmb)) Then
        ws.Cells(linha, col).Value = CLng(ValorCombo(cmb))
    Else
        ws.Cells(linha, col).Value = Empty
    End If
End Sub

Private Sub Feedback(msg As String, isErro As Boolean)
    lblFeedback.Caption = msg
    lblFeedback.ForeColor = IIf(isErro, &HFF&, &H8000&)
End Sub

' Ajusta altura e mostra um ListBox overlay
Private Sub MostrarListBox(lst As MSForms.ListBox, contItens As Long)
    If contItens > 0 Then
        Dim alt As Single: alt = contItens * 16
        If alt < 32 Then alt = 32
        If alt > 260 Then alt = 260
        lst.Height = alt
        lst.Visible = True
    Else
        lst.Visible = False
    End If
End Sub

' Ultima linha com dados na BD_Alunos (checa col A e B)
Private Function UltimaLinhaAlunos(ws As Worksheet) As Long
    Dim a As Long: a = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim b As Long: b = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    UltimaLinhaAlunos = IIf(b > a, b, a)
End Function
