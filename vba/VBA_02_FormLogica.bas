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

' --- Historico: ID sendo editado (DblClick) ---
Private mHistoricoEditandoID As Long  ' 0 = modo adicionar, >0 = modo editar

' --- Livro antigo (para detectar mudanca ao salvar) ---
Private mLivroAntigoID As Variant

' --- Flags anti-recursao ---
Private mSuprimirBusca As Boolean
Private mSuprimirLivro As Boolean
Private mSuprimirDiaChange As Boolean
Private mSuprimirContratoChange As Boolean
Private mSuprimirDataFormat As Boolean
Private mBackspacing As Boolean

' --- Dirty flag: detecta alteracoes nao salvas ---
Private mFormModificado As Boolean

' ===========================================================
' INICIALIZACAO
' ===========================================================

Private Sub UserForm_Initialize()
    mEditando = False
    mLinhaAtual = 0
    mIDLivroSelecionado = Empty
    mLivroAntigoID = Empty
    mAgendaEditandoID = 0
    mHistoricoEditandoID = 0
    mSuprimirBusca = False
    mSuprimirLivro = False
    mSuprimirDiaChange = False
    mSuprimirContratoChange = False
    mSuprimirDataFormat = False
    mFormModificado = False
    CarregarLookups
    LimparForm
End Sub

' Click fora dos overlays: fechar
Private Sub UserForm_Click()
    FecharOverlays
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
    
    ' Funcionarios (ListBox MultiSelect — filtrar por Funcao=Professor)
    Set ws = ThisWorkbook.Sheets("BD_Funcionarios")
    lstFuncionarios.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LCase(CStr(ws.Cells(r, 3).Value)) = "professor" Then
            lstFuncionarios.AddItem
            lstFuncionarios.List(lstFuncionarios.ListCount - 1, 0) = ws.Cells(r, 1).Value
            lstFuncionarios.List(lstFuncionarios.ListCount - 1, 1) = ws.Cells(r, 2).Value
        End If
    Next r
    
    ' Historico: TipoOcorrencia
    Set ws = ThisWorkbook.Sheets("BD_TipoOcorrencia")
    cmbTipoOcorrencia.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbTipoOcorrencia.AddItem
        cmbTipoOcorrencia.List(cmbTipoOcorrencia.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbTipoOcorrencia.List(cmbTipoOcorrencia.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    ' Historico: Responsavel (dropdown de todos os funcionarios)
    Set ws = ThisWorkbook.Sheets("BD_Funcionarios")
    cmbResponsavel.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbResponsavel.AddItem
        cmbResponsavel.List(cmbResponsavel.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbResponsavel.List(cmbResponsavel.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    ' Historico: Livro
    Set ws = ThisWorkbook.Sheets("BD_Livros")
    cmbLivroHist.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        cmbLivroHist.AddItem
        cmbLivroHist.List(cmbLivroHist.ListCount - 1, 0) = ws.Cells(r, 1).Value
        cmbLivroHist.List(cmbLivroHist.ListCount - 1, 1) = ws.Cells(r, 2).Value
    Next r
    
    cmbDia.Clear
    Dim dias As Variant: dias = Array("2ª", "3ª", "4ª", "5ª", "6ª", "Sáb")
    Dim d As Variant
    For Each d In dias: cmbDia.AddItem d: Next d
    
    ' cmbHora é populado dinamicamente em cmbDia_Change
    cmbHora.Clear
End Sub

' ===========================================================
'  DROPDOWN DEPENDENTE: DIA -> HORA
'  Usa GetHorariosDisponiveis() de Mod_Horarios (VBA_05)
' ===========================================================

Private Sub cmbDia_Change()
    If mSuprimirDiaChange Then Exit Sub
    
    ' Guardar hora selecionada antes de limpar
    Dim horaAnterior As String
    If cmbHora.ListIndex >= 0 Then horaAnterior = cmbHora.Value Else horaAnterior = ""
    
    cmbHora.Clear
    If cmbDia.ListIndex = -1 Then Exit Sub
    
    Dim horas As Variant
    horas = GetHorariosDisponiveis(cmbDia.Value)
    
    Dim h As Variant
    For Each h In horas
        ' Não adicionar mensagens de erro como item
        If Left(CStr(h), 4) <> "Erro" And Left(CStr(h), 6) <> "Nenhum" Then
            cmbHora.AddItem CStr(h)
        End If
    Next h
    
    ' Restaurar hora anterior se disponivel no novo dia
    If Len(horaAnterior) > 0 Then
        Dim i As Long
        For i = 0 To cmbHora.ListCount - 1
            If CStr(cmbHora.List(i)) = horaAnterior Then
                cmbHora.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    If cmbHora.ListCount = 0 Then
        Feedback "Nenhum horário disponível para " & cmbDia.Value & ".", True
    End If
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
    Dim altItem As Single: altItem = lstSugestoes.Font.Size + 3
    Dim idx As Long: idx = lstSugestoes.TopIndex + Int(Y / altItem)
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
    Dim altItem As Single: altItem = lstLivroSugestoes.Font.Size + 3
    Dim idx As Long: idx = lstLivroSugestoes.TopIndex + Int(Y / altItem)
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
        lstLivroSugestoes.SetFocus
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
    
    ' Popular ListBox (max 50)
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
        If cont >= 50 Then Exit For
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
        mLivroAntigoID = CLng(idLivro)
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
        mLivroAntigoID = Empty
        mSuprimirLivro = True: txtLivro.Value = "": mSuprimirLivro = False
    End If
    
    SelecionarCombo cmbExperiencia, ws.Cells(linhaAluno, 6).Value
    SelecionarCombo cmbModalidade, ws.Cells(linhaAluno, 7).Value
    
    ' VIP
    Dim vipVal As Variant: vipVal = ws.Cells(linhaAluno, 8).Value
    chkVIP.Value = IIf(IsEmpty(vipVal) Or IsNull(vipVal), False, CBool(vipVal))
    
    ' Active State (Col 12)
    Dim ativoVal As Variant: ativoVal = ws.Cells(linhaAluno, 12).Value
    If IsEmpty(ativoVal) Or IsNull(ativoVal) Then
        optAtivo.Value = True
    Else
        If CBool(ativoVal) Then optAtivo.Value = True Else optInativo.Value = True
    End If
    VerificarBloqueioAtivo
    
    SelecionarCombo cmbStatus, ws.Cells(linhaAluno, 3).Value
    ' Default Status for Imported: if empty or invalid, default to Ativo (1)
    If cmbStatus.ListIndex = -1 Then
         ' Check if cell has text "Ativo/Desistente" and map it, otherwise default Ativo
         Dim stVal As String: stVal = LCase(Trim(CStr(ws.Cells(linhaAluno, 3).Value)))
         If stVal = "desistente" Then
             SelecionarCombo cmbStatus, 2 ' Assuming 2 is Desistente
         Else
             SelecionarCombo cmbStatus, 1 ' Ativo
         End If
    End If

    mSuprimirContratoChange = True
    SelecionarCombo cmbContrato, ws.Cells(linhaAluno, 4).Value
    ' Default Contrato: Matricula (1) if empty
    If cmbContrato.ListIndex = -1 Then SelecionarCombo cmbContrato, 1
    mSuprimirContratoChange = False
    
    ' Default Modalidade: Presencial (1) if empty
    If cmbModalidade.ListIndex = -1 Then SelecionarCombo cmbModalidade, 1

    ' Default Experiencia: Interactive (1) (User didn't explicitly say, but usually Presencial needs it)
    ' User said "presencial preenchido... sem necessidade de escolher".
    ' Let's set Experiencia only if Modalidade is Presencial and Exp is empty.
    If ValorCombo(cmbModalidade) = 1 And cmbExperiencia.ListIndex = -1 Then
        SelecionarCombo cmbExperiencia, 1 ' Interactive default
    End If
    
    ' Funcionarios/Professores (N:N via BD_Vinculo_Professor)
    CarregarFuncionariosAluno ws.Cells(linhaAluno, 1).Value
    
    ' Data
    If IsDate(ws.Cells(linhaAluno, 10).Value) Then
        txtData.Value = Format(ws.Cells(linhaAluno, 10).Value, "dd/mm/yyyy")
    Else
        ' Default Data: Today if empty (imported)
        txtData.Value = Format(Date, "dd/mm/yyyy")
    End If
    
    ' Obs
    txtObs.Value = IIf(IsEmpty(ws.Cells(linhaAluno, 11).Value), "", CStr(ws.Cells(linhaAluno, 11).Value))
    
    ' Agenda + Preview + Historico
    CarregarAgenda ws.Cells(linhaAluno, 1).Value
    CarregarHistorico ws.Cells(linhaAluno, 1).Value
    AtualizarTipoPreview
    mFormModificado = False
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
            ' Dia primeiro, Hora depois
            lstAgenda.List(lstAgenda.ListCount - 1, 1) = CStr(ws.Cells(r, 4).Value)
            lstAgenda.List(lstAgenda.ListCount - 1, 2) = FormatarHora(ws.Cells(r, 3).Value)
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
    SelecionarCombo cmbExperiencia, 1  ' Interactive
    SelecionarCombo cmbStatus, 1       ' Ativo
    SelecionarCombo cmbContrato, 1     ' Matricula
    SelecionarCombo cmbModalidade, 1   ' Presencial
    txtData.Value = Format(Date, "dd/mm/yyyy")   ' Data de hoje
    
    txtID.SetFocus
    Feedback "Novo Aluno. Exp=Interactive, Mod=Presencial, Status=Ativo, Contrato=Matricula, Data=hoje.", False
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
    
    ' ============================================================
    ' DETECTAR MUDANCA DE LIVRO + CONTRATO (Matricula/Rematricula)
    ' ============================================================
    Dim livroMudou As Boolean: livroMudou = False
    Dim nomeContratoAtual As String: nomeContratoAtual = ""
    Dim ehMatriculaOuRematricula As Boolean: ehMatriculaOuRematricula = False
    
    ' Verificar se livro mudou (editando OU novo aluno com livro)
    If mEditando Then
        Dim livroAntigoNum As Variant: livroAntigoNum = mLivroAntigoID
        Dim livroNovoNum As Variant: livroNovoNum = mIDLivroSelecionado
        
        If IsEmpty(livroAntigoNum) And Not IsEmpty(livroNovoNum) Then
            livroMudou = True
        ElseIf Not IsEmpty(livroAntigoNum) And IsEmpty(livroNovoNum) Then
            livroMudou = True
        ElseIf Not IsEmpty(livroAntigoNum) And Not IsEmpty(livroNovoNum) Then
            livroMudou = (CLng(livroAntigoNum) <> CLng(livroNovoNum))
        End If
    Else
        ' Novo aluno: se tem livro selecionado, considerar como "mudou"
        If Not IsEmpty(mIDLivroSelecionado) Then livroMudou = True
    End If
    
    ' Verificar contrato (independente de mEditando)
    If cmbContrato.ListIndex >= 0 Then
        nomeContratoAtual = CStr(cmbContrato.List(cmbContrato.ListIndex, 1))
        Dim contratoLower As String: contratoLower = LCase(nomeContratoAtual)
        If InStr(1, contratoLower, "matricula") > 0 Or _
           InStr(1, contratoLower, "rematricula") > 0 Or _
           InStr(1, contratoLower, "matr" & ChrW(237) & "cula") > 0 Or _
           InStr(1, contratoLower, "rematr" & ChrW(237) & "cula") > 0 Then
            ehMatriculaOuRematricula = True
        End If
    End If
    
    ' FORCE HISTORY LOGIC for Matricula/Rematricula even if book didn't change (e.g. imported student first save)
    ' User said: "se o aluno ele é importado... salvar tudo isso, já vai para o histórico. O que aconteceu? Se é matrícula..."
    ' So if it's Matricula/Rematricula, we should check if we need to log it.
    ' We can use a heuristic: If History is empty? Or just always if it's a save action?
    ' "Quando eu salvar tudo isso...".
    ' Let's rely on ehMatriculaOuRematricula.
    ' But we need to distinguish "Update" vs "New/Imported First Save".
    ' If it's "Imported" (meaning we just defaulted fields), we should log.
    ' How to know if we just defaulted fields?
    ' Maybe just relax the "livroMudou" condition if it's Matricula/Rematricula AND History is Empty?
    ' User said: "Se não tiver um livro anterior registrado ali no histórico... palavra desconhecido".
    
    Dim historicoVazio As Boolean: historicoVazio = (lstHistorico.ListCount = 0)
    Dim deveRegistrarHistorico As Boolean
    deveRegistrarHistorico = (livroMudou And ehMatriculaOuRematricula) Or (ehMatriculaOuRematricula And historicoVazio)
    
    ' Se deve registrar -> confirmacao
    If deveRegistrarHistorico Then
        Dim livroAntigoNome As String: livroAntigoNome = NomeLivroPorID(mLivroAntigoID)
        Dim livroNovoNome As String: livroNovoNome = NomeLivroPorID(mIDLivroSelecionado)
        If Len(livroAntigoNome) = 0 Then livroAntigoNome = "(nenhum)"
        If Len(livroNovoNome) = 0 Then livroNovoNome = "(nenhum)"
        
        Dim dataStr As String
        If Len(Trim(txtData.Value)) > 0 Then dataStr = txtData.Value Else dataStr = Format(Date, "dd/mm/yyyy")
        
        Dim msgConf As String
        msgConf = "Confirma a atualizacao do aluno?" & vbCrLf & vbCrLf & _
                  "Aluno: " & Trim(txtNome.Value) & vbCrLf & _
                  "Livro: " & livroAntigoNome & "  " & ChrW(8594) & "  " & livroNovoNome & vbCrLf & _
                  "Contrato: " & nomeContratoAtual & vbCrLf & _
                  "Data: " & dataStr & vbCrLf & vbCrLf & _
                  "Um registro sera adicionado automaticamente ao Historico."
        
        If MsgBox(msgConf, vbQuestion + vbYesNo, "Confirmar Alteracao") = vbNo Then
            Feedback "Alteracao cancelada.", False: Exit Sub
        End If
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
    
    ' Validacao Active State antes de gravar
    If optInativo.Value Then
        If cmbStatus.ListIndex >= 0 Then
            Dim stTx As String: stTx = LCase(Trim(cmbStatus.List(cmbStatus.ListIndex, 1)))
            If stTx = "ativo" Then
                MsgBox "O aluno está com status 'Ativo'." & vbCrLf & _
                       "Para desativar o cadastro, altere o status para 'Desistente' ou 'Trancado'.", vbExclamation, "Validação"
                Exit Sub
            End If
        End If
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
    ' Funcionario/Professor N:N (grava em BD_Vinculo_Professor)
    SalvarFuncionariosAluno CLng(idStr)
    
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
    ws.Cells(linhaGravar, 12).Value = optAtivo.Value
    
    ' ============================================================
    ' ============================================================
    ' AUTO-HISTORICO: inserir evento
    ' ============================================================
    If deveRegistrarHistorico Then
        Dim idTipoOcorrencia As Long
        idTipoOcorrencia = BuscarTipoOcorrenciaPorNome(nomeContratoAtual)
        
        Dim dataEvento As Date
        If Len(Trim(txtData.Value)) > 0 And IsDate(txtData.Value) Then
            dataEvento = CDate(txtData.Value)
        Else
            dataEvento = Date
        End If
        
        Dim obsAuto As String
        Dim livAntNome As String: livAntNome = NomeLivroPorID(mLivroAntigoID)
        Dim livNovNome As String: livNovNome = NomeLivroPorID(mIDLivroSelecionado)
        
        ' Validacao: so mostrar "de -> para" se aluno tinha livro anterior
        
        Dim isMatricula As Boolean
        isMatricula = (InStr(1, LCase(nomeContratoAtual), "rematr") = 0)
        
        If isMatricula Then
             obsAuto = livNovNome
        Else
             ' Rematricula
             If Len(livAntNome) > 0 Then
                 obsAuto = "De " & livAntNome & " " & ChrW(8594) & " " & livNovNome
             Else
                 obsAuto = "De (desconhecido) " & ChrW(8594) & " " & livNovNome
             End If
        End If
        
        InserirHistoricoAuto CLng(idStr), mIDLivroSelecionado, idTipoOcorrencia, dataEvento, obsAuto
        
        ' === AUTO ENTREGA DE MATERIAL ===
        If Not IsEmpty(mIDLivroSelecionado) Then
            Dim idTipoEntrega As Long
            idTipoEntrega = BuscarTipoOcorrenciaPorNome("Entrega de Material")
            If idTipoEntrega > 0 Then
                Dim obsEntrega As String
                obsEntrega = livNovNome
                InserirHistoricoAuto CLng(idStr), mIDLivroSelecionado, idTipoEntrega, dataEvento, obsEntrega
            Else
                MsgBox "Tipo 'Entrega de Material' nao encontrado em BD_TipoOcorrencia." & vbCrLf & _
                       "Adicione este tipo para que o lancamento automatico funcione.", _
                       vbExclamation, "Wizped Office"
            End If
        End If
    End If
    
    ' Atualizar estado
    mLivroAntigoID = mIDLivroSelecionado
    mEditando = True: mLinhaAtual = linhaGravar: txtID.Enabled = False
    mFormModificado = False
    
    If deveRegistrarHistorico Then
        CarregarHistorico CLng(idStr)
        Feedback "Aluno salvo + historico registrado automaticamente!", False
    Else
        Feedback "Aluno salvo! (linha " & linhaGravar & ")", False
    End If
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
    
    ' Validacao Active State
    Dim ativoVal As Boolean: ativoVal = True
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Alunos")
    If Not IsEmpty(ws.Cells(mLinhaAtual, 12).Value) Then ativoVal = CBool(ws.Cells(mLinhaAtual, 12).Value)
    
    If ativoVal Then
        MsgBox "Este aluno está com cadastro ATIVO." & vbCrLf & _
               "Para excluir, primeiro DESATIVE o cadastro.", vbExclamation, "Segurança"
        Exit Sub
    End If
    
    If MsgBox("ATENCAO: O cadastro está desativado." & vbCrLf & _
              "Deseja EXCLUIR DEFINITIVAMENTE todos os dados?" & vbCrLf & _
              "(Histórico, Agenda, Vínculos serão apagados)", vbCritical + vbYesNo, "Confirmar Exclusão") = vbNo Then Exit Sub
    
    ' Cascata: Remover horarios
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("BD_Agenda")
    Dim idAluno As String: idAluno = CStr(txtID.Value)
    Dim rr As Long
    For rr = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If CStr(wsA.Cells(rr, 2).Value) = idAluno Then wsA.Rows(rr).Delete
    Next rr
    
    ' Cascata: Remover vinculos de professor
    Dim wsVP As Worksheet: Set wsVP = ThisWorkbook.Sheets("BD_Vinculo_Professor")
    For rr = wsVP.Cells(wsVP.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If CStr(wsVP.Cells(rr, 2).Value) = idAluno Then wsVP.Rows(rr).Delete
    Next rr
    
    ' Cascata: Remover historico
    Dim wsH As Worksheet: Set wsH = ThisWorkbook.Sheets("BD_Historico")
    For rr = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If CStr(wsH.Cells(rr, 2).Value) = idAluno Then wsH.Rows(rr).Delete
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
    
    Dim nl As Long: nl = ProximaLinhaVazia(wsA, 1)
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
    
    mAgendaEditandoID = CLng(lstAgenda.List(lstAgenda.ListIndex, 0))
    
    ' Colunas: 1=Dia, 2=Hora
    Dim diaVal As String: diaVal = lstAgenda.List(lstAgenda.ListIndex, 1)
    Dim horaVal As String: horaVal = lstAgenda.List(lstAgenda.ListIndex, 2)
    
    ' 1. Setar Dia SEM disparar cmbDia_Change (que limparia cmbHora)
    mSuprimirDiaChange = True
    Dim i As Long
    For i = 0 To cmbDia.ListCount - 1
        If CStr(cmbDia.List(i)) = diaVal Then cmbDia.ListIndex = i: Exit For
    Next i
    mSuprimirDiaChange = False
    
    ' 2. Popular horas manualmente para este dia
    cmbHora.Clear
    If cmbDia.ListIndex >= 0 Then
        Dim horas As Variant
        horas = GetHorariosDisponiveis(cmbDia.Value)
        Dim h As Variant
        For Each h In horas
            If Left(CStr(h), 4) <> "Erro" And Left(CStr(h), 6) <> "Nenhum" Then
                cmbHora.AddItem CStr(h)
            End If
        Next h
    End If
    
    ' 3. Selecionar a hora correta
    For i = 0 To cmbHora.ListCount - 1
        If CStr(cmbHora.List(i)) = horaVal Then cmbHora.ListIndex = i: Exit For
    Next i
    
    Feedback "Editando horario. Altere Dia/Hora e clique + Adicionar.", False
End Sub

' ===========================================================
'  PREVIEW DO TIPO (tempo real)
' ===========================================================

' NOTA: cmbExperiencia_Change, cmbModalidade_Change e chkVIP_Change
' estao na secao DIRTY FLAG abaixo (combinados com AtualizarTipoPreview)

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
    If mFormModificado Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Dados nao salvos. Salvar antes de fechar?", vbQuestion + vbYesNoCancel, "Wizped")
        If resp = vbYes Then
            btnSalvar_Click
            If mFormModificado Then Exit Sub  ' Se salvar falhou (validacao)
        ElseIf resp = vbCancel Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then  ' Fechou pelo X
        If mFormModificado Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("Dados nao salvos. Salvar antes de fechar?", vbQuestion + vbYesNoCancel, "Wizped")
            If resp = vbYes Then
                btnSalvar_Click
                If mFormModificado Then Cancel = 1: Exit Sub
            ElseIf resp = vbCancel Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
End Sub

' ===========================================================
'  UTILITARIOS
' ===========================================================

Private Sub LimparForm()
    mEditando = False: mLinhaAtual = 0
    mIDLivroSelecionado = Empty: mLivroAntigoID = Empty: mAgendaEditandoID = 0
    mHistoricoEditandoID = 0: mFormModificado = False
    mSuprimirDiaChange = False: mSuprimirContratoChange = False: mSuprimirDataFormat = False
    
    txtID.Value = "": txtID.Enabled = True
    txtNome.Value = "": txtData.Value = "": txtObs.Value = ""
    optAtivo.Value = True
    VerificarBloqueioAtivo
    
    mSuprimirBusca = True: txtBusca.Value = "": mSuprimirBusca = False
    mSuprimirLivro = True: txtLivro.Value = "": mSuprimirLivro = False
    
    cmbExperiencia.ListIndex = -1: cmbModalidade.ListIndex = -1
    cmbStatus.ListIndex = -1: cmbContrato.ListIndex = -1
    cmbDia.ListIndex = -1: cmbHora.Clear
    
    ' Deselecionar todos os funcionarios
    Dim i As Long
    For i = 0 To lstFuncionarios.ListCount - 1
        lstFuncionarios.Selected(i) = False
    Next i
    
    chkVIP.Value = False
    lstAgenda.Clear
    lstSugestoes.Clear: lstSugestoes.Visible = False
    lstLivroSugestoes.Clear: lstLivroSugestoes.Visible = False
    lblTipoPreview.Caption = "": lblFeedback.Caption = ""
    
    ' Historico
    lstHistorico.Clear
    cmbTipoOcorrencia.ListIndex = -1: cmbLivroHist.ListIndex = -1
    cmbResponsavel.ListIndex = -1: txtObsHist.Value = ""
    txtDataHist.Value = "" ' Format(Now, "dd/mm/yyyy hh:mm")

    mBackspacing = False
    lblFeedbackHist.Caption = ""
    mHistoricoEditandoID = 0
    btnAddHist.Caption = "+ Registrar"
End Sub

Private Sub FecharOverlays()
    lstSugestoes.Visible = False
    lstLivroSugestoes.Visible = False
End Sub

' === FECHAR OVERLAYS AO FOCAR OUTROS CONTROLES ===
Private Sub txtNome_Enter(): FecharOverlays: End Sub
Private Sub txtID_Enter(): FecharOverlays: End Sub
Private Sub cmbExperiencia_Enter(): FecharOverlays: End Sub
Private Sub cmbModalidade_Enter(): FecharOverlays: End Sub
Private Sub cmbStatus_Enter(): FecharOverlays: End Sub
Private Sub cmbContrato_Enter(): FecharOverlays: End Sub
Private Sub txtData_Enter(): FecharOverlays: End Sub
Private Sub txtObs_Enter(): FecharOverlays: End Sub
Private Sub chkVIP_Enter(): FecharOverlays: End Sub
Private Sub lstFuncionarios_Enter(): FecharOverlays: End Sub
Private Sub lstAgenda_Enter(): FecharOverlays: End Sub
Private Sub cmbDia_Enter(): FecharOverlays: End Sub
Private Sub cmbHora_Enter(): FecharOverlays: End Sub
Private Sub btnAddHora_Enter(): FecharOverlays: End Sub
Private Sub lstHistorico_Enter(): FecharOverlays: End Sub
Private Sub cmbResponsavel_Enter(): FecharOverlays: End Sub
Private Sub cmbTipoOcorrencia_Enter(): FecharOverlays: End Sub
Private Sub cmbLivroHist_Enter(): FecharOverlays: End Sub
Private Sub txtDataHist_Enter(): FecharOverlays: End Sub
Private Sub txtObsHist_Enter(): FecharOverlays: End Sub

' === DIRTY FLAG: marcar formulario como modificado ===

Private Sub txtObs_Change(): mFormModificado = True: End Sub
Private Sub chkVIP_Change(): mFormModificado = True: AtualizarTipoPreview: End Sub
Private Sub cmbExperiencia_Change(): mFormModificado = True: AtualizarTipoPreview: End Sub
Private Sub cmbModalidade_Change(): mFormModificado = True: AtualizarTipoPreview: End Sub
Private Sub cmbStatus_Change(): mFormModificado = True: End Sub

' === CONTRATO: auto-data ao mudar para Matricula/Rematricula ===
Private Sub cmbContrato_Change()
    mFormModificado = True
    If mSuprimirContratoChange Then Exit Sub
    If cmbContrato.ListIndex < 0 Then Exit Sub
    
    Dim nomeContrato As String: nomeContrato = LCase(CStr(cmbContrato.List(cmbContrato.ListIndex, 1)))
    If InStr(1, nomeContrato, "matricula") > 0 Or _
       InStr(1, nomeContrato, "rematricula") > 0 Or _
       InStr(1, nomeContrato, "matr" & ChrW(237) & "cula") > 0 Or _
       InStr(1, nomeContrato, "rematr" & ChrW(237) & "cula") > 0 Then
        mSuprimirDataFormat = True
        txtData.Value = Format(Date, "dd/mm/yyyy")
        mSuprimirDataFormat = False
    End If
End Sub

' === AUTO-FORMATACAO DE DATA (barras automaticas) ===
Private Sub AutoFormatarData(txt As MSForms.TextBox)
    If mBackspacing Then Exit Sub
    
    Dim s As String: s = txt.Value
    Dim oldSelStart As Long: oldSelStart = txt.SelStart
    Dim oldLen As Long: oldLen = Len(s)
    
    ' Remover tudo que nao seja digito
    Dim digits As String: digits = ""
    Dim c As String
    Dim ii As Long
    For ii = 1 To Len(s)
        c = Mid(s, ii, 1)
        If c >= "0" And c <= "9" Then digits = digits & c
    Next ii
    
    ' Limitar a 12 digitos (ddmmyyyyhhmm)
    If Len(digits) > 12 Then digits = Left(digits, 12)
    
    ' Montar com barras e hora
    Dim resultado As String: resultado = ""
    If Len(digits) >= 1 Then resultado = Left(digits, IIf(Len(digits) >= 2, 2, Len(digits)))
    If Len(digits) >= 3 Then resultado = resultado & "/" & Mid(digits, 3, IIf(Len(digits) >= 4, 2, Len(digits) - 2))
    If Len(digits) >= 5 Then resultado = resultado & "/" & Mid(digits, 5, IIf(Len(digits) >= 8, 4, Len(digits) - 4))
    
    ' Hora (apos 8 digitos de data)
    If Len(digits) >= 9 Then
        resultado = resultado & " " & Mid(digits, 9, IIf(Len(digits) >= 10, 2, Len(digits) - 8))
    End If
    If Len(digits) >= 11 Then
        resultado = resultado & ":" & Mid(digits, 11)
    End If
    
    If resultado <> s Then
        mSuprimirDataFormat = True
        txt.Value = resultado
        
        ' Restaurar posicao do cursor considerando caracteres inseridos/removidos
        Dim newLen As Long: newLen = Len(resultado)
        Dim newSelStart As Long
        newSelStart = oldSelStart + (newLen - oldLen)
        
        If newSelStart < 0 Then newSelStart = 0
        If newSelStart > newLen Then newSelStart = newLen
        
        txt.SelStart = newSelStart
        mSuprimirDataFormat = False
    End If
End Sub


Private Sub txtData_Change()
    mFormModificado = True
    If mSuprimirDataFormat Then Exit Sub
    AutoFormatarData txtData
End Sub

Private Sub txtDataHist_Change()
    If mSuprimirDataFormat Then Exit Sub
    AutoFormatarData txtDataHist
End Sub

Private Sub txtDataHist_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mBackspacing = (KeyCode = 8) ' Backspace
End Sub

Private Sub SelecionarCombo(cmb As MSForms.ComboBox, valor As Variant)
    If IsEmpty(valor) Or IsNull(valor) Then cmb.ListIndex = -1: Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If CStr(cmb.List(i, 0)) = CStr(valor) Then cmb.ListIndex = i: Exit Sub
    Next i
    cmb.ListIndex = -1
End Sub

' Converte Variant para String de forma segura (Null/Empty -> "")
Private Function SafeStr(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

' Busca nome do livro pelo ID em BD_Livros
Private Function NomeLivroPorID(idLivro As Variant) As String
    NomeLivroPorID = ""
    If IsEmpty(idLivro) Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Livros")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CLng(ws.Cells(r, 1).Value) = CLng(idLivro) Then
            NomeLivroPorID = CStr(ws.Cells(r, 2).Value): Exit Function
        End If
    Next r
End Function

' Busca ID da TipoOcorrencia por nome (accent-insensitive)
Private Function BuscarTipoOcorrenciaPorNome(nome As String) As Long
    BuscarTipoOcorrenciaPorNome = 0
    If Len(nome) = 0 Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_TipoOcorrencia")
    Dim nomeNorm As String: nomeNorm = LCase(RemoverAcentos(nome))
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LCase(RemoverAcentos(CStr(ws.Cells(r, 2).Value))) = nomeNorm Then
            BuscarTipoOcorrenciaPorNome = CLng(ws.Cells(r, 1).Value): Exit Function
        End If
    Next r
End Function

' Insere registro automatico no BD_Historico
Private Sub InserirHistoricoAuto(idAluno As Long, idLivro As Variant, idTipo As Long, dataEvento As Date, obs As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Historico")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Proximo ID
    Dim maxID As Long: maxID = 0
    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, 1).Value > maxID Then maxID = ws.Cells(r, 1).Value
    Next r
    
    Dim nl As Long: nl = ProximaLinhaVazia(ws, 1)
    ws.Cells(nl, 1).Value = maxID + 1
    ws.Cells(nl, 2).Value = idAluno
    
    If Not IsEmpty(idLivro) Then
        ws.Cells(nl, 3).Value = CLng(idLivro)
    End If
    
    ws.Cells(nl, 4).Value = idTipo
    ws.Cells(nl, 5).Value = Now             ' timestamp com hora
    ws.Cells(nl, 5).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(nl, 6).Value = obs
    ' col 7 = ID_Funcionario (vazio em auto — preenchido via UI)
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
        Dim altItem As Single: altItem = lst.Font.Size + 3
        Dim alt As Single: alt = contItens * altItem
        If alt < 32 Then alt = 32
        If alt > 220 Then alt = 220
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

' Encontra a proxima linha vazia para inserir dados.
' Trata o caso de tabelas novas onde row 2 existe porém está vazia.
Private Function ProximaLinhaVazia(ws As Worksheet, col As Long) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    If lastRow <= 1 Then
        ' Somente header ou planilha vazia
        ProximaLinhaVazia = 2
        Exit Function
    End If
    
    ' Checar se row 2 (primeira linha de dados) esta vazia
    If IsEmpty(ws.Cells(2, col).Value) Or Len(Trim(CStr(ws.Cells(2, col).Value))) = 0 Then
        ProximaLinhaVazia = 2
    Else
        ProximaLinhaVazia = lastRow + 1
    End If
End Function

' ===========================================================
'  FUNCIONARIOS N:N (BD_Vinculo_Professor)
' ===========================================================

' Carrega e seleciona os funcionarios vinculados ao aluno
Private Sub CarregarFuncionariosAluno(idAluno As Variant)
    If IsEmpty(idAluno) Then Exit Sub
    
    Dim wsVP As Worksheet: Set wsVP = ThisWorkbook.Sheets("BD_Vinculo_Professor")
    Dim lastRow As Long: lastRow = wsVP.Cells(wsVP.Rows.Count, 1).End(xlUp).Row
    
    ' Deselecionar todos
    Dim i As Long
    For i = 0 To lstFuncionarios.ListCount - 1
        lstFuncionarios.Selected(i) = False
    Next i
    
    ' Selecionar os vinculados
    Dim r As Long
    For r = 2 To lastRow
        If CStr(wsVP.Cells(r, 2).Value) = CStr(idAluno) Then
            Dim idFunc As Long: idFunc = CLng(wsVP.Cells(r, 3).Value)
            For i = 0 To lstFuncionarios.ListCount - 1
                If CLng(lstFuncionarios.List(i, 0)) = idFunc Then
                    lstFuncionarios.Selected(i) = True: Exit For
                End If
            Next i
        End If
    Next r
End Sub

' Salva os funcionarios selecionados (delete + insert)
Private Sub SalvarFuncionariosAluno(idAluno As Long)
    Dim wsVP As Worksheet: Set wsVP = ThisWorkbook.Sheets("BD_Vinculo_Professor")
    
    ' 1. Deletar vinculos atuais (reverso)
    Dim lastRow As Long: lastRow = wsVP.Cells(wsVP.Rows.Count, 1).End(xlUp).Row
    Dim rr As Long
    For rr = lastRow To 2 Step -1
        If Not IsEmpty(wsVP.Cells(rr, 2).Value) Then
            If CLng(wsVP.Cells(rr, 2).Value) = idAluno Then wsVP.Rows(rr).Delete
        End If
    Next rr
    
    ' 2. Encontrar proximo ID
    lastRow = wsVP.Cells(wsVP.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    Dim maxID As Long: maxID = 0
    For rr = 2 To lastRow
        If Not IsEmpty(wsVP.Cells(rr, 1).Value) Then
            If wsVP.Cells(rr, 1).Value > maxID Then maxID = CLng(wsVP.Cells(rr, 1).Value)
        End If
    Next rr
    
    ' 3. Inserir novos vinculos + concatenar nomes para col denormalizada
    Dim nextRow As Long: nextRow = lastRow + 1
    If lastRow = 1 And IsEmpty(wsVP.Cells(2, 1).Value) Then nextRow = 2
    Dim nomesConcatenados As String: nomesConcatenados = ""
    Dim i As Long
    For i = 0 To lstFuncionarios.ListCount - 1
        If lstFuncionarios.Selected(i) Then
            maxID = maxID + 1
            wsVP.Cells(nextRow, 1).Value = maxID
            wsVP.Cells(nextRow, 2).Value = idAluno
            wsVP.Cells(nextRow, 3).Value = CLng(lstFuncionarios.List(i, 0))
            nextRow = nextRow + 1
            ' Concatenar nome do funcionario
            If Len(nomesConcatenados) > 0 Then nomesConcatenados = nomesConcatenados & ", "
            nomesConcatenados = nomesConcatenados & lstFuncionarios.List(i, 1)
        End If
    Next i
    
    ' 4. Atualizar coluna denormalizada em BD_Alunos (col 9 = "Professores")
    Dim wsAlunos As Worksheet: Set wsAlunos = ThisWorkbook.Sheets("BD_Alunos")
    Dim rAluno As Long
    For rAluno = 2 To wsAlunos.Cells(wsAlunos.Rows.Count, 1).End(xlUp).Row
        If CLng(wsAlunos.Cells(rAluno, 1).Value) = idAluno Then
            wsAlunos.Cells(rAluno, 9).Value = nomesConcatenados
            Exit For
        End If
    Next rAluno
End Sub

' ===========================================================
'  HISTORICO (BD_Historico + BD_TipoOcorrencia)
' ===========================================================

Private Sub CarregarHistorico(idAluno As Variant)
    lstHistorico.Clear
    lstHistorico.ColumnCount = 6 ' ID, Data, Hora, Evento, Detalhes, Responsavel
    ' Restaurando alinhamento com Labels:
    ' Data (60) + Hora (60) = 120 (Total do Header)
    lstHistorico.ColumnWidths = "0 pt;60 pt;60 pt;120 pt;340 pt;120 pt"
    mHistoricoEditandoID = 0
    btnAddHist.Caption = "+ Registrar"
    If IsEmpty(idAluno) Then Exit Sub
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Historico")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim wsT As Worksheet: Set wsT = ThisWorkbook.Sheets("BD_TipoOcorrencia")
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("BD_Funcionarios")
    
    ' 1. Coletar linhas deste aluno num array temporario
    Dim count As Long: count = 0
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).Value) = CStr(idAluno) Then count = count + 1
    Next r
    If count = 0 Then Exit Sub
    
    ' Array:
    ' (1)=row
    ' (2)=TimestampParaOrdenacao (Double)
    ' (3)=OriginalTimestamp (Double)
    ' (4)=IsChild (Boolean) - Se True, eh filho de Matricula/Contrato
    ' (5)=NomeTipo (String) - Para identificar Matricula/Entrega
    Dim arr() As Variant: ReDim arr(1 To count, 1 To 5)
    Dim k As Long: k = 0
    
    ' Cache Tipos de Ocorrencia para performance e identificar nomes
    ' Dictionary seria ideal, mas array simples serve
    Dim lastRowT As Long: lastRowT = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).Value) = CStr(idAluno) Then
            k = k + 1
            arr(k, 1) = r
            
            Dim ts As Double: ts = 0
            If IsDate(ws.Cells(r, 5).Value) Then ts = CDbl(CDate(ws.Cells(r, 5).Value))
            arr(k, 2) = ts 
            arr(k, 3) = ts
            arr(k, 4) = False ' Default Parent
            
            ' Pegar Nome do Tipo
            Dim idTipo As Long: idTipo = CLng(ws.Cells(r, 4).Value)
            Dim nomeTipo As String: nomeTipo = ""
            Dim rt As Long
            For rt = 2 To lastRowT
                If CLng(wsT.Cells(rt, 1).Value) = idTipo Then
                    nomeTipo = LCase(wsT.Cells(rt, 2).Value): Exit For
                End If
            Next rt
            arr(k, 5) = nomeTipo
        End If
    Next r
    
    ' 2. Agrupamento (TreeView Logic)
    ' Se houver "Matricula" e "Entrega de Material" no MESMO DIA:
    ' Entrega vira Child, SortTimestamp = Timestamp da Matricula (para ficarem juntos)
    ' Se tiver varias matriculas no mesmo dia? Assume a ultima? Logica simples: Matchear pares.
    Dim i As Long, j As Long
    For i = 1 To count
        ' Se eh Entrega, procurar Matricula pai
        If InStr(arr(i, 5), "entrega") > 0 Then
            Dim tsEntrega As Double: tsEntrega = arr(i, 3)
            Dim diaEntrega As Long: diaEntrega = Int(tsEntrega)
            
            ' Procurar Matricula no mesmo dia
            For j = 1 To count
                If i <> j Then
                    If InStr(arr(j, 5), "matr" & ChrW(237) & "cula") > 0 Or InStr(arr(j, 5), "matricula") > 0 Or InStr(arr(j, 5), "contrato") > 0 Then
                        Dim tsMatr As Double: tsMatr = arr(j, 3)
                        If Int(tsMatr) = diaEntrega Then
                            ' Encontrou pai!
                            arr(i, 2) = tsMatr ' Herda timestamp do pai para ordenar junto
                            arr(i, 4) = True   ' Marca como filho (indentar)
                            Exit For ' Acha o primeiro (ou ultimo?) pai do dia. Assume 1 matricula/dia.
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
    ' 3. Bubble sort Customizado
    ' Criteria: SortTimestamp DESC.
    ' Tie-breaker: Parent antes de Child (se mesmo SortTimestamp).
    Dim tmp As Variant
    For i = 1 To count - 1
        For j = 1 To count - i
            Dim swap As Boolean: swap = False
            
            ' Sort DESC por Timestamp Agrupado
            If arr(j, 2) < arr(j + 1, 2) Then
                swap = True
            ElseIf arr(j, 2) = arr(j + 1, 2) Then
                ' Empate: Parent (IsChild=False) deve vir antes de Child (IsChild=True)
                ' Como eh DESC, "Maior" vem primeiro. Parent deve ser "Maior".
                ' Entao IsChild=False > IsChild=True?
                ' Se arr(j) eh Child e arr(j+1) eh Parent -> Swap para subir Parent
                If arr(j, 4) And Not arr(j + 1, 4) Then
                    swap = True
                End If
            End If
            
            If swap Then
                ' Trocar todas as colunas
                Dim c As Integer
                For c = 1 To 5
                    tmp = arr(j, c): arr(j, c) = arr(j + 1, c): arr(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
    
    ' 4. Preencher listbox
    For k = 1 To count
        r = CLng(arr(k, 1))
        Dim isChild As Boolean: isChild = arr(k, 4)
        
        lstHistorico.AddItem
        Dim idx As Long: idx = lstHistorico.ListCount - 1
        
        ' Col 0: ID
        lstHistorico.List(idx, 0) = ws.Cells(r, 1).Value
        
        ' Col 1: Data (dd/mm/yyyy) ou L (Visual tree)
        ' Col 2: Hora (hh:mm) - Alinhamento garantido por coluna separada
        Dim dtVal As Variant: dtVal = ws.Cells(r, 5).Value
        If IsDate(dtVal) Then
            If isChild Then
                ' Visual Tree para filho: L + 6 tracos na Coluna de Data
                lstHistorico.List(idx, 1) = ChrW(9492) & String(6, ChrW(9472))
                ' Hora na Coluna de Hora (Col 2)
                lstHistorico.List(idx, 2) = Format(dtVal, "hh:mm")
            Else
                ' Pai: Data completa na Col 1 e Hora na Col 2 (opcional, ou tudo na 1?)
                ' Melhor separar tambem para alinhar as horas dos pais e filhos
                lstHistorico.List(idx, 1) = Format(dtVal, "dd/mm/yyyy")
                lstHistorico.List(idx, 2) = Format(dtVal, "hh:mm")
            End If
        Else
            lstHistorico.List(idx, 1) = CStr(dtVal)
            lstHistorico.List(idx, 2) = ""
        End If
        
        ' Col 3: Evento (tipo)
        Dim nomeEvento As String: nomeEvento = ""
        ' Buscar nome correto (com casing original da tabela)
        Dim idT As Long: idT = CLng(ws.Cells(r, 4).Value)
        For rt = 2 To lastRowT
            If CLng(wsT.Cells(rt, 1).Value) = idT Then
                nomeEvento = wsT.Cells(rt, 2).Value: Exit For
            End If
        Next rt
        
        ' SEM indentacao no nome do evento, alinhado a esquerda
        lstHistorico.List(idx, 3) = nomeEvento

        
        ' Col 4: Detalhes
        lstHistorico.List(idx, 4) = IIf(IsEmpty(ws.Cells(r, 6).Value), "", CStr(ws.Cells(r, 6).Value))

        
        ' Col 5: Responsavel
        Dim idFunc As Variant: idFunc = ws.Cells(r, 7).Value
        If Not IsEmpty(idFunc) Then
            Dim rf As Long
            For rf = 2 To wsF.Cells(wsF.Rows.Count, 1).End(xlUp).Row
                If CStr(wsF.Cells(rf, 1).Value) = CStr(idFunc) Then
                    lstHistorico.List(idx, 5) = wsF.Cells(rf, 2).Value: Exit For
                End If
            Next rf
        End If
    Next k
    
    ' Limpar default date field (UI Requirement)
    txtDataHist.Value = ""
End Sub


' DblClick no historico: preenche campos para editar
Private Sub lstHistorico_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstHistorico.ListIndex = -1 Then Exit Sub
    Dim idx As Long: idx = lstHistorico.ListIndex
    
    mHistoricoEditandoID = CLng(lstHistorico.List(idx, 0))
    
    ' Tipo (col 3 = Evento) - Era Col 2
    Dim tipoNome As String: tipoNome = SafeStr(lstHistorico.List(idx, 3))

    Dim i As Long
    cmbTipoOcorrencia.ListIndex = -1
    For i = 0 To cmbTipoOcorrencia.ListCount - 1
        If CStr(cmbTipoOcorrencia.List(i, 1)) = tipoNome Then
            cmbTipoOcorrencia.ListIndex = i: Exit For
        End If
    Next i
    
    ' Obs (col 4 = Detalhes) - Era Col 3
    txtObsHist.Value = SafeStr(lstHistorico.List(idx, 4))

    
    ' Responsavel (col 5) - Era Col 4
    Dim responsavelNome As String: responsavelNome = SafeStr(lstHistorico.List(idx, 5))

    cmbResponsavel.ListIndex = -1
    If Len(responsavelNome) > 0 Then
        For i = 0 To cmbResponsavel.ListCount - 1
            If CStr(cmbResponsavel.List(i, 1)) = responsavelNome Then
                cmbResponsavel.ListIndex = i: Exit For
            End If
        Next i
    End If
    
    ' Data (col 1) -> Buscar da planilha (para evitar pegar o texto formatado com L)
    ' txtDataHist.Value sera preenchido no loop abaixo

    
    ' Livro (nao esta na listbox, buscar na planilha)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Historico")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CLng(ws.Cells(r, 1).Value) = mHistoricoEditandoID Then
            Dim idLivro As Variant: idLivro = ws.Cells(r, 3).Value
             SelecionarCombo cmbLivroHist, idLivro
            
            ' Carregar Data original da celula
            If IsDate(ws.Cells(r, 5).Value) Then
                txtDataHist.Value = Format(ws.Cells(r, 5).Value, "dd/mm/yyyy hh:mm")
            Else
                txtDataHist.Value = CStr(ws.Cells(r, 5).Value)
            End If
            
            Exit For

        End If
    Next r
    
    btnAddHist.Caption = "Atualizar"
    FeedbackHist "Editando evento. Clique 'Atualizar' para salvar.", False
End Sub

Private Sub btnAddHist_Click()
    If Len(Trim(txtID.Value)) = 0 Then
        FeedbackHist "Carregue um aluno primeiro.", True: Exit Sub
    End If
    If cmbTipoOcorrencia.ListIndex = -1 Then
        FeedbackHist "Selecione o tipo.", True: Exit Sub
    End If
    
    ' Valida Data (Estrita: Dias, Meses, Ano)
    If Not ValidarDataEstrita(txtDataHist) Then Exit Sub
    
    ' Valida Logica (Entrega vs Matricula)
    Dim dataNova As Date: dataNova = CDate(txtDataHist.Value)
    Dim tipoNovo As String: tipoNovo = cmbTipoOcorrencia.List(cmbTipoOcorrencia.ListIndex, 1)
    
    If Not ValidarSequenciaLogica(CLng(txtID.Value), dataNova, tipoNovo) Then Exit Sub

    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Historico")
    
    ' === MODO EDITAR ===
    If mHistoricoEditandoID > 0 Then
        Dim rr As Long
        For rr = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CLng(ws.Cells(rr, 1).Value) = mHistoricoEditandoID Then
            ' Livro
            If Not IsEmpty(ValorCombo(cmbLivroHist)) Then
                 ws.Cells(rr, 3).Value = CLng(ValorCombo(cmbLivroHist))
            Else
                 ws.Cells(rr, 3).Value = Empty
            End If
            ' Tipo
            ws.Cells(rr, 4).Value = CLng(cmbTipoOcorrencia.List(cmbTipoOcorrencia.ListIndex, 0))
            ' Data
            If IsDate(txtDataHist.Value) Then
                ws.Cells(rr, 5).Value = CDate(txtDataHist.Value)
                ws.Cells(rr, 5).NumberFormat = "dd/mm/yyyy hh:mm:ss"
            End If
            ' Obs
            ws.Cells(rr, 6).Value = Trim(txtObsHist.Value)
            ' Responsavel
            If cmbResponsavel.ListIndex >= 0 Then
                ws.Cells(rr, 7).Value = CLng(cmbResponsavel.List(cmbResponsavel.ListIndex, 0))
            End If
            Exit For
        End If
        Next rr
        
        mHistoricoEditandoID = 0
        CarregarHistorico CLng(txtID.Value)
        ' Limpar campos Historico (via LimparForm parcial ou manual)
        cmbTipoOcorrencia.ListIndex = -1: cmbLivroHist.ListIndex = -1
        txtObsHist.Value = "": cmbResponsavel.ListIndex = -1
        txtDataHist.Value = "" ' Format(Now, "dd/mm/yyyy hh:mm")

        btnAddHist.Caption = "+ Registrar"
        FeedbackHist "Evento atualizado.", False
        Exit Sub
    End If
    
    ' === MODO ADICIONAR ===
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long: maxID = 0
    Dim rA As Long
    For rA = 2 To lastRow
        If ws.Cells(rA, 1).Value > maxID Then maxID = ws.Cells(rA, 1).Value
    Next rA
    
    Dim nl As Long: nl = ProximaLinhaVazia(ws, 1)
    ws.Cells(nl, 1).Value = maxID + 1
    ws.Cells(nl, 2).Value = CLng(txtID.Value)
    
    ' Livro (da combo de historico, nao do aluno)
    If Not IsEmpty(ValorCombo(cmbLivroHist)) Then
        ws.Cells(nl, 3).Value = CLng(ValorCombo(cmbLivroHist))
    Else
        ' Fallback: se nao selecionou nada no historico, usa o do aluno?
        ' Nao, melhor explicito. Se vazio, vazio.
        ws.Cells(nl, 3).Value = Empty
    End If
    
    ws.Cells(nl, 4).Value = CLng(cmbTipoOcorrencia.List(cmbTipoOcorrencia.ListIndex, 0))
    ' Data manual
    ws.Cells(nl, 5).Value = CDate(txtDataHist.Value)
    ws.Cells(nl, 5).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
    ws.Cells(nl, 6).Value = Trim(txtObsHist.Value)
    ' Responsavel
    If cmbResponsavel.ListIndex >= 0 Then
        ws.Cells(nl, 7).Value = CLng(cmbResponsavel.List(cmbResponsavel.ListIndex, 0))
    End If
    
    CarregarHistorico CLng(txtID.Value)
    
    ' Limpar campos
    cmbTipoOcorrencia.ListIndex = -1: cmbLivroHist.ListIndex = -1
    txtObsHist.Value = "": cmbResponsavel.ListIndex = -1
    txtDataHist.Value = "" ' Format(Now, "dd/mm/yyyy hh:mm")

    
    FeedbackHist "Evento registrado.", False
End Sub

Private Sub btnRemHist_Click()
    If lstHistorico.ListIndex = -1 Then
        FeedbackHist "Selecione um evento para remover.", True: Exit Sub
    End If
    
    Dim idHist As Long: idHist = CLng(lstHistorico.List(lstHistorico.ListIndex, 0))
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Historico")
    Dim rr As Long
    For rr = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CLng(ws.Cells(rr, 1).Value) = idHist Then
            ws.Rows(rr).Delete: Exit For
        End If
    Next rr
    
    CarregarHistorico CLng(txtID.Value)
    FeedbackHist "Evento removido.", False
End Sub

Private Sub FeedbackHist(msg As String, isErro As Boolean)
    lblFeedbackHist.Caption = msg
    lblFeedbackHist.ForeColor = IIf(isErro, &HFF&, &H8000&)
End Sub

' ===========================================================
'  ACTIVE STATE BLOCKING LOGIC
' ===========================================================

Private Sub optInativo_Click()
    ' Regra: O aluno nao pode ser desativado no cadastro (Inativo)
    ' se o Status do Contrato for Ativo (1 = Matriculado).
    ' Necessario alterar o Status para Trancado, Desistente ou Concluido antes.
    
    Dim idStatus As Long: idStatus = -1
    If cmbStatus.ListIndex >= 0 Then
        idStatus = CLng(cmbStatus.List(cmbStatus.ListIndex, 0))
    End If
    
    ' ID 1 = Ativo / Matriculado (conforme padrao em btnNovo_Click e CarregarLookups)
    If idStatus = 1 Then
        MsgBox "O aluno est" & ChrW(225) & " com Status 'Ativo' (ou Matriculado)." & vbCrLf & vbCrLf & _
               "Para desativar o cadastro, altere primeiro o Status para" & vbCrLf & _
               "Trancado, Desistente ou Conclu" & ChrW(237) & "do.", _
               vbExclamation, "Opera" & ChrW(231) & ChrW(227) & "o Inv" & ChrW(225) & "lida"
        
        ' Reverter para Ativo
        optAtivo.Value = True
    End If
End Sub


Private Sub txtID_Change()
    mFormModificado = True
    VerificarBloqueioAtivo
End Sub

Private Sub txtNome_Change()
    mFormModificado = True
    VerificarBloqueioAtivo
End Sub

Private Sub VerificarBloqueioAtivo()
    Dim bloqueado As Boolean: bloqueado = True
    If Len(Trim(txtID.Value)) > 0 And Len(Trim(txtNome.Value)) > 0 Then bloqueado = False
    
    lblBloqueioAtivo.Visible = bloqueado
    frmAtivo.Enabled = Not bloqueado
    ' Force repaint if needed, but not critical
End Sub

Private Sub lblBloqueioAtivo_Click()
    MsgBox "Preencha os campos ID e Nome primeiro.", vbInformation, "Ação Necessária"
End Sub

' ===========================================================
'  VALIDACAO ESTRITA DE DATA E LOGICA
' ===========================================================

Private Function ValidarDataEstrita(txt As MSForms.TextBox) As Boolean
    Dim s As String: s = Trim(txt.Value)
    If Len(s) = 0 Then ValidarDataEstrita = False: Exit Function
    If Not IsDate(s) Then
        MsgBox "Data inv" & ChrW(225) & "lida. Formato: dd/mm/yyyy hh:mm", vbExclamation
        ValidarDataEstrita = False: Exit Function
    End If
    
    ' Parse parts manual para checar dias exatos
    ' Assumindo dd/mm/yyyy hh:mm (mask garante barras e espaco)
    Dim parts() As String: parts = Split(s, " ")
    Dim dataPart As String: dataPart = parts(0)
    Dim dateParts() As String: dateParts = Split(dataPart, "/")
    
    If UBound(dateParts) < 2 Then ValidarDataEstrita = False: Exit Function
    Dim d As Integer: d = CInt(dateParts(0))
    Dim m As Integer: m = CInt(dateParts(1))
    Dim y As Integer: y = CInt(dateParts(2))
    
    ' Checar limites de ano
    ' User: "1002" ou "20060" invalidos. Faixa razoavel: 1900 a 2100.
    If y < 1900 Or y > 2100 Then
        MsgBox "Ano " & y & " inv" & ChrW(225) & "lido no contexto.", vbExclamation
        ValidarDataEstrita = False: Exit Function
    End If
    
    ' Checar limites de mes
    If m < 1 Or m > 12 Then
        MsgBox "M" & ChrW(234) & "s " & m & " inv" & ChrW(225) & "lido.", vbExclamation
        ValidarDataEstrita = False: Exit Function
    End If

    ' Checar dias no mes (considerando bissexto via DateSerial)
    Dim diasNoMes As Integer
    diasNoMes = Day(DateSerial(y, m + 1, 0))
    If d > diasNoMes Then
        MsgBox "O m" & ChrW(234) & "s " & m & "/" & y & " s" & ChrW(243) & " tem " & diasNoMes & " dias.", vbExclamation
        ValidarDataEstrita = False: Exit Function
    End If
    
    ' Checar Futuro (Mes > Mes Atual do Ano Atual)
    ' User: "Se a pessoa colocar mes 3 (e estamos no 2)... alerta: futuro, confirma?"
    ' Comparar data inteira com Now
    Dim dt As Date: dt = CDate(s)
    If dt > Now Then
        If MsgBox("A data informada (" & s & ") " & ChrW(233) & " futura." & vbCrLf & "Confirma?", vbQuestion + vbYesNo) = vbNo Then
            ValidarDataEstrita = False: Exit Function
        End If
    End If
    
    ValidarDataEstrita = True
End Function

Private Function ValidarSequenciaLogica(idAluno As Long, dataNova As Date, tipoNovo As String) As Boolean
    ' Se tipo = Entrega de Material, verificar se existe Matricula ou Contrato ANTERIOR
    If InStr(1, LCase(tipoNovo), "entrega") = 0 Then ValidarSequenciaLogica = True: Exit Function
    
    ' Buscar data de matricula no historico (BD_Historico)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Historico")
    Dim wsT As Worksheet: Set wsT = ThisWorkbook.Sheets("BD_TipoOcorrencia")
    
    Dim dataMatricula As Date: dataMatricula = 0
    Dim encontrouMatricula As Boolean: encontrouMatricula = False
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).Value) = CStr(idAluno) Then
            Dim idTipo As Long: idTipo = CLng(ws.Cells(r, 4).Value)
            
            ' Buscar nome do tipo na tabela Lookups (ou cache se lento, mas ok aqui)
            Dim nomeTipo As String: nomeTipo = ""
            Dim rt As Long
            Dim lastRowT As Long: lastRowT = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
            For rt = 2 To lastRowT
                If CLng(wsT.Cells(rt, 1).Value) = idTipo Then
                    nomeTipo = LCase(wsT.Cells(rt, 2).Value)
                    Exit For
                End If
            Next rt
            
            If InStr(nomeTipo, "matricula") > 0 Or _
               InStr(nomeTipo, "matr" & ChrW(237) & "cula") > 0 Or _
               InStr(nomeTipo, "contrato") > 0 Then
               
                If IsDate(ws.Cells(r, 5).Value) Then
                    Dim d As Date: d = CDate(ws.Cells(r, 5).Value)
                    ' Pega a maior data de matricula encontrada (ultima matricula)
                    If d > dataMatricula Then dataMatricula = d
                    encontrouMatricula = True
                End If
            End If
        End If
    Next r
    
    If encontrouMatricula Then
        ' Se Entrega for ANTES da Matricula (dataNova < dataMatricula)
        ' User: "Entrega nao pode ser antes... saltar alerta"
        If dataNova < dataMatricula Then
             If MsgBox("A 'Entrega de Material' est" & ChrW(225) & " datada ANTES da 'Matr" & ChrW(237) & "cula' encontrada em " & Format(dataMatricula, "dd/mm/yyyy") & "." & vbCrLf & vbCrLf & _
                       "Deseja continuar mesmo assim?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                 ValidarSequenciaLogica = False: Exit Function
             End If
        End If
    End If
    
    ValidarSequenciaLogica = True
End Function

