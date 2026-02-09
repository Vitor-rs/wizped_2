' ===========================================================
' MODULO: CriarFormulario (descartavel)
' Execute CriarFormulario() UMA VEZ, depois delete este modulo.
'
' PRE-REQUISITOS:
'   - "Microsoft Visual Basic for Applications Extensibility 5.3"
'   - "Confiar acesso ao modelo de objeto do projeto VBA"
'
' LAYOUT v5:
'   - Sem frames. Secoes divididas por label titulo + linha
'   - Labels de campo ACIMA dos controles (secao Dados)
'   - Fonte Segoe UI 10 padronizada
'   - Controles com Height=28 (melhor centralização vertical do texto)
'   - Botoes centralizados, flat
'   - Livro com TextBox autocomplete + botao dropdown
'   - Overlays lstSugestoes e lstLivroSugestoes por ultimo (z-order)
' ===========================================================

Sub CriarFormulario()

    Dim VBProj As Object, VBComp As Object, frm As Object, ctrl As Object
    
    Set VBProj = ThisWorkbook.VBProject
    On Error Resume Next
    VBProj.VBComponents.Remove VBProj.VBComponents("frmAlunos")
    On Error GoTo 0
    
    Set VBComp = VBProj.VBComponents.Add(3)
    VBComp.Name = "frmAlunos"
    
    ' --- Dimensoes e cor do form ---
    VBComp.Properties("Caption") = "Wizped Office - Gerenciamento de Alunos"
    VBComp.Properties("Width") = 645
    VBComp.Properties("Height") = 620
    VBComp.Properties("BackColor") = &HFAF8F5
    
    Set frm = VBComp.Designer
    
    ' ===========================================================
    ' CONSTANTES DE ESTILO
    ' ===========================================================
    Const FN As String = "Segoe UI"
    Const FS As Long = 10         ' Fonte padrao (labels e valores)
    Const FS_TITLE As Long = 11   ' Titulo de secao
    Const FS_PREVIEW As Long = 12 ' Preview do tipo
    Const CH As Long = 28         ' Altura de controles (TextBox, ComboBox)
    Const LH As Long = 14         ' Altura de labels
    Const BG As Long = &HFAF8F5   ' Bege claro
    Const GOLD As Long = &H6AA8D5 ' Dourado Wizard (BGR)
    Const WHITE As Long = &HFFFFFF
    Const LINE_COLOR As Long = &HD0D0D0  ' Cinza claro para linhas
    Const RED_BG As Long = &HE0E0FF      ' Rosa claro (BGR) para excluir
    
    Dim y As Single  ' Posicao vertical corrente
    y = 10
    
    ' ===========================================================
    ' ===  SECAO: BUSCAR ALUNO  =================================
    ' ===========================================================
    
    ' Titulo da secao
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblSecaoBusca")
    With ctrl
        .Caption = "Buscar Aluno"
        .Left = 10: .Top = y: .Width = 200: .Height = 16
        .Font.Name = FN: .Font.Size = FS_TITLE: .Font.Bold = True
        .BackStyle = 0
    End With
    y = y + 18
    
    ' Linha separadora (label com h=1)
    Set ctrl = frm.Controls.Add("Forms.Label.1", "linBusca")
    With ctrl
        .Caption = "": .Left = 10: .Top = y: .Width = 620: .Height = 1
        .BackColor = LINE_COLOR: .BackStyle = 1
    End With
    y = y + 8
    
    ' Campo de busca + botoes (mesma linha)
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtBusca")
    With ctrl
        .Left = 10: .Top = y: .Width = 310: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnBuscar")
    With ctrl
        .Caption = "Buscar"
        .Left = 328: .Top = y: .Width = 70: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .BackColor = GOLD
    End With
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnLimpar")
    With ctrl
        .Caption = "Limpar"
        .Left = 404: .Top = y: .Width = 70: .Height = CH
        .Font.Name = FN: .Font.Size = FS
    End With
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnNovo")
    With ctrl
        .Caption = "+ Novo Aluno"
        .Left = 502: .Top = y: .Width = 128: .Height = CH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True
        .BackColor = GOLD
    End With
    
    y = y + CH + 14  ' Espaco antes da proxima secao
    
    ' ===========================================================
    ' ===  SECAO: DADOS DO ALUNO  ===============================
    ' ===========================================================
    
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblSecaoDados")
    With ctrl
        .Caption = "Dados do Aluno"
        .Left = 10: .Top = y: .Width = 200: .Height = 16
        .Font.Name = FN: .Font.Size = FS_TITLE: .Font.Bold = True
        .BackStyle = 0
    End With
    y = y + 18
    
    Set ctrl = frm.Controls.Add("Forms.Label.1", "linDados")
    With ctrl
        .Caption = "": .Left = 10: .Top = y: .Width = 620: .Height = 1
        .BackColor = LINE_COLOR: .BackStyle = 1
    End With
    y = y + 6
    
    ' ----------------------------------------------------------
    ' ROW 1: ID + Nome (labels ACIMA)
    ' ----------------------------------------------------------
    ' Labels
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblID")
    With ctrl
        .Caption = "ID (SponteWeb)"
        .Left = 10: .Top = y: .Width = 95: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblNome")
    With ctrl
        .Caption = "Nome"
        .Left = 115: .Top = y: .Width = 50: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    y = y + LH + 2
    
    ' Controles
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtID")
    With ctrl
        .Left = 10: .Top = y: .Width = 95: .Height = CH: .MaxLength = 5
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtNome")
    With ctrl
        .Left = 115: .Top = y: .Width = 515: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    y = y + CH + 8
    
    ' ----------------------------------------------------------
    ' ROW 2: Livro + Experiencia + VIP
    ' ----------------------------------------------------------
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblLivro")
    With ctrl
        .Caption = "Livro"
        .Left = 10: .Top = y: .Width = 50: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblExperiencia")
    With ctrl
        .Caption = "Experiencia"
        .Left = 252: .Top = y: .Width = 80: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    y = y + LH + 2
    
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtLivro")
    With ctrl
        .Left = 10: .Top = y: .Width = 206: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    ' Botao dropdown do livro
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnLivroDD")
    With ctrl
        .Caption = ChrW(9660)   ' Triangulo para baixo
        .Left = 216: .Top = y: .Width = 24: .Height = CH
        .Font.Name = FN: .Font.Size = 8
    End With
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbExperiencia")
    With ctrl
        .Left = 252: .Top = y: .Width = 148: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;143"
        .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    Set ctrl = frm.Controls.Add("Forms.CheckBox.1", "chkVIP")
    With ctrl
        .Caption = "VIP"
        .Left = 415: .Top = y + 2: .Width = 55: .Height = 20
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    y = y + CH + 8
    
    ' ----------------------------------------------------------
    ' ROW 3: Modalidade + Status + Contrato
    ' ----------------------------------------------------------
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblModalidade")
    With ctrl
        .Caption = "Modalidade"
        .Left = 10: .Top = y: .Width = 72: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblStatus")
    With ctrl
        .Caption = "Status"
        .Left = 170: .Top = y: .Width = 50: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblContrato")
    With ctrl
        .Caption = "Contrato"
        .Left = 345: .Top = y: .Width = 60: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    y = y + LH + 2
    
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbModalidade")
    With ctrl
        .Left = 10: .Top = y: .Width = 148: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;143"
        .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbStatus")
    With ctrl
        .Left = 170: .Top = y: .Width = 163: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;158"
        .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbContrato")
    With ctrl
        .Left = 345: .Top = y: .Width = 163: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;158"
        .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    y = y + CH + 8
    
    ' ----------------------------------------------------------
    ' ROW 4: Professor + Data Inicio
    ' ----------------------------------------------------------
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblProfessor")
    With ctrl
        .Caption = "Professor"
        .Left = 10: .Top = y: .Width = 65: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblData")
    With ctrl
        .Caption = "Data Inicio"
        .Left = 200: .Top = y: .Width = 76: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    y = y + LH + 2
    
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbProfessor")
    With ctrl
        .Left = 10: .Top = y: .Width = 178: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;173"
        .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtData")
    With ctrl
        .Left = 200: .Top = y: .Width = 105: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    y = y + CH + 6
    
    ' ----------------------------------------------------------
    ' ROW 5: Tipo (ficha) — inline
    ' ----------------------------------------------------------
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTipoCaption")
    With ctrl
        .Caption = "Tipo (ficha):"
        .Left = 10: .Top = y + 2: .Width = 84: .Height = 16
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTipoPreview")
    With ctrl
        .Caption = ""
        .Left = 98: .Top = y: .Width = 250: .Height = 20
        .Font.Name = FN: .Font.Size = FS_PREVIEW: .Font.Bold = True
        .ForeColor = GOLD: .BackStyle = 0
    End With
    y = y + 24
    
    ' ----------------------------------------------------------
    ' ROW 6: Obs
    ' ----------------------------------------------------------
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblObs")
    With ctrl
        .Caption = "Obs"
        .Left = 10: .Top = y: .Width = 30: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    y = y + LH + 2
    
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtObs")
    With ctrl
        .Left = 10: .Top = y: .Width = 620: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    y = y + CH + 14
    
    ' ===========================================================
    ' ===  SECAO: AGENDA DE HORARIOS  ===========================
    ' ===========================================================
    
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblSecaoAgenda")
    With ctrl
        .Caption = "Agenda de Horarios"
        .Left = 10: .Top = y: .Width = 200: .Height = 16
        .Font.Name = FN: .Font.Size = FS_TITLE: .Font.Bold = True
        .BackStyle = 0
    End With
    y = y + 18
    
    Set ctrl = frm.Controls.Add("Forms.Label.1", "linAgenda")
    With ctrl
        .Caption = "": .Left = 10: .Top = y: .Width = 620: .Height = 1
        .BackColor = LINE_COLOR: .BackStyle = 1
    End With
    y = y + 6
    
    ' ListBox de agenda (esquerda)
    Dim yAgendaTop As Single: yAgendaTop = y
    Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstAgenda")
    With ctrl
        .Left = 10: .Top = y: .Width = 310: .Height = 96
        .ColumnCount = 3: .ColumnWidths = "0;148;148"
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    
    ' Controles de agenda (direita)
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblDia")
    With ctrl
        .Caption = "Dia"
        .Left = 335: .Top = yAgendaTop: .Width = 28: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblHora")
    With ctrl
        .Caption = "Hora"
        .Left = 468: .Top = yAgendaTop: .Width = 36: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0
    End With
    
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbDia")
    With ctrl
        .Left = 335: .Top = yAgendaTop + LH + 2: .Width = 120: .Height = CH
        .Style = 2: .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbHora")
    With ctrl
        .Left = 468: .Top = yAgendaTop + LH + 2: .Width = 160: .Height = CH
        .Style = 2: .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE
    End With
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnAddHora")
    With ctrl
        .Caption = "+ Adicionar"
        .Left = 335: .Top = yAgendaTop + LH + 2 + CH + 6: .Width = 140: .Height = 26
        .Font.Name = FN: .Font.Size = FS: .BackColor = GOLD
    End With
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnRemHora")
    With ctrl
        .Caption = "- Remover"
        .Left = 483: .Top = yAgendaTop + LH + 2 + CH + 6: .Width = 145: .Height = 26
        .Font.Name = FN: .Font.Size = FS
    End With
    
    y = yAgendaTop + 96 + 12   ' Abaixo do lstAgenda
    
    ' ===========================================================
    ' ===  FEEDBACK + BOTOES  ===================================
    ' ===========================================================
    
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblFeedback")
    With ctrl
        .Caption = ""
        .Left = 10: .Top = y: .Width = 450: .Height = 18
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True
        .ForeColor = &H8000&: .BackStyle = 0
    End With
    y = y + 22
    
    ' Botoes CENTRALIZADOS: 3 botoes x 92px + 2 gaps x 8px = 292px
    ' Centro: (645 - 292) / 2 = 176
    Dim bx As Single: bx = 176
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnSalvar")
    With ctrl
        .Caption = "Salvar"
        .Left = bx: .Top = y: .Width = 92: .Height = 32
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True
        .BackColor = GOLD
    End With
    bx = bx + 100
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnExcluir")
    With ctrl
        .Caption = "Excluir"
        .Left = bx: .Top = y: .Width = 92: .Height = 32
        .Font.Name = FN: .Font.Size = FS
        .ForeColor = &HFF&          ' Texto vermelho
        .BackColor = RED_BG          ' Fundo rosa claro
    End With
    bx = bx + 100
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnFechar")
    With ctrl
        .Caption = "Fechar"
        .Left = bx: .Top = y: .Width = 92: .Height = 32
        .Font.Name = FN: .Font.Size = FS
    End With
    
    ' ===========================================================
    ' OVERLAYS AUTOCOMPLETE (POR ULTIMO = z-order acima de tudo)
    ' ===========================================================
    
    ' lstSugestoes: abaixo do txtBusca (y=36, txtBusca termina em 36+28=64)
    Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstSugestoes")
    With ctrl
        .Left = 10: .Top = 64: .Width = 310: .Height = 200
        .ColumnCount = 3: .ColumnWidths = "0;50;255"
        .Font.Name = FN: .Font.Size = FS
        .Visible = False: .BackColor = WHITE
        .SpecialEffect = 0: .BorderStyle = 1
    End With
    
    ' lstLivroSugestoes: abaixo do txtLivro
    ' Com CH=28: txtLivro.Top ~ 170, termina em 198
    Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstLivroSugestoes")
    With ctrl
        .Left = 10: .Top = 198: .Width = 230: .Height = 180
        .ColumnCount = 2: .ColumnWidths = "0;225"
        .Font.Name = FN: .Font.Size = FS
        .Visible = False: .BackColor = WHITE
        .SpecialEffect = 0: .BorderStyle = 1
    End With
    
    MsgBox "Formulario frmAlunos v5 criado!" & vbCrLf & _
           "1. Clique 2x em frmAlunos > F7" & vbCrLf & _
           "2. Cole VBA_02_FormLogica.bas" & vbCrLf & _
           "3. Delete este modulo", vbInformation, "Wizped"

End Sub
