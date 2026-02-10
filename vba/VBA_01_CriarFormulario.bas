Attribute VB_Name = "mod_CriarFormulario"
' ===========================================================
' MODULO: CriarFormulario (descartavel)
' Execute CriarFormulario() UMA VEZ, depois delete este modulo.
'
' LAYOUT v8: Compacto (750x470), Tipo ao lado de Contrato,
'            Obs ao lado de Professores, Agenda corrigida
' ===========================================================

Sub CriarFormulario()
    Dim VBProj As Object, VBComp As Object, frm As Object, ctrl As Object
    Set VBProj = ThisWorkbook.VBProject
    On Error Resume Next
    VBProj.VBComponents.Remove VBProj.VBComponents("frmAlunos")
    On Error GoTo 0
    
    Set VBComp = VBProj.VBComponents.Add(3)
    VBComp.Name = "frmAlunos"
    VBComp.Properties("Caption") = "Wizped Office - Gerenciamento de Alunos"
    VBComp.Properties("Width") = 750
    VBComp.Properties("Height") = 470
    VBComp.Properties("BackColor") = &HFAF8F5
    Set frm = VBComp.Designer
    
    Const FN As String = "Segoe UI"
    Const FS As Long = 10
    Const FS_T As Long = 11
    Const CH As Long = 24
    Const LH As Long = 14
    Const BG As Long = &HFAF8F5
    Const GOLD As Long = &H6AA8D5
    Const WHITE As Long = &HFFFFFF
    Const LC As Long = &HD0D0D0
    Dim y As Single
    
    ' =====================================================
    ' BUSCAR ALUNO (global)
    ' =====================================================
    y = 8
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblSecaoBusca")
    With ctrl: .Caption = "Buscar Aluno": .Left = 10: .Top = y: .Width = 200: .Height = 16
        .Font.Name = FN: .Font.Size = FS_T: .Font.Bold = True: .BackStyle = 0: End With
    y = y + 18
    Set ctrl = frm.Controls.Add("Forms.Label.1", "linBusca")
    With ctrl: .Caption = "": .Left = 10: .Top = y: .Width = 725: .Height = 1
        .BackColor = LC: .BackStyle = 1: End With
    y = y + 6
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtBusca")
    With ctrl: .Left = 10: .Top = y: .Width = 340: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnBuscar")
    With ctrl: .Caption = "Buscar": .Left = 358: .Top = y: .Width = 70: .Height = CH
        .Font.Name = FN: .Font.Size = FS: .BackColor = GOLD: End With
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnLimpar")
    With ctrl: .Caption = "Limpar": .Left = 434: .Top = y: .Width = 70: .Height = CH
        .Font.Name = FN: .Font.Size = FS: End With
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnNovo")
    With ctrl: .Caption = "+ Novo Aluno": .Left = 540: .Top = y: .Width = 145: .Height = CH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackColor = GOLD: End With
    y = y + CH + 6
    
    ' =====================================================
    ' MULTIPAGE
    ' =====================================================
    Dim mpTop As Single: mpTop = y
    Dim mpH As Single: mpH = 330
    Set ctrl = frm.Controls.Add("Forms.MultiPage.1", "mpAbas")
    With ctrl: .Left = 6: .Top = mpTop: .Width = 733: .Height = mpH
        .Font.Name = FN: .Font.Size = FS: .BackColor = BG
        .Pages(0).Caption = "Cadastro"
        .Pages(1).Caption = "Historico"
    End With
    Dim mp As Object: Set mp = frm.Controls("mpAbas")
    
    ' =====================================================
    ' TAB 0: CADASTRO
    ' =====================================================
    Dim pg0 As Object: Set pg0 = mp.Pages(0)
    y = 4
    
    ' --- ROW 1: ID + Nome ---
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblID")
    With ctrl: .Caption = "ID (SponteWeb)": .Left = 4: .Top = y: .Width = 80: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblNome")
    With ctrl: .Caption = "Nome": .Left = 88: .Top = y: .Width = 50: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    y = y + LH + 2
    Set ctrl = pg0.Controls.Add("Forms.TextBox.1", "txtID")
    With ctrl: .Left = 4: .Top = y: .Width = 75: .Height = CH: .MaxLength = 5
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.TextBox.1", "txtNome")
    With ctrl: .Left = 88: .Top = y: .Width = 620: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    y = y + CH + 4
    
    ' --- ROW 2: Livro + Experiencia + VIP + Data Inicio ---
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblLivro")
    With ctrl: .Caption = "Livro": .Left = 4: .Top = y: .Width = 40: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblExperiencia")
    With ctrl: .Caption = "Experiencia": .Left = 180: .Top = y: .Width = 76: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblData")
    With ctrl: .Caption = "Data Inicio": .Left = 445: .Top = y: .Width = 76: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    y = y + LH + 2
    Set ctrl = pg0.Controls.Add("Forms.TextBox.1", "txtLivro")
    With ctrl: .Left = 4: .Top = y: .Width = 145: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.CommandButton.1", "btnLivroDD")
    With ctrl: .Caption = ChrW(9660): .Left = 153: .Top = y: .Width = 20: .Height = CH
        .Font.Name = FN: .Font.Size = 8: End With
    Set ctrl = pg0.Controls.Add("Forms.ComboBox.1", "cmbExperiencia")
    With ctrl: .Left = 180: .Top = y: .Width = 110: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;105": .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.CheckBox.1", "chkVIP")
    With ctrl: .Caption = "VIP": .Left = 300: .Top = y + 2: .Width = 50: .Height = 20
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.TextBox.1", "txtData")
    With ctrl: .Left = 445: .Top = y: .Width = 90: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    y = y + CH + 4
    
    ' --- ROW 3: Modalidade + Status + Contrato + Tipo(ficha) ---
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblModalidade")
    With ctrl: .Caption = "Modalidade": .Left = 4: .Top = y: .Width = 72: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblStatus")
    With ctrl: .Caption = "Status": .Left = 134: .Top = y: .Width = 50: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblContrato")
    With ctrl: .Caption = "Contrato": .Left = 254: .Top = y: .Width = 60: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblTipoCaption")
    With ctrl: .Caption = "Tipo (ficha):": .Left = 390: .Top = y: .Width = 84: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    y = y + LH + 2
    Set ctrl = pg0.Controls.Add("Forms.ComboBox.1", "cmbModalidade")
    With ctrl: .Left = 4: .Top = y: .Width = 120: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;115": .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.ComboBox.1", "cmbStatus")
    With ctrl: .Left = 134: .Top = y: .Width = 110: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;105": .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.ComboBox.1", "cmbContrato")
    With ctrl: .Left = 254: .Top = y: .Width = 120: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;115": .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblTipoPreview")
    With ctrl: .Caption = "": .Left = 478: .Top = y: .Width = 230: .Height = CH
        .Font.Name = FN: .Font.Size = 12: .Font.Bold = True
        .ForeColor = GOLD: .BackStyle = 0: End With
    y = y + CH + 4
    
    ' --- ROW 4: Professores (left) + Obs (right) ---
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblProfessor")
    With ctrl: .Caption = "Professores": .Left = 4: .Top = y: .Width = 80: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblObs")
    With ctrl: .Caption = "Obs": .Left = 140: .Top = y: .Width = 30: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    y = y + LH + 2
    Set ctrl = pg0.Controls.Add("Forms.ListBox.1", "lstProfessores")
    With ctrl: .Left = 4: .Top = y: .Width = 126: .Height = 72
        .ColumnCount = 2: .ColumnWidths = "0;105"
        .MultiSelect = 1: .ListStyle = 1
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.TextBox.1", "txtObs")
    With ctrl: .Left = 140: .Top = y: .Width = 568: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    y = y + 76
    
    ' --- AGENDA ---
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblSecaoAgenda")
    With ctrl: .Caption = "Agenda de Horarios": .Left = 4: .Top = y: .Width = 200: .Height = 16
        .Font.Name = FN: .Font.Size = FS_T: .Font.Bold = True: .BackStyle = 0: End With
    y = y + 18
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "linAgenda")
    With ctrl: .Caption = "": .Left = 4: .Top = y: .Width = 708: .Height = 1
        .BackColor = LC: .BackStyle = 1: End With
    y = y + 4
    Dim yAg As Single: yAg = y
    ' lstAgenda height = labels(14+2) + combos(24+4) + buttons(24) = 68
    Dim agH As Single: agH = 72
    Set ctrl = pg0.Controls.Add("Forms.ListBox.1", "lstAgenda")
    With ctrl: .Left = 4: .Top = yAg: .Width = 300: .Height = agH
        .ColumnCount = 3: .ColumnWidths = "0;145;145"
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblDia")
    With ctrl: .Caption = "Dia": .Left = 314: .Top = yAg: .Width = 28: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.Label.1", "lblHora")
    With ctrl: .Caption = "Hora": .Left = 430: .Top = yAg: .Width = 36: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg0.Controls.Add("Forms.ComboBox.1", "cmbDia")
    With ctrl: .Left = 314: .Top = yAg + LH + 2: .Width = 108: .Height = CH: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.ComboBox.1", "cmbHora")
    With ctrl: .Left = 430: .Top = yAg + LH + 2: .Width = 140: .Height = CH: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg0.Controls.Add("Forms.CommandButton.1", "btnAddHora")
    With ctrl: .Caption = "+ Adicionar": .Left = 314: .Top = yAg + LH + CH + 6: .Width = 120: .Height = 24
        .Font.Name = FN: .Font.Size = FS: .BackColor = GOLD: End With
    Set ctrl = pg0.Controls.Add("Forms.CommandButton.1", "btnRemHora")
    With ctrl: .Caption = "- Remover": .Left = 442: .Top = yAg + LH + CH + 6: .Width = 128: .Height = 24
        .Font.Name = FN: .Font.Size = FS: End With
    
    ' =====================================================
    ' TAB 1: HISTORICO
    ' =====================================================
    Dim pg1 As Object: Set pg1 = mp.Pages(1)
    y = 4
    Set ctrl = pg1.Controls.Add("Forms.ListBox.1", "lstHistorico")
    With ctrl: .Left = 4: .Top = y: .Width = 708: .Height = 160
        .ColumnCount = 5: .ColumnWidths = "0;70;110;200;320"
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    y = y + 164
    Set ctrl = pg1.Controls.Add("Forms.Label.1", "lblSecaoAddHist")
    With ctrl: .Caption = "Registrar Evento": .Left = 4: .Top = y: .Width = 200: .Height = 16
        .Font.Name = FN: .Font.Size = FS_T: .Font.Bold = True: .BackStyle = 0: End With
    y = y + 18
    Set ctrl = pg1.Controls.Add("Forms.Label.1", "linHist")
    With ctrl: .Caption = "": .Left = 4: .Top = y: .Width = 708: .Height = 1
        .BackColor = LC: .BackStyle = 1: End With
    y = y + 6
    ' Row: Tipo + Livro + Data + Obs (all on one row)
    Set ctrl = pg1.Controls.Add("Forms.Label.1", "lblTipoOcorrencia")
    With ctrl: .Caption = "Tipo": .Left = 4: .Top = y: .Width = 40: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg1.Controls.Add("Forms.Label.1", "lblLivroHist")
    With ctrl: .Caption = "Livro": .Left = 180: .Top = y: .Width = 40: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg1.Controls.Add("Forms.Label.1", "lblDataHist")
    With ctrl: .Caption = "Data": .Left = 356: .Top = y: .Width = 40: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    Set ctrl = pg1.Controls.Add("Forms.Label.1", "lblObsHist")
    With ctrl: .Caption = "Obs": .Left = 460: .Top = y: .Width = 30: .Height = LH
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackStyle = 0: End With
    y = y + LH + 2
    Set ctrl = pg1.Controls.Add("Forms.ComboBox.1", "cmbTipoOcorrencia")
    With ctrl: .Left = 4: .Top = y: .Width = 165: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;160": .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg1.Controls.Add("Forms.ComboBox.1", "cmbLivroHist")
    With ctrl: .Left = 180: .Top = y: .Width = 165: .Height = CH
        .ColumnCount = 2: .ColumnWidths = "0;160": .BoundColumn = 1: .TextColumn = 2: .Style = 2
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg1.Controls.Add("Forms.TextBox.1", "txtDataHist")
    With ctrl: .Left = 356: .Top = y: .Width = 90: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    Set ctrl = pg1.Controls.Add("Forms.TextBox.1", "txtObsHist")
    With ctrl: .Left = 460: .Top = y: .Width = 252: .Height = CH
        .Font.Name = FN: .Font.Size = FS
        .SpecialEffect = 0: .BorderStyle = 1: .BackColor = WHITE: End With
    y = y + CH + 6
    ' Botoes
    Set ctrl = pg1.Controls.Add("Forms.CommandButton.1", "btnAddHist")
    With ctrl: .Caption = "+ Registrar": .Left = 4: .Top = y: .Width = 120: .Height = 26
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackColor = GOLD: End With
    Set ctrl = pg1.Controls.Add("Forms.CommandButton.1", "btnRemHist")
    With ctrl: .Caption = "- Remover": .Left = 132: .Top = y: .Width = 120: .Height = 26
        .Font.Name = FN: .Font.Size = FS: End With
    Set ctrl = pg1.Controls.Add("Forms.Label.1", "lblFeedbackHist")
    With ctrl: .Caption = "": .Left = 268: .Top = y + 4: .Width = 440: .Height = 18
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True
        .ForeColor = &H8000&: .BackStyle = 0: End With
    
    ' =====================================================
    ' FEEDBACK + BOTOES GLOBAIS
    ' =====================================================
    y = mpTop + mpH + 4
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblFeedback")
    With ctrl: .Caption = "": .Left = 10: .Top = y: .Width = 500: .Height = 18
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True
        .ForeColor = &H8000&: .BackStyle = 0: End With
    y = y + 22
    Dim bx As Single: bx = 230
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnSalvar")
    With ctrl: .Caption = "Salvar": .Left = bx: .Top = y: .Width = 92: .Height = 32
        .Font.Name = FN: .Font.Size = FS: .Font.Bold = True: .BackColor = GOLD: End With
    bx = bx + 100
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnExcluir")
    With ctrl: .Caption = "Excluir": .Left = bx: .Top = y: .Width = 92: .Height = 32
        .Font.Name = FN: .Font.Size = FS: .ForeColor = &HFF&: .BackColor = &HE0E0FF: End With
    bx = bx + 100
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnFechar")
    With ctrl: .Caption = "Fechar": .Left = bx: .Top = y: .Width = 92: .Height = 32
        .Font.Name = FN: .Font.Size = FS: End With
    
    ' =====================================================
    ' OVERLAYS
    ' =====================================================
    Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstSugestoes")
    With ctrl: .Left = 10: .Top = 58: .Width = 340: .Height = 200
        .ColumnCount = 3: .ColumnWidths = "0;50;285"
        .Font.Name = FN: .Font.Size = FS
        .Visible = False: .BackColor = WHITE
        .SpecialEffect = 0: .BorderStyle = 1: End With
    ' lstLivroSugestoes movido para dentro de pg0 (ver abaixo)
    Set ctrl = pg0.Controls.Add("Forms.ListBox.1", "lstLivroSugestoes")
    With ctrl: .Left = 4: .Top = 88: .Width = 165: .Height = 200
        .ColumnCount = 2: .ColumnWidths = "0;160"
        .Font.Name = FN: .Font.Size = FS
        .Visible = False: .BackColor = WHITE
        .SpecialEffect = 0: .BorderStyle = 1: End With
    
    MsgBox "frmAlunos v8 criado!" & vbCrLf & _
           "1. F7 > cole VBA_02" & vbCrLf & _
           "2. Delete este modulo", vbInformation, "Wizped"
End Sub
