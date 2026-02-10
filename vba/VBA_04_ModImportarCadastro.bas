Attribute VB_Name = "mod_ImportarCadastro"
' ===========================================================
' MODULO: mod_ImportarCadastro
' PROPÓSITO: Importar dados de alunos do relatório Sponte PDF
'
' COMO USAR:
'   1. Cole este código em um módulo padrão no VBE
'   2. O botão "Importar Cadastro" na Ribbon chama OnImportarCadastro
'   3. Requer: Python 3 + pdfplumber instalado (pip install pdfplumber)
'   4. O script wizped_import.py deve estar na mesma pasta do .xlsm
'
' FLUXO:
'   1. Dialog: "Abrir Sponte Web" ou "Buscar PDF no computador"
'   2. Se Buscar PDF: FileDialog → Python parseia → CSV → Merge
'   3. Comparação inteligente: por ID e por nome (fuzzy)
'   4. Resultado: quantos adicionados, atualizados, ignorados
'
' COLUNAS BD_Alunos (nova ordem):
'   1=ID_Aluno  2=Nome  3=ID_Status  4=ID_Contrato  5=ID_Livro
'   6=ID_Experiencia  7=ID_Modalidade  8=VIP  9=ID_Professor
'   10=Data_Inicio  11=Obs
' ===========================================================

' URL do relatório no Sponte Web
Private Const SPONTE_URL As String = "https://www.sponteweb.com.br/SPRel/Alunos/DadosCadastro.aspx"

' Mapeamento de Situação → ID_Status
' Ativo=1, Trancado=2, Desistente=3, Interessado=4
Private Function MapearStatus(situacao As String) As Variant
    Select Case LCase(Trim(situacao))
        Case "ativo": MapearStatus = 1
        Case "trancado": MapearStatus = 2
        Case "desistente": MapearStatus = 3
        Case "interessado": MapearStatus = 4
        Case Else: MapearStatus = 1  ' Padrão: Ativo
    End Select
End Function

' ===========================================================
' PONTO DE ENTRADA (chamado pela Ribbon)
' ===========================================================
Public Sub OnImportarCadastro(control As IRibbonControl)
    ImportarCadastroSponte
End Sub

Public Sub ImportarCadastroSponte()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Importar Cadastro Sponte" & vbCrLf & vbCrLf & _
                  "SIM = Buscar PDF já baixado no computador" & vbCrLf & _
                  "NÃO = Abrir Sponte Web para gerar o relatório" & vbCrLf & _
                  "CANCELAR = Voltar", _
                  vbYesNoCancel + vbQuestion, "Wizped Office - Importar")
    
    Select Case resp
        Case vbYes
            ImportarPDF
        Case vbNo
            AbrirSponteWeb
        Case vbCancel
            ' Nada
    End Select
End Sub

' ===========================================================
' ABRIR SPONTE WEB NO NAVEGADOR
' ===========================================================
Private Sub AbrirSponteWeb()
    On Error Resume Next
    ' Método 1: Shell com navegador padrão
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")
    shellObj.Run SPONTE_URL, 1, False
    Set shellObj = Nothing
    
    If Err.Number <> 0 Then
        Err.Clear
        ' Método 2: FollowHyperlink
        ThisWorkbook.FollowHyperlink SPONTE_URL
    End If
    On Error GoTo 0
    
    MsgBox "O Sponte Web foi aberto no navegador." & vbCrLf & _
           "Gere o relatório 'Dados do Cadastro - Layout Padrão'" & vbCrLf & _
           "e salve o PDF. Depois clique em Importar novamente" & vbCrLf & _
           "e escolha 'Buscar PDF'.", vbInformation, "Wizped Office"
End Sub

' ===========================================================
' IMPORTAR PDF (fluxo principal)
' ===========================================================
Private Sub ImportarPDF()
    ' 1. Pedir o arquivo PDF
    Dim pdfPath As String
    pdfPath = SelecionarPDF()
    If pdfPath = "" Then Exit Sub
    
    ' 2. Verificar se Python está disponível
    If Not PythonDisponivel() Then
        MsgBox "Python 3 não encontrado no sistema." & vbCrLf & vbCrLf & _
               "Instale Python 3 e pdfplumber:" & vbCrLf & _
               "  pip install pdfplumber" & vbCrLf & vbCrLf & _
               "Tentando importação alternativa (texto bruto)...", _
               vbExclamation, "Wizped Office"
        ' Alternativa sem Python: tentar ler o CSV se já existir
        Dim csvAlt As String
        csvAlt = Replace(pdfPath, ".pdf", "_parsed.csv", , , vbTextCompare)
        If Dir(csvAlt) = "" Then
            MsgBox "Nenhum CSV encontrado. Execute manualmente:" & vbCrLf & _
                   "python wizped_import.py """ & pdfPath & """", _
                   vbCritical, "Wizped Office"
            Exit Sub
        End If
        ProcessarCSV csvAlt
        Exit Sub
    End If
    
    ' 3. Chamar Python para parsear o PDF
    Application.StatusBar = "Processando PDF do Sponte..."
    DoEvents
    
    Dim csvPath As String
    csvPath = ChamarPython(pdfPath)
    
    If csvPath = "" Then
        Application.StatusBar = False
        MsgBox "Erro ao processar o PDF." & vbCrLf & _
               "Verifique se o pdfplumber está instalado:" & vbCrLf & _
               "  pip install pdfplumber", vbCritical, "Wizped Office"
        Exit Sub
    End If
    
    ' 4. Processar o CSV gerado
    ProcessarCSV csvPath
    
    Application.StatusBar = False
End Sub

' ===========================================================
' FILE DIALOG PARA SELECIONAR PDF
' ===========================================================
Private Function SelecionarPDF() As String
    SelecionarPDF = ""
    
    Dim fd As Object
    Set fd = Application.FileDialog(1) ' msoFileDialogFilePicker
    
    With fd
        .Title = "Selecionar Relatório Sponte (PDF)"
        .Filters.Clear
        .Filters.Add "Arquivos PDF", "*.pdf"
        .FilterIndex = 1
        .AllowMultiSelect = False
        
        ' Tentar abrir na pasta Downloads
        Dim downloadsPath As String
        downloadsPath = Environ("USERPROFILE") & "\Downloads"
        If Dir(downloadsPath, vbDirectory) <> "" Then
            .InitialFileName = downloadsPath & "\"
        End If
        
        If .Show = -1 Then
            SelecionarPDF = .SelectedItems(1)
        End If
    End With
End Function

' ===========================================================
' VERIFICAR SE PYTHON ESTÁ DISPONÍVEL
' ===========================================================
Private Function PythonDisponivel() As Boolean
    On Error Resume Next
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")
    Dim result As Long
    result = shellObj.Run("python --version", 0, True)
    PythonDisponivel = (Err.Number = 0 And result = 0)
    If Not PythonDisponivel Then
        Err.Clear
        result = shellObj.Run("python3 --version", 0, True)
        PythonDisponivel = (Err.Number = 0 And result = 0)
    End If
    Set shellObj = Nothing
    On Error GoTo 0
End Function

' ===========================================================
' CHAMAR PYTHON PARA PARSEAR O PDF
' ===========================================================
Private Function ChamarPython(pdfPath As String) As String
    ChamarPython = ""
    
    ' O script Python deve estar na mesma pasta do workbook
    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & "\wizped_import.py"
    
    ' Se não encontrar na pasta do workbook, tentar pasta do PDF
    If Dir(scriptPath) = "" Then
        scriptPath = Left(pdfPath, InStrRev(pdfPath, "\")) & "wizped_import.py"
    End If
    
    If Dir(scriptPath) = "" Then
        MsgBox "Script wizped_import.py não encontrado." & vbCrLf & _
               "Coloque-o na mesma pasta do arquivo Excel:" & vbCrLf & _
               ThisWorkbook.Path, vbCritical, "Wizped Office"
        Exit Function
    End If
    
    ' Montar comando
    Dim cmd As String
    cmd = "python """ & scriptPath & """ """ & pdfPath & """"
    
    ' Executar e aguardar
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")
    
    On Error Resume Next
    Dim exitCode As Long
    exitCode = shellObj.Run("cmd /c " & cmd, 0, True)  ' vbHide=0, Wait=True
    
    If Err.Number <> 0 Or exitCode <> 0 Then
        ' Tentar com python3
        cmd = "python3 """ & scriptPath & """ """ & pdfPath & """"
        Err.Clear
        exitCode = shellObj.Run("cmd /c " & cmd, 0, True)
    End If
    On Error GoTo 0
    
    Set shellObj = Nothing
    
    ' Verificar se o CSV foi criado
    Dim csvPath As String
    csvPath = Replace(pdfPath, ".pdf", "_parsed.csv", , , vbTextCompare)
    If Len(csvPath) < 5 Then csvPath = pdfPath & "_parsed.csv"
    
    If Dir(csvPath) <> "" Then
        ChamarPython = csvPath
    End If
End Function

' ===========================================================
' PROCESSAR CSV E FAZER O MERGE COM BD_ALUNOS
' ===========================================================
Private Sub ProcessarCSV(csvPath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BD_Alunos")
    
    ' Ler CSV
    Dim csvData() As Variant
    Dim csvCount As Long
    csvCount = LerCSV(csvPath, csvData)
    
    If csvCount = 0 Then
        MsgBox "Nenhum dado encontrado no CSV.", vbExclamation, "Wizped Office"
        Exit Sub
    End If
    
    Application.StatusBar = "Comparando " & csvCount & " alunos do Sponte com BD_Alunos..."
    DoEvents
    
    ' Construir índices dos alunos existentes
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row  ' Col 2 = Nome (mais confiável)
    If lastRow < 2 Then lastRow = 1
    
    ' Índice por ID
    Dim existingIDs As Object
    Set existingIDs = CreateObject("Scripting.Dictionary")
    
    ' Índice por nome normalizado → linha
    Dim existingNames As Object
    Set existingNames = CreateObject("Scripting.Dictionary")
    
    Dim r As Long
    For r = 2 To lastRow
        Dim cellID As Variant: cellID = ws.Cells(r, 1).Value
        Dim cellName As String: cellName = Trim(CStr(ws.Cells(r, 2).Value))
        
        If Not IsEmpty(cellID) And cellID <> "" Then
            existingIDs(CStr(CLng(cellID))) = r
        End If
        
        If cellName <> "" Then
            Dim normName As String: normName = NormalizarNome(cellName)
            If Not existingNames.Exists(normName) Then
                existingNames(normName) = r
            End If
        End If
    Next r
    
    ' Processar cada aluno do CSV
    Dim statsIDExistente As Long, statsNomeMatch As Long
    Dim statsNomeUpdate As Long, statsNovo As Long, statsIDCorrigido As Long
    
    Dim i As Long
    For i = 1 To csvCount
        Dim csvID As Long: csvID = CLng(csvData(i, 1))
        Dim csvNome As String: csvNome = Trim(CStr(csvData(i, 2)))
        Dim csvSituacao As String: csvSituacao = Trim(CStr(csvData(i, 3)))
        Dim csvStatusID As Variant: csvStatusID = MapearStatus(csvSituacao)
        
        ' Progresso
        If i Mod 10 = 0 Then
            Application.StatusBar = "Processando aluno " & i & " de " & csvCount & "..."
            DoEvents
        End If
        
        Dim idStr As String: idStr = CStr(csvID)
        
        ' CASO 1: ID já existe na planilha
        If existingIDs.Exists(idStr) Then
            Dim existRow As Long: existRow = existingIDs(idStr)
            ' Verificar se nome precisa atualizar (nome mais completo)
            Dim existName As String: existName = Trim(CStr(ws.Cells(existRow, 2).Value))
            If Len(csvNome) > Len(existName) Then
                ws.Cells(existRow, 2).Value = csvNome
                statsNomeUpdate = statsNomeUpdate + 1
            Else
                statsIDExistente = statsIDExistente + 1
            End If
            GoTo ProximoAluno
        End If
        
        ' CASO 2: Nome existe mas sem ID (ou com ID diferente)
        Dim csvNormName As String: csvNormName = NormalizarNome(csvNome)
        Dim matchRow As Long: matchRow = 0
        
        ' Busca exata por nome normalizado
        If existingNames.Exists(csvNormName) Then
            matchRow = existingNames(csvNormName)
        Else
            ' Busca fuzzy: percorrer todos e comparar
            matchRow = BuscarNomeFuzzy(ws, lastRow, csvNome)
        End If
        
        If matchRow > 0 Then
            Dim matchID As Variant: matchID = ws.Cells(matchRow, 1).Value
            
            If IsEmpty(matchID) Or matchID = "" Then
                ' Sem ID → adicionar
                ws.Cells(matchRow, 1).Value = csvID
                statsNomeMatch = statsNomeMatch + 1
            Else
                ' Tem ID diferente → corrigir para o do Sponte
                ws.Cells(matchRow, 1).Value = csvID
                ' Atualizar índice
                existingIDs.Remove CStr(CLng(matchID))
                statsIDCorrigido = statsIDCorrigido + 1
            End If
            
            ' Atualizar nome se mais completo
            Dim curName As String: curName = Trim(CStr(ws.Cells(matchRow, 2).Value))
            If Len(csvNome) > Len(curName) Then
                ws.Cells(matchRow, 2).Value = csvNome
            End If
            
            ' Registrar novo ID no índice
            existingIDs(idStr) = matchRow
            GoTo ProximoAluno
        End If
        
        ' CASO 3: Aluno novo → adicionar linha
        lastRow = lastRow + 1
        ws.Cells(lastRow, 1).Value = csvID         ' ID
        ws.Cells(lastRow, 2).Value = csvNome        ' Nome
        ws.Cells(lastRow, 3).Value = csvStatusID    ' Status
        ' Demais campos ficam vazios (livro, experiência, etc.)
        statsNovo = statsNovo + 1
        existingIDs(idStr) = lastRow
        
ProximoAluno:
    Next i
    
    Application.StatusBar = False
    
    ' Relatório
    Dim msg As String
    msg = "Importação concluída!" & vbCrLf & vbCrLf & _
          "Total no PDF: " & csvCount & " alunos" & vbCrLf & _
          "─────────────────────────" & vbCrLf
    
    If statsIDExistente > 0 Then msg = msg & "Já existiam (sem alteração): " & statsIDExistente & vbCrLf
    If statsNomeMatch > 0 Then msg = msg & "ID adicionado (match por nome): " & statsNomeMatch & vbCrLf
    If statsNomeUpdate > 0 Then msg = msg & "Nome atualizado (mais completo): " & statsNomeUpdate & vbCrLf
    If statsIDCorrigido > 0 Then msg = msg & "ID corrigido (Sponte prevalece): " & statsIDCorrigido & vbCrLf
    If statsNovo > 0 Then msg = msg & "Alunos novos adicionados: " & statsNovo & vbCrLf
    
    MsgBox msg, vbInformation, "Wizped Office - Importação"
End Sub

' ===========================================================
' LER CSV COM SEPARADOR ;
' ===========================================================
Private Function LerCSV(csvPath As String, ByRef dados() As Variant) As Long
    Dim ff As Integer: ff = FreeFile
    Dim linha As String
    Dim linhas() As String
    Dim conteudo As String
    
    ' Ler arquivo inteiro
    Open csvPath For Binary As #ff
    conteudo = Space$(LOF(ff))
    Get #ff, , conteudo
    Close #ff
    
    ' Remover BOM UTF-8 se existir
    If Left(conteudo, 3) = Chr(239) & Chr(187) & Chr(191) Then
        conteudo = Mid(conteudo, 4)
    End If
    
    linhas = Split(conteudo, vbLf)
    
    Dim count As Long: count = 0
    Dim maxRows As Long: maxRows = UBound(linhas)
    
    ' Primeira passagem: contar linhas válidas (pular header)
    Dim j As Long
    For j = 1 To maxRows
        If Trim(Replace(linhas(j), vbCr, "")) <> "" Then count = count + 1
    Next j
    
    If count = 0 Then
        LerCSV = 0
        Exit Function
    End If
    
    ReDim dados(1 To count, 1 To 3) ' ID, Nome, Situacao
    
    Dim idx As Long: idx = 0
    For j = 1 To maxRows
        linha = Trim(Replace(linhas(j), vbCr, ""))
        If linha = "" Then GoTo ProxLinha
        
        Dim campos() As String
        campos = Split(linha, ";")
        
        If UBound(campos) >= 2 Then
            idx = idx + 1
            dados(idx, 1) = campos(0)  ' ID
            dados(idx, 2) = campos(1)  ' Nome
            dados(idx, 3) = campos(2)  ' Situacao
        End If
ProxLinha:
    Next j
    
    LerCSV = idx
End Function

' ===========================================================
' NORMALIZAR NOME (remover acentos, lowercase, trim)
' ===========================================================
Private Function NormalizarNome(nome As String) As String
    Dim s As String: s = LCase(Trim(nome))
    ' Remover acentos manualmente (VBA não tem unicode normalizer)
    s = Replace(s, "á", "a"): s = Replace(s, "à", "a"): s = Replace(s, "ã", "a"): s = Replace(s, "â", "a"): s = Replace(s, "ä", "a")
    s = Replace(s, "é", "e"): s = Replace(s, "è", "e"): s = Replace(s, "ê", "e"): s = Replace(s, "ë", "e")
    s = Replace(s, "í", "i"): s = Replace(s, "ì", "i"): s = Replace(s, "î", "i"): s = Replace(s, "ï", "i")
    s = Replace(s, "ó", "o"): s = Replace(s, "ò", "o"): s = Replace(s, "õ", "o"): s = Replace(s, "ô", "o"): s = Replace(s, "ö", "o")
    s = Replace(s, "ú", "u"): s = Replace(s, "ù", "u"): s = Replace(s, "û", "u"): s = Replace(s, "ü", "u")
    s = Replace(s, "ç", "c"): s = Replace(s, "ñ", "n")
    ' Remover espaços duplicados
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizarNome = Trim(s)
End Function

' ===========================================================
' BUSCA FUZZY POR NOME
' Compara primeiro+último nome e tokens em comum
' ===========================================================
Private Function BuscarNomeFuzzy(ws As Worksheet, lastRow As Long, nomePDF As String) As Long
    BuscarNomeFuzzy = 0
    
    Dim pdfNorm As String: pdfNorm = NormalizarNome(nomePDF)
    Dim pdfTokens() As String: pdfTokens = Split(pdfNorm, " ")
    If UBound(pdfTokens) < 1 Then Exit Function  ' Precisa de pelo menos 2 partes
    
    Dim pdfFirst As String: pdfFirst = pdfTokens(0)
    Dim pdfLast As String: pdfLast = pdfTokens(UBound(pdfTokens))
    
    Dim bestScore As Long: bestScore = 0
    Dim bestRow As Long: bestRow = 0
    
    Dim r As Long
    For r = 2 To lastRow
        Dim cellName As String: cellName = Trim(CStr(ws.Cells(r, 2).Value))
        If cellName = "" Then GoTo ProxRow
        
        Dim exNorm As String: exNorm = NormalizarNome(cellName)
        
        ' Match exato
        If exNorm = pdfNorm Then
            BuscarNomeFuzzy = r
            Exit Function
        End If
        
        Dim exTokens() As String: exTokens = Split(exNorm, " ")
        If UBound(exTokens) < 1 Then GoTo ProxRow
        
        Dim exFirst As String: exFirst = exTokens(0)
        Dim exLast As String: exLast = exTokens(UBound(exTokens))
        
        ' Primeiro + último nome devem bater
        If exFirst = pdfFirst And exLast = pdfLast Then
            ' Contar tokens em comum
            Dim score As Long: score = 0
            Dim t1 As Long, t2 As Long
            For t1 = 0 To UBound(exTokens)
                Dim tok As String: tok = exTokens(t1)
                If Len(tok) <= 2 Then GoTo ProxToken  ' Ignorar "de", "da", etc.
                For t2 = 0 To UBound(pdfTokens)
                    If tok = pdfTokens(t2) Then
                        score = score + 1
                        Exit For
                    End If
                    ' Abreviações: "C." matches "Cardoso"
                    If Right(tok, 1) = "." And Len(tok) >= 2 Then
                        If Left(pdfTokens(t2), Len(tok) - 1) = Left(tok, Len(tok) - 1) Then
                            score = score + 1
                            Exit For
                        End If
                    End If
                Next t2
ProxToken:
            Next t1
            
            ' Score mínimo: primeiro + último + pelo menos 1 intermediário
            If score >= 2 And score > bestScore Then
                bestScore = score
                bestRow = r
            End If
        End If
ProxRow:
    Next r
    
    BuscarNomeFuzzy = bestRow
End Function
