Attribute VB_Name = "mod_RibbonCallbacks"
' ===========================================================
' MODULO: mod_RibbonCallbacks
' Callbacks para os botoes da Ribbon Wizped v4
'
' COMO USAR:
'   1. Cole este codigo em um modulo padrao no VBE
'   2. Aplique o RibbonX_Wizped.xml via Custom UI Editor
' ===========================================================

Option Explicit

' -----------------------------------------------
' Grupo: Cadastro
' -----------------------------------------------

' Abrir formulario de gerenciamento (modo busca)
Public Sub OnGerenciarAlunos(control As IRibbonControl)
    frmAlunos.Show vbModeless
End Sub

' Abrir formulario ja em modo novo aluno
Public Sub OnNovoAluno(control As IRibbonControl)
    frmAlunos.Show vbModeless
    frmAlunos.btnNovo_Click
End Sub

' Importar cadastro do Sponte
' NOTA: OnImportarCadastro ja esta definido em VBA_04_ModImportarCadastro.bas
'       Nao duplicar aqui.

' -----------------------------------------------
' Grupo: Relatorios
' -----------------------------------------------

' Gerar fichas mensais (placeholder - sera implementado)
Public Sub OnGerarFichas(control As IRibbonControl)
    MsgBox "Gerar Fichas de Frequencia" & vbCrLf & _
           "(Funcionalidade em desenvolvimento)", _
           vbInformation, "Wizped Office"
End Sub

' Dashboard de indicadores (placeholder)
Public Sub OnDashboard(control As IRibbonControl)
    MsgBox "Dashboard de Indicadores" & vbCrLf & _
           "(Funcionalidade em desenvolvimento)" & vbCrLf & vbCrLf & _
           "Futuro: metricas de alunos ativos, inativos," & vbCrLf & _
           "distribuicao por professor, livro, etc.", _
           vbInformation, "Wizped Office"
End Sub

' -----------------------------------------------
' Grupo: Ferramentas
' -----------------------------------------------

' Mostrar todas as planilhas BD_*
Public Sub OnMostrarPlanilhas(control As IRibbonControl)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 3) = "BD_" Then ws.Visible = xlSheetVisible
    Next ws
    MsgBox "Planilhas de dados agora visiveis.", vbInformation, "Wizped Office"
End Sub

' Esconder todas as planilhas BD_*
Public Sub OnEsconderPlanilhas(control As IRibbonControl)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 3) = "BD_" Then ws.Visible = xlSheetVeryHidden
    Next ws
    MsgBox "Planilhas de dados ocultadas.", vbInformation, "Wizped Office"
End Sub

' Recriar formulario de alunos (deleta e recria frmAlunos)
Public Sub OnRecriarFormulario(control As IRibbonControl)
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Deseja recriar o formulario frmAlunos do zero?" & vbCrLf & _
                  "Isso ira deletar o formulario atual e recria-lo.", _
                  vbQuestion + vbYesNo, "Wizped Office")
    If resp = vbNo Then Exit Sub
    
    ' Deletar formulario existente
    On Error Resume Next
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name = "frmAlunos" Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            Exit For
        End If
    Next vbComp
    On Error GoTo 0
    
    ' Recriar via CriarFormularioAlunos
    CriarFormularioAlunos
    
    MsgBox "Formulario frmAlunos recriado com sucesso!", _
           vbInformation, "Wizped Office"
End Sub

' Sobre o Wizped Office
Public Sub OnSobre(control As IRibbonControl)
    Dim info As String
    info = "Wizped Office v4.0" & vbCrLf & vbCrLf & _
           "Sistema de Gestao de Alunos" & vbCrLf & _
           "Desenvolvido para Wizped Idiomas" & vbCrLf & vbCrLf & _
           "Alunos cadastrados: " & _
           (ThisWorkbook.Sheets("BD_Alunos").Cells( _
            ThisWorkbook.Sheets("BD_Alunos").Rows.Count, 1).End(xlUp).Row - 1) & vbCrLf & _
           "Professores: " & _
           (ThisWorkbook.Sheets("BD_Professores").Cells( _
            ThisWorkbook.Sheets("BD_Professores").Rows.Count, 1).End(xlUp).Row - 1) & vbCrLf & _
           "Livros ativos: " & _
           Application.WorksheetFunction.CountIf( _
            ThisWorkbook.Sheets("BD_Livros").Columns(6), True)
    
    MsgBox info, vbInformation, "Wizped Office"
End Sub
