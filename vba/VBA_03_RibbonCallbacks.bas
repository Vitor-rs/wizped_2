Attribute VB_Name = "mod_RibbonCallbacks"
' ===========================================================
' MODULO: mod_RibbonCallbacks
' Callbacks para os botoes da Ribbon Wizped
'
' COMO USAR:
'   1. Cole este codigo em um modulo padrao no VBE
'   2. Aplique o RibbonX_Wizped.xml via Custom UI Editor
' ===========================================================

Option Explicit

' -----------------------------------------------
' Grupo: Alunos
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
' NOTA: OnImportarCadastro já está definido em VBA_04_ModImportarCadastro.bas
'       Não duplicar aqui.

' -----------------------------------------------
' Grupo: Fichas
' -----------------------------------------------

' Gerar fichas mensais (placeholder - sera implementado)
Public Sub OnGerarFichas(control As IRibbonControl)
    MsgBox "Gerar Fichas de Frequencia" & vbCrLf & _
           "(Funcionalidade em desenvolvimento)", _
           vbInformation, "Wizped"
End Sub

' -----------------------------------------------
' Grupo: Planilhas
' -----------------------------------------------

' Mostrar todas as planilhas BD_*
Public Sub OnMostrarPlanilhas(control As IRibbonControl)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 3) = "BD_" Then ws.Visible = xlSheetVisible
    Next ws
    MsgBox "Planilhas de dados agora visiveis.", vbInformation, "Wizped"
End Sub

' Esconder todas as planilhas BD_*
Public Sub OnEsconderPlanilhas(control As IRibbonControl)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 3) = "BD_" Then ws.Visible = xlSheetVeryHidden
    Next ws
    MsgBox "Planilhas de dados ocultadas.", vbInformation, "Wizped"
End Sub
