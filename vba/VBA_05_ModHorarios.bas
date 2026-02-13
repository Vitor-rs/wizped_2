' Attribute VB_Name = "Mod_Horarios"
Option Explicit

' Função para retornar um Array de horários disponíveis dado um dia da semana
' Usado para popular ComboBox dependente
Public Function GetHorariosDisponiveis(ByVal DiaSemana As String) As Variant
    Dim tbl As ListObject
    Dim colDia As ListColumn
    Dim colHora As ListColumn
    Dim arrHoras() As String
    Dim i As Long, count As Long
    Dim cellDia As Range, cellHora As Range
    
    ' Inicializa
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets("BD_Horarios").ListObjects("Tbl_Horarios")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        GetHorariosDisponiveis = Array("Erro: Tbl_Horarios não encontrada")
        Exit Function
    End If
    
    ' Tenta encontrar a coluna do dia (Ex: "2ª", "3ª", "Sáb")
    ' A string DiaSemana deve bater com o cabeçalho da tabela
    On Error Resume Next
    Set colDia = tbl.ListColumns(DiaSemana)
    On Error GoTo 0
    
    If colDia Is Nothing Then
        GetHorariosDisponiveis = Array("Erro: Dia '" & DiaSemana & "' não encontrado")
        Exit Function
    End If
    
    Set colHora = tbl.ListColumns("Hora")
    ReDim arrHoras(0 To tbl.ListRows.count - 1)
    count = 0
    
    ' Loop para filtrar
    For i = 1 To tbl.ListRows.count
        Set cellDia = colDia.DataBodyRange.Cells(i, 1)
        Set cellHora = colHora.DataBodyRange.Cells(i, 1)
        
        ' Verifica se tem "X" ou valor (não vazio)
        If UCase(Trim(cellDia.Value)) = "X" Then
            arrHoras(count) = Format(cellHora.Value, "hh:mm") ' Formata a hora
            count = count + 1
        End If
    Next i
    
    ' Redimensiona para o tamanho real
    If count > 0 Then
        ReDim Preserve arrHoras(0 To count - 1)
        GetHorariosDisponiveis = arrHoras
    Else
        GetHorariosDisponiveis = Array("Nenhum horário disponível")
    End If
End Function

' Teste Imediato (pode rodar no VBE)
Sub TesteHorario()
    Dim horas As Variant
    Dim h As Variant
    
    horas = GetHorariosDisponiveis("2ª")
    
    Debug.Print "--- Horários para 2ª ---"
    For Each h In horas
        Debug.Print h
    Next h
End Sub
