Option Explicit
Dim Folder_Standart As String
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function


Sub Block_Dedication()

'Codigo para atualizacao do block Dedication, funcao de importar logs brutos, organizar os dados e inserir a formula de dT

'Workbooks que serao usados
Dim Control_Book As Workbook
Dim Source_Book As Workbook

'Worksheets que serao usados
Dim Source_Sheet As Worksheet
Dim Logs_Sheet As Worksheet
Dim Gabarito_Sheet As Worksheet
Dim Resultado_Sheet As Worksheet
Dim Grafico_Sheet As Worksheet
Dim Aluno_Sheet As Worksheet
Dim Planilha As Worksheet

'A localização do Log no PC
Dim Logs_Path As FileDialog

'Graficos
Dim Graf_Obj As ChartObject
Dim Controle_Obj As ChartObject
Dim Controle_chart As Chart

'Variaveis
Dim Atividade As Variant
Dim dTMax As Variant
Dim Atividades_Analisadas As Double
Dim Atividade_Total As Variant
Dim Alunos As Variant
Dim Nome As Variant
Dim Dedicacao_Especifica As Variant
Dim Dedicacao_Especifica_Total As Variant
Dim Dedicacao_Geral As Variant
Dim Max_Dedicacao_Especifica As Double
Dim Max_Dedicacao_Geral As Double
Dim Media As Double
Dim Amostra As Double

'Index
Dim Source_Row As Long
Dim Atividade_Row As Long

'Contador de loops
Dim i As Long
Dim j As Long
Dim k As Long
Dim t As Long
Dim m As Long
Dim n As Long

'Dicionarios
Dim dTMax_Dict As New Scripting.Dictionary
Dim Dedicacao_Especifica_Total_Dict As New Scripting.Dictionary
Dim Dedicacao_Geral_Dict As New Scripting.Dictionary
Dim Media_Atividade_Dict As New Scripting.Dictionary
Dim Soma_DesvP_Dict As New Scripting.Dictionary

'Arrays
Dim Alunos_Lista() As Variant
Dim Atividades_Lista() As Variant

'Timer do início
Dim StartTime As Double

'Desativar atualizacao de tela, alertas e muda o modo de calculo pra manual
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'Define a planilha controle e a aba dados
Set Control_Book = ActiveWorkbook
Set Logs_Sheet = Control_Book.Sheets("Logs")
Set Gabarito_Sheet = Control_Book.Sheets("Gabarito")

Atividades_Analisadas = Gabarito_Sheet.Cells(Rows.Count, 1).End(xlUp).Row

If Atividades_Analisadas = 1 Then
    MsgBox "Sem Atividades no Gabarito!", vbCritical, "Block Dedication"
    GoTo Cancel
End If

ReDim Atividades_Lista(1 To Atividades_Analisadas - 2)

For i = 2 To Atividades_Analisadas

    Atividade = Gabarito_Sheet.Cells(i, 1).Value2
    dTMax = Gabarito_Sheet.Cells(i, 2).Value2
    
    If Atividade <> "Dedica" & ChrW(231) & ChrW(227) & "o Geral Independente da Atividade estar no Gabarito" Then
        Atividades_Lista(i - 1) = Atividade
    End If
    
    If dTMax = 0 Or dTMax = "" Then
        MsgBox "Operation Cancelled!" & "Insira valor de dT", vbCritical, "Block Dedication"
        GoTo Cancel
    Else
        dTMax_Dict(Atividade) = Gabarito_Sheet.Cells(i, 2).Value2
    End If
    
Next i

'Pede ao usuario para selecionar o log de atividades
Set Logs_Path = Application.FileDialog(msoFileDialogFilePicker)

Folder_Standart = Gabarito_Sheet.Cells(1, 9).Value2

With Logs_Path
    .AllowMultiSelect = False
    .Title = "Selecionar Log de Atividades"
    If IsEmpty(Folder_Standart) = True Then
        .InitialFileName = "C:\*"
    Else
        .InitialFileName = Folder_Standart
    End If
    .Filters.Clear
    .Filters.Add "Excel files", "*.xls;*.xlsx;*.xlsm,*.xlsb"
    .Filters.Add "All files", "*.*"
End With

'Caso tenha sido selecionado, seguir com o codigo
If Logs_Path.Show = -1 Then

    'Inicia o timer do código
    StartTime = Timer
    'Define a planilha log e a aba unica
    Set Source_Book = Workbooks.Open(Logs_Path.SelectedItems(1))
    Set Source_Sheet = Source_Book.Worksheets(1)
    
    Source_Row = Source_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Limpar a aba de dados
    t = Logs_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    Logs_Sheet.Cells(1, 1).Value2 = "Hora"
    Logs_Sheet.Cells(1, 2).Value2 = "dT atividade"
    Logs_Sheet.Cells(1, 3).Value2 = "dT Geral"
    Logs_Sheet.Cells(1, 4).Value2 = "Nome completo"
    Logs_Sheet.Cells(1, 5).Value2 = "Usuário afetado"
    Logs_Sheet.Cells(1, 6).Value2 = "Contexto do Evento"
    Logs_Sheet.Cells(1, 7).Value2 = "Componente"
    Logs_Sheet.Cells(1, 8).Value2 = "Nome do evento"
    Logs_Sheet.Cells(1, 9).Value2 = "Descri" & ChrW(231) & ChrW(227) & "o"
    Logs_Sheet.Cells(1, 10).Value2 = "Origem"
    Logs_Sheet.Cells(1, 11).Value2 = "Endereço IP"
    
    If t > 1 Then
        Logs_Sheet.Rows("2:" & t).Delete
    End If
    
    'Transpor as informacoes de uma planilha pra outra
    Logs_Sheet.Range("A2:A" & Source_Row).FormulaLocal = Source_Sheet.Range("A2:A" & Source_Row).Value
    Logs_Sheet.Range("D2:K" & Source_Row).Value2 = Source_Sheet.Range("B2:I" & Source_Row).Value2
    
    'Marcar o local
    Gabarito_Sheet.Cells(1, 9).Value2 = Trim(Left(Logs_Path.SelectedItems(1), InStrRev(Logs_Path.SelectedItems(1), "\")))
    
    'Fecha a planilha de logs
    Source_Book.Close SaveChanges:=False
    
Else

    GoTo Cancel
    
End If

'Limpar sheets de atividades passadas
For Each Planilha In Control_Book.Sheets
    If Planilha.Name <> "Logs" And Planilha.Name <> "Gabarito" Then
        Planilha.Delete
    End If
Next Planilha

'Ativar filtros

    If Not Logs_Sheet.AutoFilterMode Then
        Logs_Sheet.Range("A1").AutoFilter
    End If
    
    With Logs_Sheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("A1:A" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("F1:F" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("D1:D" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Adicionar a formula de dt

For i = 2 To Source_Row
    Atividade = Logs_Sheet.Cells(i, 6).Value2
    
    If dTMax_Dict.Exists(Atividade) Then
        dTMax = dTMax_Dict(Atividade)
    Else
        dTMax = dTMax_Dict("Dedica" & ChrW(231) & ChrW(227) & "o Geral Independente da Atividade estar no Gabarito")
    End If
    
    Logs_Sheet.Cells(i, 2).FormulaR1C1 = "=IFERROR(IF(AND(RC[2]=R[-1]C[2],RC[4]=R[-1]C[4],RC[-1]-R[-1]C[-1]>0,RC[-1]-R[-1]C[-1]<TIME(0," & dTMax & ",0)),RC[-1]-R[-1]C[-1],0),0)"
    Logs_Sheet.Cells(i, 3).FormulaR1C1 = "=IFERROR(IF(AND(RC[1]=R[-1]C[1],RC[-2]-R[-1]C[-2]>0,RC[-2]-R[-1]C[-2]<TIME(0," & dTMax_Dict("Dedica" & ChrW(231) & ChrW(227) & "o Geral Independente da Atividade estar no Gabarito") & ",0)),RC[-2]-R[-1]C[-2],0),0)"
Next i

'Criar a lista de alunos
Alunos = 0
ReDim Alunos_Lista(0)

For i = 2 To Source_Row
    Nome = Logs_Sheet.Cells(i, 4).Value2
    If IsInArray(Nome, Alunos_Lista) = False Then
        Alunos = Alunos + 1
        ReDim Preserve Alunos_Lista(Alunos)
        Alunos_Lista(Alunos) = Nome
    End If
Next i

'Para cada aluno, preencher os dados de cada atividade
For i = 1 To UBound(Alunos_Lista)
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = Left(Alunos_Lista(i), 30)
    Set Aluno_Sheet = ActiveSheet
    Aluno_Sheet.Cells(1, 1).Value2 = "Atividade"
    Aluno_Sheet.Cells(1, 2).Value2 = "Dedica" & ChrW(231) & ChrW(227) & "o (min)"
    Dedicacao_Especifica = 0
    Dedicacao_Especifica_Total = 0
    Dedicacao_Geral = 0
    
    'Organizar por atividade
    With Logs_Sheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("A1:A" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("F1:F" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("D1:D" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Logs_Sheet.UsedRange.Columns("B:B").Calculate
    
    'Loop para dedicacao especifica
    For j = LBound(Atividades_Lista) To UBound(Atividades_Lista)
        m = j + 1   'Variavel para facilitar o loop
        Aluno_Sheet.Cells(m, 1).Value2 = Atividades_Lista(j) 'O nome da atividade
        Atividade = Aluno_Sheet.Cells(m, 1).Value2
        Dedicacao_Especifica = Application.WorksheetFunction.SumIfs(Logs_Sheet.Range("$B:$B"), Logs_Sheet.Range("$D:$D"), Alunos_Lista(i), Logs_Sheet.Range("$F:$F"), Atividade) * 1440
        Aluno_Sheet.Cells(m, 2).Value2 = Dedicacao_Especifica
        Aluno_Sheet.Cells(m, 2).NumberFormat = "0"
        Media_Atividade_Dict(Atividade) = Media_Atividade_Dict(Atividade) + Dedicacao_Especifica
        Dedicacao_Especifica_Total = Dedicacao_Especifica_Total + Dedicacao_Especifica
    Next j
    
    'Organizar para analise geral
    With Logs_Sheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("F1:F" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("A1:A" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("D1:D" & Source_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Logs_Sheet.UsedRange.Columns("C:C").Calculate
    
    Aluno_Sheet.Cells(Atividades_Analisadas, 1).Value2 = "Total das atividades acima"
    Aluno_Sheet.Cells(Atividades_Analisadas, 2).Value2 = Dedicacao_Especifica_Total
    Aluno_Sheet.Cells(Atividades_Analisadas, 2).NumberFormat = "0"
    Dedicacao_Especifica_Total_Dict(Alunos_Lista(i)) = Dedicacao_Especifica_Total
    Aluno_Sheet.Cells(Atividades_Analisadas + 1, 1).Value2 = "Dedica" & ChrW(231) & ChrW(227) & "o Geral Independente da Atividade estar no Gabarito"
    Dedicacao_Geral = Application.WorksheetFunction.SumIfs(Logs_Sheet.Range("$C:$C"), Logs_Sheet.Range("$D:$D"), Alunos_Lista(i)) * 1440
    Aluno_Sheet.Cells(Atividades_Analisadas + 1, 2).Value2 = Dedicacao_Geral
    Aluno_Sheet.Cells(Atividades_Analisadas + 1, 2).NumberFormat = "0"
    Dedicacao_Geral_Dict(Alunos_Lista(i)) = Dedicacao_Geral
    
    Aluno_Sheet.Columns("A:B").EntireColumn.AutoFit
Next i

'Organizar as abas
For i = 1 To Sheets.Count
   For j = 1 To Sheets.Count - 1
         If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
            Sheets(j).Move After:=Sheets(j + 1)
         End If
   Next j
Next i
Logs_Sheet.Move Before:=Control_Book.Sheets(1)
Gabarito_Sheet.Move After:=Control_Book.Sheets(1)

'Preencher a tabela de resultados
Sheets.Add(After:=Sheets("Gabarito")).Name = "Resultado"
Set Resultado_Sheet = ActiveSheet
Resultado_Sheet.Cells(1, 1).Value2 = "Aluno"
Resultado_Sheet.Cells(1, 2).Value2 = "Tempo Dedicado Total Espec" & ChrW(237) & "fico (min)"
Resultado_Sheet.Cells(1, 3).Value2 = "Tempo Dedicado Total Relativo (%)"
Resultado_Sheet.Cells(1, 4).Value2 = "Tempo Dedicado Geral Absoluto (min)"
Resultado_Sheet.Cells(1, 5).Value2 = "Tempo Dedicado Geral Relativo (%)"

For i = 1 To UBound(Alunos_Lista)
    n = i + 1
    Resultado_Sheet.Cells(n, 1).Value2 = Alunos_Lista(i)
    Resultado_Sheet.Cells(n, 2).Value2 = Dedicacao_Especifica_Total_Dict(Alunos_Lista(i))
    Resultado_Sheet.Cells(n, 2).NumberFormat = "0"
    Resultado_Sheet.Cells(n, 4).Value2 = Dedicacao_Geral_Dict(Alunos_Lista(i))
    Resultado_Sheet.Cells(n, 4).NumberFormat = "0"
Next i

Max_Dedicacao_Especifica = Application.WorksheetFunction.Max(Resultado_Sheet.Range("B2:B" & UBound(Alunos_Lista) + 1))
Max_Dedicacao_Geral = Application.WorksheetFunction.Max(Resultado_Sheet.Range("D2:D" & UBound(Alunos_Lista) + 1))

For i = 1 To UBound(Alunos_Lista)
    n = i + 1
    Resultado_Sheet.Cells(n, 1).Value2 = Alunos_Lista(i)
    Resultado_Sheet.Cells(n, 3).Value2 = Dedicacao_Especifica_Total_Dict(Alunos_Lista(i)) / Max_Dedicacao_Especifica
    Resultado_Sheet.Cells(n, 3).Style = "Percent"
    Resultado_Sheet.Cells(n, 3).NumberFormat = "0.00%"
    Resultado_Sheet.Cells(n, 5).Value2 = Dedicacao_Geral_Dict(Alunos_Lista(i)) / Max_Dedicacao_Geral
    Resultado_Sheet.Cells(n, 5).Style = "Percent"
    Resultado_Sheet.Cells(n, 5).NumberFormat = "0.00%"
Next i

Resultado_Sheet.Range("A1").AutoFilter

With Resultado_Sheet.AutoFilter.Sort
    .SortFields.Clear
    .SortFields.Add2 Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Resultado_Sheet.Columns("A:E").EntireColumn.AutoFit

'Calcular Media e Desvio Padrao
Gabarito_Sheet.Activate
For i = 2 To Atividades_Analisadas - 1
    Atividade = Gabarito_Sheet.Cells(i, 1).Value2
    Media = Media_Atividade_Dict(Atividade) / Atividades_Analisadas
    Gabarito_Sheet.Cells(i, 3).Value2 = Media
    Gabarito_Sheet.Cells(i, 3).NumberFormat = "0"
    
    For j = 1 To UBound(Alunos_Lista)
    Alunos = Alunos_Lista(j)
    Set Aluno_Sheet = Control_Book.Sheets(Left(Alunos_Lista(j), 30))
    Atividade_Row = Aluno_Sheet.Range("A:A").Find(What:=Atividade, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
    Amostra = Aluno_Sheet.Cells(Atividade_Row, 2).Value2
    Soma_DesvP_Dict(Atividade) = Soma_DesvP_Dict(Atividade) + (Amostra - Media) ^ 2
    Next j
    
    Gabarito_Sheet.Cells(i, 4).Value2 = (Soma_DesvP_Dict(Atividade) / (Atividades_Analisadas - 2)) ^ (1 / 2)
    Gabarito_Sheet.Cells(i, 4).NumberFormat = "0"
    
Next i

Gabarito_Sheet.Cells(Atividades_Analisadas, 3).Value2 = Application.WorksheetFunction.Average(Resultado_Sheet.Range(Resultado_Sheet.Cells(2, 4), Resultado_Sheet.Cells(UBound(Alunos_Lista) + 1, 4)))
Gabarito_Sheet.Cells(Atividades_Analisadas, 3).NumberFormat = "0"
Gabarito_Sheet.Cells(Atividades_Analisadas, 4).Value2 = Application.WorksheetFunction.StDev(Resultado_Sheet.Range(Resultado_Sheet.Cells(2, 4), Resultado_Sheet.Cells(UBound(Alunos_Lista) + 1, 4)))
Gabarito_Sheet.Cells(Atividades_Analisadas, 4).NumberFormat = "0"

'----------Graficos

Sheets.Add(After:=Sheets("Resultado")).Name = "Grafico"
Set Grafico_Sheet = ActiveSheet

'Apagar grafico anterior
For Each Graf_Obj In Grafico_Sheet.ChartObjects
    Graf_Obj.Delete
Next

Set Controle_Obj = Grafico_Sheet.ChartObjects.Add(Top:=0, Left:=0, Width:=800, Height:=500)
Controle_Obj.Chart.SetSourceData Sheets("Gabarito").Range(Sheets("Gabarito").Cells(1, 1), Sheets("Gabarito").Cells(Atividades_Analisadas, 4))
Set Controle_chart = Grafico_Sheet.ChartObjects(1).Chart
Controle_chart.ChartType = xlColumnClustered
Controle_chart.HasTitle = True
Controle_chart.ChartTitle.Text = "Dedicação por Atividade"

'Voltar a aba de resumo
Resultado_Sheet.Activate

'Ativar atualizacao de tela, alertas e muda o modo de calculo pra manual
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

'Exibir uma mensagem de operação ok e o tempo que se passou

MsgBox "Operation Successful!" & vbNewLine & "Run Time: " & Format((Timer - StartTime) / 86700, "hh:mm:ss"), vbInformation, "Block Dedication"

Exit Sub

Cancel:

'Ativar atualizacao de tela, alertas e muda o modo de calculo pra manual
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

'Exibir uma mensagem de erro

MsgBox "Operation Cancelled!", vbCritical, "Block Dedication"

End Sub
