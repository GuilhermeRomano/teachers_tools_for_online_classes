from vb2py.vbfunctions import *
from vb2py.vbdebug import *



def Import():
    Control_Book = Workbook()

    Logs_Book = Workbook()

    Dados_Sheet = Worksheet()

    Logs_Sheet = Worksheet()

    Control_Sheet = Worksheet()

    DataH_Sheet = Worksheet()

    Grafico_Sheet = Worksheet()

    Log_Loc = String()

    log_row = Long()

    resumo_row = Long()

    Activ_row = Long()

    Graf_Obj = ChartObject()

    Controle_Obj = ChartObject()

    Controle_chart = Chart()

    dtMAX = Long()

    aluno_test = Variant()

    ativ_test = Variant()

    Vencedor = Variant()

    Tempo_Conclusao = Variant()

    Atividade = Variant()

    Evento = Variant()

    Tempo_aluno = Variant()

    Contexto = Variant()

    Diferenca = Variant()

    planilha = Worksheet()

    Horario = Double()

    t = Long()

    k = Long()

    i = Long()

    j = Long()

    l = Long()

    Lista = Range()

    StartTime = Double()
    #Codigo para atualizacao do block Dedication, funcao de importar logs brutos, organizar os dados e inserir a formula de dT
    #Workbooks que serao usados
    #Worksheets que serao usados
    #A localização do Log no PC
    #Numero de linhas pra abas usadas
    #Graficos
    #Valor do dT max em minutos
    #Contador de alunos e atividades
    #variaveis genericas
    #indices para coluna e linha
    #Contador de loops
    #range para remoção
    #Timer do início
    #Desativar atualizacao de tela, alertas e muda o modo de calculo pra manual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    #Define a planilha controle e a aba dados
    Control_Book = ActiveWorkbook
    Dados_Sheet = Control_Book.Sheets('Dados')
    Control_Sheet = Control_Book.Sheets('Resumo')
    DataH_Sheet = Control_Book.Sheets('Data Hidden')
    Grafico_Sheet = Control_Book.Sheets('Grafico')
    dtMAX = DataH_Sheet.Cells(1, 5)
    #Tira os filtros ativos
    #If (Dados_Sheet.AutoFilterMode And Dados_Sheet.FilterMode) Or Dados_Sheet.FilterMode Then
    #    Dados_Sheet.ShowAllData
    #End If
    #--------------Inicio do import
    #Pede ao usuario para selecionar o log
    Log_Loc = Application.GetOpenFilename('Excel compatible Files (*.xls;*.csv;*.xlsx;*.xlsb), *.xls;*.csv;*.xlsx;*.xlsb', VBGetMissingArgument(Application.GetOpenFilename, 1), 'Selecione o log')
    if Log_Loc == 'False':
        GoTo(Cancel)
    #Inicia o timer do código
    StartTime = Timer()
    #Define a planilha log e a aba unica
    Logs_Book = Workbooks.Open(Log_Loc)
    Logs_Sheet = Logs_Book.Worksheets(1)
    log_row = Logs_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    #Limpar a aba de dados
    t = Dados_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dados_Sheet.Cells[1, 1].Value2 = 'Hora'
    Dados_Sheet.Cells[1, 2].Value2 = 'Dt'
    Dados_Sheet.Cells[1, 3].Value2 = 'Nome completo'
    Dados_Sheet.Cells[1, 4].Value2 = 'Usuário afetado'
    Dados_Sheet.Cells[1, 5].Value2 = 'Contexto do Evento'
    Dados_Sheet.Cells[1, 6].Value2 = 'Componente'
    Dados_Sheet.Cells[1, 7].Value2 = 'Nome do evento'
    Dados_Sheet.Cells[1, 8].Value2 = 'Descrição'
    Dados_Sheet.Cells[1, 9].Value2 = 'Origem'
    Dados_Sheet.Cells[1, 10].Value2 = 'Endereço IP'
    if t > 1:
        Dados_Sheet.Rows('2:' + t).Delete()
    #Transpor as informacoes de uma planilha pra outra
    Dados_Sheet.Range['A2:A' + log_row].FormulaLocal = Logs_Sheet.Range('A2:A' + log_row).Value
    Dados_Sheet.Range['C2:J' + log_row].Value2 = Logs_Sheet.Range('B2:I' + log_row).Value2
    #Fecha a planilha de logs
    Logs_Book.Close(False)
    #limpa sheets de atividades passadas
    for planilha in Control_Book.Sheets:
        if Left(planilha.Name, 5) == 'Ativ ':
            planilha.Delete()
    #Ativar filtros
    if not Dados_Sheet.AutoFilterMode:
        Dados_Sheet.Range('A1').AutoFilter()
    Dados_Sheet.AutoFilter.Sort.SortFields.Clear()
    Dados_Sheet.AutoFilter.Sort.SortFields.Add2(Key=Range('A1:A' + log_row), SortOn=xlSortOnValues, Order=xlAscending, DataOption=xlSortNormal)
    with_variable0 = ActiveWorkbook.Worksheets('Dados').AutoFilter.Sort
    with_variable0.Header = xlYes
    with_variable0.MatchCase = False
    with_variable0.Orientation = xlTopToBottom
    with_variable0.SortMethod = xlPinYin
    with_variable0.Apply()
    Dados_Sheet.AutoFilter.Sort.SortFields.Clear()
    Dados_Sheet.AutoFilter.Sort.SortFields.Add2(Key=Range('E1:E' + log_row), SortOn=xlSortOnValues, Order=xlAscending, DataOption=xlSortNormal)
    with_variable1 = Dados_Sheet.AutoFilter.Sort
    with_variable1.Header = xlYes
    with_variable1.MatchCase = False
    with_variable1.Orientation = xlTopToBottom
    with_variable1.SortMethod = xlPinYin
    with_variable1.Apply()
    Dados_Sheet.AutoFilter.Sort.SortFields.Clear()
    Dados_Sheet.AutoFilter.Sort.SortFields.Add2(Key=Range('C1:C' + log_row), SortOn=xlSortOnValues, Order=xlAscending, DataOption=xlSortNormal)
    with_variable2 = Dados_Sheet.AutoFilter.Sort
    with_variable2.Header = xlYes
    with_variable2.MatchCase = False
    with_variable2.Orientation = xlTopToBottom
    with_variable2.SortMethod = xlPinYin
    with_variable2.Apply()
    #Adicionar a formula de dt
    Dados_Sheet.Range['B2'].Value2 = 0
    Dados_Sheet.Range['B3:B' + log_row].FormulaR1C1 = '=IF(AND(RC[1]=R[-1]C[1],RC[3]=R[-1]C[3],RC[-1]-R[-1]C[-1]>0,RC[-1]-R[-1]C[-1]<TIME(0,' + dtMAX + ',0)),RC[-1]-R[-1]C[-1],0)'
    #-----------Inicio do controle
    #Limpar dados anteriores
    t = Control_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    k = Control_Sheet.Cells(3, Columns.Count).End(xlToLeft).Column
    Control_Sheet.Rows('3:' + t).Delete()
    Control_Sheet.Cells[3, 1].Value2 = 'Nome'
    #Adicionar alunos novos
    t = Control_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    for i in vbForRange(2, log_row):
        aluno_test = Application.VLookup(Dados_Sheet.Cells(i, 3).Value2, Control_Sheet.Range('A:A'), 1, False)
        if IsError(aluno_test) == True:
            Control_Sheet.Rows(t + 1).Insert(Shift=xlShiftDown)
            Control_Sheet.Cells[t + 1, 1].Value2 = Dados_Sheet.Cells(i, 3).Value2
            t = t + 1
    t = Control_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    Control_Sheet.Cells[t + 1, 1].Value2 = 'MEDIA'
    Control_Sheet.Cells[t + 2, 1].Value2 = 'DESVIO PADRAO'
    Control_Sheet.Cells[3, 2].Value2 = 'TOTAL'
    #Limpa atividades antigas da legenda
    DataH_Sheet.Activate()
    Rows('2:2').Select()
    Range(Rows('2:2'), Rows('2:2').End(xlDown)).Delete(Shift=xlUp)
    Selection.Delete(Shift=xlUp)
    #Adiciona atividades novas
    Control_Sheet.Activate()
    k = Control_Sheet.Cells(3, Columns.Count).End(xlToLeft).Column
    for i in vbForRange(2, log_row):
        ativ_test = Application.VLookup(Dados_Sheet.Cells(i, 5).Value2, DataH_Sheet.Range('B:B'), 1, False)
        if IsError(ativ_test) == True:
            Control_Sheet.Cells[3, k + 1].Value2 = k - 1
            DataH_Sheet.Cells[k, 1].Value2 = k - 1
            DataH_Sheet.Cells[k, 2].Value2 = Dados_Sheet.Cells(i, 5).Value2
            k = k + 1
    #Inserir fórmulas
    t = Control_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    k = Control_Sheet.Cells(3, Columns.Count).End(xlToLeft).Column
    #valores
    Control_Sheet.Range[Cells(4, 3), Cells(t - 2, k)].FormulaR1C1 = '=MINUTE(SUMIFS(Dados!C2,Dados!C2,"<>#VALUE!",Dados!C2,"<>#VALOR!",Dados!C3,Resumo!RC1,Dados!C5,VLOOKUP(Resumo!R3C,\'Data Hidden\'!C1:C2,2,FALSE)))'
    #Total
    Control_Sheet.Range[Cells(4, 2), Cells(t - 2, 2)].FormulaR1C1 = '=SUM(RC[1]:RC[' + k - 2 + '])'
    #Media
    Control_Sheet.Range[Cells(t - 1, 2), Cells(t - 1, k)].FormulaR1C1 = '=AVERAGE(R4C:R' + t - 2 + 'C)'
    #Desvio Padrão
    Control_Sheet.Range[Cells(t, 2), Cells(t, k)].FormulaR1C1 = '=STDEV.P(R4C:R' + t - 2 + 'C)'
    #----------Graficos
    #Apagar grafico anterior
    Grafico_Sheet.Activate()
    for Graf_Obj in Grafico_Sheet.ChartObjects:
        Graf_Obj.Delete()
    Controle_Obj = Grafico_Sheet.ChartObjects.Add(Top= 0, Left= 0, Width= 800, Height= 500)
    Controle_Obj.Chart.SetSourceData(Sheets('Resumo').Range(Sheets('Resumo').Cells(3, 1), Sheets('Resumo').Cells(t - 2, k)))
    Controle_chart = Grafico_Sheet.ChartObjects(1).Chart
    Controle_chart.ChartType = xlColumnClustered
    Controle_chart.HasTitle = True
    Controle_chart.ChartTitle.Text = 'Dedicação por Atividade por Aluno'
    #Formatação
    Control_Sheet.Activate()
    #Nome (cabecalho)
    with_variable3 = Cells(3, 1).Interior
    with_variable3.Pattern = xlSolid
    with_variable3.PatternColorIndex = xlAutomatic
    with_variable3.ThemeColor = xlThemeColorAccent4
    with_variable3.TintAndShade = - 0.249977111117893
    with_variable3.PatternTintAndShade = 0
    #Nomes
    with_variable4 = Range(Cells(4, 1), Cells(t - 2, 1)).Interior
    with_variable4.Pattern = xlSolid
    with_variable4.PatternColorIndex = xlAutomatic
    with_variable4.ThemeColor = xlThemeColorAccent4
    with_variable4.TintAndShade = 0.399975585192419
    with_variable4.PatternTintAndShade = 0
    #Atividades (cabecalho)
    with_variable5 = Range(Cells(3, 3), Cells(3, k)).Interior
    with_variable5.Pattern = xlSolid
    with_variable5.PatternColorIndex = xlAutomatic
    with_variable5.ThemeColor = xlThemeColorAccent1
    with_variable5.TintAndShade = - 0.249977111117893
    with_variable5.PatternTintAndShade = 0
    #Valores
    with_variable6 = Range(Cells(4, 3), Cells(t - 2, k)).Interior
    with_variable6.Pattern = xlSolid
    with_variable6.PatternColorIndex = xlAutomatic
    with_variable6.ThemeColor = xlThemeColorAccent1
    with_variable6.TintAndShade = 0.399975585192419
    with_variable6.PatternTintAndShade = 0
    #media e desvio (cabecalhos)
    with_variable7 = Range(Cells(t - 1, 1), Cells(t, 1)).Interior
    with_variable7.Pattern = xlSolid
    with_variable7.PatternColorIndex = xlAutomatic
    with_variable7.ThemeColor = xlThemeColorAccent6
    with_variable7.TintAndShade = - 0.249977111117893
    with_variable7.PatternTintAndShade = 0
    #medias e desvios
    with_variable8 = Range(Cells(t - 1, 2), Cells(t, k)).Interior
    with_variable8.Pattern = xlSolid
    with_variable8.PatternColorIndex = xlAutomatic
    with_variable8.ThemeColor = xlThemeColorAccent6
    with_variable8.TintAndShade = 0.399975585192419
    with_variable8.PatternTintAndShade = 0
    #total
    with_variable9 = Cells(3, 2).Interior
    with_variable9.Pattern = xlSolid
    with_variable9.PatternColorIndex = xlAutomatic
    with_variable9.ThemeColor = xlThemeColorAccent6
    with_variable9.TintAndShade = - 0.249977111117893
    with_variable9.PatternTintAndShade = 0
    #valores totais
    with_variable10 = Range(Cells(4, 2), Cells(t, 2)).Interior
    with_variable10.Pattern = xlSolid
    with_variable10.PatternColorIndex = xlAutomatic
    with_variable10.ThemeColor = xlThemeColorAccent6
    with_variable10.TintAndShade = 0.399975585192419
    with_variable10.PatternTintAndShade = 0
    #Achar os vencedores
    #Loop para cada atividade
    for i in vbForRange(2, k - 1):
        Vencedor = 'Ninguem'
        Tempo_Conclusao = 47484
        Atividade = DataH_Sheet.Cells(i, 2).Value2
        #loop pelos logs para comparar valores de tempo de conclusão
        for j in vbForRange(2, log_row):
            Evento = Dados_Sheet.Cells(j, 7).Value2
            Contexto = Dados_Sheet.Cells(j, 5).Value2
            Tempo_aluno = CVDate(Dados_Sheet.Cells(j, 1).Value2)
            if Evento == 'Conclus' + ChrW(227) + 'o da atividade do curso atualizada' and Contexto == Atividade and Tempo_aluno < Tempo_Conclusao:
                Vencedor = Dados_Sheet.Cells(j, 3).Value2
                Tempo_Conclusao = Tempo_aluno
        if Vencedor != 'Ninguem':
            #caso exista um vencedor, pintar no controle que ele eh vencedor
            with_variable11 = Control_Sheet.Cells(Control_Sheet.Range('A:A').Find(Vencedor, LookIn= xlValues).Row, i + 1).Interior
            with_variable11.Pattern = xlSolid
            with_variable11.PatternColorIndex = xlAutomatic
            with_variable11.Color = 49407
            with_variable11.TintAndShade = 0
            with_variable11.PatternTintAndShade = 0
            with_variable12 = Control_Sheet.Cells(Control_Sheet.Range('A:A').Find(Vencedor, LookIn= xlValues).Row, i + 1).Font
            with_variable12.ColorIndex = xlAutomatic
            with_variable12.TintAndShade = 0
            with_variable12.Bold = True
            #caso exista um vencedor deve ser criada uma nova aba, mostrando quem concluiu e o atraso em relacao ao campeao
            #criacao da aba
            Sheets.Add[After= Sheets(Sheets.Count)].Name = 'Ativ ' + i - 1
            ActiveSheet.Cells[1, 1].Value2 = 'NOME'
            ActiveSheet.Cells[1, 2].Value2 = 'HORARIO'
            ActiveSheet.Cells[1, 3].Value2 = 'ATRASO'
            #trazer dos logs quem completou
            l = 2
            for j in vbForRange(2, log_row):
                Evento = Dados_Sheet.Cells(j, 7).Value2
                Contexto = Dados_Sheet.Cells(j, 5).Value2
                Tempo_aluno = CVDate(Dados_Sheet.Cells(j, 1).Value2)
                if Evento == 'Conclus' + ChrW(227) + 'o da atividade do curso atualizada' and Contexto == Atividade:
                    ActiveSheet.Cells[l, 1] = Dados_Sheet.Cells(j, 3).Value2
                    ActiveSheet.Cells[l, 2] = Tempo_aluno
                    l = l + 1
            #remover duplicatas
            Activ_row = ActiveSheet.Range('A' + Rows.Count).End(xlUp).Row
            Lista = ActiveSheet.Range('A1:C' + Activ_row)
            Lista.RemoveDuplicates(Columns=1, Header=xlYes)
            #Por filtro e ordernar
            ActiveSheet.Range('A1').Select()
            Selection.AutoFilter()
            ActiveSheet.AutoFilter.Sort.SortFields.Clear()
            ActiveSheet.AutoFilter.Sort.SortFields.Add2(Key=Range('B1:B' + Activ_row), SortOn=xlSortOnValues, Order=xlAscending, DataOption=xlSortNormal)
            with_variable13 = ActiveSheet.AutoFilter.Sort
            with_variable13.Header = xlYes
            with_variable13.MatchCase = False
            with_variable13.Orientation = xlTopToBottom
            with_variable13.SortMethod = xlPinYin
            with_variable13.Apply()
            #Definir os atrasos
            Activ_row = ActiveSheet.Range('A' + Rows.Count).End(xlUp).Row
            ActiveSheet.Cells[2, 3] = '-'
            if Activ_row > 2:
                for l in vbForRange(3, Activ_row):
                    Diferenca = DateDiff('n', ActiveSheet.Cells(2, 2).Value2, ActiveSheet.Cells(l, 2).Value2)
                    ActiveSheet.Cells[l, 3].Value2 = Diferenca + ' minutos'
            Columns('A:C').EntireColumn.AutoFit()
    #Voltar a aba de resumo
    Control_Sheet.Activate()
    #Ativar atualizacao de tela, alertas e muda o modo de calculo pra manual
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    #Exibir uma mensagem de operação ok e o tempo que se passou
    MsgBox('Operation Successful!' + vbNewLine + 'Run Time: ' + Format(( Timer() - StartTime )  / 86700, 'hh:mm:ss'), vbInformation, 'Block Dedication')
    return
    #Ativar atualizacao de tela, alertas e muda o modo de calculo pra manual
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    #Exibir uma mensagem de erro
    MsgBox('Operation Cancelled!', vbCritical, 'Block Dedication')
