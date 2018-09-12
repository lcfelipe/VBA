'@author=jonathan_conzatti
'Macro que verifica quanto tempo "útil" se passou entre duas datas
'Exemplo: verificar quanto tempo se passou entre duas datas considerando somente as horas/minutos entre 08:30 e 17:30
'Recebe:
'   Data Inicial e Data Final
'Devolve:
'   Devolve em minutos a diferenca entre as datas
'Modo de uso
'   na planilha em excel colocar as duas datas
'Conceitos importantes: manipulação de datas, intervalos, constantes
Function CalculaComHorario(dtInicial As Date, dtFinal As Date)
    Const HoraInicio = 8
    Const minutoInicio = 30
    Const HoraTermino = 17
    Const minutoTermino = 30
    
    Dim horaInicial, horaFinal, minutoInicial, minutoFinal, diainicial, diafinal, mesinicial, mesfinal, anoinicial, anofinal As Integer
    Dim diferenca As Double
    
    diainicial = Day(dtInicial)
    diafinal = Day(dtFinal)
    mesinicial = Month(dtInicial)
    mesfinal = Month(dtFinal)
    anoinicial = Year(dtInicial)
    anofinal = Year(dtFinal)
    horaInicial = Hour(dtInicial)
    horaFinal = Hour(dtFinal)
    minutoInicial = Minute(dtInicial)
    minutoFinal = Minute(dtFinal)
        
    Do While 1
        If Not (minutoInicial = minutoFinal And horaInicial = horaFinal And diainicial = diafinal And mesinicial = mesfinal And anoinicial = anofinal) Then
            minutoInicial = minutoInicial + 1
            If Not (horaInicial < HoraInicio Or horaInicial > HoraTermino) Then
                If Not ((horaInicial = HoraInicio And minutoInicial < minutoInicio) Or (horaInicial = HoraTermino And minutoInicial >= minutoTermino)) Then
                    diferenca = diferenca + 1
                End If
            End If
            If minutoInicial = 60 Then
                minutoInicial = 0
                horaInicial = horaInicial + 1
            End If
            If horaInicial = 25 Then
                horaInicial = 0
                diainicial = diainicial + 1
            End If
            If (diainicial = 31 And (mesinicial = 4 Or mesinicial = 6 Or mesinicial = 9 Or mesinicial = 11)) Or (mesinicial = 2 And diainicial = 29) Or diainicial = 32 Then
                diainicial = 1
                mesinicial = mesinicial + 1
            End If
            If mesinicial = 13 Then
                mesinicial = 1
                anoinicial = anoinicial + 1
            End If
        Else
            GoTo Fim
        End If
    Loop
Fim:
    CalculaComHorario = diferenca
End Function