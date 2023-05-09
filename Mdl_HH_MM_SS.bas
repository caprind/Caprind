Attribute VB_Name = "Mdl_HH_MM_SS"
'===========================
'=   VARIAVEIS DE TEXTO    =
'===========================
Public ULTIDESC     As String 'Ultimo evento da lista - OK
Public HoraTotal    As String 'Total em horas minutos e segundos retornados pela função ElapsedTime ( 123:12:33 ) - OK
Public Operador     As String 'OK

'===========================
'=   VARIAVEIS DE NÚMERO   =
'===========================
Public ULTICOD          As Integer 'Ultimo código do evento da lista - OK
Public TOK              As Long 'Peças conforme do evento - OK
Public TNC              As Long 'Peças não conforme do evento - OK
Public LOTE             As Long 'Total de peças a ser produzidas - OK
Public QTNC             As Long 'Quantidade de peças não conforme - OK
Public QTPC             As Long 'OK
Public Eficiencia_prep  As Double 'OK
Public Eficiencia_exec  As Double 'OK
Public Eficiencia       As Double 'OK
Public SegPrev          As Double 'Segundos utilizados - OK
Public Segutil          As Double 'Segundos previstos - OK
Public D, H, M, s       As Double 'Controles da função ElapsedTime - OK
Public TotalSegundos    As Double 'Tempo total em segundos - OK
Public CMSEG            As Double 'Custo máquina em reais - OK
Public CTTLOTE          As Double 'Custo total do lote - OK
Public CTTPECA          As Double 'Custo total da peça - OK

Public TPPSEG           As Double 'Tempo de preparação previsto em segundos
Public TPUSEG           As Double 'Tempo de preparação utilizado em segundos - OK
Public TPUSEGDECS       As Double 'Tempo de preparação utilizado em segundos - OK

Public TEPSEG           As Double 'Tempo de execução previsto em segundos - OK
Public TEUSEG           As Double 'Tempo de execução Utilizado em segundos - OK
Public TTUTILSEG        As Double 'Tempo total Utilizado na OS em segundos (execução + preparação)
Public TTEUTILS         As Double 'Tempo total utilizado na OS em segundos
Public Dias             As Double 'OK
'Public QTLOTE           As Double 'Quantidade do lote - OK
Public TTNC             As Double 'Total de peças não conforme - OK
Public TTOK             As Double 'Total de peças conforme - OK

'===========================
'=   VARIAVEIS DE DATA     =
'===========================
Public Data                 As Date 'OK
Public DataConclusaoOS      As Date 'OK
Public DataConclusaoOrdem   As Date 'OK
Public TEP                  As Date 'Tempo de execucao previsto - OK
Public TPP                  As Date 'Tempo previsto por peça - OK
Public ValorDia             As Date 'OK
Public TotalHora            As Date 'Controles da função ElapsedTime - OK
Public TotalMinuto          As Date 'Controles da função ElapsedTime - OK
Public TotalSegundo         As Date 'Controles da função ElapsedTime - OK
Public TempoInicio          As Date 'Hora de inicio do evento - OK
Public TempoFinal           As Date 'Hora final do evento - OK
Public TempoTotal           As Date 'Total de tempo do evento - OK
Public TempoTotalProd       As Date 'Somatório de tempo total produzindo - OK
Public TempoTotalPrep       As Date 'Somatório de tempo total preparando a máquina - OK
Public TPPREV               As Date 'Tempo de preparação previsto - OK
Public TTPUTIL              As Variant 'Tempo total de preparação previsto
Public TPUTIL               As Date 'Tempo de preparação utilizado - OK
Public TEPREV               As Date 'Tempo de execução previsto - OK
Public TEUTIL               As Date 'Tempo de execução utilizado - OK
Public TempoTotalUtil       As Variant ' Somatorio dos total de execução e preparação
Public TTUTIL               As Variant 'Tempo total utilizado na OS

'===========================
'=   VARIAVEIS DE DECISÃO  =
'===========================
Public OSControlada As Boolean 'Controle de Ordem de serviço controlada sim/não - OK
Public Processo_controlado As Boolean 'OK
Public Gravar       As Boolean 'OK

'===========================
'=   VARIAVEIS INDEFINIDA  =
'===========================
Public TEMPODISP    As Variant 'Totalizacao de Tempo total disponível - OK
Public TotalDia     As Variant ' Controles da função ElapsedTime - OK
Public Horas        As Variant ' Controles da função ElapsedTime - OK
Public Minutos      As Variant ' Controles da função ElapsedTime - OK
Public Segundos     As Variant ' Controles da função ElapsedTime - OK
Public Minuto       As Variant 'OK
Public Segundo      As Variant 'OK
Public UltiOperador As Variant 'Ultimo operador a apontar evento da O.s
Public Ultimo       As Variant 'OK
Public Penultimo    As Variant 'OK
Public Resultado    As Variant 'OK

Public TempoUltimo As Date

Public Function FormataTempo(TotalSeg As Double)
On Error GoTo tratar_erro

TotalHoras = 0
TotalMin = 0
Do While TotalSeg >= 60
    TotalMin = TotalMin + 1
    TotalSeg = TotalSeg - 60
Loop
If TotalSeg = 60 Then
    TotalSeg = 0
    TotalMin = TotalMin + 1
End If
Do While TotalMin >= 60
    TotalHoras = TotalHoras + 1
    TotalMin = TotalMin - 60
Loop

TotalSeg = Format(TotalSeg, "###,##0.0000000000")
If TotalSeg < 10 Then
    FormataTempo = IIf(Len(TotalHoras) < 2, "0" & TotalHoras, TotalHoras) & ":" & IIf(Len(TotalMin) = 2, TotalMin, "0" & TotalMin) & ":" & IIf(Len(TotalSeg) = 2, TotalSeg, "0" & TotalSeg)
Else
    FormataTempo = IIf(Len(TotalHoras) < 2, "0" & TotalHoras, TotalHoras) & ":" & IIf(Len(TotalMin) = 2, TotalMin, "0" & TotalMin) & ":" & TotalSeg
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Public Function ElapsedTime(Interval)
On Error GoTo tratar_erro

D = Int(CSng(Interval))
H = Format(Int(CSng(Interval * 24)), "###,###,##0")
M = Format(Int(CSng(Interval * 24 * 60)), "###,###,##0")
s = Format(Int(CSng(Interval * 24 * 3600)), "###,###,##0")

'Debug.print "Dia(s) = " & D
'Debug.print "hora(s) = " & H
'Debug.print "minuto(s) = " & M
'Debug.print "Segundo(s) = " & s

Horas = H
Minutos = M Mod 60
Segundos = s Mod 60

Hr = Horas
Mn = Minutos
Sg = Segundos

HoraTotal = IIf(Len(Hr) = 1, "0" & Hr, Hr) & ":" & IIf(Len(Mn) = 1, "0" & Mn, Mn) & ":" & IIf(Len(Sg) = 1, "0" & Sg, Sg)

'Debug.print HoraTotal

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcFormataHora(HoraFormato As String)
On Error GoTo tratar_erro
Dim Hora As Long 'OK
Dim DataMinuto As Date 'OK
Dim DataMinuto1 As Date 'OK

DataResultado = 0
DecimoSegundos = 0
Texto = ""
Numero = 0
Numero1 = Len(HoraFormato)
Hora = 0
If Numero1 <> 1 Then
    Do While Numero1 <> 0
        If Texto = ":" Then GoTo Pula
        Texto = Left(HoraFormato, (Numero + 1))
        Texto = Right(Texto, Len(Texto) - Numero)
        Numero = Numero + 1
        Numero1 = Numero1 - 1
    Loop
Pula:
    Hora = Left(HoraFormato, (Numero - 1))
    Texto1 = Hora
    Numero2 = Len(Texto1)
    If Hora >= 24 Then
        Do While Hora >= 24
            Hora = Hora - 24
            DataResultado = DataResultado + #11:59:59 PM# + #12:00:01 AM#
        Loop
    End If
        
    'Verifica qtde. de horas
    Texto = ""
    Numero = 0
    Numero1 = Len(HoraFormato)
    Do While Numero1 <> 0
        If Texto = ":" Then GoTo Pula1
        Texto = Left(HoraFormato, (Numero + 1))
        Texto = Right(Texto, Len(Texto) - Numero)
        Numero = Numero + 1
        Numero1 = Numero1 - 1
    Loop
Pula1:
    MinutoSeg = Right(HoraFormato, Len(HoraFormato) - Numero)
    If Len(MinutoSeg) = 5 Then
        DataMinuto = FormataTempo(Right(MinutoSeg, 2))
        DataMinuto1 = "00:" & Left(MinutoSeg, 2) & ":00"
        MinutoSeg = Right(DataMinuto + DataMinuto1, 5)
    End If
    If Hora < 10 Then Texto = "0" & Hora & ":" & MinutoSeg Else Texto = Hora & ":" & MinutoSeg
    DecimoSegundos = IIf(Len(Texto) > 8, Right(Texto, Len(Texto) - 8), 0)
    If DataResultado <> "00:00:00" Then
        If Hora < 10 Then DataResultado = DataResultado & " " & "0" & Hora & ":" & Left(MinutoSeg, 5) Else DataResultado = DataResultado & " " & Hora & ":" & Left(MinutoSeg, 5)
    Else
        DataResultado = Left(Texto, 8)
    End If
End If
    ElapsedTime (DataResultado)
    'Debug.print HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function GetElapsedTime(Interval)
On Error GoTo tratar_erro
Dim TotalHours As Long, TotalMinutes As Long, TotalSeconds As Long 'OK
Dim Days As Long, Hours As Long, Minutes As Long, Seconds As Long 'OK

Days = Int(CSng(Interval))
TotalHours = Int(CSng(Interval * 24))
TotalMinutes = Int(CSng(Interval * 1440))
TotalSeconds = Int(CSng(Interval * 86400))
Hours = TotalHours Mod 24
Minutes = TotalMinutes Mod 60
Seconds = TotalSeconds Mod 60
GetElapsedTime = Days & " Dias " & Hours & " Horas " & Minutes & " Minutos " & Seconds & " Segundos "
        
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
        
Function calculaIntervaloTempo(intervalo)
On Error GoTo tratar_erro
Dim TotalHoras As Long 'OK
Dim TotalMinutos As Long 'OK
Dim TotalSegundos As Long 'OK

Dias = Int(CSng(intervalo))
TotalHoras = Int(CSng(intervalo * 24))
TotalMinutos = Int(CSng(intervalo * 1440))
TotalSegundos = Int(CSng(intervalo * 86400))
Horas = TotalHoras Mod 24
Minutos = TotalMinutos Mod 60
Segundos = TotalSegundos Mod 60

calculaIntervaloTempo = Dias & " dias " & Horas & " Horas " & Minutos & " Minutos " & Segundos & " Segundos "

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function CalculaIntervaloHoras(intervalo)
On Error GoTo tratar_erro

CalculaIntervaloHoras = Int(CSng(intervalo * 24))

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function CalculaDHMS(Periodo As Date, qt As Currency)
On Error GoTo tratar_erro

TotalDias = "23:59:59"
Dias = 0
CalculaDHMS = 0
If qt > 0 Then
    Do While qt > 0
        CalculaDHMS = CalculaDHMS + Periodo
        qt = qt - 1
    Loop
Else
    CalculaDHMS = Periodo
End If
Do While CalculaDHMS > TotalDias
    CalculaDHMS = CalculaDHMS - 1
    Dias = Dias + 1
Loop
Horas = Hour(CalculaDHMS)
Minutos = Minute(CalculaDHMS)
Segundos = Second(CalculaDHMS)

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function CalculaIntervaloSegundos(Dias, Horas, Minutos, Segundos)
On Error GoTo tratar_erro
 
CalculaIntervaloSegundos = Int(CSng(Dias * 24 * 3600)) + Int(CSng(Horas * 3600)) + Int(CSng(Minutos * 60)) + Segundos
'Debug.print "Dia = " & CalculaIntervaloSegundos

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function CalculaEficiencia(segdisp, Segutil)
On Error GoTo tratar_erro

If Segutil > 0 And segdisp > 0 Then CalculaEficiencia = (segdisp / Segutil) * 100

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
