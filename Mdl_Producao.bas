Attribute VB_Name = "Mdl_Producao"
Public TotalSegundosUtilizadosExecucaoUnidade As Integer
Public TotalSegundosUtilizadosPreparacaoOS As Integer
Public TotalSegundosUtilizadosExecucaoOS As Integer
Public TotalSegundosUtilizadosOS As Integer
Public PrazoEntrega As Date
Public PrazoFinal As Date


Public Sub ProcSomaSegundosOS(OS As Integer)
On Error GoTo tratar_erro
'=============================================================================
' Tempo de Setup
'=============================================================================
Set TBAbrir = CreateObject("adodb.recordset")

StrSql = "select sum(tempototalseg) as TotalSegundosPreparacaoOS from producaofases where CodigoDesc = '2' and IDFase = '" & OS & "'"

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = False Then
TotalSegundosUtilizadosPreparacaoOS = TBAbrir!TotalSegundosPreparacaoOS
End If

TBAbrir.Close

'=============================================================================
' Tempo de execução
'=============================================================================
Set TBAbrir = CreateObject("adodb.recordset")

StrSql = "select sum(tempototalseg) as TotalSegundosExecucaoOS, sum(quant) as Totalaprovado, SUM(reprovada) as Reprovado, Sum(QTCD) as Condicional from producaofases where CodigoDesc = '1' and IDFase = '" & OS & "'"

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = False Then
TotalProduzido = TBAbrir!TotalAprovado + TBAbrir!Reprovado
TotalSegundosUtilizadosExecucaoOS = IIf(IsNull(TBAbrir!TotalSegundosExecucaoOS), 0, TBAbrir!TotalSegundosExecucaoOS)

'=============================================================================
' Tempo de execução por unidade
'=============================================================================
If TotalProduzido > 1 Then
TotalSegundosUtilizadosExecucaoUnidade = TotalSegundosExecucaoOS / TotalProduzido
Else
If TotalSegundosExecucaoOS > 1 Then
TotalSegundosUtilizadosExecucaoUnidade = (TotalSegundosExecucaoOS * (TotalProduzido * 100)) / 100
End If
End If

End If
TBAbrir.Close

'=============================================================================
' Tempo total utilizado na OS
'=============================================================================

TotalSegundosUtilizadosOS = TotalSegundosUtilizadosPreparacaoOS + TotalSegundosUtilizadosExecucaoOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
