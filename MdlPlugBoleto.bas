Attribute VB_Name = "MdlPlugBoleto"
Public Const PlugCNPJ = "05272563000152"
Public Const PlugToken = "e2e2a4aa55eedc5bf0386bd8ebb0a876"

Public FBoletoX As spdBoletoX

Public IDDuplicata As Long

Public txtRetorno As String
Public txtEnvio As String
Public IDIntegracao As String
Public listaIdsIntegracao As String
Public BoletoPDF As String

'Dados do cedente
Public Escritural As Boolean
Public TipoSacado As String
Public Emissor As String
Public IDBanco As Integer
Public NossoNumero As Long
Public NumeroRemessa As Integer
Public DiretorioRemessa As String
Public DiretorioBoleto As String

Public CedenteRazaoSocial As String
Public CedenteNomeFantasia As String
Public CedenteCpfCnpj As String
Public CedenteEnderecoLogradouro As String
Public CedenteEnderecoNumero As String
Public CedenteEnderecoComplemento As String
Public CedenteEnderecoBairro As String
Public CedenteEnderecoCEP As String
Public CedenteEnderecoCidadeIBGE As String
Public CedenteTelefone As String
Public CedenteEmail As String

'Dados conta cedente
Public ContaCodigoBanco As String
Public ContaAgencia As String
Public ContaAgenciaDV As String
Public ContaNumero As String
Public ContaNumeroDV As String
Public ContaTipo As String
Public ContaCodigoBeneficiario As String
Public ContaValidacaoAtiva As Boolean

'Dados boleto
Public CedenteContaNumero As String
Public CedenteContaNumeroDV As String
Public CedenteConvenioNumero As String
Public CedenteContaCodigoBanco As String
Public SacadoEmail As String
Public SacadoNome As String
Public SacadoCPFCNPJ As String
Public SacadoCelular As String
Public SacadoEnderecoLogradouro As String
Public SacadoEnderecoNumero As String
Public SacadoEnderecoBairro As String
Public SacadoEnderecoCidade As String
Public SacadoEnderecoComplemento As String
Public SacadoEnderecoPais As String
Public SacadoEnderecoUF As String
Public SacadoEnderecoCEP As String

Public TituloNossoNumero As String
Public TituloNumeroDocumento As String
Public TituloDataVencimento As String
Public TituloDataEmissao As String
Public TituloValor As String

Public TituloCodigoJuros As String
Public TituloValorJuros As String
Public TituloDataJuros As String

Public TituloCodigoMulta As String
Public TituloValorMultaTaxa As String
Public TituloDataMulta As String

Public TituloMensagem01 As String
Public TituloMensagem02 As String
Public TituloMensagem03 As String
Public TituloInformacoesAdicionais As String
Public TituloInstrucoes As String

'Dados email
Public EmailNomeRemetente As String
Public EmailRemetente As String
Public EmailAssunto As String
Public EmailMensagem As String
Public EmailDestinatario As String
Public TipoImpressao As String

'SacadoCPFCNPJ = 1001001000113#
'SacadoEmail=email@email.com.br
'SacadoEnderecoLogradouro=Avenida Brasil
'SacadoEnderecoNumero = 54
'SacadoEnderecoBairro = Centro
'SacadoEnderecoCEP = 86890000
'SacadoEnderecoCidade = Maringá
'SacadoEnderecoComplemento=Ato. 704
'SacadoEnderecoPais = Brasil
'SacadoEnderecoUF = PR
'SacadoNome=João da Silva
'SacadoTelefone = 4430379500#
'SacadoCelular = 44998214397#
'CedenteContaCodigoBanco = 341
'CedenteContaNumero = 1324
'CedenteContaNumeroDV = 2
'CedenteConvenioNumero = 221324
'TituloNossoNumero = 341
'TituloValor=10,00
'TituloNumeroDocumento=NFe 12345
'TituloDataEmissao = 27 / 12 / 18
'TituloDataVencimento = 27 / 12 / 18
'TituloAceite = N
'TituloDocEspecie = 1
'TituloLocalPagamento=Pagavel em qualquer banco até o vencimento.
'TituloCodDesconto = 1
'TituloDataDesconto = 27 / 12 / 18
'TituloValorDescontoTaxa=2,00
'TituloCodDesconto2 = 1
'TituloDataDesconto2 = 27 / 12 / 18
'TituloValorDescontoTaxa2=2,00
'TituloValorDesconto=2,00
'TituloCodigoJuros = 1
'TituloDataJuros = 27 / 12 / 18
'TituloValorJuros=2,00
'TituloCodigoMulta = 1
'TituloDataMulta = 27 / 12 / 18
'TituloValorMultaTaxa=2,00
'TituloValorMulta=2,00
'TituloCodProtesto = 2
'TituloPrazoProtesto = 30
'TituloCodBaixaDevolucao = 2
'TituloPrazoBaixa = 30
'TituloMensagem01=Conceder desconto de R$2,oo até 27/12/2018
'TituloMensagem02=Cobrar R$2,00 ao dia após o vencimento
'TituloSacadorAvalista=Joa Silva
'TituloSacadorAvalistaEndereco=Avenida Brasil
'TituloSacadorAvalistaCidade = Maringa
'TituloSacadorAvalistaCEP = 86890000
'TituloSacadorAvalistaUF = PR
'TituloInscricaoSacadorAvalista = 1001001000113#
'TituloEmissaoBoleto = A
'TituloCategoria = 2
'TituloPostagemBoleto = N
'TituloCodEmissaoBloqueto = 2
'TituloOutrosAcrescimos=2,00
'TituloInformacoesAdicionais=Entregar para setor financeiro.
'TituloInstrucoes=Não receber após 30 dias do vencimento.
'TituloParcela = 3 / 10
'TituloVariacaoCarteira = 17
'SALVARBOLETO

Public Function PlugIncluirCedente() As String
On Error GoTo tratar_erro

    IncluirCedente = _
    "INCLUIRCEDENTE" & vbCrLf & _
    "CedenteRazaoSocial=" & CedenteRazaoSocial & vbCrLf & _
    "CedenteNomeFantasia=" & CedenteNomeFantasia & vbCrLf & _
    "CedenteCpfCnpj=" & CedenteCpfCnpj & vbCrLf & _
    "CedenteEnderecoLogradouro=" & CedenteEnderecoLogradouro & vbCrLf & _
    "CedenteEnderecoNumero=" & CedenteEnderecoNumero & vbCrLf & _
    "CedenteEnderecoComplemento=" & CedenteEnderecoComplemento & vbCrLf & _
    "CedenteEnderecoBairro=" & CedenteEnderecoBairro & vbCrLf & _
    "CedenteEnderecoCEP=" & CedenteEnderecoCEP & vbCrLf & _
    "CedenteEnderecoCidadeIBGE=" & CedenteEnderecoCidadeIBGE & vbCrLf & _
    "CedenteTelefone=" & CedenteTelefone & vbCrLf & _
    "CedenteEmail=" & CedenteEmail & vbCrLf & _
    "SALVARCEDENTE"

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function PlugIncluirEmail() As String
On Error GoTo tratar_erro

    INCLUIREMAIL = _
    "INCLUIREMAIL" & vbCrLf & _
    "IdIntegracao=" & IDIntegracao & vbCrLf & _
    "EmailNomeRemetente=" & EmailNomeRemetente & vbCrLf & _
    "EmailRemetente=" & EmailRemetente & vbCrLf & _
    "EmailAssunto=" & EmailAssunto & vbCrLf & _
    "EmailMensagem=" & EmailMensagem & vbCrLf & _
    "EmailDestinatario=" & EmailDestinatario & vbCrLf & _
    "TipoImpressao= 99" & vbCrLf & _
    "SALVAREMAIL"
    

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function


Public Function btnEnviarEmailLote(IDIntegracao As String)
On Error GoTo tratar_erro

Dim EnviarLoteResposta As spdRetEnvioEmailLote

txtRetorno = ""

Set EnviarLoteResposta = FBoletoX.EnviarEmailLote(txtTx2)

txtRetorno = ".:: Enviar E-mail Lote ::."
txtRetorno = txtRetorno & vbNewLine & " Mensagem : " & EnviarLoteResposta.Mensagem
txtRetorno = txtRetorno & vbNewLine & " Status : " & EnviarLoteResposta.status
txtRetorno = txtRetorno & vbNewLine & "Protocolo : " & EnviarLoteResposta.protocolo

edtProtocoloEmailLote = EnviarLoteResposta.protocolo

If (EnviarLoteResposta.status = "ERRO") Then
    txtRetorno = txtRetorno & "ErroClasse: " & EnviarLoteResposta.ErroClasse
End If

If (EnviarLoteResposta.ErroConexao <> "") Then
    txtRetorno = txtRetorno & "Erro Classe: " + EnviarLoteResposta.ErroClasse
End If

txtRetorno = txtRetorno & vbNewLine

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function PlugIncluirConta() As String
On Error GoTo tratar_erro

    IncluirConta = _
    "INCLUIRCEDENTECONTA" & vbCrLf & _
    "ContaCodigoBanco=" & ContaCodigoBanco & vbCrLf & _
    "ContaAgencia=" & ContaAgencia & vbCrLf & _
    "ContaAgenciaDV=" & ContaAgenciaDV & vbCrLf & _
    "ContaNumero=" & ContaNumero & vbCrLf & _
    "ContaNumeroDV=" & ContaNumeroDV & vbCrLf & _
    "ContaTipo=" & ContaTipo & vbCrLf & _
    "ContaCodigoBeneficiario=" & ContaCodigoBeneficiario & vbCrLf & _
    "ContaValidacaoAtiva=" & ContaValidacaoAtiva & vbCrLf & _
    "SALVARCEDENTECONTA"

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub PlugGerarPDFBoleto(protocolo As String)
On Error GoTo tratar_erro
Dim SalvarPDFLote As spdRetSalvarLoteImpressaoPDF
Set FBoletoX = New BoletoX.spdBoletoX
  
FBoletoX.Config.SalvarLogs = True
FBoletoX.Config.url = "http://cobrancabancaria.tecnospeed.com.br"
FBoletoX.ConfigurarSoftwareHouse PlugCNPJ, PlugToken
FBoletoX.Config.CedenteCpfCnpj = CedenteCpfCnpj
txtRetorno = ""
txtEnvio = ""

DiretorioPDF = Localrel & "\Boletos\ArquivosPDF"

If DS.FileOrDirExists(DiretorioPDF) = False Then
    MkDir (DiretorioPDF)
End If


Arquivo = BoletoPDF

If DS.FileOrDirExists(Arquivo) = True Then
    DS.FileDelete (Arquivo)
End If

Set SalvarPDFLote = FBoletoX.SalvarLoteImpressaoPDF(protocolo, Arquivo)

txtRetorno = ".:: Consultar Protocolo Lote Impress?o::."
txtRetorno = txtRetorno & vbNewLine & "Mensagem : " & SalvarPDFLote.Mensagem
txtRetorno = txtRetorno & vbNewLine & "Status : " & SalvarPDFLote.status

If (SalvarPDFLote.status = "ERRO") Then
    USMsgBox "Erro ao gerar pdf de impressão do boleto, tente novamente", vbCritical, "CAPRIND v5.0"
    Sit_REG = 1
    Exit Sub
End If

txtRetorno = txtRetorno & vbNewLine
ShellExecute 0, "open", Arquivo, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Function PlugGerarProtocoloBoleto(IDIntegracao As String)
On Error GoTo tratar_erro
Dim ImprimirLoteList As spdRetImprimirLote
Set FBoletoX = New BoletoX.spdBoletoX
  
FBoletoX.Config.SalvarLogs = True
FBoletoX.Config.url = "http://cobrancabancaria.tecnospeed.com.br"
FBoletoX.ConfigurarSoftwareHouse PlugCNPJ, PlugToken
FBoletoX.Config.CedenteCpfCnpj = CedenteCpfCnpj

'Busca protocolo de impressao
1:

Set ImprimirLoteList = FBoletoX.ImprimirLote(IDIntegracao, "99")
  
If (ImprimirLoteList.protocolo <> "") Then
  PlugGerarProtocoloBoleto = ImprimirLoteList.protocolo
  If Financeiro_Contas_Receber = True Then
    StrSql = "update tbl_Detalhes_Recebimento set Protocolo = '" & ImprimirLoteList.protocolo & "' where IDContaReceber = " & IDDuplicata
    Conexao.Execute (StrSql)
    StrSql = "update tbl_contas_Receber set Protocolo = '" & ImprimirLoteList.protocolo & "' where IDIntconta = " & IDDuplicata
    Conexao.Execute (StrSql)
 Else
    StrSql = "update tbl_Detalhes_Recebimento set Protocolo = '" & ImprimirLoteList.protocolo & "' where id = " & IDDuplicata
    Conexao.Execute (StrSql)
 End If
    
Else
    frmFaturamento_Prod_serv_boleto.txtRetorno.Text = ImprimirLoteList.ErroClasse
    frmFaturamento_Prod_serv_boleto.txtRetorno.Text = ImprimirLoteList.ErroConexao
    Sit_REG = 1
  'GoTo 1
End If

'PlugGerarProtocoloBoleto = ImprimirLoteList.Protocolo

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub PlugDescartarBoleto(IDIntegracao As String)
On Error GoTo tratar_erro

Dim DescarteList As BoletoX.spdRetDescartarLista
  Dim DescarteItem As BoletoX.spdRetDescartarTituloItem
Set FBoletoX = New BoletoX.spdBoletoX
'
FBoletoX.Config.SalvarLogs = True
FBoletoX.Config.url = "http://cobrancabancaria.tecnospeed.com.br"
FBoletoX.ConfigurarSoftwareHouse PlugCNPJ, PlugToken
FBoletoX.Config.CedenteCpfCnpj = CedenteCpfCnpj
  
  Set DescarteList = FBoletoX.Descartar(IDIntegracao)
  With frmFaturamento_Prod_serv_boleto.txtRetorno
  .Text = ""
  .Text = .Text & ".:: Descartar ::." & vbNewLine
  .Text = .Text & "Mensagem: " & DescarteList.Mensagem & vbNewLine
  .Text = .Text & "Status: " & DescarteList.status & vbNewLine
  
  If (DescarteList.status = "ERRO") Then
    .Text = .Text & "ErroClasse: " & DescarteList.ErroClasse
  End If
  
  If (DescarteList.status = "SUCESSO") Then
    frmFaturamento_Prod_serv_boleto.txtIDIntegracao = ""
    frmFaturamento_Prod_serv_boleto.txtStatus = ""
    frmFaturamento_Prod_serv_boleto.txtProtocolo = ""
    
    StrSql = "UPDATE tbl_Detalhes_Recebimento set IDIntegracao = null, Protocolo = null, Status = null, seq_remessa = null, Nosso_numero = null where IDIntegracao = '" & IDIntegracao & "'"
    Conexao.Execute StrSql
    
    StrSql = "UPDATE tbl_contas_receber set IDIntegracao = null, Protocolo = null where IDIntegracao = '" & IDIntegracao & "'"
    Conexao.Execute StrSql
    
    USMsgBox "Boleto excluido com sucesso", vbInformation, "CAPRIND v5.0"
    If Financeiro_Contas_Receber = True Then
    frmContas_Receber.ProcCarregaLista 1
    Else
    
    End If
  End If
  
  For i = 0 To DescarteList.Count - 1
    Set DescarteItem = DescarteList.Item(i)
    .Text = .Text & "Item: " & (i + 1) & vbNewLine
    .Text = .Text & "IdIntegracao: " & DescarteItem.IDIntegracao & vbNewLine
    .Text = .Text & "Erro: " & DescarteItem.Erro & vbNewLine
    .Text = .Text & "ErroClasse: " & DescarteItem.ErroClasse & vbNewLine
  Next i
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub PlugGerarRemessa(IDIntegracao As String)
On Error GoTo tratar_erro

Dim retornoLista As BoletoX.spdRetGerarRemessaLista
Dim retornoItem  As BoletoX.spdRetGerarRemessaItem

Set FBoletoX = New BoletoX.spdBoletoX
  
FBoletoX.Config.SalvarLogs = True
FBoletoX.Config.url = "http://cobrancabancaria.tecnospeed.com.br"
FBoletoX.ConfigurarSoftwareHouse PlugCNPJ, PlugToken
FBoletoX.Config.CedenteCpfCnpj = CedenteCpfCnpj
  
  'Dim Titulos As BoletoX.spdGerarRemessaTitulos

  Set retornoLista = FBoletoX.GerarRemessa(IDIntegracao)
 
  
  If retornoLista.status <> "SUCESSO" Then
    USMsgBox "Não foi possivel gerar o arquivo de remessa, tente novamente. | Erro: " & retornoLista.Mensagem, vbCritical, "CAPRIND v5.0"
    Exit Sub
  End If
  
  
  
  For i = 0 To retornoLista.Count - 1
    Set retornoItem = retornoLista.Item(i)
    'Set Titulos = retornoItem.Titulos
    txtRetorno = txtRetorno & "Item : " & CStr(i + 1) & vbNewLine
    txtRetorno = txtRetorno & "Remessa : " & retornoItem.Remessa & vbNewLine
    txtRetorno = txtRetorno & "Banco : " & retornoItem.Banco & vbNewLine
    txtRetorno = txtRetorno & "Conta : " & retornoItem.Conta & vbNewLine
    txtRetorno = txtRetorno & "Erro : " & retornoItem.Erro & vbNewLine
    
    For j = 0 To retornoItem.Titulos.Count - 1
      txtRetorno = txtRetorno & "Id Integração : " & CStr(j + 1) & " - " & retornoItem.Titulos.Item(j) & vbNewLine
    Next j
  
  Next i

    strarquivo = NumeroRemessa + 1
    NumeroRemessa = strarquivo
    
'Banco Bradesco
  If CedenteContaCodigoBanco = 237 Then 'Bradesco
    
        Mes = Month(Date)
        Dia = Day(Date)
        If Len(Mes) = 1 Then Mes = "0" & Mes
        If Len(Dia) = 1 Then Dia = "0" & Dia
        If Len(strarquivo) = 1 Then strarquivo = "0" & strarquivo
        Arquivo = "CB" & Dia & Mes & strarquivo & ".REM"
  End If
  
  'Banco Itaú
  If CedenteContaCodigoBanco = 341 Then 'Itau
    Select Case Len(strarquivo)
      Case 1: Arquivo = "0000000" & strarquivo & ".txt"
      Case 2: Arquivo = "000000" & strarquivo & ".txt"
      Case 4: Arquivo = "00000" & strarquivo & ".txt"
      Case 5: Arquivo = "0000" & strarquivo & ".txt"
      Case 6: Arquivo = "000" & strarquivo & ".txt"
      Case 6: Arquivo = "00" & strarquivo & ".txt"
      Case 7: Arquivo = "0" & strarquivo & ".txt"
    End Select
  End If

''' Salvar a remessa em UTF-8(SEM BOM)
   Set objStream = CreateObject("ADODB.Stream")
   Set objStreamNoBOM = CreateObject("ADODB.Stream")

   With objStream
      .Open
      .Charset = "UTF-8"
      .WriteText retornoItem.Remessa
      .Position = 0
   End With
   


   With objStreamNoBOM
      .Charset = "Windows-1252"
      .Open
      .Type = 2
      .WriteText objStream.ReadText
      
      If (Right(DiretorioRemessa, 1)) = "\" Then
      .SaveToFile DiretorioRemessa & Arquivo, 2
      Else
      .SaveToFile DiretorioRemessa & "\" & Arquivo, 2
      End If

      .Close
   End With
   StrSql = "update tbl_Instituicoes set seq_remessa = " & NumeroRemessa & " where id = " & IDBanco
   Conexao.Execute (StrSql)
   If Financeiro_Contas_Receber = True Then
    StrSql = "update tbl_Detalhes_Recebimento set seq_remessa = " & NumeroRemessa & " where idcontareceber = " & IDConta
    Conexao.Execute (StrSql)
   Else
    StrSql = "update tbl_Detalhes_Recebimento set seq_remessa = " & NumeroRemessa & " where id = " & IDDuplicata
    Conexao.Execute (StrSql)
   End If
   

   objStream.Close
   USMsgBox "Arquivo remessa gerado com sucesso", vbInformation, "CAPRIND v5.0"

   frmFaturamento_Prod_serv_boleto.txtRetorno = txtRetorno
   
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Function PlugConsultarBoleto(IDIntegracao As String) As String
On Error GoTo tratar_erro

Dim ConsultarList As BoletoX.spdRetConsultarLista
Dim ConsultarItem As BoletoX.spdRetConsultarTituloItem

 Set FBoletoX = New BoletoX.spdBoletoX

 
FBoletoX.Config.SalvarLogs = True
FBoletoX.Config.url = "http://cobrancabancaria.tecnospeed.com.br"
FBoletoX.ConfigurarSoftwareHouse PlugCNPJ, PlugToken
FBoletoX.Config.CedenteCpfCnpj = CedenteCpfCnpj
 Set ConsultarList = FBoletoX.Consultar(IDIntegracao)
 
  txtRetorno = ".:: Consultar Título ::."
  txtRetorno = txtRetorno & vbNewLine & "Mensagem : " & ConsultarList.Mensagem
  txtRetorno = txtRetorno & vbNewLine & "Status : " & ConsultarList.status
  txtRetorno = txtRetorno & vbNewLine
  

  Do While ConsultarList.Count <> 0
  
     For i = 0 To ConsultarList.Count - 1
        Set ConsultarItem = ConsultarList.Item(i)
        txtRetorno = txtRetorno & "ITEM: " & str((i + 1)) & vbNewLine
        txtRetorno = txtRetorno & "IdIntegracao: " & ConsultarItem.IDIntegracao & vbNewLine
        txtRetorno = txtRetorno & "Situacao: " + ConsultarItem.situacao & vbNewLine
        txtRetorno = txtRetorno & "Motivo: " + ConsultarItem.motivo & vbNewLine
        txtRetorno = txtRetorno & vbNewLine

        txtRetorno = txtRetorno & "CEDENTE:" & vbNewLine
        txtRetorno = txtRetorno & "Agencia: " + ConsultarItem.CedenteAgencia & vbNewLine
        txtRetorno = txtRetorno & "AgenciaDV: " + ConsultarItem.CedenteAgenciaDV & vbNewLine
        txtRetorno = txtRetorno & "Código Banco: " + ConsultarItem.CedenteCodigoBanco & vbNewLine
        txtRetorno = txtRetorno & "Carteira: " + ConsultarItem.CedenteCarteira & vbNewLine
        txtRetorno = txtRetorno & "Conta: " + ConsultarItem.CedenteCarteira & vbNewLine
        txtRetorno = txtRetorno & "Numero Convênio: " + ConsultarItem.CedenteNumeroConvenio & vbNewLine
        txtRetorno = txtRetorno & vbNewLine

        txtRetorno = txtRetorno & "SACADO:" & vbNewLine
        txtRetorno = txtRetorno & "CPFCNPJ: " + ConsultarItem.SacadoCPFCNPJ & vbNewLine
        txtRetorno = txtRetorno & "Nome: " + ConsultarItem.SacadoNome & vbNewLine
        txtRetorno = txtRetorno & "Telefone: " + ConsultarItem.SacadoTelefone & vbNewLine
        txtRetorno = txtRetorno & "Email: " + ConsultarItem.SacadoEmail & vbNewLine
        txtRetorno = txtRetorno & "Endereço Número: " + ConsultarItem.SacadoEnderecoNumero & vbNewLine
        txtRetorno = txtRetorno & "Endereço Bairro: " + ConsultarItem.SacadoEnderecoBairro & vbNewLine
        txtRetorno = txtRetorno & "Endereço CEP: " + ConsultarItem.SacadoEnderecoCEP & vbNewLine
        txtRetorno = txtRetorno & "Endereço Cidade: " + ConsultarItem.SacadoEnderecoCidade & vbNewLine
        txtRetorno = txtRetorno & "Endereço Complemento: " + ConsultarItem.SacadoEnderecoComplemento & vbNewLine
        txtRetorno = txtRetorno & "Endereço Logradouro: " + ConsultarItem.SacadoEnderecoLogradouro & vbNewLine
        txtRetorno = txtRetorno & "Endereço País: " + ConsultarItem.SacadoEnderecoPais & vbNewLine
        txtRetorno = txtRetorno & "Endereço UF: " + ConsultarItem.SacadoEnderecoUF & vbNewLine
        txtRetorno = txtRetorno & vbNewLine

'        txtRetorno = txtRetorno & "TÍTULO:" & vbNewLine
'        txtRetorno = txtRetorno & "Número Documento: " + ConsultarItem.TituloNumeroDocumento & vbNewLine
'        txtRetorno = txtRetorno & "Origem Documento: " + ConsultarItem.TituloOrigemDocumento & vbNewLine
'        txtRetorno = txtRetorno & "Nosso Número: " + ConsultarItem.TituloNossoNumero & vbNewLine
'        txtRetorno = txtRetorno & "Data Emissão: " + ConsultarItem.TituloDataEmissao & vbNewLine
'        txtRetorno = txtRetorno & "Data Vencimento: " + ConsultarItem.TituloDataVencimento & vbNewLine
'        txtRetorno = txtRetorno & "Data Desconto: " + CStr(ConsultarItem.TituloDataDesconto) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Desconto: " + CStr(ConsultarItem.TituloValorDesconto) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Juros: " + CStr(ConsultarItem.TituloValorJuros) & vbNewLine
'        txtRetorno = txtRetorno & "Prazo Protesto: " + ConsultarItem.TituloPrazoProtesto & vbNewLine
'        txtRetorno = txtRetorno & "Mensagem 1: " + ConsultarItem.TituloMensagem01 & vbNewLine
'        txtRetorno = txtRetorno & "Mensagem 2: " + ConsultarItem.TituloMensagem02 & vbNewLine
'        txtRetorno = txtRetorno & "Mensagem 3: " + ConsultarItem.TituloMensagem03 & vbNewLine
'        txtRetorno = txtRetorno & "Valor: " + CStr(ConsultarItem.TituloValor) & vbNewLine
'        txtRetorno = txtRetorno & "Data Crédito: " + ConsultarItem.PagamentoDataCredito & vbNewLine
'        txtRetorno = txtRetorno & "Título Pago: " + CStr(ConsultarItem.PagamentoRealizado) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Crédito: " + CStr(ConsultarItem.PagamentoValorCredito) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Outros Acréscimos: " + CStr(ConsultarItem.TituloValorOutrosAcrescimos) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Pago: " + CStr(ConsultarItem.PagamentoValorPago) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Taxa Cobrança: " + CStr(ConsultarItem.PagamentoValorTaxaCobranca) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Abatimento: " + CStr(ConsultarItem.TituloValorAbatimento) & vbNewLine
'        txtRetorno = txtRetorno & "Valor Outras Despesas: " + CStr(ConsultarItem.PagamentoValorOutrasDespesas) & vbNewLine
'        txtRetorno = txtRetorno & "Valor IOF: " + CStr(ConsultarItem.PagamentoValorIOF) & vbNewLine
'        txtRetorno = txtRetorno & "Data Pagamento: " + ConsultarItem.PagamentoData & vbNewLine
'        txtRetorno = txtRetorno & "Valor Outros Créditos: " + CStr(ConsultarItem.PagamentoValorOutrosCreditos) & vbNewLine
'        txtRetorno = txtRetorno & "Pagamento Valor Desconto: " + CStr(ConsultarItem.PagamentoValorDesconto) & vbNewLine
'        txtRetorno = txtRetorno & "Pagamento Valor Acréscimos: " + CStr(ConsultarItem.PagamentoValorAcrescimos) & vbNewLine
'        txtRetorno = txtRetorno & "Pagamento Valor Abatimento: " + CStr(ConsultarItem.PagamentoValorAbatimento) & vbNewLine
'        txtRetorno = txtRetorno & "Impressão Visualizada: " + CStr(ConsultarItem.ImpressaoVisualizada) & vbNewLine

'        txtRetorno = txtRetorno & "OCORRÊNCIAS:" & vbNewLine
'        txtRetorno = txtRetorno & vbNewLine
'
'        txtRetorno = txtRetorno & vbNewLine
'
'        If ConsultarItem.TituloOcorrencias.Count > 0 Then
'            txtRetorno = txtRetorno & "OCORRÊNCIAS:" & vbNewLines
'
'            For j = 0 To ConsultarItem.TituloOcorrencias.Count - 1
'                txtRetorno = txtRetorno & "  Código: " & ConsultarItem.TituloOcorrencias.Item(j).CODIGO   'Código da ocorrência
'                txtRetorno = txtRetorno & " - " & ConsultarItem.TituloOcorrencias.Item(j).Mensagem & vbNewLine    'Mensagem de ocorrência
'            Next j
'
'            txtRetorno = txtRetorno & vbNewLine
'
'        End If
        Mensagem = ConsultarItem.situacao

      Next i
        
      ConsultarList.PaginaSeguinte

  Loop
  
  'Debug.print Mensagem
  frmFaturamento_Prod_serv_boleto.txtRetorno = txtRetorno
  frmFaturamento_Prod_serv_boleto.txtStatus = Mensagem
        
Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub PlugEmitirBoleto()
On Error GoTo tratar_erro

Dim retorno As BoletoX.spdRetIncluirLista

Set FBoletoX = New BoletoX.spdBoletoX

FBoletoX.Config.SalvarLogs = True
FBoletoX.Config.url = "http://cobrancabancaria.tecnospeed.com.br"
FBoletoX.ConfigurarSoftwareHouse PlugCNPJ, PlugToken

'Imprimirpdf:
txtEnvio = PlugIncluirBoleto()
FBoletoX.Config.CedenteCpfCnpj = CedenteCpfCnpj
'Sleep 3000

Set retorno = FBoletoX.Incluir(txtEnvio)

For i = 0 To retorno.Count - 1
    If i = 0 Then
        listaIdsIntegracao = retorno.Item(i).IDIntegracao
    Else
        listaIdsIntegracao = listaIdsIntegracao & "," & retorno.Item(i).IDIntegracao
    End If
    
    If listaIdsIntegracao <> "" Then
    frmFaturamento_Prod_serv_boleto.txtIDIntegracao.Text = listaIdsIntegracao
    
    If Financeiro_Contas_Receber = True Then
        StrSql = "update tbl_Detalhes_Recebimento set IDIntegracao = '" & listaIdsIntegracao & "' where IDContaReceber = " & IDDuplicata
        Conexao.Execute (StrSql)
        StrSql = "update tbl_contas_Receber set IDIntegracao = '" & listaIdsIntegracao & "' where IDIntconta = " & IDDuplicata
        Conexao.Execute (StrSql)
    Else
        StrSql = "update tbl_Detalhes_Recebimento set IDIntegracao = '" & listaIdsIntegracao & "' where id = " & IDDuplicata
        Conexao.Execute (StrSql)
    End If
        
    End If

    
    If retorno.Item(i).situacao <> "" Then
        Mensagem = retorno.Item(i).situacao
        
    If Mensagem <> "EMITIDO" Then
            PlugConsultarBoleto (listaIdsIntegracao)
            If Mensagem = "EMITIDO" Then
            Sit_REG = 1
            Else
                Sit_REG = 0
                'USMsgBox txtRetorno, vbInformation, "CAPRIND v5.0"
            Exit Sub
            End If
    Else
    frmFaturamento_Prod_serv_boleto.txtStatus = Mensagem
    End If
        
     '   MsgBox retorno.Item(i).Erro

    If Financeiro_Contas_Receber = True Then
        StrSql = "update tbl_Detalhes_Recebimento set Status = '" & Mensagem & "' where IDContaReceber = " & IDDuplicata
        Conexao.Execute (StrSql)
    Else
        StrSql = "update tbl_Detalhes_Recebimento set Status = '" & Mensagem & "' where id = " & IDDuplicata
        Conexao.Execute (StrSql)
    End If
        

        
        'frmFaturamento_Prod_serv_boleto.txtStatus = retorno.Item(i).situacao
           '     USMsgBox retorno.Item(i).Erro, vbCritical, "CAPRIND v5.0"
    Else
    Sit_REG = 1
        USMsgBox (retorno.Item(i).Erro), vbOKOnly, "CAPRIND v5.0"
    End If
Next i



'PlugEmitirBoleto = listaIdsIntegracao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function PlugIncluirBoleto() As String
On Error GoTo tratar_erro

If Escritural = False Then

    PlugIncluirBoleto = _
    "INCLUIRBOLETO" & vbCrLf & vbCrLf & _
    "CedenteContaNumero=" & CedenteContaNumero & vbCrLf & vbCrLf & _
    "CedenteContaNumeroDV=" & CedenteContaNumeroDV & vbCrLf & vbCrLf & _
    "CedenteConvenioNumero=" & CedenteConvenioNumero & vbCrLf & vbCrLf & _
    "CedenteContaCodigoBanco=" & CedenteContaCodigoBanco & vbCrLf & vbCrLf & _
    "SacadoEmail=" & SacadoEmail & vbCrLf & vbCrLf & _
    "SacadoNome=" & SacadoNome & vbCrLf & vbCrLf & _
    "SacadoCPFCNPJ=" & SacadoCPFCNPJ & vbCrLf & vbCrLf & _
    "SacadoCelular=" & SacadoCelular & vbCrLf & vbCrLf & _
    "SacadoEnderecoLogradouro=" & SacadoEnderecoLogradouro & vbCrLf & vbCrLf & _
    "SacadoEnderecoNumero=" & SacadoEnderecoNumero & vbCrLf & vbCrLf & _
    "SacadoEnderecoBairro=" & SacadoEnderecoBairro & vbCrLf & vbCrLf & _
    "SacadoEnderecoCEP=" & SacadoEnderecoCEP & vbCrLf & vbCrLf & _
    "TituloNossoNumero=" & TituloNossoNumero & vbCrLf & vbCrLf & _
    "TituloNumeroDocumento=" & TituloNumeroDocumento & vbCrLf & vbCrLf & _
    "TituloDataVencimento=" & TituloDataVencimento & vbCrLf & vbCrLf & _
    "TituloDataEmissao=" & TituloDataEmissao & vbCrLf & vbCrLf & _
    "TituloValor=" & TituloValor & vbCrLf & vbCrLf & vbCrLf & _
    "TituloCodigoJuros=" & TituloCodigoJuros & vbCrLf & vbCrLf & _
    "TituloDataJuros=" & TituloDataJuros & vbCrLf & vbCrLf & _
    "TituloValorJuros=" & TituloValorJuros & vbCrLf & vbCrLf & _
    "TituloCodigoMulta=" & TituloCodigoMulta & vbCrLf & vbCrLf & _
    "TituloDataMulta=" & TituloDataMulta & vbCrLf & vbCrLf & _
    "TituloValorMultaTaxa=" & TituloValorMultaTaxa & vbCrLf & vbCrLf
    
    PlugIncluirBoleto = PlugIncluirBoleto & _
    "TituloMensagem01=" & TituloMensagem01 & vbCrLf & vbCrLf & _
    "TituloMensagem02=" & TituloMensagem02 & vbCrLf & vbCrLf & _
    "TituloMensagem03=" & TituloMensagem03 & vbCrLf & vbCrLf & _
    "TituloInformacoesAdicionais=" & TituloInformacoesAdicionais & vbCrLf & vbCrLf & _
    "TituloInstrucoes=" & TituloInstrucoes & vbCrLf & vbCrLf & _
    "SALVARBOLETO"
Else
    PlugIncluirBoleto = _
    "INCLUIRBOLETO" & vbCrLf & vbCrLf & _
    "CedenteContaNumero=" & CedenteContaNumero & vbCrLf & vbCrLf & _
    "CedenteContaNumeroDV=" & CedenteContaNumeroDV & vbCrLf & vbCrLf & _
    "CedenteConvenioNumero=" & CedenteConvenioNumero & vbCrLf & vbCrLf & _
    "CedenteContaCodigoBanco=" & CedenteContaCodigoBanco & vbCrLf & vbCrLf & _
    "SacadoEmail=" & SacadoEmail & vbCrLf & vbCrLf & _
    "SacadoNome=" & SacadoNome & vbCrLf & vbCrLf & _
    "SacadoCPFCNPJ=" & SacadoCPFCNPJ & vbCrLf & vbCrLf & _
    "SacadoCelular=" & SacadoCelular & vbCrLf & vbCrLf & _
    "SacadoEnderecoLogradouro=" & SacadoEnderecoLogradouro & vbCrLf & vbCrLf & _
    "SacadoEnderecoNumero=" & SacadoEnderecoNumero & vbCrLf & vbCrLf & _
    "SacadoEnderecoBairro=" & SacadoEnderecoBairro & vbCrLf & vbCrLf & _
    "SacadoEnderecoCEP=" & SacadoEnderecoCEP & vbCrLf & vbCrLf & _
    "TituloNumeroDocumento=" & TituloNumeroDocumento & vbCrLf & vbCrLf & _
    "TituloDataVencimento=" & TituloDataVencimento & vbCrLf & vbCrLf & _
    "TituloDataEmissao=" & TituloDataEmissao & vbCrLf & vbCrLf & _
    "TituloValor=" & TituloValor & vbCrLf & vbCrLf & vbCrLf & _
    "TituloCodigoJuros=" & TituloCodigoJuros & vbCrLf & vbCrLf & _
    "TituloDataJuros=" & TituloDataJuros & vbCrLf & vbCrLf & _
    "TituloValorJuros=" & TituloValorJuros & vbCrLf & vbCrLf & _
    "TituloCodigoMulta=" & TituloCodigoMulta & vbCrLf & vbCrLf & _
    "TituloDataMulta=" & TituloDataMulta & vbCrLf & vbCrLf & _
    "TituloValorMultaTaxa=" & TituloValorMultaTaxa & vbCrLf & vbCrLf
    
    PlugIncluirBoleto = PlugIncluirBoleto & _
    "TituloMensagem01=" & TituloMensagem01 & vbCrLf & vbCrLf & _
    "TituloMensagem02=" & TituloMensagem02 & vbCrLf & vbCrLf & _
    "TituloMensagem03=" & TituloMensagem03 & vbCrLf & vbCrLf & _
    "TituloInformacoesAdicionais=" & TituloInformacoesAdicionais & vbCrLf & vbCrLf & _
    "TituloInstrucoes=" & TituloInstrucoes & vbCrLf & vbCrLf & _
    "SALVARBOLETO"

End If


Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function


