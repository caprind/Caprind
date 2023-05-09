Attribute VB_Name = "Mdl_NFe"
Public DataEmissao As Date


Function FunBuscaDadosCNPJ(CnpjDados As String) As Boolean
On Error GoTo tratar_erro

Dim resposta As String
Dim obj As MSXML2.ServerXMLHTTP50
Set obj = New MSXML2.ServerXMLHTTP50
obj.Open "GET", "https://www.receitaws.com.br/v1/cnpj/" & CnpjDados

conteudo = CnpjDados
obj.send conteudo

resposta = obj.responseText
'Debug.print resposta

Nome = LerDadosJSON(resposta, "nome", "", "")
UF = LerDadosJSON(resposta, "uf", "", "")
telefone = LerDadosJSON(resposta, "telefone", "", "")
Bairro = LerDadosJSON(resposta, "bairro", "", "")
logradouro = LerDadosJSON(resposta, "logradouro", "", "")
Numero = LerDadosJSON(resposta, "numero", "", "")
CEP = LerDadosJSON(resposta, "cep", "", "")
municipio = LerDadosJSON(resposta, "municipio", "", "")



      cmbuf.Text = Trim(p.Item("retConsCad").Item("infCons").Item("UF"))
      
      If cmbPessoa.Text = "Jurídica" Then
      txtRG_IE = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("IE"))
      End If
      
      txtnomerazao.Text = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      'txtnomefantasia = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      txtendereco = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xLgr"))
      txtNumero = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("nro"))
      txtBairro = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xBairro"))
      If Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xMun")) <> "" Then
      cmbCidade.Text = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xMun"))
      End If
      txtCEP = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("CEP"))
      cmbRegimeTributario.Text = IIf(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xRegApur") = "NORMAL - REGIME PERIÓDICO DE APURAÇÃO", "Lucro presumido", "Simples Nacional")
      'txtnomefantasia = p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome")
      If txtCEP <> "" Then
      'Cmd_buscarCEP_Click
      End If
      txtCategoria.Text = "A"
      USMsgBox "Consulta relizada com sucesso, dados carregados", vbInformation, "CAPRIND v5.0"

USMsgBox "Nome = " & Nome & vbCrLf & "UF: " & UF & vbCrLf & "Telefone: " & telefone & vbCrLf & "Bairro: " & Bairro & vbCrLf

Exit Function
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("Não foi possível carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Function
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

