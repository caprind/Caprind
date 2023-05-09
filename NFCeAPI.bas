Attribute VB_Name = "NFCeAPI"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references

'Atributo privado da classe
Private Const tempoResposta = 500
Private Const impressaoParam = """impressao"":{" & """tipo"":""pdf""," & """ecologica"":false," & """itemLinhas"":""1""," & """itemDesconto"":false," & """larguraPapel"":""80mm""}"
'=======================================================================
'                      Token Caprind Sistemas                          =
'=======================================================================
'Private Const token = "Q0FQUklORCBTSVNURU1BSEs5c1o="
'=======================================================================
'                        Token FNL Tecnologia                          =
'=======================================================================
Private Const token = "RkFCSU8gQ0FSRE9TTyBSc2ZGcTI="
'=======================================================================

'Esta função envia um conteudo para uma URL, em requisições do tipo POST
Function NFCe_enviaConteudoParaAPI(conteudo As String, url As String, tpConteudo As String) As String
On Error GoTo SAI
    Dim contentType As String
    
    If (tpConteudo = "txt") Then
        contentType = "text/plain;charset=utf-8"
    ElseIf (tpConteudo = "xml") Then
        contentType = "application/xml;charset=utf-8"
    Else
        contentType = "application/json;charset=utf-8"
    End If
    
    Dim obj As MSXML2.ServerXMLHTTP50
    Set obj = New MSXML2.ServerXMLHTTP50
    obj.Open "POST", url
    obj.setRequestHeader "Content-Type", contentType
    If Trim(token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", token
    End If
    obj.send conteudo
    Dim resposta As String
    resposta = obj.responseText
    
    Select Case obj.status
        Case 401
            MsgBox ("Token não enviado ou inválido")
        Case 403
            MsgBox ("Token sem permissão")
    End Select
    
    NFCe_enviaConteudoParaAPI = resposta

Exit Function
SAI:
  NFCe_enviaConteudoParaAPI = "{" & """status"":""" & Err.Number & """," & """motivo"":""" & Err.Description & """" & "}"
End Function

'Esta função realiza o processo completo de emissãoo: envio e download do documento
Public Function NFCe_emitirSincrono(conteudo As String, tpConteudo As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro

   statusEnvio = ""
    statusDownload = ""
    motivo = ""
    erros = ""
    chNFe = ""
    cStat = ""
    nProt = ""
    
    gravaLinhaLog ("[EMISSAO_SINCRONA_INICIO]")
    
    resposta = NFCe_emitir(conteudo, tpConteudo)

    statusEnvio = NFCe_LerDadosJSON(resposta, "status", "", "")

    'Testa se o envio foi feito com sucesso (200) ou se ï¿½ para reconsultar (-6)
    If (statusEnvio = "100") Or (statusEnvio = "-100") Then
    
        cStat = NFCe_LerDadosJSON(resposta, "nfeProc", "cStat", "")

        'Testa se o cStat ï¿½ igual a 100 ou 150, pois ambos significam "Autorizado"
        If (cStat = "100") Or (cStat = "150") Then
        
            chNFe = NFCe_LerDadosJSON(resposta, "nfeProc", "chNFe", "")
            nProt = NFCe_LerDadosJSON(resposta, "nfeProc", "nProt", "")
            motivo = NFCe_LerDadosJSON(resposta, "nfeProc", "xMotivo", "")

            Sleep (tempoResposta)

            resposta = NFCe_downloadESalvar(chNFe, tpAmb, caminho, exibeNaTela)
            statusDownload = NFCe_LerDadosJSON(resposta, "status", "", "")
            
            'Testa se houve problema no download
            If (statusDownload <> "100") Then
            
                motivo = NFCe_LerDadosJSON(resposta, "motivo", "", "")
                
            End If
        Else
        
            motivo = NFCe_LerDadosJSON(resposta, "nfeProc", "xMotivo", "")
            
        End If
        
    ElseIf (status = "-995") Then

        motivo = NFCe_LerDadosJSON(resposta, "motivo", "", "")
        erros = NFCe_LerDadosJSON(resposta, "erros", "", "")
        
    Else
    
        motivo = NFCe_LerDadosJSON(resposta, "motivo", "", "")
        
    End If
    
    'Monta o JSON de retorno
    retorno = "{"
    retorno = retorno & """statusEnvio"":""" & statusEnvio & ""","
    retorno = retorno & """statusDownload"":""" & statusDownload & ""","
    retorno = retorno & """cStat"":""" & cStat & ""","
    retorno = retorno & """chNFe"":""" & chNFe & ""","
    retorno = retorno & """nProt"":""" & nProt & ""","
    retorno = retorno & """motivo"":""" & motivo & ""","
    retorno = retorno & """erros"":""" & erros & """"
    retorno = retorno & "}"
    
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("[EMISSAO_SINCRONA_FIM]")

    NFCe_emitirSincrono = retorno

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o envio de uma NFC-e
Public Function NFCe_emitir(conteudo As String, tpConteudo As String) As String
On Error GoTo tratar_erro

    url = "https://nfce.ns.eti.br/v1/nfce/issue"

    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
        
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    NFCe_emitir = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o download de documentos de uma NFC-e
Public Function NFCe_download(chNFe As String, tpAmb As String) As String
On Error GoTo tratar_erro
Dim JSON As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & ""","
    JSON = JSON & impressaoParam
    JSON = JSON & "}"

    url = "https://nfce.ns.eti.br/v1/nfce/get"

    gravaLinhaLog ("[DOWNLOAD_NFCE_DADOS]")
    gravaLinhaLog (JSON)
        
    resposta = NFCe_enviaConteudoParaAPI(JSON, url, "json")
    
    status = NFCe_LerDadosJSON(resposta, "status", "", "")
        
    'O retorno da API serão gravado somente em caso de erro,
    'para não gerar um log extenso com o PDF e XML
    If (status <> "100") Then
    
        gravaLinhaLog ("[DOWNLOAD_NFCE_RESPOSTA]")
        gravaLinhaLog (resposta)
        
    Else

        gravaLinhaLog ("[DOWNLOAD_NFCE_STATUS]")
        gravaLinhaLog (status)
        
    End If

    NFCe_download = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o download de documentos de uma NFC-e e salva-os
Public Function NFCe_downloadESalvar(chNFe As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro
 Dim XML As String
 
    resposta = NFCe_download(chNFe, tpAmb)
    status = NFCe_LerDadosJSON(resposta, "status", "", "")

    If status = "100" Then
        
        'Cria o diretório, caso não exista
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
    
        XML = NFCe_LerDadosJSON(resposta, "nfeProc", "xml", "")
        
        If FileOrDirExists(DiretorioXML & chNFe & "-procNFCe.xml") = False Then
        Call NFCe_salvarXML(XML, DiretorioXML, chNFe)
        End If
        
        
        If InStr(1, impressaoParam, "pdf") Then
        
            pdf = NFCe_LerDadosJSON(resposta, "pdf", "", "")
        If FileOrDirExists(DiretorioXML & chNFe & "-procNFCe.pdf") = False Then
            Call NFCe_salvarPDF(pdf, DiretorioDanfe, chNFe)
          End If
          
            If exibeNaTela Then
            Documento = DiretorioDanfe & chNFe & "-procNFCe.pdf"
            FileExecute (Documento)
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informaï¿½ï¿½es")
        gravaLinhaLog ("[Ocorreu um erro, veja o Retorno da API para mais informações  - Metodo: downloadNFCeESalvar]")
    End If

    NFCe_downloadESalvar = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function


'Esta função realiza o download de eventos de uma NFC-e e salva-os
Public Function NFCe_downloadEventoESalvar(chNFe As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro
Dim XML As String
 
    resposta = NFCe_download(chNFe, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "100" Then
    
        'Cria o diretório, caso não exista
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        XML = LerDadosJSON(resposta, "retEvento", "xml", "")
        chNFeCanc = LerDadosJSON(resposta, "retEvento", "chNFeCanc", "")
        Call salvarXML(XML, caminho, chNFeCanc, "CANC")

        If InStr(1, impressaoParam, "pdf") Then
        
            pdf = LerDadosJSON(resposta, "pdfCancelamento", "", "")
            Call salvarPDF(pdf, caminho, chNFeCanc, "CANC")
            
            If exibeNaTela Then
    
                ShellExecute 0, "open", caminho & chNFeCanc & "CANC-procEvenNFe.pdf", "", "", vbNormalFocus
            
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informaï¿½ï¿½es")
         gravaLinhaLog ("[Ocorreu um erro, veja o Retorno da API para mais informações  - Metodo: downloadEventoNFCeESalvar]")
    End If

    NFCe_downloadEventoESalvar = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o cancelamento de uma NFC-e
Public Function NFCe_cancelar(chNFe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro
Dim JSON As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & ""","
    JSON = JSON & """dhEvento"":""" & dhEvento & ""","
    JSON = JSON & """nProt"":""" & nProt & ""","
    JSON = JSON & """xJust"":""" & xJust & """"
    JSON = JSON & "}"
    
    url = "https://nfce.ns.eti.br/v1/nfce/cancel"
    
    gravaLinhaLog ("[CANCELAMENTO_DADOS]")
    gravaLinhaLog (JSON)
    
    resposta = enviaConteudoParaAPI(JSON, url, "json")
        
    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    'Se houve sucesso no evento, realiza o download
    If (status = "135") Then
    
        respostaDownload = NFCe_downloadEventoESalvar(chNFe, tpAmb, caminho, exibeNaTela)
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "100") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    NFCe_cancelar = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza a consulta de situação de uma NFC-e
Public Function NFCe_consultarSituacao(chNFe As String, tpAmb As String) As String
On Error GoTo tratar_erro
Dim JSON As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & """"
    JSON = JSON & "}"

    url = "https://nfce.ns.eti.br/v1/nfce/status"
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_DADOS]")
    gravaLinhaLog (JSON)

    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
   NFCe_consultarSituacao = resposta
   
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza a inutilização de um intervalo de numeração de NFC-e
Public Function NFCe_inutilizar(cUF As String, tpAmb As String, Ano As String, CNPJ As String, Serie As String, nNFIni As String, nNFFin As String, xJust As String) As String
On Error GoTo tratar_erro
Dim JSON As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """cUF"":""" & cUF & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & ""","
    JSON = JSON & """ano"":""" & Ano & ""","
    JSON = JSON & """CNPJ"":""" & CNPJ & ""","
    JSON = JSON & """serie"":""" & Serie & ""","
    JSON = JSON & """nNFIni"":""" & nNFIni & ""","
    JSON = JSON & """nNFFin"":""" & nNFFin & ""","
    JSON = JSON & """xJust"":""" & xJust & """"
    JSON = JSON & "}"

    url = "https://nfce.ns.eti.br/v1/nfce/inut"
    
    gravaLinhaLog ("[INUTILIZACAO_DADOS]")
    gravaLinhaLog (JSON)
        
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    gravaLinhaLog ("[INUTILIZACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    NFCe_inutilizar = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o envio de e-mail de uma NFC-e
Public Function NFCe_enviarEmail(chNFe As String, enviaEmailDoc As String, Email) As String
On Error GoTo tratar_erro
Dim JSON As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """enviaEmailDoc"":" & enviaEmailDoc & ","
    JSON = JSON & """email"":["
    
    Dim emails() As String
    Dim i, quantidade As Integer
    
    emails = Split(Email, ",")
    
    quantidade = UBound(emails)
    
    For i = 0 To quantidade
        If (i = quantidade) Then
            JSON = JSON & """" & emails(i) & """"
        Else
            JSON = JSON & """" & emails(i) & ""","
        End If
    Next
    
    JSON = JSON & "]"
    JSON = JSON & "}"

    url = "https://nfce.ns.eti.br/v1/util/resendemail"
    
    gravaLinhaLog ("[ENVIO_EMAIL_DADOS]")
    gravaLinhaLog (JSON)
        
    resposta = enviaConteudoParaAPI(JSON, url, "json")

    gravaLinhaLog ("[ENVIO_EMAIL_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    NFCe_enviarEmail = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função salva um XML
Public Sub NFCe_salvarXML(XML As String, caminho As String, chNFe As String, Optional Tipo As String = "")
On Error GoTo tratar_erro
    
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String
    Dim extensao As String
    
    If (Tipo = "CANC") Then
        extensao = "-procEvenNFCe.xml"
    Else
        extensao = "-procNFCe.xml"
    End If
    'Seta o caminho para o arquivo XML
    localParaSalvar = caminho & chNFe & nSeqEvento & extensao

    'Remove as contrabarras
    conteudoSalvar = Replace(XML, "\""", "")

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

'Esta função salva um PDF
Public Function NFCe_salvarPDF(pdf As String, caminho As String, chNFe As String, Optional Tipo As String = "") As Boolean
On Error GoTo SAI

    Dim conteudoSalvar  As String
    Dim localParaSalvar As String
    Dim extensao As String
    If (Tipo = "CANC") Then
        extensao = "-procEvenNFCe.pdf"
    Else
        extensao = "-procNFCe.pdf"
    End If
    'Seta o caminho para o arquivo PDF
    localParaSalvar = caminho & chNFe & extensao

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function

'Esta função lê os dados de um JSON
Public Function NFCe_LerDadosJSON(sJsonString As String, key1 As String, key2 As String, key3 As String, Optional key4 As String, Optional key5 As String) As String
On Error GoTo err_handler
    
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" And key5 <> "" Then
        NFCe_LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet), key5, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" Then
        NFCe_LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" Then
        NFCe_LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet)
    ElseIf key1 <> "" And key2 <> "" Then
        NFCe_LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet)
    ElseIf key1 <> "" Then
        NFCe_LerDadosJSON = VBA.CallByName(objJSON, key1, VbGet)
    End If
Err_Exit:

Exit Function
err_handler:
    NFCe_LerDadosJSON = "Error: " & Err.Description
    Resume Err_Exit
End Function

'Esta função lê os dados de um XML
Public Function NFCe_LerDadosXML(sXml As String, key1 As String, key2 As String) As String
On Error GoTo tratar_erro

    NFCe_LerDadosXML = ""
    
    Set XML = New DOMDocument50
    XML.async = False
    
    If XML.loadXML(sXml) Then
        'Tentar pegar o strCampoXML
        Set objNodeList = XML.getElementsByTagName(key1 & "//" & key2)
        Set objNode = objNodeList.NextNode
        
        Dim valor As String
        valor = objNode.Text
        
        If Len(Trim(valor)) > 0 Then 'CONSEGUI LER O XML NODE
            NFCe_LerDadosXML = valor
        End If
        Else
        MsgBox "Não foi possível ler o conteúdo do XML da NFe especificado para leitura.", vbCritical, "ERRO"
    End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função grava uma linha de texto em um arquivo de log
Public Sub NFCe_gravaLinhaLog(conteudoSalvar As String)
On Error GoTo tratar_erro

    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim localParaSalvar As String
    Dim Data As String
    
    'Diretório para salvar os logs
    localParaSalvar = App.Path & "\log\"
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(localParaSalvar, vbDirectory) = "" Then
        MkDir (localParaSalvar)
    End If
    
    'Pega data atual
    Data = Format(Date, "yyyyMMdd")
    
    'Diretório + nome do arquivo para salvar os logs
    localParaSalvar = App.Path & "\log\" & Data & ".txt"
    
    'Pega data e hora atual
    Data = DateTime.Now
    
    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Append Shared As #fnum
    Print #fnum, Data & " - " & conteudoSalvar
    Close fnum

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
