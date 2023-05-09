Attribute VB_Name = "NFeAPI"
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'=======================================================================
'                      Token Caprind Sistemas                          =
'=======================================================================
'Private Const token = "Q0FQUklORCBTSVNURU1BSEs5c1o="
'=======================================================================
'                        Token FNL Tecnologia                          =
'=======================================================================
Private Const token = "RkFCSU8gQ0FSRE9TTyBSc2ZGcTI="
'=======================================================================

'Esta função salva um PDF
Public Function salvarPDF(pdf As String, caminho As String, chNFe As String, nSeqEvento As String) As Boolean
On Error GoTo SAI

    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo PDF
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procNFe.pdf"
    Else
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procEvenNFe.pdf"
    End If

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum

Exit Function
SAI:
    USMsgBox (Err.Number & " - " & Err.Description), vbCritical, "CAPRIND v5.0"
End Function

Public Function previaNFeESalvar(conteudo As String, tpConteudo As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro

    Dim resposta As String
    Dim status As String
    Dim pdf As String

    resposta = previaNFe(conteudo, tpConteudo)

    status = LerDadosJSON(resposta, "status", "", "")
    pdf = LerDadosJSON(resposta, "pdf", "", "")

    If (status = "200") Then
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        Call salvarPreviaPDF(pdf, caminho, "Previa", "", "")
        If exibeNaTela Then

            ShellExecute 0, "open", caminho & "Previa" & "-procNFe.pdf", "", "", vbNormalFocus

        End If
    Else
        MsgBox ("Ocorreu um erro ao fazer a requisicao de previa da NFe. Verifique os logs para obter mais informacoes.")
    End If

    previaNFeESalvar = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Function salvarPreviaPDF(pdf As String, caminho As String, chNFe As String, tpEvento As String, nSeqEvento As String) As Boolean
On Error GoTo SAI
    Dim conteudoSalvar  As String
    Dim tpEventoSalvar As String
    Dim localParaSalvar As String

    If (nSeqEvento = "") Then
        localParaSalvar = caminho & tpEvento & chNFe & nSeqEvento & "-procNFe.pdf"
    Else
        localParaSalvar = caminho & tpEvento & chNFe & nSeqEvento & "-procEvenNFe.pdf"
    End If

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum
    
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function


Public Function previaNFe(conteudo As String, tpConteudo As String) As String
On Error GoTo tratar_erro

    Dim url As String
    Dim resposta As String

    url = "https://nfe.ns.eti.br/util/preview/nfe"

    gravaLinhaLog ("[PREVIA_DADOS]")
    gravaLinhaLog (conteudo)
        
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[PREVIA_RESPOSTA]")
    gravaLinhaLog (resposta)

    previaNFe = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Function enviaConteudoParaAPI(conteudo As String, url As String, tpConteudo As String) As String
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
'=====================================
' Envia dados do XML ou Json pra API
'=====================================
'Debug.print conteudo
obj.send conteudo
'=====================================
Dim resposta As String
resposta = obj.responseText
'Debug.print resposta
'Debug.print conteudo

Select Case obj.status
    'Se o token não for enviado ou for inválido
    Case 401
        USMsgBox ("Token não enviado ou inválido")
    'Se o token informado for inválido 403
    Case 403
        USMsgBox ("Token sem permissão")
    Case 200
        EmailEnviado = True
End Select

enviaConteudoParaAPI = resposta
    
Exit Function
SAI:
  enviaConteudoParaAPI = "{" & """status"":""" & Err.Number & """," & """motivo"":""" & Err.Description & """" & "}"
End Function

'Esta função realiza o processo completo de emissão: envio, consulta e download do documento
Public Function emitirNFeSincrono(conteudo As String, tpConteudo As String, CNPJ As String, tpDown As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro



'Inicia as variáveis vazias
statusEnvio = ""
statusConsulta = ""
statusDownload = ""
motivo = ""
xMotivo = ""
erros = ""
nsNRec = ""
chNFe = ""
cStat = ""
nProt = ""

    'Envia a NF-e para processamento
    resposta = emitirNFe(conteudo, tpConteudo)

    'Lê o status de retorno
    statusEnvio = LerDadosJSON(resposta, "status", "", "")
    
    'Se o documento já foi emitido carrega o nsNrec na caixa de texto
    If (statusEnvio = "-7") Then
    'Se sim, lê o nsNRec
     nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")
     frmFaturamento_Prod_Serv_NFe_NS.txtnsNrec.Text = nsNRec
     'frmFaturamento_Prod_Serv_NFe_NS.procConsultarNFE
    End If
   'Debug.print resposta
    'Testa se o envio foi feito com sucesso
    'Se for -6, precisa ler o campo nsNRec, pois quer dizer que o documento está pendente de consulta

    If (statusEnvio = "200") Or (statusEnvio = "-6") Or (statusEnvio = "100") Then
        'Se sim, lê o nsNRec
        'Debug.print resposta
        
         If NFCe = False Then
            nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")
            frmFaturamento_Prod_Serv_NFe_NS.txtnsNrec.Text = nsNRec
        End If
        
        'Aguarda 2000 (2 segundos) milisegundos antes de consultar o status de processamento
        Sleep (6000)

        'Consulta o status de processamento
        
        If NFCe = False Then ' Se não for NFCe consulta status NFe
            resposta = consultarStatusProcessamento(CNPJ, nsNRec, tpAmb)
        Else 'Se for NFCe consulta status
            chNFe = frmFaturamento_Prod_Serv_NFe_NS.txtchNFe.Text
            resposta = NFCe_consultarSituacao(chNFe, tpAmb)
        End If

        'Lê o status de retorno
        statusConsulta = LerDadosJSON(resposta, "status", "", "")

        'Testa se a consulta foi feita com sucesso (200)
        If (statusConsulta = "200" Or statusConsulta = "100") Then
            'Se sim, lê o cStat da nota
            
            If NFCe = False Then
                cStat = LerDadosJSON(resposta, "cStat", "", "")
            Else
                cStat = LerDadosJSON(resposta, "nfeProc", "cStat", "")
            End If
            
            'Debug.print resposta
            'Testa se o cStat é igual a 100 ou 150, pois ambos significam "Autorizado"
            If (cStat = "100") Or (cStat = "150") Then
            '=============================
            ' NFE AUTORIZADA COM SUCESSO =
            '=============================
            
                'Lê a chave
                If NFCe = False Then
                    chNFe = LerDadosJSON(resposta, "chNFe", "", "")
                Else
                    chNFe = LerDadosJSON(resposta, "nfeProc", "chNFe", "")
                End If
                
                frmFaturamento_Prod_Serv_NFe_NS.txtchNFe.Text = chNFe

                'Lê o protocolo
                
                If NFCe = False Then
                    nProt = LerDadosJSON(resposta, "nProt", "", "")
                Else
                    nProt = LerDadosJSON(resposta, "nfeProc", "nProt", "")
                End If
                
                frmFaturamento_Prod_Serv_NFe_NS.txt_nProt.Text = nProt

                'Lê o motivo
                If NFCe = False Then
                    motivo = LerDadosJSON(resposta, "xMotivo", "", "")
                Else
                    motivo = LerDadosJSON(resposta, "nfeProc", "xMotivo", "")
                End If
                
                'Faz o download do documento
'                If NFCe = False Then
'                    resposta = downloadNFeAndSave(chNFe, tpAmb, tpDown, caminho, exibeNaTela)
'                Else
'                    resposta = NFCe_downloadESalvar(chNFe, tpAmb, caminho, exibeNaTela)
'                End If
                
                'Lê o status de retorno
'                statusDownload = LerDadosJSON(resposta, "status", "", "")
'
'                If (statusDownload <> "200") Then
'                    motivo = LerDadosJSON(resposta, "motivo", "", "")
'                End If
            Else
                'Se o cStat for diferente de 100, lê o motivo da rejeição (Sefaz)
                'Debug.print resposta
                If NFCe = False Then
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")
                Else
                motivo = LerDadosJSON(resposta, "nfeProc", "xMotivo", "")
                End If
            End If
        Else
            'Se a consulta for diferente de 200, lê o motivo (Erro de API)
            motivo = LerDadosJSON(resposta, "motivo", "", "")
        End If
    'Se o status for -7, quer dizer que o documento já foi processado anteriormente e autorizado, então, podemos ler o nsNRec
    ElseIf (status = "-7") Then
        'Lê o motivo
        motivo = LerDadosJSON(resposta, "motivo", "", "")
            
        'Lê o nsNRec
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")
        
'===========================
ElseIf (statusEnvio = "-4") Then
 'Lê o motivo
 motivo = LerDadosJSON(resposta, "motivo", "", "")

ElseIf (statusEnvio = "-2") Then
 'Lê o motivo
 motivo = LerDadosJSON(resposta, "motivo", "", "")
 
 Dim errosMsg() As String
 errosMsg = Split(resposta, "[""")
 errosMsg(1) = Left(errosMsg(1), Len(errosMsg(1)) - 2)
 errosMsg(1) = Replace(errosMsg(1), """,""", " - ")
 errosMsg(1) = Replace(errosMsg(1), """", "")
 erros = errosMsg(1)
'==============================================================
' Substituido
'==============================================================
'    'Se for -4 ou -2, precisa ler o campo "erros" também, pois ambos significam erro de validação contra o schema
'    ElseIf (statusEnvio = "-4") Or (statusEnvio = "-2") Then
'        'Lê o motivo
'        motivo = LerDadosJSON(resposta, "motivo", "", "")
'
'        'Lê os erros
'        erros = LerDadosJSON(resposta, "erros", "", "")
'==========================================================
    'Se for -999 é erro interno, precisa ler o campo erro
    ElseIf (statusEnvio = "-999") Or (statusEnvio = "-5") Then
        'Lê o objeto erro
        erros = Split(resposta, """erro"":""")

        'Lê o motivo do erro
        erros = LerDadosJSON(resposta, "erro", "", "")

        erros = LerDadosJSON(erros, "xMotivo", "", "")
    Else
        'Qualquer outro retorno, lê o motivo
    If NFCe = False Then
        statusEnvio = LerDadosJSON(resposta, "status", "", "")
'        cStat = LerDadosJSON(resposta, "cStat", "", "")
        motivo = LerDadosJSON(resposta, "motivo", "", "")
    Else
        statusEnvio = LerDadosJSON(resposta, "status", "", "")
        cStat = LerDadosJSON(resposta, "nfeProc", "cStat", "")
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        xMotivo = LerDadosJSON(resposta, "nfeProc", "xMotivo", "")
    End If
        
    End If
    
    
    'Grava dados de retorno
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("")

    emitirNFeSincrono = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o envio de uma NF-e
Public Function emitirNFe(conteudo As String, tpConteudo As String) As String
On Error GoTo tratar_erro

    Dim url As String
    Dim resposta As String

    'Informa a url para onde deve ser enviado
    If NFCe = False Then
        url = "https://nfe.ns.eti.br/nfe/issue" 'Emitir NF-e
    Else
        url = "https://nfce.ns.eti.br/v1/nfce/issue" 'Emitir NFC-e
    End If
    
    'Grava dados envio
    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
        
    'Envia o conteudo para a URL
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    'Grava resposta API
    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    emitirNFe = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza a consulta o status de processamento de uma NF-e
Public Function consultarStatusProcessamento(CNPJ As String, nsNRec As String, tpAmb As String) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """CNPJ"":""" & CNPJ & ""","
    JSON = JSON & """nsNRec"":""" & nsNRec & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & """"
    JSON = JSON & "}"

    'Informa a url para onde deve ser enviado
    If NFCe = False Then
    url = "https://nfe.ns.eti.br/nfe/issue/status"
    Else
    url = "https://nfce.ns.eti.br/v1/nfce/status"
    End If
    
    'Grava dados envio
    gravaLinhaLog ("[CONSULTA_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    'Grava resposta API
    gravaLinhaLog ("[CONSULTA_RESPOSTA]")
    gravaLinhaLog (resposta)

    consultarStatusProcessamento = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza a consulta o status de processamento de uma NF-e
Public Function consultarStatusWS(CNPJ As String, tpAmb As String, nsUF As String, nsVersao As String) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """CNPJCont"":""" & CNPJ & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & ""","
    JSON = JSON & """UF"":""" & nsUF & ""","
    JSON = JSON & """versao"":""" & nsVersao & """"
    JSON = JSON & "}"

    'Informa a url para onde deve ser enviado
    url = "https://nfe.ns.eti.br/util/wssefazstatus"
    
    'Grava dados envio
    gravaLinhaLog ("[CONSULTA_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    'Grava resposta API
    gravaLinhaLog ("[CONSULTA_RESPOSTA]")
    gravaLinhaLog (resposta)

    consultarStatusWS = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o download de documentos de uma NF-e
Public Function downloadNFe(chNFe As String, tpDown As String, tpAmb As String) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String
        Dim status As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """tpDown"":""" & tpDown & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & """"
    JSON = JSON & "}"

    'Informa a url para onde deve ser enviado
    If NFCe = False Then
    url = "https://nfe.ns.eti.br/nfe/get"
    Else
    url = "https://nfce.ns.eti.br/v1/nfce/get"
    End If

    'Grava dados envio
    gravaLinhaLog ("[DOWNLOAD_NFE_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    'Lê o status de retorno
    status = LerDadosJSON(resposta, "status", "", "")
        
    'A resposta do download só será gravada em caso de erro, para evitar de gerar um log muito grande
    If (status <> "200") Then
        'Grava resposta API
        gravaLinhaLog ("[DOWNLOAD_NFE_RESPOSTA]")
        gravaLinhaLog (resposta)
    'Se der sucesso, grava apenas o status
    Else
        'Grava status API
        gravaLinhaLog ("[DOWNLOAD_NFE_STATUS]")
        gravaLinhaLog (status)
    End If

    downloadNFe = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o download de documentos de uma NF-e e salva-os
Public Function downloadNFeAndSave(chNFe As String, tpAmb As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro

Dim baixarXML As Boolean
Dim baixarPDF As Boolean
Dim baixarJSON As Boolean
Dim XML As String
Dim JSON As String
Dim pdf As String
Dim status As String
Dim resposta As String
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(caminho, vbDirectory) = "" Then
        MkDir (caminho)
    End If

    'Checa o que baixar com base no tpDown informado
    If InStr(1, tpDown, "X") Then
        baixarXML = True
    End If
    If InStr(1, tpDown, "P") Then
        baixarPDF = True
    End If
    If InStr(1, tpDown, "J") Then
        baixarJSON = True
    End If

    'Requisição de download do documento
    resposta = downloadNFe(chNFe, tpDown, tpAmb)
    
    'Pega o status de RetornoXML da requisição
    status = LerDadosJSON(resposta, "status", "", "")

    'Se o RetornoXML da API for positivo, salva o que foi solicitado
    If status = "200" Or status = "100" Then
        'Checa se deve baixar XML
        If baixarXML = True Then
        'Debug.print resposta
            If NFCe = False Then
            XML = LerDadosJSON(resposta, "xml", "", "")
            Else
            XML = LerDadosJSON(resposta, "nfeProc", "xml", "")
            End If
            Call salvarXML(XML, caminho, chNFe, "")
        End If
        'Checa se deve baixar JSON
        If baixarJSON = True Then
            Dim conteudoJSON() As String
            'Separa o JSON da NFe
            conteudoJSON = Split(resposta, """nfeProc"":{")
            JSON = "{""nfeProc"":{" & conteudoJSON(1)
            Call salvarJSON(JSON, caminho, chNFe, "")
        End If
        'Checa se deve baixar PDF
        If baixarPDF = True Then

           ' pdf = LerDadosJSON(resposta, "pdf", "", "")
            pdf = NFCe_LerDadosJSON(resposta, "pdf", "", "")

        
            Call salvarPDF(pdf, caminho, chNFe, "")
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & chNFe & "-procNFe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        USMsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadNFeAndSave = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o download de eventos de uma NF-e
Public Function downloadEventoNFe(chNFe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & ""","
    JSON = JSON & """tpDown"":""" & tpDown & ""","
    JSON = JSON & """tpEvento"":""" & tpEvento & ""","
    JSON = JSON & """nSeqEvento"":""" & nSeqEvento & """"
    JSON = JSON & "}"

    'Informa a url para onde deve ser enviado
    url = "https://nfe.ns.eti.br/nfe/get/event"
    
    'Grava dados envio
    gravaLinhaLog ("[DOWNLOAD_EVENTO_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")

    'Lê o status de retorno
    status = LerDadosJSON(resposta, "status", "", "")
    
    'A resposta do download só será gravada em caso de erro, para evitar de gerar um log muito grande
    If (status <> "200") Then
        'Grava resposta API
        gravaLinhaLog ("[DOWNLOAD_EVENTO_RESPOSTA]")
        gravaLinhaLog (resposta)
    'Se der sucesso, grava apenas o status
    Else
        'Grava status API
        gravaLinhaLog ("[DOWNLOAD_EVENTO_STATUS]")
        gravaLinhaLog (status)
    End If

    downloadEventoNFe = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o download de eventos de uma NF-e e salva-os
Public Function downloadEventoNFeAndSave(chNFe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String, caminho As String, exibeNaTela As Boolean) As String
On Error GoTo tratar_erro

    Dim baixarXML As Boolean
    Dim baixarPDF As Boolean
    Dim baixarJSON As Boolean
    Dim XML As String
    Dim JSON As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(caminho, vbDirectory) = "" Then
        MkDir (caminho)
    End If

    'Checa o que baixar com base no tpDown informado
    If InStr(1, tpDown, "X") Then
        baixarXML = True
    End If
    If InStr(1, tpDown, "P") Then
        baixarPDF = True
    End If
    If InStr(1, tpDown, "J") Then
        baixarJSON = True
    End If

    'Requisição de download do documento
    resposta = downloadEventoNFe(chNFe, tpAmb, tpDown, tpEvento, nSeqEvento)
    
    'Pega o status de retorno da requisição
    status = LerDadosJSON(resposta, "status", "", "")

    'Se o retorno da API for positivo, salva o que foi solicitado
    If status = "200" Then
        'Checa se deve baixar XML
        If baixarXML = True Then
            XML = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(XML, caminho, chNFe, nSeqEvento)
        End If
        'Checa se deve baixar JSON
        If baixarJSON = True Then
            JSON = LerDadosJSON(resposta, "json", "", "")
            Call salvarJSON(JSON, caminho, chNFe, nSeqEvento)
        End If
        'Checa se deve baixar PDF
        If baixarPDF = True Then
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chNFe, nSeqEvento)
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & chNFe & nSeqEvento & "-procEvenNFe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        USMsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadEventoNFeAndSave = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza o cancelamento de uma NF-e
Public Function cancelarNFe(chNFe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, tpDown As String, caminho As String, exibeNaTela As Boolean, Optional ByVal a3 As Boolean = False, Optional ByVal cnpjEmitente As String = "") As String
On Error GoTo tratar_erro

   Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    Dim objDom As DOMDocument50
    Dim objEnvioEvento As IXMLDOMElement
    Dim objEvento As IXMLDOMElement
    Dim objEveInf As IXMLDOMElement
    Dim objEvedet As IXMLDOMElement
    
    Set objDom = New DOMDocument50

        
    'Informa a URL
    If NFCe = False Then
        url = "https://nfe.ns.eti.br/nfe/cancel"
    Else
        url = "https://nfce.ns.eti.br/v1/nfce/cancel"
    End If
        
    If frmFaturamento_Prod_Serv_NFe_NS.txtTPCertificado.Text = "A1" Then
        Dim JSON As String
        
        'Monta o JSON
        JSON = "{"
        JSON = JSON & """chNFe"":""" & chNFe & ""","
        JSON = JSON & """tpAmb"":""" & tpAmb & ""","
        JSON = JSON & """dhEvento"":""" & dhEvento & ""","
        JSON = JSON & """nProt"":""" & nProt & ""","
        JSON = JSON & """xJust"":""" & xJust & """"
        JSON = JSON & "}"
        
        'Grava dados envio
        gravaLinhaLog ("[CANCELAMENTO_DADOS]")
        gravaLinhaLog (JSON)
        
        'Envia o json para a URL
        resposta = enviaConteudoParaAPI(JSON, url, "json")
    Else
        Dim XML As String
        'Monta o XML
        XML = "<?xml version=""1.0"" encoding=""utf-8""?>"
        XML = XML & "<evento versao=""1.00"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
        XML = XML & "<infEvento Id=""ID110111" & chNFe & "01"">"
        XML = XML & "<cOrgao>" & Left(chNFe, 2) & "</cOrgao>"
        XML = XML & "<tpAmb>" & tpAmb & "</tpAmb>"
        XML = XML & "<CNPJ>" & CnpjNF & "</CNPJ>"
        XML = XML & "<chNFe>" & chNFe & "</chNFe>"
        XML = XML & "<dhEvento>" & dhEvento & "</dhEvento>"
        XML = XML & "<tpEvento>110111</tpEvento>"
        XML = XML & "<nSeqEvento>1</nSeqEvento>"
        XML = XML & "<verEvento>1.00</verEvento>"
        XML = XML & "<detEvento versao=""1.00"">"
        XML = XML & "<descEvento>Cancelamento</descEvento>"
        XML = XML & "<nProt>" & nProt & "</nProt>"
        XML = XML & "<xJust>" & xJust & "</xJust>"
        XML = XML & "</detEvento>"
        XML = XML & "</infEvento>"
        XML = XML & "</evento>"
        
        objDom.loadXML (XML)
        '===============================================================
        If frmFaturamento_Prod_Serv_NFe_NS.txtD1 <> "" Then
        objDom.Save (frmFaturamento_Prod_Serv_NFe_NS.txtD1 & NomeArquivo & "CA.xml")
        End If
        frmFaturamento_Prod_Serv_NFe_NS.txtRetorno.Text = objDom.XML
        'frmFaturamento_Prod_Serv_NFe_NS.Text = objDom.xml
        'Debug.print XML
        '===============================================================

        Dim AssinaturaXML2 As New AssinaturaXML2.Principal
        RetornoXML = AssinaturaXML2.assinarXML(XML, "infEvento", CnpjNF)
        
       If RetornoXML = "Certificado Digital não encontrado" Then
         RetornoXML = AssinaturaXML2.assinarXML(XML, "infEvento", frmFaturamento_Prod_Serv_NFe_NS.txtSerialCertificado)
       End If
       
       If RetornoXML = "Certificado Digital não encontrado" Then
         USMsgBox "Verifique seu certificado digital e tente novamente!", vbCritical, "CAPRIND v5.0"
         Var1 = 0
         Exit Function
       End If
       Var1 = 1
       
       XML = RetornoXML
       
        'Grava dados envio
        gravaLinhaLog ("[CANCELAMENTO_DADOS]")
        gravaLinhaLog (XML)
        
        'Envia o json para a URL
        resposta = enviaConteudoParaAPI(XML, url, "xml")
    End If
        
    'Grava resposta API
    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    'Lê o status
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200" Or status = "135") Then
        'Aguarda 2000 (2 segundos) milisegundos antes de consultar o status de processamento
        Sleep (2000)
        
        If NFCe = False Then
            respostaDownload = downloadEventoNFeAndSave(chNFe, tpAmb, tpDown, "CANC", "1", caminho, exibeNaTela)
        Else
            respostaDownload = NFCe_downloadEventoESalvar(chNFe, tpAmb, caminho, exibeNaTela)
        End If
        
        status = LerDadosJSON(respostaDownload, "status", "", "")
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        frmFaturamento_RetornoSEFAZ.txtRetorno = motivo
        
        'Se der problema na consulta do Status
        If (status <> "200") Then
            'Retorna a resposta da Consulta
            resposta = respostaDownload
        End If
    '========================================
    'ACERTA TABELA DE DADOS COM O STATUS DA NFE
    '========================================
    StatusNFe = "101"
    Var1 = 2

    ID_nota = frmFaturamento_Prod_Serv_NFe_NS.txtID_nota.Text
'    ID_NFe = frmFaturamento_Prod_Serv_NFe_NS.txtID_nota.Text
'=======================================================================
    ProcExcluirArquivosRemessa frmFaturamento_Prod_Serv_NFe_NS.txtID_nota.Text
    ProcExcluirContas frmFaturamento_Prod_Serv_NFe_NS.txtID_nota.Text, frmFaturamento_Prod_Serv.opt_Saida, frmFaturamento_Prod_Serv.txttipocliente
    Conexao.Execute "DELETE from ECEV from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = ECEV.ID_faturamento where NFP.ID_nota = " & ID_nota
'=======================================================================
    Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = '" & StatusNFe & "' where id_nota = " & ID_nota
    Conexao.Execute "Update tbl_dados_nota_fiscal Set int_status = '" & Var1 & "' where id = " & ID_nota
    frmFaturamento_Prod_Serv_NFe_NS.procCancelarTabelas
    ProcExcluirSaida_NFe ID_nota
    
'=======================================================================================
' Excluir contas geradas pela nota
'=======================================================================================
    ID_nota = frmFaturamento_Prod_Serv_NFe_NS.txtID_nota
    Conexao.Execute "DELETE from CC from CC_realizado CC INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = CC.ID_financeiro Where CR.ID_Nota = " & ID_nota & " and CC.Operacao = 'Crédito'"
    Conexao.Execute "DELETE from FF from Familia_financeiro FF INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = FF.IDconta Where CR.ID_Nota = " & ID_nota & " and FF.Tipoconta = 'R' and (CR.Proposta IS NULL or CR.Proposta = N'')"
    Conexao.Execute "DELETE from FC from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON CR.IDFluxo = FC.IDFluxo Where CR.ID_Nota = " & ID_nota & " and (CR.Proposta IS NULL or CR.Proposta = N'')"
    Conexao.Execute "DELETE from tbl_contas_receber where ID_Nota = " & ID_nota & " and (Proposta IS NULL or Proposta = N'')"
    
'=======================================================================================
' Excluir relacionamentos na nota
'=======================================================================================
    Conexao.Execute "DELETE from Faturamento_Relacionamento where ID_Nota = " & ID_nota
'=======================================================================================
    USMsgBox "Nota fiscal cancelada com sucesso!", vbInformation, "CAPRIND v5.0"
'=======================================================================================
    Else
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        xMotivo = LerDadosJSON(resposta, "erro", "xMotivo", "")
        USMsgBox xMotivo, vbInformation, "CAPRIND v5.0"
        frmFaturamento_RetornoSEFAZ.txtRetorno = xMotivo
        frmFaturamento_Prod_Serv_NFe_NS.txtRetorno = xMotivo
    End If
    
    cancelarNFe = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Function corrigirNFe(chNFe As String, tpAmb As String, dhEvento As String, nSeqEvento As String, xCorrecao As String, tpDown As String, caminho As String, exibeNaTela As Boolean, Optional ByVal a3 As Boolean = False, Optional ByVal cnpjEmitente As String = "") As String
On Error GoTo tratar_erro

Dim JSON As String
Dim url As String
Dim resposta As String
Dim status As String
Dim respostaDownload As String
Dim XML As String

'Informa a URL
url = "https://nfe.ns.eti.br/nfe/cce"

'==========================================================================
' Se for certificado tipo A1 então usa sistema de envio JSON
'==========================================================================
If TPCertificado = "A1" Then
'Monta o JSON
JSON = "{"
JSON = JSON & """chNFe"":""" & chNFe & ""","
JSON = JSON & """tpAmb"":""" & tpAmb & ""","
JSON = JSON & """dhEvento"":""" & dhEvento & ""","
JSON = JSON & """nSeqEvento"":""" & nSeqEvento & ""","
JSON = JSON & """xCorrecao"":""" & xCorrecao & """"
JSON = JSON & "}"

'Grava dados envio
gravaLinhaLog ("[CCE_DADOS]")
gravaLinhaLog (JSON)

'Envia o json para a URL
resposta = enviaConteudoParaAPI(JSON, url, "json")
End If

'==========================================================================
' Se for certificado A3 envia em sistema XML e assina
'==========================================================================
If TPCertificado = "A3" Then

If SerialCertificado = "" Then
  USMsgBox "Certificado digital não encontrado", vbCritical, "CAPRIND v5.0"
  Exit Function
End If

Dim nSeqEvento_Aux
If (CInt(nSeqEvento) < 10) Then
nSeqEvento_Aux = "0" & nSeqEvento
End If

'=====================================================
'Monta o XML
'=====================================================
XML = "<?xml version=""1.0"" encoding=""utf-8""?>"
XML = XML & "<evento versao=""1.00"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
XML = XML & "<infEvento Id=""ID110110" & chNFe & nSeqEvento_Aux & """>"
XML = XML & "<cOrgao>" & Left(chNFe, 2) & "</cOrgao>"
XML = XML & "<tpAmb>" & tpAmb & "</tpAmb>"
XML = XML & "<CNPJ>" & chCNPJ & "</CNPJ>"
XML = XML & "<chNFe>" & chNFe & "</chNFe>"
XML = XML & "<dhEvento>" & dhEvento & "</dhEvento>"
XML = XML & "<tpEvento>110110</tpEvento>"
XML = XML & "<nSeqEvento>" & nSeqEvento & "</nSeqEvento>"
XML = XML & "<verEvento>1.00</verEvento>"
XML = XML & "<detEvento versao=""1.00"">"
XML = XML & "<descEvento>Carta de Correcao</descEvento>"
XML = XML & "<xCorrecao>" & xCorrecao & "</xCorrecao>"
XML = XML & "<xCondUso>A Carta de Correcao e disciplinada pelo paragrafo 1o-A do art. 7o do Convenio S/N, de 15 de dezembro de 1970 e pode ser utilizada para regularizacao de erro ocorrido na emissao de documento fiscal, desde que o erro nao esteja relacionado com: I - as variaveis que determinam o valor do imposto tais como: base de calculo, aliquota, diferenca de preco, quantidade, valor da operacao ou da prestacao; II - a correcao de dados cadastrais que implique mudanca do remetente ou do destinatario; III - a data de emissao ou de saida.</xCondUso>"
XML = XML & "</detEvento>"
XML = XML & "</infEvento>"
XML = XML & "</evento>"
'====================================================
' Assina o XML com certificado A3 usando o CNPJ
'====================================================
Dim AssinaturaXML2 As New AssinaturaXML2.Principal
XML = AssinaturaXML2.assinarXML(XML, "infEvento", chCNPJ)
'====================================================
' Se não der certo a assinatura buscando por cnpj o certificado
'====================================================
If XML = "Certificado Digital não encontrado" Then
'=====================================================
' Monta novamente o XML
'=====================================================
XML = "<?xml version=""1.0"" encoding=""utf-8""?>"
XML = XML & "<evento versao=""1.00"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
XML = XML & "<infEvento Id=""ID110110" & chNFe & nSeqEvento_Aux & """>"
XML = XML & "<cOrgao>" & Left(chNFe, 2) & "</cOrgao>"
XML = XML & "<tpAmb>" & tpAmb & "</tpAmb>"
XML = XML & "<CNPJ>" & chCNPJ & "</CNPJ>"
XML = XML & "<chNFe>" & chNFe & "</chNFe>"
XML = XML & "<dhEvento>" & dhEvento & "</dhEvento>"
XML = XML & "<tpEvento>110110</tpEvento>"
XML = XML & "<nSeqEvento>" & nSeqEvento & "</nSeqEvento>"
XML = XML & "<verEvento>1.00</verEvento>"
XML = XML & "<detEvento versao=""1.00"">"
XML = XML & "<descEvento>Carta de Correcao</descEvento>"
XML = XML & "<xCorrecao>" & xCorrecao & "</xCorrecao>"
XML = XML & "<xCondUso>A Carta de Correcao e disciplinada pelo paragrafo 1o-A do art. 7o do Convenio S/N, de 15 de dezembro de 1970 e pode ser utilizada para regularizacao de erro ocorrido na emissao de documento fiscal, desde que o erro nao esteja relacionado com: I - as variaveis que determinam o valor do imposto tais como: base de calculo, aliquota, diferenca de preco, quantidade, valor da operacao ou da prestacao; II - a correcao de dados cadastrais que implique mudanca do remetente ou do destinatario; III - a data de emissao ou de saida.</xCondUso>"
XML = XML & "</detEvento>"
XML = XML & "</infEvento>"
XML = XML & "</evento>"

'====================================================
' Assina o XML com certificado A3 usando o serial do certificado
'====================================================
  XML = AssinaturaXML2.assinarXML(XML, "infEvento", SerialCertificado)
'====================================================
'Debug.print XML
'====================================================
' Verifica se houve erro na assinatura
'====================================================
  If Left(XML, 4) = "Erro" Then
   USMsgBox XML, vbCritical, "CAPRIND v5.0"
   Exit Function
  End If
'====================================================
End If
'====================================================
'Grava dados envio
gravaLinhaLog ("[CCE_DADOS]")
gravaLinhaLog (XML)
'====================================================
'Envia o XML assinado para a URL
resposta = enviaConteudoParaAPI(XML, url, "xml")
frmFaturamento_CartaCorrecao_NS.txtRetorno.Text = resposta
'====================================================
End If
'====================================================
'Grava resposta API
gravaLinhaLog ("[CCE_RESPOSTA]")
gravaLinhaLog (resposta)
'====================================================
'Lê o status
status = LerDadosJSON(resposta, "status", "", "")

If (status = "200") Then
respostaDownload = downloadEventoNFeAndSave(chNFe, tpAmb, tpDown, "CCE", nSeqEvento, caminho, exibeNaTela)

status = LerDadosJSON(respostaDownload, "status", "", "")
'=====================================================
'Se der problema no download
If (status <> "200") Then
'Retorna a resposta do download
resposta = respostaDownload
End If
End If

corrigirNFe = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function


'Esta função realiza a consulta de cadastro de contribuinte
Public Function consultarCadastroContribuinte(CNPJCont As String, UF As String, documentoConsulta As String, tpConsulta As String) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """CNPJCont"":""" & CNPJCont & ""","
    JSON = JSON & """UF"":""" & UF & ""","
    JSON = JSON & """" & tpConsulta & """:""" & documentoConsulta & """"
    JSON = JSON & "}"

    'Informa a url para onde deve ser enviado
    url = "https://nfe.ns.eti.br/util/conscad"

    'Grava dados envio
    gravaLinhaLog ("[CONSULTA_CADASTRO_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    'Grava resposta API
    gravaLinhaLog ("[CONSULTA_CADASTRO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarCadastroContribuinte = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza a consulta de situação de uma NF-e
Public Function consultarSituacao(licencaCnpj As String, chNFe As String, tpAmb As String, versao As String) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """licencaCnpj"":""" & licencaCnpj & ""","
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """tpAmb"":""" & tpAmb & ""","
    JSON = JSON & """versao"":""" & versao & """"
    JSON = JSON & "}"

    'Informa a url para onde deve ser enviado
    url = "https://nfe.ns.eti.br/nfe/stats"
    
    'Grava dados envio
    gravaLinhaLog ("[CONSULTA_SITUACAO_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    'Grava resposta API
    gravaLinhaLog ("[CONSULTA_SITUACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarSituacao = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Function enviarEmail(chNFe As String, enviaEmailDoc As String, anexarPDF As String, Email) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String
    'Dim anexarPDF As String 'para facilidar a montagem do json, deixar como string

    anexarPDF = "true" 'aqui definimos como padrão que o pdf será anexado no email, mas pode ser alterado conforme a necessidade do parceiro

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & ""","
    JSON = JSON & """enviaEmailDoc"":" & enviaEmailDoc & ","
    If tpAmb = 2 Then
    JSON = JSON & """tpAmb"":" & tpAmb & ","
    End If
    JSON = JSON & """anexarPDF"":" & anexarPDF & "," 'adicionado aqui a concatenação da tag "anexarPDF" "tpAmb":"2",
    JSON = JSON & """email"":["
    
    'Tratamento para caso tenha sido passado mais de um e-mail no parâmetro "email"
    Dim emails() As String
    Dim i, quantidade As Integer
    
    'Divide por ocorrência de vírgula
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

    'Informa a url para onde deve ser enviado
    If NFCe = False Then
        url = "https://nfe.ns.eti.br/util/resendemail"
    Else
        url = "https://nfce.ns.eti.br/v1/util/resendemail"
    End If
    
    'Grava dados envio
    gravaLinhaLog ("[ENVIO_EMAIL_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    'Grava resposta API
    gravaLinhaLog ("[ENVIO_EMAIL_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    enviarEmail = resposta
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função realiza a inutilização de um intervalo de numeração de NF-e
Public Function inutilizar(cUF As String, tpAmb As String, Ano As String, CNPJ As String, Serie As String, nNFIni As String, nNFFin As String, xJust As String) As String
On Error GoTo tratar_erro
    Dim XML As String
    Dim JSON As String
    Dim url As String
    Dim resposta As String
'===========================================
'Monta o JSON
'===========================================
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
'==================================================
' Se for certificado tipo A1
'==================================================
If frmFaturamento_Prod_Serv_NFe_NS.txtTPCertificado = "A1" Then
'==================================================
'Informa a url para onde deve inutilizar
'==================================================

If NFCe = False Then
    url = "https://nfe.ns.eti.br/nfe/inut"
Else
    url = "https://nfce.ns.eti.br/v1/nfce/inut"
End If

'==================================================
'Grava dados envio no log
'==================================================
    gravaLinhaLog ("[INUTILIZACAO_DADOS]")
    gravaLinhaLog (JSON)
'==================================================
'Envia o json para a URL
'==================================================
    resposta = enviaConteudoParaAPI(JSON, url, "json")
'==================================================
'Grava resposta API
'==================================================
    gravaLinhaLog ("[INUTILIZACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    inutilizar = resposta
'==================================================
' Se for certificado A3
'==================================================
Else
'==================================================
' Dados da url pra criar o XML
'==================================================
url = "https://nfe.ns.eti.br/util/generateinut"
'==================================================
'Envia o json para a URL criar XML
'==================================================
resposta = enviaConteudoParaAPI(JSON, url, "json")
'==================================================
'Grava resposta API
'==================================================
gravaLinhaLog ("[INUTILIZACAO_RESPOSTA]")
gravaLinhaLog (resposta)
status = LerDadosJSON(resposta, "status", "", "")
'==================================================
' Se gerou o XML
'==================================================
If status = "200" Then 'Se gerou o xml
    XML = LerDadosJSON(resposta, "xml", "", "")
    'Debug.print XML
    XML = "<?xml version=""1.0"" encoding=""utf-8""?>" & XML
'==================================================
' Assina XML
'==================================================
Dim AssinaturaXML2 As New AssinaturaXML2.Principal
'==================================================
' Envia XML por CNPJ para DLL assinar
'==================================================
RetornoXML = AssinaturaXML2.assinarXML(XML, "infInut", CNPJ)
  'Debug.print retorno
'==================================================
' Se for certificado sem cnpj
'==================================================
    If RetornoXML = "Certificado Digital não encontrado" Then
'==================================================
' Envia XML por serial do certificado para DLL assinar
'==================================================
      RetornoXML = AssinaturaXML2.assinarXML(XML, "infInut", frmFaturamento_Prod_Serv_NFe_NS.txtSerialCertificado)
'==================================================
    End If
'==================================================
' Se der algum erro na assinatura
    If RetornoXML = "Certificado Digital não encontrado" Then
      USMsgBox "Verifique seu certificado digital e tente novamente!", vbCritical, "CAPRIND v5.0"
         Var1 = 0
         Exit Function
    End If
'==================================================
     Var1 = 1
'==================================================
' Assinatura com sucesso
'==================================================
XML = RetornoXML
'==================================================
'Grava dados envio
'==================================================
gravaLinhaLog ("[INUTILIZACAO_DADOS]")
  gravaLinhaLog (XML)
'=================================================
'Envia o XML assinado para a URL de inutilização
'=================================================
    url = "https://nfe.ns.eti.br/nfe/inut"
    resposta = enviaConteudoParaAPI(XML, url, "xml")
End If
inutilizar = resposta
End If
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função faz a listagem de nsNRec vinculados a uma chave de NF-e
Public Function listarNSNRecs(chNFe As String) As String
On Error GoTo tratar_erro

    Dim JSON As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    JSON = "{"
    JSON = JSON & """chNFe"":""" & chNFe & """"
    JSON = JSON & "}"

    'Informa a url para onde deve ser enviado
    url = "https://nfe.ns.eti.br/util/list/nsnrecs"
    
    'Grava dados envio
    gravaLinhaLog ("[LISTA_NSNRECS_DADOS]")
    gravaLinhaLog (JSON)
        
    'Envia o json para a URL
    resposta = enviaConteudoParaAPI(JSON, url, "json")
    
    'Grava resposta API
    gravaLinhaLog ("[LISTA_NSNRECS_RESPOSTA]")
    gravaLinhaLog (resposta)

    listarNSNRecs = resposta

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função salva um XML
Public Sub salvarXML(XML As String, caminho As String, chNFe As String, nSeqEvento As String)
On Error GoTo tratar_erro

    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo XML
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procNFe.xml"
    Else
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procEvenNFe.xml"
    End If

    'Remove as contrabarras
    conteudoSalvar = Replace(XML, "\""", "")

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    If DS.FileOrDirExists(localParaSalvar) = False Then
    fsT.SaveToFile localParaSalvar
    Else
    USMsgBox "Já existe o arquivo XML salvo no local " & caminho & "!", vbCritical, "CAPRIND v5.0"
    End If
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

'Esta função salva um JSON
Public Sub salvarJSON(JSON As String, caminho As String, chNFe As String, nSeqEvento As String)
On Error GoTo tratar_erro

    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo JSON
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procNFe.json"
    Else
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procEvenNFe.json"
    End If

    conteudoSalvar = JSON

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

'Esta função lê os dados de um JSON
Public Function LerDadosJSON(sJsonString As String, key1 As String, key2 As String, key3 As String, Optional key4 As String, Optional key5 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" And key5 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet), key5, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet)
    ElseIf key1 <> "" And key2 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet)
    ElseIf key1 <> "" Then
        LerDadosJSON = VBA.CallByName(objJSON, key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    LerDadosJSON = "Error: " & Err.Description
    'Debug.print LerDadosJSON
    Resume Err_Exit
End Function

'Esta função lê os dados de um XML
Public Function LerDadosXML(sXml As String, key1 As String, key2 As String) As String
On Error GoTo tratar_erro

    On Error Resume Next
    LerDadosXML = ""
    
    Set XML = New DOMDocument50
    XML.async = False
    
    If XML.loadXML(sXml) Then
        'Tentar pegar o strCampoXML
        Set objNodeList = XML.getElementsByTagName()
        'Set objNodeList = XML.getElementsByTagName("descEvento")
        Set objNode = objNodeList.NextNode
        
        Dim valor As String
        valor = objNode.Text
        'Debug.print valor
        If Len(Trim(valor)) > 0 Then 'CONSEGUI LER O XML NODE
            LerDadosXML = valor
        End If
        Else
        USMsgBox "Não foi possível ler o conteúdo do XML da NFe especificado para leitura.", vbCritical, "CAPRIND v5.0", "ERRO"
    End If
  
  
 Set objXML = CreateObject("Microsoft.XMLDOM")
objXML.async = False

objXML.Load (sXml)

Dim Root, i

Set Root = objXML.documentElement
USMsgBox "Numeros de TAGs (Nodes): " & Root.childNodes.Length

For i = 0 To (Root.childNodes.Length) - 1
  MsgBox (Root.childNodes.Item(i).selectSingleNode("title").nodeTypedValue)
Next
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

'Esta função grava uma linha de texto em um arquivo de log
Public Sub gravaLinhaLog(conteudoSalvar As String)
On Error GoTo tratar_erro

    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim localParaSalvar As String
    Dim Data1 As String
    
    'Diretório para salvar os logs
    localParaSalvar = App.Path & "\log\"
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(localParaSalvar, vbDirectory) = "" Then
        MkDir (localParaSalvar)
    End If
    
    localParaSalvar = App.Path & "\log\" & nfDocumento & ".txt"
    TextoRetorno = TextoRetorno & Data & " - " & conteudoSalvar & vbCrLf
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
