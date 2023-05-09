Attribute VB_Name = "mdl_XML"
'====================================================
' Variaveis de importação XML
'====================================================
Public CST As String
Public CSTICMS As String
Public CSTPIS As String
Public CSTCOFINS As String
Public CSTIPI As String
Public orig As String
Public CSTOrig As String
Public CRT As String
Public natOp As String
Public tpNF As String ' Tipo da nota 0 = entrada 1 = saida
Public nNF As String
Public retorno As Boolean

Public NFentrada As Boolean 'Nota fiscal de entrada
Public NFSaida As Boolean 'Nota fiscal de saida
Public TPNota As String 'Tipo da nota própria(P) ou terceiros(T)

Public strPedido As String
Public ImportarXML As Boolean
Public strCaminho As String


Public PosicaoAntiga As Long
Public PosicaoLimite As Long
Public PosicaoBase As Long
Public lngPosicaoInicial As Long
Public lngPosicaoFinal As Long
Public lngPosicaoAuxiliar As Long
Public strarquivo As String
Public n As Long
Public lLinha As Integer

'Variaveis importacao do XML
'Dados da nota fiscal
Public NFE_cUF
Public NFE_cNF
Public NFE_natOp
Public NFE_modNFe As String
Public NFE_Serie
Public NFE_nNF
Public NFE_dhEmi
Public NFE_dhSaiEnt
Public NFE_tpNF
Public NFE_idDest
Public NFE_cMunFG
Public NFE_tpImp
Public NFE_tpEmis
Public NFE_cDV
Public NFE_tpAmb
Public NFE_finNFe
Public NFE_indFinal
Public NFE_indPres
Public NFE_procEmi
Public NFE_verProc

'Dados do emitente>
Public EMITCNPJ
Public EMITxNome
Public EMITxFant
Public EMITenderEmit
Public EMITxLgr
Public EMITnro
Public EMITxBairro
Public EMITcMun
Public EMITxMun
Public EMITUF
Public EMITCEP
Public EMITcPais
Public EMITxPais
Public EMITfone

Public EMITie
Public EMITCRT

'Dados do destinatário
Public DestidEstrangeiro
Public DestxNome

'endereço destinatário
Public DestxLgr
Public Destnro
Public DestxCpl
Public DestxBairro
Public DestcMun
Public DestxMun
Public DestUF
Public DestcPais
Public DestxPais


Public DestindIEDest

'autorizado receber XML
Public AutCNPJ

'Detalhes dos produtos
Public cProd
Public cEAN
Public xProd
Public NCM
Public CFOP
Public uCom
Public qCom
Public vUnCom
Public vProd
Public cEANTrib
Public uTrib
Public qTrib
Public vUnTrib
Public vFrete
Public vOutro
Public indTot

'Documento de importação
Public ImpnDI
Public ImpdDI
Public ImpxLocDesemb
Public ImpUFDesemb
Public ImpdDesemb
Public ImptpViaTransp
Public ImpvAFRMM
Public ImptpIntermedio
Public ImpCNPJ
Public ImpUFTerceiro
Public ImpcExportador

'adi
Public ImpAdinAdicao
Public ImpAdinSeqAdic
Public ImpAdicFabricante

'Dados dos impostos
'ICMS>
'ICMSSN900>
Public ICMSorig
Public ICMSCSOSN
Public ICMSmodBC
Public ICMSvBC
Public ICMSpRedBC
Public ICMSpICMS
Public ICMSvICMS
Public ICMSmodBCST
Public ICMSpMVAST
Public ICMSpRedBCST
Public ICMSvBCST
Public ICMSpICMSST
Public ICMSvICMSST
Public ICMSpCredSN
Public ICMSvCredICMSSN

'Imposto IPI
Public IPIcEnq
'IPITrib>
Public IPICST
Public IPIvBC
Public IPIpIPI
Public IPIvIPI

'Imposto de importação
Public IIvBC
Public IIvDespAdu
Public IIvII
Public IIvIOF

'Imposto PIS
'PISOutr>
Public PISCST
Public PISvBC
Public PISpPIS
Public PISvPIS

'Imposto COFINS
'COFINSOutr>
Public COFINSCST
Public COFINSvBC
Public COFINSpCofins
Public COFINSvCOFINS

'Total Nota fiscal
'ICMSTot>
Public TOTALvBC
Public TOTALvICMS
Public TOTALvICMSDeson
Public TOTALvFCP
Public TOTALvBCST
Public TOTALvST
Public TOTALvFCPST
Public TOTALvFCPSTRet
Public TOTALvProd
Public TOTALvFrete
Public TOTALvSeg
Public TOTALvDesc
Public TOTALvII
Public TOTALvIPI
Public TOTALvIPIDevol
Public TOTALvPIS
Public TOTALvCOFINS
Public TOTALvOutro
Public TOTALvNF
Public TOTALvTotTrib

'Dados Frete
Public FRETEEmodFrete

'Dados Transportadora
Public FRETECNPJ
Public FRETExNome
Public FRETEie
Public FRETExEnder
Public FRETExMun
Public FRETEUF

'Dados dos volumes
Public FRETEqVol
Public FRETEesp
Public FRETEmarca
Public FRETEnVol
Public FRETEpesoL
Public FRETEpesoB


'Dados das cobranças
'faturas
Public FATURAnFat
Public FATURAvOrig
Public FATURAvDesc
Public FATURAvLiq

'Pagamentos
'Detalhes Pagamentos
Public FATURAtPag
Public FATURAvPag

'informações Adicionais
Public ADICIONAISinfCpl

Public Function ProcImportarXMLCarregacampo(V1 As String, V2 As String, V3 As Integer)
On Error GoTo tratar_erro

        lngPosicaoInicial = InStr(IIf(PosicaoBase > 0, PosicaoBase, 1), strarquivo, V1, 1)
        PosicaoBase = lngPosicaoInicial
        lngPosicaoFinal = InStr(IIf(PosicaoBase > 0, PosicaoBase, 1), strarquivo, V2, 1)

        
    If lngPosicaoFinal > 0 And lngPosicaoInicial > 0 Then
        If lngPosicaoFinal > lngPosicaoInicial Then
            ProcImportarXMLCarregacampo = Mid(strarquivo, lngPosicaoInicial + V3, (lngPosicaoFinal - (lngPosicaoInicial + V3)))
            PosicaoBase = lngPosicaoFinal
        End If
    End If
    
    'Debug.print PosicaoBase
    

 
    
Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub ProcImportarXMLDadosNota()
On Error GoTo tratar_erro
ID_nota = 0

If strCaminho = "" Then
    frmEstoque_Recebimento_Menu.CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml"
    frmEstoque_Recebimento_Menu.CommonDialog1.ShowOpen
    
    strCaminho = frmEstoque_Recebimento_Menu.CommonDialog1.filename
        If strCaminho = "" Then
            Var = 0
            ImportarXML = False
            Exit Sub
        Else
            Var = 1
        End If
    Else
        Var = 1
 End If

PosicaoBase = 0
    frmfaturamento_Nova_Nota.lblXML.Caption = "Importação de arquivo XML sendo executado, aguarde..."
    frmfaturamento_Nova_Nota.Refresh
    
    ' Ler arquivo XML
    n = FreeFile()
    Open strCaminho For Input As #n
    strarquivo = Input(LOF(n), n)
    strarquivo = Replace(strarquivo, "ï»¿", "")
    Close #n
    
    infNFe = ProcImportarXMLCarregacampo("<infNFe", "/infNFe>", Len("<infNFe"))
    'Debug.print Left$(infNFe, 4)
    If Left$(infNFe, 4) = " Id=" Then
    infNFe = Left$(infNFe, 52)
    infNFe = Right$(infNFe, 44)
    Else
    infNFe = Left$(infNFe, 66)
    infNFe = Right$(infNFe, 44)
    End If
'========================================================
' Chave de cesso
'========================================================
PosicaoBase = 0
    chNF = infNFe
    
'Dados da nota fiscal
    V1 = "ide"
    PosicaoBase = InStr(1, strarquivo, V1, 1)

    'Dados da nota fiscal
    cUF = ProcImportarXMLCarregacampo("<cUF>", "</cUF>", Len("<cUF>"))
    
    natOp = UCase(ProcImportarXMLCarregacampo("<natOp>", "</natOp>", Len("<natOp>")))
    indmod = UCase(ProcImportarXMLCarregacampo("<mod>", "</mod>", Len("<mod>")))
    Serie = ProcImportarXMLCarregacampo("<serie>", "</serie>", Len("<serie>"))
    nNF = ProcImportarXMLCarregacampo("<nNF>", "</nNF>", Len("<nNF>"))
    nNF = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(nNF), 9)
    dhEmi = ProcImportarXMLCarregacampo("<dhEmi>", "</dhEmi>", Len("<dhEmi>"))
    dhEmi = Replace(dhEmi, "T", " ")
    dhEmi = Left$(dhEmi, 19)
    dhEmi = Format(dhEmi, "General date")
   
    dhSaiEnt = ProcImportarXMLCarregacampo("<dhSaiEnt>", "</dhSaiEnt>", Len("<dhSaiEnt>"))
    dhSaiEnt = Replace(dhSaiEnt, "T", " ")
    dhSaiEnt = Left$(dhSaiEnt, 19)
    dhSaiEnt = Format(dhSaiEnt, "General date")
    
    tpNF = UCase(ProcImportarXMLCarregacampo("<tpNF>", "</tpNF>", Len("<tpNF>")))
    idDest = UCase(ProcImportarXMLCarregacampo("<idDest>", "</idDest>", Len("<idDest>")))
    cMunFG = UCase(ProcImportarXMLCarregacampo("<cMunFG>", "</cMunFG>", Len("<cMunFG>")))
    
    tpImp = UCase(ProcImportarXMLCarregacampo("<tpImp>", "</tpImp>", Len("<tpImp>")))
    tpEmis = UCase(ProcImportarXMLCarregacampo("<tpEmis>", "</tpEmis>", Len("<tpEmis>")))
    'cMunFG  = UCase(ProcImportarXMLCarregacampo("<cMunFG>", "</cMunFG>", Len("<cMunFG>")))
    cDV = UCase(ProcImportarXMLCarregacampo("<cDV>", "</cDV>", Len("<cDV>")))
    tpAmb = UCase(ProcImportarXMLCarregacampo("<tpAmb>", "</tpAmb>", Len("<tpAmb>")))
    
    finNFe = UCase(ProcImportarXMLCarregacampo("<finNFe>", "</finNFe>", Len("<finNFe>")))
    
    Select Case finNFe
        Case "1"
        finNFe = "NF-e normal"
        Case "2"
        finNFe = "NF-e complementar"
        Case "3"
        finNFe = "NF-e de ajuste"
        Case "4"
        finNFe = "Devolução/Retorno"
    End Select
   
    
    indFinal = UCase(ProcImportarXMLCarregacampo("<indFinal>", "</indFinal>", Len("<indFinal>")))
    Select Case indFinal
        Case "0"
        indFinal = "Não"
        Case "1"
        indFinal = "Consumidor final"
    End Select
    
    indPres = UCase(ProcImportarXMLCarregacampo("<indPres>", "</indPres>", Len("<indPres>")))
    
    procEmi = UCase(ProcImportarXMLCarregacampo("<procEmi>", "</procEmi>", Len("<procEmi>")))
    verProc = UCase(ProcImportarXMLCarregacampo("<verProc>", "</verProc>", Len("<verProc>")))
    
    
    Select Case indPres
        Case "0"
        indPres = "Não se aplica"  ' (por exemplo, para a Nota Fiscal complementar ou de ajuste);
        Case "1"
        indPres = "Operação presencial"
        Case "2"
        indPres = "Operação não presencial, pela Internet"
        Case "3"
        indPres = "Operação não presencial, Teleatendimento;"
        Case "4"
        indPres = "NFC-e em operação com entrega em domicílio;"
        Case "5"
        indPres = "Operação presencial, fora do estabelecimento"
        Case "9"
        indPres = "Operação não presencial, outros."
    End Select

    'Dados do emitente
'    CNPJ  = LerDadosXML(strarquivo, "emit", "CNPJ")
    CNPJ = ProcImportarXMLCarregacampo("<CNPJ>", "</CNPJ>", Len("<CNPJ>"))
    CNPJ = Format(CNPJ, "@@.@@@.@@@/@@@@-@@")
    xNome = UCase(ProcImportarXMLCarregacampo("<xNome>", "</xNome>", Len("<xNome>")))
    xFant = UCase(ProcImportarXMLCarregacampo("<xFant>", "</xFant>", Len("<xFant>")))
    
    'Endereço emitente
    xLgr = UCase(ProcImportarXMLCarregacampo("<xLgr>", "</xLgr>", Len("<xLgr>")))
    nro = ProcImportarXMLCarregacampo("<nro>", "</nro>", Len("<nro>"))
    xBairro = UCase(ProcImportarXMLCarregacampo("<xBairro>", "</xBairro>", Len("<xBairro>")))
    cMun = UCase(ProcImportarXMLCarregacampo("<cMun>", "</cMun>", Len("<cMun>")))
    xMun = UCase(ProcImportarXMLCarregacampo("<xMun>", "</xMun>", Len("<xMun>")))
    UF = UCase(ProcImportarXMLCarregacampo("<UF>", "</UF>", Len("<UF>")))
    CEP = ProcImportarXMLCarregacampo("<CEP>", "</CEP>", Len("<CEP>"))
'    cPais = UCase(ProcImportarXMLCarregacampo("<cPais>", "</cPais>", Len("<cPais>")))
    xPais = UCase(ProcImportarXMLCarregacampo("<xPais>", "</xPais>", Len("<xPais>")))
    'Var1 = "fone"
    'fone = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
        
    Var1 = "IE"
    Var_IE = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

'    Var_IE = UCase(ProcImportarXMLCarregacampo("<IE>", "</IE>", Len("<IE>")))
    
    Var1 = "CRT"
    CRT = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
 '   CRT = UCase(ProcImportarXMLCarregacampo("<CRT>", "</CRT>", Len("<CRT>")))

    'Dados do Destinatário
 PosicaoBase = 0
 PosicaoInicial = 0
 PosicaoFinal = 0
 lngPosicaoFinal = 0
 
    V1 = "<dest>"
    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
    'Fornecedor = ProcImportarXMLCarregacampo("<" & V1 & ">", "</" & V1 & ">", Len("<" & V1 & ">"))
    V1 = "</dest>"
    
    PosicaoLimite = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
    PosicaoAntiga = PosicaoBase
    
    dest_CNPJ = ProcImportarXMLCarregacampo("<CPF>", "</CPF>", Len("<CPF>"))
    dest_CNPJ = Format(dest_CNPJ, "@@@.@@@.@@@-@@")
    
    If PosicaoBase = 0 Then
    PosicaoBase = 0
    PosicaoInicial = 0
    PosicaoFinal = 0
    V1 = "<dest>"
    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)

    dest_CNPJ = ProcImportarXMLCarregacampo("<CNPJ>", "</CNPJ>", Len("<CNPJ>"))
    dest_CNPJ = Format(dest_CNPJ, "@@.@@@.@@@/@@@@-@@")
    
    If PosicaoBase > PosicaoLimite Then
    PosicaoBase = 0
    PosicaoInicial = 0
    lngPosicaoFinal = PosicaoAntiga
    V1 = "<dest>"
    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
    End If
    
    End If
    
    dest_xNome = UCase(ProcImportarXMLCarregacampo("<xNome>", "</xNome>", Len("<xNome>")))
    
    'Endereço Destinatario
    dest_xLgr = UCase(ProcImportarXMLCarregacampo("<xLgr>", "</xLgr>", Len("<xLgr>")))
    dest_nro = ProcImportarXMLCarregacampo("<nro>", "</nro>", Len("<nro>"))
    dest_xBairro = UCase(ProcImportarXMLCarregacampo("<xBairro>", "</xBairro>", Len("<xBairro>")))
    dest_xMun = UCase(ProcImportarXMLCarregacampo("<xMun>", "</xMun>", Len("<xMun>")))
    dest_UF = UCase(ProcImportarXMLCarregacampo("<UF>", "</UF>", Len("<UF>")))
    dest_CEP = ProcImportarXMLCarregacampo("<CEP>", "</CEP>", Len("<CEP>"))
    dest_xPais = UCase(ProcImportarXMLCarregacampo("<xPais>", "</xPais>", Len("<xPais>")))
    dest_indIEDest = UCase(ProcImportarXMLCarregacampo("<indIEDest>", "</indIEDest>", Len("<indIEDest>")))
   
    Select Case dest_indIEDest
      Case "1": dest_indIEDest = "1 - Contribuinte ICMS (informar a IE do destinatário)"
      Case "2": dest_indIEDest = "2 - Contribuinte isento de Inscrição no cadastro de Contribuintes"
      Case "9": dest_indIEDest = "9 - Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS."
    End Select
'=========================================================================
'  Gravação dos dados na NFe
'=========================================================================
'If frmFaturamento_Prod_Serv.opt_Saida = True Then
'TPNota = "P"
'Else
TPNota = "T"
'End If

'   If TPNota = "P" Then
'    tpNF = "1"
'    Aplicacao = "P"
'    TipoNF = "M1"
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select CAST(int_NotaFiscal AS int) AS NF, Serie FROM tbl_Dados_Nota_Fiscal where tipoNF = '" & TipoNF & "' and Aplicacao = 'P' and ID_empresa = '1' and int_NotaFiscal IS NOT NULL order by  NF desc,dt_DataEmissao desc", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
   ' TBAbrir.MoveLast
'        QuantsolicitadoN1 = TBAbrir!NF + 1
'        FamiliaAntiga = QuantsolicitadoN1
'        nNF = FunTamanhoTextoZeroEsq(FamiliaAntiga, 9)
'        Serie = IIf(IsNull(TBAbrir!Serie), 1, TBAbrir!Serie)
'    End If
    
'   Else
    Aplicacao = "T"
    tpNF = "2"
'   End If

Set TBGravar = CreateObject("adodb.recordset")
StrSql = "Select * from tbl_Dados_Nota_Fiscal where int_NotaFiscal = '" & nNF & "' and txt_CNPJ_CPF = '" & CNPJ & "' and int_TipoNota = '" & tpNF & "' and Serie = '" & Serie & "' and ID_empresa = '" & IDempresa & "' "

'Debug.print StrSql

TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        ImportarXML = True
    Else
        ImportarXML = False
    Exit Sub
End If


    TBGravar!Aplicacao = Aplicacao
    TBGravar!TabelaSN = 2
    TBGravar!Regime = FunVerifRegimeEmpresa(IDempresa)
    TBGravar!pedido_interno = False
    TBGravar!DtValidacaoOF = Now
    TBGravar!RespValidacaoOF = pubUsuario
    TBGravar!ID_empresa = IDempresa
    TBGravar!int_NotaFiscal = nNF
    TBGravar!Serie = Serie
    TBGravar!int_TipoNota = tpNF
    TBGravar!TipoNF = "M1"
    TBGravar!dt_DataEmissao = dhEmi
    TBGravar!dt_Saida_Entrada = dhEmi  'dhSaiEnt
    TBGravar!Hora_emissao = Format(dhEmi, "hh:mm")
    TBGravar!Modelo = indmod
    TBGravar!DtValidacao = Date
    TBGravar!RespValidacao = pubUsuario

'===========================================
' Dados do destinatário
'===========================================
'  If tpNF = "1" And TBGravar!Aplicacao = "P" Then
Set TBClientes = CreateObject("adodb.recordset")
   TBClientes.Open "Select * from Clientes where CPF_CNPJ = '" & CNPJ & "'", Conexao, adOpenKeyset, adLockOptimistic
     If TBClientes.EOF = False Then
        TBGravar!Id_Int_Cliente = TBClientes!IDCliente
        NomeRazao = TBClientes!NomeRazao
        IDCliente = TBClientes!IDCliente
        TBClientes.Close
     Else
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from Compras_fornecedores where CPF_CNPJ = '" & CNPJ & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = False Then
            TBGravar!Id_Int_Cliente = TBFornecedor!IDCliente
            NomeRazao = TBFornecedor!Nome_Razao
            IDCliente = TBFornecedor!IDCliente
           TBFornecedor.Close
        End If
     End If
   
    TBGravar!txt_Razao_Nome = IIf(NomeRazao <> "", NomeRazao, xNome)
    TBGravar!txt_Endereco = IIf(xLgr <> "", xLgr, "Sem endereço")
    TBGravar!txt_Bairro = xBairro
    TBGravar!txt_tipocliente = IIf(Len(CNPJ) <> 14, "JP", "FP")
    TBGravar!txt_UF = UF
    TBGravar!txt_CNPJ_CPF = CNPJ
    TBGravar!Txt_CEP = CEP
    TBGravar!txt_Municipio = xMun
    TBGravar!txt_Hora_Saida = Format(dhEmi, "hh:mm")
    TBGravar!Int_status = "1"
    TBGravar!Numero = nro

    
    TBGravar.Update
    
ID_nota = TBGravar!ID

TBGravar.Close
'===========================================================================================
'Gravar chave de acesso
'===========================================================================================
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!ID_nota = ID_nota
    TBGravar!Chave_acesso = chNF
    TBGravar!Finalidade_emissao = 1 'finNFe
    'TBGravar!status = "100"
    
    Select Case indFinal
        Case "Não" '"0"
        TBGravar!Consumidor_final = "0"
        Case "Consumidor final" '"1"
        TBGravar!Consumidor_final = "1"
    End Select
       
       
    Select Case indPres
        Case "Não se aplica" '"0"
        TBGravar!Presenca_comprador = "0"
        Case "Operação presencial" '"1"
         TBGravar!Presenca_comprador = "1"
        Case "Operação não presencial, pela Internet" '"2"
         TBGravar!Presenca_comprador = "2"
        Case "Operação não presencial, Teleatendimento;" '"3"
         TBGravar!Presenca_comprador = "3"
        Case "NFC-e em operação com entrega em domicílio;" '"4"
         TBGravar!Presenca_comprador = "4"
        Case "Operação presencial, fora do estabelecimento" '"5"
         TBGravar!Presenca_comprador = "5"
        Case "Operação não presencial, outros." '"9"
         TBGravar!Presenca_comprador = "9"
    End Select
       
    TBGravar.Update
End If
TBGravar.Close

''===========================================================================================
''Gravar Pedido de compras se for nota de entrada
''===========================================================================================
If tpNF = 2 And strPedido <> "" Then
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_proposta_nota where NF = '" & nNF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
End If
Do While TBGravar.EOF = False
    TBGravar!ID_nota = ID_nota
    TBGravar!Proposta = strPedido
    TBGravar!NF = nNF
    TBGravar!Revisao = 0 'frmEstoque_Recebimento.txtrev.Text
TBGravar.Update
TBGravar.MoveNext
Loop

TBGravar.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST51()
On Error GoTo tratar_erro


        Var1 = "modBC"
        modBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If
    
        Var1 = "pRedBC"
        p_RedBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If

        Var1 = "vbc"
        v_bc = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If
    
        Var1 = "pICMS"
        p_ICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If
    
        Var1 = "vICMSOp"
        v_ICMSOp = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If

        Var1 = "pDif"
        p_Dif = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If

        Var1 = "vICMSDif"
        v_ICMSDif = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If

        Var1 = "vICMS"
        v_ICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
        PosicaoBase = PosicaoAntiga
        End If
        
    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST30()
On Error GoTo tratar_erro


        Var1 = "modBCST"
        modBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pMVAST"
        p_MVAST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        Var1 = "pRedBCST"
        p_RedBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
          
        Var1 = "vBCST"
        v_BCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        

        Var1 = "pICMSST"
        p_ICMSST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMSST"
        v_ICMSST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST40()
On Error GoTo tratar_erro

    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST50()
On Error GoTo tratar_erro

    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST20()
On Error GoTo tratar_erro



        Var1 = "modBC"
        modBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        
        Var1 = "pRedBC"
        p_RedBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vBC"
        vbcIMCS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        Var1 = "pICMS"
        ICMSpICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMS"
        ICMSvICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST00()
On Error GoTo tratar_erro

        Var1 = "modBC"
        modBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vBC"
        IMCSpi = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pICMS"
        ICMSpICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMS"
        ICMSvICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST10()
On Error GoTo tratar_erro

        Var1 = "modBC"
        modBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vBC"
        vbcIMCS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pICMS"
        ICMSpICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMS"
        ICMSvICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "modBCST"
        ICMSmodBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

 '       Var1 = "pRedBCST"
 '       ICMSpRedBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vBCST"
        ICMSvBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pICMSST"
        ICMSpICMSST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMSST"
        ICMSvICMSST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST60()
On Error GoTo tratar_erro

        Var1 = "vBCSTRet"
        v_BCSTRet = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        Var1 = "pST"
        pST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        Var1 = "vICMSSubstituto"
        v_ICMSSubstituto = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        Var1 = "vICMSSTRet"
        v_ICMSSTRet = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcVerificaCST()
On Error GoTo tratar_erro
If CSTIPI = "" Then
CSTIPI = "99"
End If

If CSTPIS = "" Then
CSTPIS = "99"
End If

If CSTCOFINS = "" Then
CSTCOFINS = "99"
End If


If vICMSCST <> "" Then
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_NaturezaOperacao_CST where ID_CFOP = " & TBAbrir!ID_CFOP & " and CST_ICMS = '" & vICMSCST & "' and CST_IPI = '" & CSTIPI & "' and CST_PIS = '" & CSTPIS & "' and CST_COFINS = '" & CSTCOFINS & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!ID_CFOP = TBAbrir!ID_CFOP
    TBGravar!CST_ICMS = vICMSCST
    TBGravar!CST_IPI = CSTIPI
    TBGravar!CST_PIS = CSTPIS
    TBGravar!CST_Cofins = CSTCOFINS
    TBGravar.Update
End If
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Public Sub ProcBuscaXMlPasta()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
DiretorioRetorno = TBAbrir!Caminho_RetornoNfe
IDempresa = TBAbrir!CODIGO
End If
TBAbrir.Close


If USMsgBox("Deseja realmente importar todos os arquivos XML contidos na pasta " & DiretorioRetorno & "XML?", vbYesNo, "CAPRIND v5.0") = vbNo Then
Exit Sub
End If

Dim Diretorio As Folder
Dim Arquivo As File
Dim Fso As New FileSystemObject
Dim CboDiretorio As ComboBox
Dim CboArquivo As ComboBox


Set Diretorio = Fso.GetFolder(DiretorioRetorno & "XML\")

For Each Arquivo In Diretorio.files
If Right(Arquivo.Name, 3) = "xml" Then
TPNota = "P"
strCaminho = Arquivo
ProcImportarXML
FileCopy Arquivo, DiretorioRetorno & "XML\Importados\" & Arquivo.Name
Kill Arquivo
End If
Next


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcImportarXMLTotaisNota()
On Error GoTo tratar_erro
'==============================================================================
' Carregatotais da nota
'==============================================================================

    V1 = "total"
    PosicaoBase = InStr(1, strarquivo, V1, 1)

    Var1 = "vBC"
    vBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vBC = Replace(vBC, ".", ",")
    vBC = Format(vBC, "###,##0.00")

    Var1 = "vICMS"
    vICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vICMS = Replace(vICMS, ".", ",")
    vICMS = Format(vICMS, "###,##0.00")
    
    Var1 = "vICMSDeson"
    vICMSDeson = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vICMSDeson = Replace(vICMSDeson, ".", ",")
    vICMSDeson = Format(vICMSDeson, "###,##0.00")
    
    Var1 = "vFCP"
    vFCP = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFCP = Replace(vFCP, ".", ",")
    vFCP = Format(vFCP, "###,##0.00")
    
    Var1 = "vBCST"
    vBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vBCST = Replace(vBCST, ".", ",")
    vBCST = Format(vBCST, "###,##0.00")
    
    Var1 = "vST"
    vST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vST = Replace(vST, ".", ",")
    vST = Format(vST, "###,##0.00")
    
    Var1 = "vFCPST"
    vFCPST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFCPST = Replace(vFCPST, ".", ",")
    vFCPST = Format(vFCPST, "###,##0.00")
    
    Var1 = "vFCPSTRet"
    vFCPSTRet = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFCPSTRet = Replace(vFCPSTRet, ".", ",")
    vFCPSTRet = Format(vFCPSTRet, "###,##0.00")
    
    
    Var1 = "vProd"
    vProdTotal = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vProdTotal = Replace(vProdTotal, ".", ",")
    vProdTotal = Format(vProdTotal, "###,##0.00")
    
    
    Var1 = "vFrete"
    vFrete = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFrete = Replace(vFrete, ".", ",")
    vFrete = Format(vFrete, "###,##0.00")
    
    
    Var1 = "vSeg"
    vSeg = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vSeg = Replace(vSeg, ".", ",")
    vSeg = Format(vSeg, "###,##0.00")

    Var1 = "vDesc"
    vDesc = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vDesc = Replace(vDesc, ".", ",")
    vDesc = Format(vDesc, "###,##0.00")
        
    Var1 = "vIPI"
    vIPI = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vIPI = Replace(vIPI, ".", ",")
    vIPI = Format(vIPI, "###,##0.00")
    
    Var1 = "vIPIDevol"
    vIPIDevol = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vIPIDevol = Replace(vIPIDevol, ".", ",")
    vIPIDevol = Format(vIPIDevol, "###,##0.00")
    
    Var1 = "vPIS"
    vPIS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vPIS = Replace(vPIS, ".", ",")
    vPIS = Format(vPIS, "###,##0.00")
        
    Var1 = "vCOFINS"
    vCOFINS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vCOFINS = Replace(vCOFINS, ".", ",")
    vCOFINS = Format(vCOFINS, "###,##0.00")
    
    Var1 = "vOutro"
    vOutro = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vOutro = Replace(vOutro, ".", ",")
    vOutro = Format(vOutro, "###,##0.00")
    
    Var1 = "vNF"
    vNF = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vNF = Replace(vNF, ".", ",")
    vNF = Format(vNF, "###,##0.00")
    
    Var1 = "vTotTrib"
    vTotTrib = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vTotTrib = Replace(vTotTrib, ".", ",")
    vTotTrib = Format(vTotTrib, "###,##0.00")
    
    
  '========================================================================================================
' Cadastrar totais da nota fiscal
'========================================================================================================
 Set TBTotaisnota = CreateObject("adodb.recordset")
   TBTotaisnota.Open "Select * from tbl_Totais_Nota where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = True Then
    TBTotaisnota.AddNew
    End If
    
    TBTotaisnota!ID_nota = ID_nota
    TBTotaisnota!int_NotaFiscal = chNF 'Familiatext
    TBTotaisnota!dbl_Base_ICMS = vBC
    TBTotaisnota!dbl_Valor_ICMS = vICMS
    TBTotaisnota!dbl_Base_ICMS_Subst = vBCST
    TBTotaisnota!dbl_Valor_ICMS_Subst = vST
    TBTotaisnota!dbl_Valor_Total_Produtos = vProdTotal
    TBTotaisnota!dbl_Valor_Frete = vFrete
    TBTotaisnota!dbl_Valor_Seguro = vSeg
    TBTotaisnota!Valor_total_desconto = vDesc
    TBTotaisnota!dbl_Desp_Adicionais = "0.00" 'vOutro.Text
    TBTotaisnota!dbl_Valor_Total_IPI = vIPI
    TBTotaisnota!dbl_Valor_Total_Nota = vNF
    'TBTotaisnota!Valor_total_II = IIf(vImpostoImportacao.Text <> "", vImpostoImportacao.Text, "0")
    TBTotaisnota.Update
    'End If
    TBTotaisnota.Close
  
    
    
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcImportarXMLCFOP()
On Error GoTo tratar_erro

'If CFOP = "5902" Or CFOP = "6902" Then
'retorno = True
'Else
'retorno = False
'End If
If TPNota = "T" Then
If Left(CFOP, 1) <> 1 And Left(CFOP, 1) <> 2 Then
If Left(CFOP, 1) = 5 Then
CFOP = "1" & Right(CFOP, 3)
End If

If Left(CFOP, 1) = 6 Then
CFOP = "2" & Right(CFOP, 3)
End If
End If
End If

CFOP = Format(CFOP, "@.@@@")

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select * from tbl_NaturezaOperacao where id_CFOP = '" & CFOP & "'", Conexao, adOpenKeyset, adLockOptimistic
 If TBCFOP.EOF = True Then
 TBCFOP.AddNew
 TBCFOP!ID_CFOP = CFOP
 TBCFOP!Txt_descricao = natOp '"Venda de mercadoria"
 TBCFOP!Txt_ICMS = "SIM"
 TBCFOP!txt_IPI = "SIM"
 TBCFOP!Vendas = True
 TBCFOP!Data = Date
 TBCFOP!Responsavel = pubUsuario
 TBCFOP!Proprio = True
 TBCFOP!Terceiros = False
 TBCFOP!TemPIS = True
 TBCFOP!TemCOFINS = True
 TBCFOP!DtValidacao = Date
 TBCFOP!RespValidacao = pubUsuario
 TBCFOP!Tipo_CFOP = "1"
 TBCFOP!retorno = False
 TBCFOP!Remessa = False
 TBCFOP!Devolucao = False
 TBCFOP.Update
 Else
 TBAbrir!ID_CFOP = IIf(IsNull(TBCFOP!IDCountCfop), 0, TBCFOP!IDCountCfop)
 retorno = IIf(IsNull(TBCFOP!retorno), False, TBCFOP!retorno)
 Devolucao = IIf(IsNull(TBCFOP!Devolucao), False, TBCFOP!Devolucao)
 vICMSCST = orig & CSTICMS
 ProcVerificaCST
 End If
 TBCFOP.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcImportarXMLProdutos()
On Error GoTo tratar_erro
'=====================================
'Carrega Dados lista de produtos
'=====================================
 PosicaoBase = 0
 lngPosicaoInicial = 0
 lngPosicaoFinal = 0
    V1 = "prod"
    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
    
    
Inicio:
    
    If PosicaoBase > 0 Then
    

    
    Var1 = "cProd"
    cProd = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    
    V1 = "/det"
    PosicaoLimite = InStr(IIf(PosicaoBase > 0, PosicaoBase, 1), strarquivo, V1, 1)
    
 If lngPosicaoFinal = 0 Then
 GoTo FIM
 End If
 
'If PosicaoLimite = 0 Then GoTo FIM

    Var1 = "xProd"
    xProd = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "NCM"
    NCM = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "CFOP"
    CFOP = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "uCom"
    uCom = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "qCom"
    qCom = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "vUnCom"
    vUnCom = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    Var1 = "vProd"
    vProd = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
'    V1 = "vFrete"
'   PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
    
PosicaoAntiga = PosicaoBase
    
'    Var1 = "vFrete"
'    vFrete = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
' lngPosicaoInicial = 0
' lngPosicaoFinal = 0
1:
    
If PosicaoBase > PosicaoAntiga + 30 Then
PosicaoBase = PosicaoAntiga
PosicaoAntiga = PosicaoBase
GoTo 1
End If

PosicaoAntiga = PosicaoBase
'   vDesc = "0.00"
'    Var1 = "vDesc"
'    vDesc = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
' lngPosicaoInicial = 0
' lngPosicaoFinal = 0
2:
posicaocalculada = PosicaoAntiga + 141
If PosicaoBase > posicaocalculada Then
 vDesc = "0.00"
PosicaoBase = PosicaoAntiga
PosicaoAntiga = PosicaoBase
GoTo 2
End If

    Var1 = "orig"
    orig = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'3:
'If PosicaoBase > posicaoantiga + 30 Then
' vDesc = "0.00"
'PosicaoBase = posicaoantiga
'posicaoantiga = PosicaoBase
'GoTo 3
'End If
    

V1 = "</ICMS>"

PosicaoLimite = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
PosicaoAntiga = PosicaoBase

'==============================================================
' CST ICMS
'==============================================================
If CRT <> 3 Then
'========================================
'Simples nacional
'========================================
 Var1 = "CSOSN"
 Var3 = PosicaoBase
 CSTICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
     Select Case CSTICMS
     Case "101": ProcCSOSN101
     Case "102": ProcCSOSN102
    ' Case "201": ProcCSOSN101
    ' Case "202": ProcCSOSN101
     Case "400": ProcCSOSN400
    ' Case "500": ProcCSOSN101
    ' Case "900": ProcCSOSN101
     End Select
 Else
'========================================
'lucro presumido ou real
'========================================
Var1 = "CST"
ICMSpICMS = 0
ICMSvICMS = 0
ICMSpICMSST = 0
ICMSvBCST = 0
ICMSvICMSST = 0
ICMSCSOSN = 0

'PosicaoBase = PosicaoAntiga
PosicaoAntiga = PosicaoBase

 CSTICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
     Select Case CSTICMS
     Case "00": ProcCST00
     Case "10": ProcCST10
     Case "20": ProcCST20
     Case "30": ProcCST30
     Case "40": ProcCST40
     Case "51": ProcCST51
     Case "50": ProcCST50
     'Case "60": ProcCST60
     Case "70": ProcCST70
     Case "90": ProcCST90
    End Select
 End If
'==============================================================
' CST IPI
'==============================================================
'PosicaoBase = posicaoantiga
'lngPosicaoInicial = posicaoantiga
'lngPosicaoFinal = posicaoantiga

If PosicaoBase = 0 Then
PosicaoBase = PosicaoAntiga
End If

'Debug.print PosicaoBase
PosicaoAntiga = PosicaoBase
Var1 = "IPI"
 V1 = "/imposto"
 PosicaoLimite = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
 
 V1 = "IPI"
 PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)

If PosicaoBase <> 0 And PosicaoBase < PosicaoLimite Then
Var1 = "CST"
CSTIPI = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
If CSTIPI = "49" Or CSTIPI = "50" Or CSTIPI = "99" Then

Var1 = "pIPI"
IPIpIPI = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
PosicaoBase = PosicaoAntiga
End If

Var1 = "vIPI"
IPIvIPI = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
Else
IPIpIPI = 0
IPIvIPI = 0
End If
End If

If PosicaoBase = 0 Or PosicaoBase > PosicaoLimite Then
PosicaoBase = PosicaoAntiga
End If

'==============================================================
' CST PIS
'==============================================================
PosicaoAntiga = PosicaoBase

If PosicaoBase <> 0 Then
Var1 = "CST"
CSTPIS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
End If

If PosicaoBase = 0 Then
PosicaoBase = PosicaoAntiga
End If
'===================================================================================================
'CST Cofins
'===================================================================================================
PosicaoAntiga = PosicaoBase

If PosicaoBase <> 0 Then
Var1 = "CST"
CSTCOFINS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
End If

If PosicaoBase = 0 Then
PosicaoBase = PosicaoAntiga
End If
'=====================================================================================================
' Fim dos impostos
'=====================================================================================================
   
If cProd <> "" Then

If IsNumeric(vUnCom) Then
    vUnCom = Replace(vUnCom, ".", ",")
    vUnCom = "R$ " & vUnCom
End If

If IsNumeric(vProd) Then
    vProd = Replace(vProd, ".", ",")
    vProd = "R$ " & vProd
End If

ValorTotal = 0
'==============================================================
' Grava lista de produtos
'==============================================================
Codproduto = ""
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select * from item_aplicacoes where n_referencia = '" & cProd & "'", Conexao, adOpenKeyset, adLockOptimistic
      If TBComponente.EOF = False Then
      Codproduto = TBComponente!Codproduto

    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projProduto where codproduto = " & TBComponente!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
             If TBItem.EOF = False Then
                cProd = TBItem!Desenho
                cDesc = TBItem!Descricao
                Codproduto = TBItem!Codproduto
             Else
                Codproduto = 0
             End If
    End If
    
'================================================================
' Busca id da NCM
'================================================================
Set TBAliquota = CreateObject("adodb.recordset")
Dim ID_NCM As Long
'Debug.print Format(NCM, "@@.@@.@@@@")
    TBAliquota.Open "Select * from tbl_ClassificacaoFiscal where IDIntClasse = '" & Format(NCM, "@@@@.@@.@@") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAliquota.EOF = False Then
            ID_NCM = TBAliquota!Idclass
        Else
        
            ID_NCM = 0
        End If
    TBAliquota.Close
'================================================================
' Busca id da CFOP
'================================================================
    Set TBAliquota = CreateObject("adodb.recordset")
       TBAliquota.Open "Select * from tbl_NaturezaOperacao where id_CFOP = '" & Format(CFOP, "@.@@@@") & "'", Conexao, adOpenKeyset, adLockOptimistic
           If TBAliquota.EOF = False Then
               ID_CFOP = TBAliquota!IDCountCfop
           Else
               ID_CFOP = 0
           End If
       TBAliquota.Close
       
    
 If Right(CFOP, 3) = "901" Then
        Consignacao = True
        StatusRE = "CONSIGNAÇÃO RECEBIDA"
Else
        Consignacao = True
        StatusRE = "ENTRADA_NOTA_FISCAL"
End If
'================================================================
'Verifica se tem o item cadastrado
'================================================================
Teste = xProd
Dim strRef As String
Dim unProd As String
Dim Manual As Boolean


Set TBItem = CreateObject("adodb.recordset")
 TBItem.Open "Select * from projProduto where Desenho = '" & cProd & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        cProd = TBItem!Desenho 'TBItem!Codproduto
        Codproduto = TBItem!Codproduto
        cDesc = TBItem!Descricao
    Else
        If Familia = "" And LocalArmazenamento = "" Then
        Inspecao_recebimento = True
        Estoque = 1
            SubTipoItem = 3
         frmProduto_Novo.Show 1
        End If
            strRef = cProd
            DescricaoProduto = xProd
            unProd = uCom
            nReferencia = cProd
            Manual = CodManual
            cProd = FunCriaNovoProdServ(Manual, "codmanual = " & IIf(CodManual = False, 0, 1) & " and subtipoitem = " & SubTipoItem, "", strRef, 0, DescricaoProduto, DescricaoProduto, Familia, 0, 0, 0, unProd, unProd, ID_NCM, True, False, False, False, SubTipoItem, "P", "", 0, 0, 0, "", IDCliente, NomeRazao, "")
            StrSql = "Update projproduto set ID_CFOP1 = " & ID_CFOP & ", Estoque = " & IIf(Estoque = False, 0, 1) & ", Insp_recebimento = " & IIf(Inspecao_recebimento = False, 0, 1) & ", ID_Tipo = " & ID_Tipo & " where Codproduto = " & Codproduto & ""
            Conexao.Execute StrSql
    End If
    
    
'===============================================================================
' Inicio dos cadastros dos itens na nota fiscal
'===============================================================================
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Detalhes_Nota where ID_nota = " & ID_nota & " and int_Cod_Produto = '" & cProd & "' and int_Qtd = " & qCom & "", Conexao, adOpenKeyset, adLockOptimistic
TBAbrir.AddNew
TBAbrir!Tipo = "P"
TBAbrir!int_Cod_Produto = cProd
TBAbrir!Codproduto = IIf(Codproduto <> "", Codproduto, 0)
TBAbrir!N_referencia = nReferencia
TBAbrir!int_NotaFiscal = nNF
TBAbrir!ID_nota = ID_nota
TBAbrir!int_Qtd = Replace(qCom, ".", ",")
TBAbrir!Saldo = Replace(qCom, ".", ",")
TBAbrir!Txt_descricao = IIf(cDesc <> "", cDesc, xProd)

'============================================================================================================
'Carrega dados da DI
'============================================================================================================
'    V1 = "DI"
'    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
'Debug.print PosicaoBase
'If PosicaoBase > 0 Then
'
'    Var1 = "nDI"
'    nDI = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "dDI"
'    DDI = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "xLocDesemb"
'    xlocDesemb = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "UFDesemb"
'    UFDesemb = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "dDesemb"
'    dDesemb = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "tpViaTransp"
'    tpViaTransp = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "tpIntermedio"
'    TPIntermedio = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "cExportador"
'    cExportador = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'End If
'
''============================================================================================================
''Carrega dados das adicoes da DI
''============================================================================================================
'    V1 = "adi"
'    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
'Debug.print PosicaoBase
'If PosicaoBase > 0 Then
'    Var1 = "nAdicao"
'    nAdicao = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "nSeqAdic"
'    nSeqAdic = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "cFabricante"
'    cFabricante = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'End If
'
'Set TBAliquota = CreateObject("adodb.recordset")
'TBAliquota.Open "Select * from tbl_Detalhes_Nota_NFe where ID_item = '" & TBAbrir!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
'If TBAliquota.EOF = True Then
'TBAliquota.AddNew
'End If
'TBAliquota!ID_nota = ID_nota
'TBAliquota!Id_Item = TBAbrir!Int_codigo
'
'TBAliquota!Documento_importacao = nDI 'Lista.ListItems.Item(contador).ListSubItems(33).Text
'TBAliquota!Data_registro = DDI 'Lista.ListItems.Item(contador).ListSubItems(34).Text
'TBAliquota!Local_desembaraco = xlocDesemb 'Lista.ListItems.Item(contador).ListSubItems(35).Text
'TBAliquota!UF_desembaraco = UFDesemb 'Lista.ListItems.Item(contador).ListSubItems(36).Text
'TBAliquota!Data_desembaraco = dDesemb 'Lista.ListItems.Item(contador).ListSubItems(37).Text
'TBAliquota!Via_transp = tpViaTransp 'Lista.ListItems.Item(contador).ListSubItems(39).Text
'TBAliquota!Codigo_exportador = cExportador 'Lista.ListItems.Item(contador).ListSubItems(40).Text
'TBAliquota!Numero_adicao = nAdicao 'Lista.ListItems.Item(contador).ListSubItems(42).Text
'TBAliquota!Numero_sequencial = nSeqAdic 'contador
'TBAliquota!Codigo_fabricante = cFabricante 'Lista.ListItems.Item(contador).ListSubItems(43).Text
'
''TBAliquota!Valor_BC_importacao = Lista.ListItems.Item(contador).ListSubItems(28).Text
''TBAliquota!Valor_imposto_importacao = Lista.ListItems.Item(contador).ListSubItems(30).Text
''TBAliquota!Valor_despesas = Lista.ListItems.Item(contador).ListSubItems(29).Text
''TBAliquota!Valor_imposto_OperacoesFinanceiras = Lista.ListItems.Item(contador).ListSubItems(31).Text
''TBAliquota!NCM = Lista.ListItems.Item(contador).ListSubItems(3).Text
'TBAliquota.Update
'TBAliquota.Close


'=======================================================================================
' Busca dados da CFOP
'=======================================================================================
ProcImportarXMLCFOP
'=======================================================================================
' Item é de retorno
'=======================================================================================
TBAbrir!retorno = retorno
'=======================================================================================
' Verifica se existe NCM cadastrada
'=======================================================================================
Set TBAliquota = CreateObject("adodb.recordset")
NCM = Replace(NCM, ".", "")
NCM = Format(NCM, "@@@@.@@.@@")
TBAliquota.Open "Select * from tbl_ClassificacaoFiscal where IDIntClasse = '" & NCM & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
TBAliquota.AddNew
TBAliquota!IDIntClasse = NCM
TBAliquota.Update
End If
TBAbrir!ID_CF = IIf(IsNull(TBAliquota!Idclass), 0, TBAliquota!Idclass)
TBAliquota.Close
'=======================================================================================
' Verifica se existe Unidade cadastrada
'=======================================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from Unidade_Medida where unidade = '" & uCom & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
TBAliquota.AddNew
TBAliquota!Unidade = uCom
TBAliquota!Data = Date
TBAliquota!Responsavel = pubUsuario
TBAliquota!Descricao = "Unidade importada"
TBAliquota.Update
TBAliquota.Close
'ProcCarregaComboUnidade Cmb_un_com, True
'ProcCarregaComboUnidade txtUN, True

End If

'=======================================================================================
TBAbrir!txt_Unid = uCom
TBAbrir!Unidade_com = uCom
'TBAbrir!Familia = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
TBAbrir!dbl_ValorUnitario = vUnCom
TBAbrir!dbl_ValorTotal = Format(vUnCom * TBAbrir!int_Qtd, "###,##0.00")
TBAbrir!int_ICMS = IIf(ICMSpICMS <> "", Replace(ICMSpICMS, ".", ","), "0")
TBAbrir!ICMS_SN = IIf(ICMSpICMS <> "", Replace(ICMSpICMS, ".", ","), "0")
TBAbrir!int_IPI = IIf(IPIpIPI <> "", Replace(IPIpIPI, ".", ","), "0")
TBAbrir!dbl_valoripi = IIf(IPIvIPI <> "", Replace(IPIvIPI, ".", ","), 0)
TBAbrir!txt_CST = vICMSCST
TBAbrir!CST_IPI = CSTIPI
TBAbrir!CST_PIS = CSTPIS
TBAbrir!CST_Cofins = CSTCOFINS
TBAbrir!Valor_frete = IIf(vFrete <> "", Replace(vFrete, ".", ","), 0)
TBAbrir!Valor_desconto = IIf(vDesc <> "", Replace(vDesc, ".", ","), 0)

TBAbrir.Update
'================================================================
'Gravar dados da NFe
'================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_NFe where ID_item = '" & TBAbrir!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
TBAliquota.AddNew
End If
TBAliquota!ID_nota = ID_nota
TBAliquota!Id_Item = TBAbrir!Int_codigo
TBAliquota.Update
TBAliquota.Close

'================================================================
'Gravar dados do icms
'================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where ID_item = '" & TBAbrir!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
TBAliquota.AddNew
End If

TBAliquota!Id_Item = TBAbrir!Int_codigo
TBAliquota!Origem_mercadoria = orig
TBAliquota!Tributacao_ICMS = IIf(CSTICMS <> "", Replace(CSTICMS, ".", ","), 0)
TBAliquota!Valor_BC = IIf(ICMSvBC <> "", Replace(ICMSvBC, ".", ","), 0)
TBAliquota!Valor_ICMS = IIf(ICMSvICMS <> "", Replace(ICMSvICMS, ".", ","), 0)
TBAliquota!ICMS_SN = IIf(P_CredSN <> "", Replace(ICMSvICMS, ".", ","), 0)
TBAliquota!Valor_ICMS_SN = IIf(v_CredICMSSN <> "", Replace(v_CredICMSSN, ".", ","), 0)

TBAliquota!Valor_BC_ST = IIf(ICMSvBCST <> "", Replace(ICMSvBCST, ".", ","), 0)
TBAliquota!Valor_ICMS_ST = IIf(ICMSvICMSST <> "", Replace(ICMSvICMSST, ".", ","), 0)
TBAliquota!Aliquota_imposto_ST = IIf(ICMSpICMSST <> "", Replace(ICMSpICMSST, ".", ","), 0)



TBAliquota.Update
TBAliquota.Close
'================================================================
'Gravar dados do IPI
'================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_IPI where ID_item = '" & TBAbrir!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
TBAliquota.AddNew
End If

TBAliquota!Id_Item = TBAbrir!Int_codigo
TBAliquota!Codigo_situacaoTributaria = IIf(CSTIPI <> "", CSTIPI, "0")
TBAliquota!Valor_BC = vBCIPI
TBAliquota.Update
TBAliquota.Close
'================================================================
'Gravar dados do PIS
'================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_PIS where ID_item = '" & TBAbrir!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
TBAliquota.AddNew
End If
TBAliquota!Id_Item = TBAbrir!Int_codigo
TBAliquota!Codigo_situacaoTributaria = CSTPIS 'Lista.ListItems.Item(Contador).ListSubItems(16).Text
TBAliquota!Valor_BC = vbcPIS 'Format(IIf(Lista.ListItems.Item(Contador).ListSubItems(20).Text <> "", Lista.ListItems.Item(Contador).ListSubItems(20).Text, "0,00"), "###,##0.00")
TBAliquota.Update
TBAliquota.Close
'================================================================
'Gravar dados do COFINS
'================================================================
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_COFINS where ID_item = '" & TBAbrir!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
TBAliquota.AddNew
End If

TBAliquota!Id_Item = TBAbrir!Int_codigo
TBAliquota!Codigo_situacaoTributaria = CSTCOFINS 'Lista.ListItems.Item(Contador).ListSubItems(7).Text
TBAliquota!Valor_BC = vBCCofins 'Format(IIf(Lista.ListItems.Item(Contador).ListSubItems(23).Text <> "", Lista.ListItems.Item(Contador).ListSubItems(23).Text, "0,00"), "###,##0.00")
TBAliquota.Update
'TBAliquota.Close
'TBAbrir.Close
Contador = Contador - 1
lLinha = lLinha + 1
GoTo Inicio
End If

End If
FIM:

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcImportarXMLDuplicatas()
On Error GoTo tratar_erro

'===================================================================================================
' Carregar a fatura da Nfe
'===================================================================================================
    Var1 = "nFat"
    fatnFat = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "vOrig"
    fatvOrig = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "vDesc"
    fatvDesc = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    fatvDesc = Replace(fatvDesc, ".", ",")
    fatvDesc = Format(fatvDesc, "###,##0.00")
  
    Var1 = "vLiq"
    fatvLiq = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    fatvLiq = Replace(fatvLiq, ".", ",")
    fatvLiq = Format(fatvLiq, "###,##0.00")
    
    
'=========================================================================
'Carregar lista de duplicatas
'=========================================================================
'ListaDuplicatas.ListItems.Clear
    
Inicio2:
    If PosicaoBase > 0 Then
    
    Var1 = "nDup"
    nDup = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    
    Var1 = "dVenc"
    dVenc = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
  
    Var1 = "vDup"
    vDup = UCase(ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    vDup = Replace(vDup, ".", ",")
    vDup = Format(vDup, "###,##0.00")
    
If nDup <> "" Then

Set TBReceber = CreateObject("adodb.recordset")
TBReceber.Open "Select * from tbl_Detalhes_Recebimento", Conexao, adOpenKeyset, adLockOptimistic
TBReceber.AddNew
TBReceber!dt_Vencimento = dVenc
TBReceber!dbl_Valor = vDup 'Replace(vDup, ".", ",")

If nDup <> "" Then
TBReceber!txt_Parcela = nDup & "/" & nDup
End If

TBReceber!int_NotaFiscal = nNF
TBReceber!ID_nota = ID_nota
TBReceber.Update
TBReceber.Close
lLinha = lLinha + 1
GoTo Inicio2
End If
End If

     Var1 = "indPag"
    fatindPag = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "tPag"
    fattPag = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "vPag"
    fatvPag = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    If fatvPag <> "" Then
    fatvPag = Replace(fatvPag, ".", ",")
    fatvPag = Format(fatvPag, "###,##0.00")
   End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

''Public Sub ProcAbrirXML()
''On Error GoTo tratar_erro
''    CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml"
''    CommonDialog1.ShowOpen
''
''    strCaminho = CommonDialog1.filename
''    If strCaminho = "" Then Exit Sub
''If USMsgBox("Deseja realmente importar o XML " & strCaminho & "", vbYesNo, "CAPRIND v5.0") = vbYes Then
''    FunImportarXML (strCaminho)
''    PosicaoBase = 1
''Else
''  USMsgBox "Importação cancelada com sucesso!", vbInformation, "CAPRIND v5.0"
''End If
''
''Exit Sub
''tratar_erro:
''    MsgBox ("Descrição do erro : " + Error()), vbCritical
''    Exit Sub
''End Sub

Public Sub ProcImportarXMLTransporte()
On Error GoTo tratar_erro

'==============================================================
' Carrega dados transporte
'==============================================================
    
    Var1 = "CNPJ"
    transpCNPJ = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    transpCNPJ = Format(transpCNPJ, "@@.@@@.@@@/@@@@-@@")
    Var1 = "xNome"
    transpxNome = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
  
    Var1 = "IE"
    transpIE = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
  
    Var1 = "xEnder"
    transpxEnder = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
 
    Var1 = "xMun"
    transpxMun = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "UF"
    transpUF = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "qVol"
    transpqVol = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    transpqVol = Replace(transpqVol, ".", ",")
    transpqVol = Format(transpqVol, "###,##0.00")

    Var1 = "esp"
    transpesp = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "marca"
    transpMarca = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "nVol"
    transpnVol = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "pesoL"
    transppesoL = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    transppesoL = Replace(transppesoL, ".", ",")
    transppesoL = Format(transppesoL, "###,##0.00")

    Var1 = "pesoB"
    transppesoB = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    transppesoB = Replace(transppesoB, ".", ",")
    transppesoB = Format(transppesoB, "###,##0.00")
    
    
     Set TBTransporte = CreateObject("adodb.recordset")
   TBTransporte.Open "Select * from tbl_Dados_Transp where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
   If TBTransporte.EOF = True Then
    TBTransporte.AddNew
    End If
    
    TBTransporte!ID_nota = ID_nota
    TBTransporte!txt_Razao = transpxNome
    TBTransporte!txt_CNPJ = transpCNPJ
    TBTransporte!txt_IE = transpIE
    TBTransporte!txt_UF = transpUF
    
    If transpxMun <> "" Then
    TBTransporte!txt_Municipio = transpxMun
    End If
    
    If transpxEnder <> "" Then
    TBTransporte!txt_Endereco = transpxEnder
    End If
    
    If transpqVol <> "" Then
    TBTransporte!int_Qtd_Transp = IIf(transpqVol < "", transpqVol, "0")
    End If
    
    If transpMarca <> "" Then
    TBTransporte!txt_Marca = transpMarca
    End If
    
    TBTransporte!dbl_Peso_Bruto = IIf(transppesoB <> "", transppesoB, "0")
    TBTransporte!dbl_Peso_Liquido = IIf(transppesoL <> "", transppesoL, "0")
    TBTransporte!Tipo_transp = "F"
    TBTransporte.Update
   'End If
TBTransporte.Close


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Public Sub ProcImportarXMLDadosAdicionais()
On Error GoTo tratar_erro
    
'====================================================================================================
'Carregar dados adicionais
'====================================================================================================
    Var1 = "infCpl"
    infCpl = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'====================================================================================================
'Carregar numero do protocolo de recebimento do SEFAZ
'====================================================================================================
    nProt = ProcImportarXMLCarregacampo("<nProt>", "</nProt>", Len("<nProt>"))
'====================================================================================================
'Carregar status de recebimento do SEFAZ
    xMotivo = ProcImportarXMLCarregacampo("<xMotivo>", "</xMotivo>", Len("<xMotivo>"))
    
 Set TBGravar_NFe = CreateObject("adodb.recordset")
   TBGravar_NFe.Open "Select * from tbl_DadosAdicionais where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
   If TBGravar_NFe.EOF = True Then
   TBGravar_NFe.AddNew
   End If
   
   TBGravar_NFe!ID_nota = ID_nota
   TBGravar_NFe!mem_DadosAdicionais = infCpl
   TBGravar_NFe.Update
   TBGravar_NFe.Close
    
   ' USMsgBox "Importação do XML efetuada com sucesso!", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Public Sub ProcImportarXML()
On Error GoTo tratar_erro

'Define que pode importar o xml
ImportarXML = True

'Verifica se existe nota cadastrada e continua a importacao
ProcImportarXMLDadosNota

'Se não existir a nota cadastrada continua a importacao
If ImportarXML = True Then
    Familia = ""
    LocalArmazenamento = ""
    ProcImportarXMLProdutos
    ProcImportarXMLTotaisNota
    ProcImportarXMLTransporte
    ProcImportarXMLDuplicatas
    ProcImportarXMLDadosAdicionais

    If tpNF = 1 Then
        NFSaida = True
        NFentrada = False
    Else
        NFSaida = False
        NFentrada = True
    End If
    
    If ID_nota <> 0 And ID_empresa <> 0 Then
        If LocalArmazenamento = "" Then
            frmEstoque_Local.Show 1
        End If
        EntrarEstoqueNF
    End If

End If

strPedido = ""
strCaminho = ""
If ImportarXML = True Then
    USMsgBox "XML importado com sucesso!", vbInformation, "CAPRIND v5.0"
    With frmFaturamento_Prod_Serv
        .Strsql_Faturamento = "Select NF.Int_TipoNota, NF.ID, NF.dt_DataEmissao, NF.dt_Saida_Entrada, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, TN.dbl_Valor_Total_Nota, NF.txt_Razao_Nome, NF.txt_CNPJ_CPF, NF.Int_status, NF.Imprimir, NF.ID_empresa, NF.Aplicacao, NF.DtValidacaoOF, NF.DtValidacao, NF.RPS, NF.Modelo, NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF LEFT OUTER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by NF.Int_TipoNota, NF.ID, NF.dt_DataEmissao, NF.dt_Saida_Entrada, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, TN.dbl_Valor_Total_Nota, NF.txt_Razao_Nome, NF.txt_CNPJ_CPF, NF.Int_status, NF.Imprimir, NF.ID_empresa, NF.Aplicacao, NF.DtValidacaoOF, NF.DtValidacao, NF.RPS, NF.Modelo, NF.int_NotaFiscal HAVING NF.Int_status = 1 and  NF.tiponf = 'M1' and NF.Aplicacao = 'T' and NF.ID = " & ID_nota & " and NF.ID_empresa = 1 order by NF.int_NotaFiscal desc"
        .ProcCarregaListaNota (1)
    End With
Else
    USMsgBox "Cancelamento de importação de XML com sucesso!", vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCSOSN101()
On Error GoTo tratar_erro

        Var1 = "pCredSN"
        P_CredSN = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vCredICMSSN"
        v_CredICMSSN = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCSOSN102()
On Error GoTo tratar_erro

'        Var1 = "pCredSN"
'        P_CredSN = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'        Var1 = "vCredICMSSN"
'        v_CredICMSSN = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Public Sub ProcCSOSN400()
On Error GoTo tratar_erro

'        Var1 = "pCredSN"
'        p_CredSN = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'        Var1 = "vCredICMSSN"
'        v_CredICMSSN = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Public Sub ProcCST70()
On Error GoTo tratar_erro

        Var1 = "modBC"
        modBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        Var1 = "pRedBC"
        p_RedBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        Var1 = "vBC"
        vbcIMCS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pICMS"
        p_ICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMS"
        v_ICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "modBCST"
        modBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pMVAST"
        p_MVAST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        Var1 = "pRedBCST"
        p_RedBCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vBCST"
        v_BCST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pICMSST"
        p_ICMSST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMSST"
        v_ICMSST = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCST90()
On Error GoTo tratar_erro

        Var1 = "modBC"
        modBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
                
        Var1 = "vBC"
        vbcIMCS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
        Var1 = "pRedBC"
        p_RedBC = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "pICMS"
        p_ICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

        Var1 = "vICMS"
        v_ICMS = ProcImportarXMLCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    orig = orig & CST


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

