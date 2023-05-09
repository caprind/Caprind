VERSION 5.00
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_RetornoSEFAZ 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Cancelar NFe SEFAZ | CAPRIND V5.0"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_RetornoSEFAZ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   4110
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USLabel USLabel3 
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   870
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   344
      Caption         =   "Motivo do cancelamento da Nota fiscal eletrônica"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      NoHTMLCaption   =   "Motivo do cancelamento da Nota fiscal eletrônica"
   End
   Begin ControlesUteis.txtA txtMotivo 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Motivo do cancelamento da Nota fiscal."
      Top             =   930
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   1296
      Text            =   ""
      Caption         =   ""
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
   End
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1770
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   344
      BackColor       =   14737632
      Caption         =   "Retorno do SEFAZ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      NoHTMLCaption   =   "Retorno do SEFAZ"
   End
   Begin DrawSuite2022.USButton cmdCancelar 
      Height          =   585
      Left            =   3420
      TabIndex        =   3
      Top             =   3450
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      DibPicture      =   "frmFaturamento_RetornoSEFAZ.frx":000C
      Caption         =   "Cancelar NFe"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      PicAlign        =   7
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin VB.TextBox txtRetorno 
      ForeColor       =   &H00000080&
      Height          =   1425
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1950
      Width           =   4365
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   688
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DrawSuite2022.USButton btnCriarXML 
      Height          =   435
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Criar XML de cancelamento"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      Caption         =   "Criar XML"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   0
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   8421504
      GradientColor1  =   0
      GradientColor2  =   0
      GradientColor3  =   0
      GradientColor4  =   0
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   8421504
      GradientColorOver2=   8421504
      GradientColorOver3=   8421504
      GradientColorOver4=   8421504
      GradientColorDown1=   4210752
      GradientColorDown2=   4210752
      GradientColorDown3=   4210752
      GradientColorDown4=   4210752
      ShowFocusRect   =   0   'False
      Theme           =   6
   End
   Begin DrawSuite2022.USButton btnAssinarXML 
      Height          =   435
      Left            =   1320
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      Caption         =   "Assinar XML"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      ShowFocusRect   =   0   'False
      Theme           =   3
   End
   Begin DrawSuite2022.USButton btnEnviarXML 
      Height          =   435
      Left            =   2400
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   767
      Caption         =   "Enviar XML"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   1154291
      BorderColorDisabled=   13160660
      BorderColorDown =   16576
      BorderColorOver =   8438015
      GradientColor1  =   1154291
      GradientColor2  =   1154291
      GradientColor3  =   1154291
      GradientColor4  =   1154291
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   8438015
      GradientColorOver2=   8438015
      GradientColorOver3=   8438015
      GradientColorOver4=   8438015
      GradientColorDown1=   16576
      GradientColorDown2=   16576
      GradientColorDown3=   16576
      GradientColorDown4=   16576
      ShowFocusRect   =   0   'False
      Theme           =   5
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o motivo do cancelamento com no mínimo 15 caracteres"
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   600
      Width           =   4545
   End
End
Attribute VB_Name = "frmFaturamento_RetornoSEFAZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAssinarXML_Click()
On Error GoTo tratar_erro
Dim objDom As DOMDocument50
Dim objEnvioEvento As IXMLDOMElement
Dim objEvento As IXMLDOMElement
Dim objEveInf As IXMLDOMElement
Dim objEvedet As IXMLDOMElement
Dim XML As String
    
Set objDom = New DOMDocument50
    
NomeArquivo = frmFaturamento_Prod_Serv_NFe_NS.txtNota
If USMsgBox("Deseja realmente assinar o xml de cancelamento da NF" & frmFaturamento_Prod_Serv_NFe_NS.txtNota.Text & " ?", vbYesNo, "CAPRIND v5.0") = vbNo Then
Exit Sub
End If

XML = txtRetorno.Text


 Dim AssinaturaXML2 As New AssinaturaXML2.Principal
 XML = AssinaturaXML2.assinarXML(XML, "infEvento", chCNPJ)
 
 objDom.loadXML (XML)
 objDom.Save (DiretorioEnvio & "NF" & frmFaturamento_Prod_Serv_NFe_NS.txtNota & "CA.xml")
 'frmFaturamento_Prod_Serv_NFe_NS.txtResultado.Text = objDom.xml
 frmFaturamento_RetornoSEFAZ.txtRetorno = objDom.XML
 'Debug.print XML
 USMsgBox "Xml da NF" & frmFaturamento_Prod_Serv_NFe_NS.txtNota.Text & " assinado com sucesso!", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnCriarXML_Click()
On Error GoTo tratar_erro

Dim objDom As DOMDocument50
Dim objEnvioEvento As IXMLDOMElement
Dim objEvento As IXMLDOMElement
Dim objEveInf As IXMLDOMElement
Dim objEvedet As IXMLDOMElement
Dim chNFe As String
Dim dtCan As String
Dim nProt As String
Dim TextoCancelamento As String
Dim XML As String
    
Set objDom = New DOMDocument50
    
NomeArquivo = frmFaturamento_Prod_Serv_NFe_NS.txtNota
If USMsgBox("Deseja realmente gerar xml de cancelamento da NF" & frmFaturamento_Prod_Serv_NFe_NS.txtNota.Text & "", vbYesNo, "CAPRIND v5.0") = vbNo Then
Exit Sub
End If

If txtMotivo.Text = "" Then
 USMsgBox "Digite o motivo do cancelamento", vbInformation, "CAPRIND v5.0"
 Exit Sub
End If

TextoCancelamento = txtMotivo.Text
If Len(TextoCancelamento) < 15 Then
 USMsgBox ("O motivo deve possuir um valor minimo de 15 caracteres!"), vbExclamation, "CAPRIND v5.0"
Exit Sub
End If

With frmFaturamento_Prod_Serv_NFe_NS
  chNFe = .txtchNFe
  dtCan = Format(Date, "yyyy-mm-dd") & "T" & Left(Time, 8) & FunVerifFusoHorario(True)
  nProt = Trim(.txt_nProt)
End With


        'Monta o XML
        XML = "<?xml version=""1.0"" encoding=""utf-8""?>"
        XML = XML & "<evento versao=""1.00"" xmlns=""http://www.portalfiscal.inf.br/nfe"">"
        XML = XML & "<infEvento Id=""ID110111" & chNFe & "01"">"
        XML = XML & "<cOrgao>" & Left(chNFe, 2) & "</cOrgao>"
        XML = XML & "<tpAmb>" & tpAmb & "</tpAmb>"
        XML = XML & "<CNPJ>" & CnpjNF & "</CNPJ>"
        XML = XML & "<chNFe>" & chNFe & "</chNFe>"
        XML = XML & "<dhEvento>" & dtCan & "</dhEvento>"
        XML = XML & "<tpEvento>110111</tpEvento>"
        XML = XML & "<nSeqEvento>1</nSeqEvento>"
        XML = XML & "<verEvento>1.00</verEvento>"
        XML = XML & "<detEvento versao=""1.00"">"
        XML = XML & "<descEvento>Cancelamento</descEvento>"
        XML = XML & "<nProt>" & nProt & "</nProt>"
        XML = XML & "<xJust>" & TextoCancelamento & "</xJust>"
        XML = XML & "</detEvento>"
        XML = XML & "</infEvento>"
        XML = XML & "</evento>"
        
        objDom.loadXML (XML)
        objDom.Save (DiretorioEnvio & "NF" & frmFaturamento_Prod_Serv_NFe_NS.txtNota & "CA.xml")
        'frmFaturamento_Prod_Serv_NFe_NS.txtResultado.Text = objDom.xml
        frmFaturamento_RetornoSEFAZ.txtRetorno = objDom.XML
        'Debug.print XML
        USMsgBox "Xml de cancelamento da NF" & frmFaturamento_Prod_Serv_NFe_NS.txtNota.Text & " gerado com sucesso!", vbInformation, "CAPRIND v5.0"
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnEnviarXML_Click()
On Error GoTo tratar_erro

Dim XML As String
Dim url As String
Dim resposta As String
Dim status As String
Dim respostaDownload As String
Dim chNFe As String
Dim dtCan As String
Dim nProt As String
Dim TextoCancelamento As String
NomeArquivo = frmFaturamento_Prod_Serv_NFe_NS.txtNota
nfDocumento = "CA" & NomeArquivo

If txtMotivo.Text = "" Then
 USMsgBox "Digite o motivo do cancelamento", vbInformation, "CAPRIND v5.0"
 Exit Sub
End If

TextoCancelamento = txtMotivo.Text
If Len(TextoCancelamento) < 15 Then
 USMsgBox ("O motivo deve possuir um valor minimo de 15 caracteres!"), vbExclamation, "CAPRIND v5.0"
Exit Sub
End If

With frmFaturamento_Prod_Serv_NFe_NS
  chNFe = .txtchNFe
  dtCan = Format(Date, "yyyy-mm-dd") & "T" & Left(Time, 8) & FunVerifFusoHorario(True)
  nProt = Trim(.txt_nProt)
  
  
End With


If USMsgBox("Deseja realmente enviar o xml de cancelamento da NF" & frmFaturamento_Prod_Serv_NFe_NS.txtNota.Text & "", vbYesNo, "CAPRIND v5.0") = vbNo Then
Exit Sub
End If
    
    'Informa a URL
 url = "https://nfe.ns.eti.br/nfe/cancel"
 XML = txtRetorno
 
 
 
 'Envia o xml para a URL
 resposta = enviaConteudoParaAPI(XML, url, "xml")
 
     status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
        'Aguarda 2000 (2 segundos) milisegundos antes de consultar o status de processamento
        Sleep (2000)
        
        respostaDownload = downloadEventoNFeAndSave(chNFe, tpAmb, "Xp", "CANC", "1", DiretorioXMLDanfe, True)
        
        status = LerDadosJSON(respostaDownload, "status", "", "")
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        frmFaturamento_RetornoSEFAZ.txtRetorno = motivo
        
        'Se der problema no download
        If (status <> "200") Then
            'Retorna a resposta do download
            resposta = respostaDownload
        End If
    '========================================
    ' AQUI ACERTA TABELA DE DADOS COM O STATUS DA NFE
    '========================================
    StatusNFe = "101"
    Var1 = 2
    ID_nota = frmFaturamento_Prod_Serv_NFe_NS.txtID_nota.Text
    
'=======================================================================
    ProcExcluirArquivosRemessa frmFaturamento_Prod_Serv_NFe_NS.txtID_nota.Text
    ProcExcluirContas frmFaturamento_Prod_Serv_NFe_NS.txtID_nota.Text, frmFaturamento_Prod_Serv.opt_Saida, frmFaturamento_Prod_Serv.txttipocliente
    Conexao.Execute "DELETE from ECEV from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = ECEV.ID_faturamento where NFP.ID_nota = " & IDConta
'=======================================================================
    Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = '" & StatusNFe & "' where id_nota = " & ID_nota
    Conexao.Execute "Update tbl_dados_nota_fiscal Set int_status = '" & Var1 & "' where id = " & ID_nota
    frmFaturamento_Prod_Serv_NFe_NS.procCancelarTabelas
    ProcExcluirSaida_NFe ID_nota
    '========================================
    Else
    frmFaturamento_RetornoSEFAZ.txtRetorno = resposta

    End If
    
   ' cancelarNFe = resposta

 
 txtRetorno = resposta

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo tratar_erro

Dim chNFe As String
Dim dtCan As String
Dim nProt As String
Dim TextoCancelamento As String
NomeArquivo = frmFaturamento_Prod_Serv_NFe_NS.txtNota
nfDocumento = "CC" & NomeArquivo

If txtMotivo.Text = "" Then
 USMsgBox "Digite o motivo do cancelamento", vbInformation, "CAPRIND v5.0"
 Exit Sub
End If

TextoCancelamento = txtMotivo.Text
If Len(TextoCancelamento) < 15 Then
 USMsgBox ("O motivo deve possuir um valor minimo de 15 caracteres!"), vbExclamation, "CAPRIND v5.0"
Exit Sub
End If

With frmFaturamento_Prod_Serv_NFe_NS

If .txtchNFe.Text = "" Then
    Exit Sub
End If

  chNFe = .txtchNFe
  dtCan = Format(Date, "yyyy-mm-dd") & "T" & Left(Time, 8) & FunVerifFusoHorario(True)
  nProt = Trim(.txt_nProt)
  
If NFCe = False Then
    frmFaturamento_Prod_Serv_NFe_NS.ProcCriarPastaXML
    Cancelamento = NFeAPI.cancelarNFe(chNFe, tpAmb, dtCan, nProt, TextoCancelamento, "Xp", DiretorioXML, True)
 Else
    frmFaturamento_Prod_Serv_NFe_NS.ProcCriarPastaXML
    frmFaturamento_Prod_Serv_NFe_NS.ProcCriarPastaDanfe
    Cancelamento = NFeAPI.cancelarNFe(chNFe, tpAmb, dtCan, nProt, TextoCancelamento, "Xp", DiretorioDanfe, True)
 End If
 
  If Var1 = "0" Then
    Unload Me
  Exit Sub
  End If
  
  .ProcPuxaDados
  .ProcCarregaListaNota (1)
End With

With frmFaturamento_Prod_Serv
  .ProcCarregaListaNota (1)
End With

Unload Me


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
DiretorioEnvio = TBAbrir!Caminho_Nfe
DiretorioRetorno = TBAbrir!Caminho_RetornoNfe
DiretorioXMLDanfe = TBAbrir!Caminho_XMLDanfe
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

