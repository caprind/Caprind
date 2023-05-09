VERSION 5.00
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_inutilizarNFe 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Inutilizar numeração de nota"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_inutilizarNFe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Top             =   3855
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   714
   End
   Begin VB.TextBox txtRetorno 
      ForeColor       =   &H00000080&
      Height          =   765
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2940
      Width           =   4185
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_inutilizarNFe.frx":000C
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados para inutilização de número de nota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2085
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   4245
      Begin VB.TextBox txtAno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Ano da emissão da(s) nota(s)."
         Top             =   1680
         Width           =   975
      End
      Begin ControlesUteis.txtA txtXmotivo 
         Height          =   1215
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   2143
         Text            =   ""
         Caption         =   "Motivo da inutilização"
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
      Begin VB.TextBox txtnNFIni 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Série da nota fiscal"
         Top             =   1680
         Width           =   1245
      End
      Begin VB.TextBox txtserie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Numero da nota inicial da sequencia de inutilização"
         Top             =   1680
         Width           =   375
      End
      Begin DrawSuite2022.USButton btn_Inutilizar 
         Height          =   585
         Left            =   2880
         TabIndex        =   10
         ToolTipText     =   "Inutilizar numero de nota"
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1032
         DibPicture      =   "frmFaturamento_inutilizarNFe.frx":7A80
         Caption         =   "Inutilizar"
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
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero nota "
         Height          =   195
         Left            =   1890
         TabIndex        =   8
         Top             =   1500
         Width           =   555
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Série"
         Height          =   195
         Left            =   1170
         TabIndex        =   7
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ano emissão"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   1500
         Width           =   1005
      End
   End
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   344
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
      ForeColor       =   128
      NoHTMLCaption   =   "Retorno do SEFAZ"
   End
End
Attribute VB_Name = "frmFaturamento_inutilizarNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Inutilizar_Click()
On Error GoTo tratar_erro
Dim UFEmit As String
Dim CNPJEmit As String
Dim xAno As String
Dim xMotivo As String
Dim nNFIni As String
Dim nNFFim As String

ID_nota = frmFaturamento_Prod_Serv_NFe_NS.txtID_nota
NomeArquivo = frmFaturamento_Prod_Serv_NFe_NS.txtNota
nfDocumento = "INU" & NomeArquivo

If txtXmotivo.Text = "" Then
 USMsgBox "Digite o motivo da inutilização", vbInformation, "CAPRIND v5.0"
 Exit Sub
End If


If USMsgBox("Deseja realmente inutilizar o numero " & txtnNFIni.Text & " no SEFAZ?", vbYesNo, "CAPRIND v5.0") = vbYes Then
  Set TBAbrir = CreateObject("adodb.recordset")
  TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
     If TBAbrir.EOF = False Then
     procVerificacodigoUF (TBAbrir!UF)
     UFEmit = CodUF
     CNPJEmit = ReturnNumbersOnly(TBAbrir!CNPJ)
     tpAmb = IIf(IsNull(TBAbrir!tpAmb) = True, "2", TBAbrir!tpAmb)
     End If
     TBAbrir.Close
     
     xAno = Right(txtAno.Text, 2)
     xMotivo = txtXmotivo.Text
     nNFIni = Int(txtnNFIni.Text)
     nNFFim = Int(txtnNFIni.Text)
     
     If UFEmit <> "" And tpAmb <> "" And xAno <> "" And CNPJEmit <> "" And nNFIni <> "" And nNFFim <> "" And xMotivo <> "" Then
     If NFCe = False Then
      retornoinut = NFeAPI.inutilizar(UFEmit, tpAmb, xAno, CNPJEmit, txtSerie.Text, nNFIni, nNFFim, xMotivo)
      Else
      retornoinut = NFCe_inutilizar(UFEmit, tpAmb, xAno, CNPJEmit, txtSerie.Text, nNFIni, nNFFim, xMotivo)
      End If
      
      txtRetorno.Text = retornoinut
      status = LerDadosJSON(txtRetorno.Text, "status", "", "")
        If status = "200" Or status = "102" Then

'Debug.print txtRetorno.Text


            
            If NFCe = False Then
                cStat = LerDadosJSON(txtRetorno.Text, "retornoInutNFe", "cStat", "")
                xMotivo = LerDadosJSON(txtRetorno.Text, "retornoInutNFe", "xMotivo", "")
                nProt = LerDadosJSON(txtRetorno.Text, "retornoInutNFe", "nProt", "")
                xmlInut = LerDadosJSON(txtRetorno.Text, "retornoInutNFe", "xmlInut", "")
                Chave = LerDadosJSON(txtRetorno.Text, "retornoInutNFe", "chave", "")
            Else
                cStat = LerDadosJSON(txtRetorno.Text, "retInutNFe", "cStat", "")
                xMotivo = LerDadosJSON(txtRetorno.Text, "retInutNFe", "xMotivo", "")
                nProt = LerDadosJSON(txtRetorno.Text, "retInutNFe", "nProt", "")
                xmlInut = LerDadosJSON(txtRetorno.Text, "retInutNFe", "xml", "")
                Chave = LerDadosJSON(txtRetorno.Text, "retInutNFe", "idInut", "")
            End If
            

           frmFaturamento_Prod_Serv_NFe_NS.txtchNFe = Chave
           frmFaturamento_Prod_Serv_NFe_NS.txt_nProt = nProt
           frmFaturamento_Prod_Serv_NFe_NS.txtcStat = cStat
           Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & Chave & "' , nProt = '" & nProt & "'  where id_nota = " & ID_nota
           Mensagem = "Motivo: " & vbCrLf & xMotivo & vbCrLf
           Mensagem = Mensagem & "Chave: " & vbCrLf & Chave & vbCrLf
           Mensagem = Mensagem & "nProt: " & vbCrLf & nProt
           txtRetorno.Text = Mensagem
'=======================================================================================
' Excluir contas geradas pela nota
'=======================================================================================
            Conexao.Execute "DELETE from CC from CC_realizado CC INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = CC.ID_financeiro Where CR.ID_Nota = " & ID_nota & " and CC.Operacao = 'Crédito'"
            Conexao.Execute "DELETE from FF from Familia_financeiro FF INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = FF.IDconta Where CR.ID_Nota = " & ID_nota & " and FF.Tipoconta = 'R' and (CR.Proposta IS NULL or CR.Proposta = N'')"
            Conexao.Execute "DELETE from FC from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON CR.IDFluxo = FC.IDFluxo Where CR.ID_Nota = " & ID_nota & " and (CR.Proposta IS NULL or CR.Proposta = N'')"
            Conexao.Execute "DELETE from tbl_contas_receber where ID_Nota = " & ID_nota & " and (Proposta IS NULL or Proposta = N'')"
            
            With frmFaturamento_Prod_Serv_NFe_NS
            .ProcPuxaDados
            .ProcCarregaListaNota (1)
            End With
            With frmFaturamento_Prod_Serv
            .ProcCarregaListaNota (1)
            End With
            
            USMsgBox "Inutilização de numero efetuado com sucesso!", vbInformation, "CAPRIND v5.0"
        Else
           USMsgBox "Não foi possivel inutilizar esse numero de NFe, veja arquivo log para mais detalhes", vbCritical, "CAPRIND v5.0"
        End If

     Else
      USMsgBox "Verifique o preenchimento dos campos necessários", vbCritical, "CAPRIND v5.0"
     End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim UFEmit As String
Dim CNPJEmit As String

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
UFEmit = TBAbrir!UF
CNPJEmit = ReturnNumbersOnly(TBAbrir!CNPJ)
tpAmb = IIf(IsNull(TBAbrir!tpAmb) = True, "2", TBAbrir!tpAmb)
End If
TBAbrir.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where int_NotaFiscal = '" & frmFaturamento_Prod_Serv_NFe_NS.txtNota.Text & "' and aplicacao = 'P' and serie  = '" & frmFaturamento_Prod_Serv_NFe_NS.txtSerie.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
txtAno.Text = Year(TBAbrir!dt_DataEmissao)
txtnNFIni.Text = TBAbrir!int_NotaFiscal
txtSerie.Text = TBAbrir!Serie
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
