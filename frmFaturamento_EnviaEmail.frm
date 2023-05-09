VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmFaturamento_EnviaEmail 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Enviar arquivo DANFE e XML por email para o cliente"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   Icon            =   "frmFaturamento_EnviaEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   5655
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USOptionButton opt1 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   4380
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   503
      BackColor       =   16382457
      Caption         =   "Enviar para destinatário"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   0   'False
   End
   Begin FlexCell.Grid GridEmail 
      Height          =   2835
      Left            =   120
      TabIndex        =   1
      Top             =   1410
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5001
      Appearance      =   0
      BackColorBkg    =   16777215
      BorderColor     =   8421504
      CellBorderColor =   8421504
      Cols            =   3
      DefaultFontSize =   8.25
      DisplayFocusRect=   0   'False
      FixedRowColStyle=   0
      GridColor       =   12632256
      Rows            =   1
      ScrollBarStyle  =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   630
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   180
      Top             =   5280
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   5250
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SearchText      =   "Enviando email..."
      ShowText        =   0   'False
      Value           =   0
   End
   Begin DrawSuite2022.USButton cmdEnviarDocumentos 
      Height          =   855
      Left            =   5580
      TabIndex        =   2
      ToolTipText     =   "Enviar Danfe e XML por email"
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      DibPicture      =   "frmFaturamento_EnviaEmail.frx":000C
      Caption         =   "Enviar DFE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      GradientColor1  =   128
      GradientColor2  =   192
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
      PicSize         =   5
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   4
      ToolTipTitle    =   "Caprind v5.0"
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   688
      DibPicture      =   "frmFaturamento_EnviaEmail.frx":36458
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DrawSuite2022.USOptionButton opt2 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   4620
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      BackColor       =   16382457
      Caption         =   "Enviar para destinatário e transportadora"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USOptionButton opt3 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   4860
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      BackColor       =   16382457
      Caption         =   "Enviar para destinatário, transportadora e enviar-me uma cópia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      ShowFocusRect   =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Abaixo lista de email para receber os documentos fiscais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   3
      Top             =   690
      Width           =   6675
   End
End
Attribute VB_Name = "frmFaturamento_EnviaEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaEmailUsuario()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Usuario,Email from Usuarios where Usuario = '" & pubUsuario & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
       If TBAbrir!Email = "" Or IsNull(TBAbrir!Email) = True Then
          USMsgBox "Atenção " & pubUsuario & " você não tem email cadastrado para receber uma cópia!!!", vbCritical, "CAPRIND v5.0"
          Exit Sub
       End If
        EmailUsuario = IIf(IsNull(TBAbrir!Email) = False, TBAbrir!Email, "")
        nomeUsuario = TBAbrir!Usuario
        nomeUsuario = LCase(nomeUsuario)
        nomeUsuario = UCase(Mid$(nomeUsuario, 1, 1)) & Right$(nomeUsuario, Len(nomeUsuario) - 1)
        GridEmail.AddItem nomeUsuario & vbTab & TBAbrir!Email
    End If
    TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdEnviarDocumentos_Click()
On Error GoTo tratar_erro

'======================================================
'Se for homologação envia só para usuário
'======================================================
If tpAmb = "2" Then
   EmailEnvioNFe = "vendas@caprind.com.br"
End If



Timer1.Enabled = True
Timer2.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcAjustaGridEmail
Opt1.Value = True
Opt1_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcAjustaGridEmail()
On Error GoTo tratar_erro

With GridEmail
    .AllowUserPaste = cellTextOnly
    .AllowUserResizing = False
    .ExtendLastCol = True
    .BoldFixedCell = False
    .DisplayDateTimeMask = True
    .DisplayFocusRect = False
    .SelectionMode = cellSelectionByRow

    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).Text = "Contato"
    .Cell(0, 2).Text = "Email"
        
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(2).CellType = cellHyperLink
    .Column(2).Alignment = cellCenterCenter
    
    .Column(0).Width = 10
    .Column(1).Width = 100
    .Column(2).Width = 180
'    .Range(2, 1, .rows - 1, 1).ForeColor = RGB(0, 0, 128)
  End With
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEmailTransportadora()
On Error GoTo tratar_erro

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select NF.*, T.*, E.Simples, E.Simples1, E.Cultural, E.CNPJ, E.CNAE, E.Razao, E.Empresa, E.IM, E.ie, E.Tipo_endereco, E.Endereco, E.Numero as numeroEmpresa, E.Complemento, E.Tipo_bairro, E.Bairro, E.Cidade, E.UF, E.CEP, E.Telefone, E.Email, NFE.Consumidor_final, NFE.Presenca_comprador, NFE.Forma_emissao, NFE.Finalidade_emissao, NFE.Enviar_Email, NFE.Forma_pagamento, NFE.FormaPagto, NFE.DA_entrega, NFE.DA_cobranca, NFE.ID_entrega, NFE.ID_Cobranca from (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota T ON NF.ID = T.ID_nota) INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE ON NFE.ID_Nota = NF.ID INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo WHERE NF.ID = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then

'Verifica se tem transportadora na NF para consultar o e-mail
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select IdIntTransp, txt_Razao from tbl_Dados_Transp where ID_nota = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then
        'Verifica se a transportadora é o Cliente
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select Email, IDCliente from Clientes where IDcliente = " & TBFIltro!IdIntTransp & " and NomeRazao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            EmailTransportadora = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
            GridEmail.AddItem "Transportadora" & vbTab & TBClientes!Email
        'Busca contatos do cliente
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select nomecontato, Email from Clientes_Contatos where IDcliente = " & TBClientes!IDCliente & "   and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                        NomeContato = TBFI!NomeContato
                        NomeContato = LCase(NomeContato)
                        NomeContato = UCase(Mid$(NomeContato, 1, 1)) & Right$(NomeContato, Len(NomeContato) - 1)
                        GridEmail.AddItem NomeContato & vbTab & TBFI!Email
                        EmailTransportadora = EmailTransportadora & ", " & TBFI!Email
                        
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
            TBClientes.Close
    Else
        'Verifica se a transportadora é tipo Fornecedor
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select Email from Compras_fornecedores where IDcliente = " & TBFIltro!IdIntTransp & " and Nome_Razao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                EmailTransportadora = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
                GridEmail.AddItem "Transportadora" & vbTab & TBClientes!Email
        'Busca contatos do fornecedor
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select nomecontato, Email from Contatos_fornecedor where IdFornecedor = " & TBClientes!IDCliente & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockReadOnly
                If TBFI.EOF = False Then
                    Do While TBFI.EOF = False
                        If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                            GridEmail.AddItem TBFI!NomeContato & vbTab & TBFI!Email
                            EmailTransportadora = EmailTransportadora & ", " & TBFI!Email
                        End If
                        TBFI.MoveNext
                    Loop
                End If
                TBFI.Close
            End If
        End If
    End If
End If

TBproducao.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCarregaEmailCliente()
On Error GoTo tratar_erro
GridEmail.rows = 1
GridEmail.Refresh

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select NF.*, T.*, E.Simples, E.Simples1, E.Cultural, E.CNPJ, E.CNAE, E.Razao, E.Empresa, E.IM, E.ie, E.Tipo_endereco, E.Endereco, E.Numero as numeroEmpresa, E.Complemento, E.Tipo_bairro, E.Bairro, E.Cidade, E.UF, E.CEP, E.Telefone, E.Email, NFE.Consumidor_final, NFE.Presenca_comprador, NFE.Forma_emissao, NFE.Finalidade_emissao, NFE.Enviar_Email, NFE.Forma_pagamento, NFE.FormaPagto, NFE.DA_entrega, NFE.DA_cobranca, NFE.ID_entrega, NFE.ID_Cobranca from (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota T ON NF.ID = T.ID_nota) INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE ON NFE.ID_Nota = NF.ID INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo WHERE NF.ID = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then

'=======================================
'Se o destinatário for Cliente
'=======================================
Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select Email from Clientes where IDcliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockReadOnly
    If TBClientes.EOF = False Then
        EmailCliente = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
        GridEmail.AddItem "Destinatário" & vbTab & TBClientes!Email
'============================================
'Buscar email dos contatos do Cliente
'============================================
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select NomeContato, Email from Clientes_Contatos where IDcliente = " & TBproducao!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                        NomeContato = TBFI!NomeContato
                        NomeContato = LCase(NomeContato)
                        NomeContato = UCase(Mid$(NomeContato, 1, 1)) & Right$(NomeContato, Len(NomeContato) - 1)
                        GridEmail.AddItem NomeContato & vbTab & TBFI!Email
                        EmailCliente = EmailCliente & ", " & TBFI!Email
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
        End If
    Else
'=======================================
'Se o destinatário for Fornecedor
'=======================================
Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select Email, Pais, Codigo_pais from Compras_fornecedores where IDcliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            EmailCliente = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
'============================================
'Buscar email dos contatos do Fornecedor
'============================================
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBproducao!Id_Int_Cliente & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                       EmailCliente = EmailCliente & "," & TBFI!Email
                        GridEmail.AddItem TBFI!NomeContato & vbTab & TBFI!Email
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
        End If
End If
TBClientes.Close

TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt1_Click()
On Error GoTo tratar_erro
GridEmail.rows = 6
GridEmail.Refresh

EmailEnvioNFe = ""
EmailCliente = ""

If Opt1.Value = True Then
ProcCarregaEmailCliente
End If

If EmailCliente = "" Then
    USMsgBox "Atenção !" & vbCrLf & vbCrLf & "Email de envio de documentos fiscais não encontrado no cadastro do cliente." & vbCrLf & "Favor verificar.", vbCritical, "CAPRIND  v5.0"
    Exit Sub
End If

EmailEnvioNFe = EmailCliente
'Debug.print Email
GridEmail.Range(1, 1, GridEmail.rows - 1, 1).ForeColor = RGB(0, 0, 128)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt2_Click()
On Error GoTo tratar_erro
GridEmail.rows = 1
GridEmail.Refresh

EmailEnvioNFe = ""
EmailCliente = ""
EmailTransportadora = ""

If Opt2.Value = True Then
ProcCarregaEmailCliente
ProcCarregaEmailTransportadora
End If

EmailEnvioNFe = EmailCliente & ", " & EmailTransportadora
GridEmail.Range(1, 1, GridEmail.rows - 1, 1).ForeColor = RGB(0, 0, 128)

'Debug.print Email

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub opt3_Click()
On Error GoTo tratar_erro
GridEmail.rows = 1
GridEmail.Refresh

EmailEnvioNFe = ""
EmailCliente = ""
EmailTransportadora = ""
EmailUsuario = ""

If opt3.Value = True Then
    ProcCarregaEmailCliente
    ProcCarregaEmailTransportadora
    ProcCarregaEmailUsuario
End If

EmailEnvioNFe = EmailCliente & ", " & EmailTransportadora & ", " & EmailUsuario
GridEmail.Range(1, 1, GridEmail.rows - 1, 1).ForeColor = RGB(0, 0, 128)

'Debug.print Email

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

If situacao = 0 Then
    PBLista.Value = PBLista.Value + 1
    situacao = 1
    Exit Sub
End If
If situacao = 1 Then
    situacao = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Timer2_Timer()
On Error GoTo tratar_erro
Dim testeEmail As String

'EmailEnvioNFe = EmailEnvioNFe & ", caprind@caprind.com.br"

If NFCe = False Then
testeEmail = enviarEmail(frmFaturamento_Prod_Serv_NFe_NS.txtchNFe.Text, "true", "true", EmailEnvioNFe)
Else
testeEmail = NFCe_enviarEmail(frmFaturamento_Prod_Serv_NFe_NS.txtchNFe.Text, "true", EmailEnvioNFe)
End If

'Debug.print testeEmail
xMotivo = LerDadosJSON(testeEmail, "motivo", "", "")

If EmailEnviado = True Then
Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Enviar_email = '2' where id_nota = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota
Conexao.Execute "Update tbl_dados_nota_fiscal Set Imprimir = '1' where id = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota
USMsgBox "A DANFE e o XML foram enviados para o(s) email(s) " & EmailEnvioNFe & " com sucesso!", vbInformation, "CAPRIND v5.0"
End If


Timer1.Enabled = False
Timer2.Enabled = False
PBLista.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
