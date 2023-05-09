VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frmFaturamentoEndereco 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Endereços"
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin DrawSuite2014.USButton btnCarregar 
      Height          =   795
      Left            =   5340
      TabIndex        =   6
      Top             =   2430
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1402
      DibPicture      =   "frmFaturamentoEndereco.frx":0000
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Carregar endereços na nota fiscal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      PicAlign        =   7
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
   Begin VB.Frame Frame2 
      Caption         =   "Endereço de entrega"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   270
      TabIndex        =   3
      Top             =   1500
      Width           =   6615
      Begin VB.ComboBox Cmb_cobranca 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmFaturamentoEndereco.frx":339B
         Left            =   120
         List            =   "frmFaturamentoEndereco.frx":339D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Endereço de cobrança"
         Top             =   300
         Width           =   6315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Endereço de entrega"
      Height          =   825
      Left            =   270
      TabIndex        =   2
      Top             =   570
      Width           =   6615
      Begin VB.ComboBox cmbEntrega 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmFaturamentoEndereco.frx":339F
         Left            =   120
         List            =   "frmFaturamentoEndereco.frx":33A1
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Endereço de entrega."
         Top             =   300
         Width           =   6315
      End
   End
   Begin DrawSuite2014.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   3285
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   714
   End
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   714
      DibPicture      =   "frmFaturamentoEndereco.frx":33A3
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
      Icon            =   "frmFaturamentoEndereco.frx":673E
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmFaturamentoEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo tratar_erro

With cmbEntrega
.Clear
'Pega ID do cliente da nota fiscal
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select ID, Id_Int_Cliente, txt_Razao_Nome, Modelo from tbl_Dados_Nota_Fiscal where id = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then

    If TBFIltro!Modelo = "65" Then
        NFCe = True
    Else
        NFCe = False
    End If

    'Verifica se é cliente ou fornecedor
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Clientes where IDCliente = " & TBFIltro!Id_Int_Cliente & " and NomeRazao = " & TBFIltro!txt_Razao_Nome, Conexao, adOpenKeyset, adLockReadOnly
        'Define como cliente ou fornecedor
        If TBFI.EOF = False Then Tipo = "C" Else Tipo = "F"
        TBFI.Close

        Permitido = True
        'Busca da proposta\Pedido o endereço de entrega
        Set TBVendas = CreateObject("adodb.recordset")
        'Se for cliente
        If Tipo = "C" Then
            TextoID = ""
            TBVendas.Open "Select VC.* from (vendas_comercial VC INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao) INNER JOIN tbl_proposta_nota PN ON PN.proposta = VP.Ncotacao and PN.Revisao = VP.Revisao where PN.ID_nota = " & TBFIltro!ID & " and VC.ID_entrega IS NOT NULL and VC.ID_entrega <> 0 order by VC.ID_entrega", Conexao, adOpenKeyset, adLockReadOnly
            If TBVendas.EOF = False Then
                Permitido = False
                Do While TBVendas.EOF = False
                    If TextoID <> TBVendas!ID_entrega Then
                        .AddItem TBVendas!Local_entrega
                        .ItemData(cmbEntrega.NewIndex) = TBVendas!ID_entrega
                        TextoID = TBVendas!ID_entrega
                    End If
                    TBVendas.MoveNext
                Loop
            End If
            TBVendas.Close
        'Se for fornecedor
        Else
            TBVendas.Open "Select CC.* from (Compras_comercial CC INNER JOIN Compras_pedido CP ON CP.IDpedido = CC.IDpedido) INNER JOIN tbl_proposta_nota PN ON PN.proposta = CP.Pedido and PN.Revisao = 0 where PN.ID_nota = " & TBFIltro!ID & " and CC.ID_entrega IS NOT NULL and CC.ID_entrega <> 0 and CC.localentrega IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
            If TBVendas.EOF = False Then
                Permitido = False
                .AddItem TBVendas!localentrega
                .ItemData(cmbEntrega.NewIndex) = TBVendas!ID_entrega
                txtID_entrega = TBVendas!ID_entrega
                .Text = TBVendas!localentrega
            End If
            TBVendas.Close
        End If

'Busca do cadastro do cliente o local de entrega
        If Tipo = "C" Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_entrega where idcliente = " & TBFIltro!Id_Int_Cliente & " and Tipo = '" & Tipo & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Do While TBClientes.EOF = False
                    If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                        Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                    Else
                        Endereco = IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
                    End If
                    If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                        Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                    Else
                        Bairro = IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
                    End If
                    Endereco2 = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_entrega), "", TBClientes!cidade_entrega) & " - " & IIf(IsNull(TBClientes!uf_entrega), "", TBClientes!uf_entrega) & " - " & IIf(IsNull(TBClientes!cep_entrega), "", TBClientes!cep_entrega)
                    .AddItem Endereco2
                    .ItemData(cmbEntrega.NewIndex) = TBClientes!identrega
                    txtID_entrega = IIf(IsNull(TBClientes!identrega), 0, TBClientes!identrega)
                    .Text = Endereco2
                    TBClientes.MoveNext
                Loop
            End If
        End If
        

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
