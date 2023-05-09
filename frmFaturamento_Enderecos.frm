VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Enderecos 
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
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton btnCarregar 
      Height          =   795
      Left            =   5340
      TabIndex        =   6
      Top             =   2430
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1402
      DibPicture      =   "frmFaturamento_Enderecos.frx":0000
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Endereço(s) de cobrança"
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
         ItemData        =   "frmFaturamento_Enderecos.frx":339B
         Left            =   120
         List            =   "frmFaturamento_Enderecos.frx":339D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Endereço de cobrança"
         Top             =   300
         Width           =   6315
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Endereço(s) de entrega"
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
         ItemData        =   "frmFaturamento_Enderecos.frx":339F
         Left            =   120
         List            =   "frmFaturamento_Enderecos.frx":33A1
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Endereço de entrega."
         Top             =   300
         Width           =   6315
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   3285
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_Enderecos.frx":33A3
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
      Icon            =   "frmFaturamento_Enderecos.frx":673E
   End
End
Attribute VB_Name = "frmFaturamento_Enderecos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcCarregaEntrega()
On Error GoTo tratar_erro



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnCarregar_Click()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv_NFe_NS

.txtEntrega.Text = cmbEntrega.Text

If cmbEntrega <> "" Then
    .txtID_entrega = cmbEntrega.ItemData(cmbEntrega.ListIndex)
Else
   .txtID_entrega = 0
End If

.txtCobranca.Text = Cmb_cobranca.Text

If Cmb_cobranca <> "" Then
    .txtID_cobranca.Text = Cmb_cobranca.ItemData(Cmb_cobranca.ListIndex)
Else
    .txtID_cobranca.Text = 0
End If

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro


'Pega ID do cliente da nota fiscal
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select ID, Id_Int_Cliente, txt_Razao_Nome, Modelo from tbl_Dados_Nota_Fiscal where id = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then

                If TBFIltro!Modelo = "65" Then
                    NFCe = True
                Else
                    NFCe = False
                End If

    'Verifica se é cliente ou fornecedor
        Set TBFI = CreateObject("adodb.recordset")
        StrSql = "Select * from Clientes where IDCliente = " & TBFIltro!Id_Int_Cliente & " and NomeRazao = '" & TBFIltro!txt_Razao_Nome & "'"
        'Debug.print StrSql
        
        TBFI.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
        'Define como cliente ou fornecedor
                    If TBFI.EOF = False Then
                        Tipo = "C"
                    Else
                    Tipo = "F"
                            TBFI.Close
                    End If
        
       End If
       

If Tipo <> "" Then
 '=================================================================
 ' Carrega local de entrega do cadastro geral
 '=================================================================
        With cmbEntrega
        .Clear
            Set TBClientes = CreateObject("adodb.recordset")
            StrSql = "Select * from clientes_entrega where idcliente = " & TBFIltro!Id_Int_Cliente & " and Tipo = '" & Tipo & "'"
            'Debug.print StrSql

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
                    .Text = Endereco2
                    TBClientes.MoveNext
                Loop
            End If
            TBClientes.Close
    End With

 '=================================================================
 ' Carrega local de cobrança do cadastro geral
 '=================================================================

    With Cmb_cobranca
    .Clear
        Set TBClientes = CreateObject("adodb.recordset")
        StrSql = "Select * from clientes_cobranca where idcliente = " & TBFIltro!Id_Int_Cliente & " and Tipo = '" & Tipo & "'"
        'Debug.print StrSql

        TBClientes.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Do While TBClientes.EOF = False
                    If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                        Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
                    Else
                        Endereco = IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
                    End If
                    If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                        Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
                    Else
                        Bairro = IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
                    End If
                    Endereco2 = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_Cobranca), "", TBClientes!cidade_Cobranca) & " - " & IIf(IsNull(TBClientes!uf_Cobranca), "", TBClientes!uf_Cobranca) & " - " & IIf(IsNull(TBClientes!cep_Cobranca), "", TBClientes!cep_Cobranca)
                    .AddItem Endereco2
                    .ItemData(Cmb_cobranca.NewIndex) = TBClientes!idCobranca
                    .Text = Endereco2
                    TBClientes.MoveNext
                Loop
            End If

   End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
