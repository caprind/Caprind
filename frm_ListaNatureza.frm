VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_ListaNatureza 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Localizar natureza de operação | CFOP"
   ClientHeight    =   9585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ListaNatureza.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   17
      Top             =   9180
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   714
      DibPicture      =   "frm_ListaNatureza.frx":1042
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frm_ListaNatureza.frx":4692
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.TextBox Txt_observacoes 
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
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Observações."
      Top             =   8025
      Width           =   15345
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   150
      TabIndex        =   13
      Top             =   540
      Width           =   15345
      Begin MSMask.MaskEdBox txtTexto 
         Height          =   435
         Left            =   6870
         TabIndex        =   1
         Top             =   330
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   767
         _Version        =   393216
         BackColor       =   12640511
         MaxLength       =   5
         Mask            =   "#.###"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3930
         TabIndex        =   15
         Top             =   240
         Width           =   2925
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1530
            TabIndex        =   5
            Top             =   180
            Width           =   585
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   870
            TabIndex        =   4
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2130
            TabIndex        =   6
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_ListaNatureza.frx":56E4
         Left            =   180
         List            =   "frm_ListaNatureza.frx":56EE
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.TextBox txt_dados_adicionais 
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
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   7815
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Texto padrão para dados adicionais da nota."
      Top             =   6735
      Width           =   7672
   End
   Begin VB.TextBox txt_corpo_nota 
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
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Texto padrão para corpo da nota."
      Top             =   6735
      Width           =   7635
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   6180
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
   Begin MSComctlLib.ListView lst_NatOp 
      Height          =   4695
      Left            =   150
      TabIndex        =   2
      Top             =   1470
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   884
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "CFOP"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Natureza da operação"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "ICMS"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "IPI"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "PIS"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "COFINS"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Red. BC ICMS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Soma IPI BC ICMS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Soma IPI BC ICMS ST"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Soma ret. nos totais"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Desc. SUFRAMA"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "MP aplic."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "T"
         Text            =   "Dest. imp."
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "Remessa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Object.Tag             =   "T"
         Text            =   "Retorno"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Observações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7350
      TabIndex        =   14
      Top             =   7800
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Texto dados adicionais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10830
      TabIndex        =   11
      Top             =   6510
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Texto corpo da nota"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3225
      TabIndex        =   10
      Top             =   6510
      Width           =   1470
   End
End
Attribute VB_Name = "frm_ListaNatureza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro
  
lst_NatOp.ListItems.Clear
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With lst_NatOp.ListItems
            .Add , , TBLISTA!IDCountCfop
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!ID_CFOP), "", TBLISTA!ID_CFOP)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Txt_ICMS), "", TBLISTA!Txt_ICMS)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!txt_IPI), "", TBLISTA!txt_IPI)
            .Item(.Count).SubItems(5) = IIf(TBLISTA!TemPIS = True, "SIM", "NÃO")
            .Item(.Count).SubItems(6) = IIf(TBLISTA!TemCOFINS = True, "SIM", "NÃO")
            .Item(.Count).SubItems(7) = IIf(TBLISTA!TemReducaoBC = True, "SIM", "NÃO")
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!txt_Somar), "", TBLISTA!txt_Somar)
            .Item(.Count).SubItems(9) = IIf(TBLISTA!Somar_IPI_BC_ICMSST = True, "SIM", "NÃO")
            .Item(.Count).SubItems(10) = IIf(TBLISTA!Soma_retorno_totalnf = True, "SIM", "NÃO")
            .Item(.Count).SubItems(11) = IIf(TBLISTA!Suframa = True, "SIM", "NÃO")
            .Item(.Count).SubItems(12) = IIf(TBLISTA!MPA = True, "SIM", "NÃO")
            .Item(.Count).SubItems(13) = IIf(TBLISTA!Retem = True, "SIM", "NÃO")
            .Item(.Count).SubItems(14) = IIf(TBLISTA!Remessa = True, "SIM", "NÃO")
            .Item(.Count).SubItems(15) = IIf(TBLISTA!retorno = True, "SIM", "NÃO")
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro
  
lst_NatOp.ListItems.Clear

If cmbfiltrarpor.Text = "CFOP" Then
    txtTexto.Mask = "#.###"
Else
    txtTexto.Mask = ""
    txtTexto.Text = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyReturn: lst_NatOp_DblClick
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfiltrarpor = "CFOP"
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_NatOp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lst_NatOp, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_NatOp_DblClick()
On Error GoTo tratar_erro
    
If lst_NatOp.ListItems.Count = 0 Then Exit Sub
If Clientes = True Then
    With frmVendas_cliente
        .txtID_cfop.Text = lst_NatOp.SelectedItem
        .txtCFOP.Text = lst_NatOp.SelectedItem.ListSubItems(1)
        .txtOperacao.Text = lst_NatOp.SelectedItem.ListSubItems(2)
    End With
ElseIf Compras_Fornecedores = True Then
    With frmCompras_fornecedores
        .txtID_cfop.Text = lst_NatOp.SelectedItem
        .txtCFOP.Text = lst_NatOp.SelectedItem.ListSubItems(1)
        .txtOperacao.Text = lst_NatOp.SelectedItem.ListSubItems(2)
    End With
ElseIf Vendas_PI = True Or Vendas_Proposta = True Then
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            If Sit_REG = 1 Then
                .Txt_ID_CFOP_prod = lst_NatOp.SelectedItem
                .Txt_CFOP_prod = lst_NatOp.SelectedItem.ListSubItems(1)
                .Txt_natureza_operacao_prod = lst_NatOp.SelectedItem.ListSubItems(2)
                If lst_NatOp.SelectedItem.ListSubItems(15) = "SIM" Then .chkRetorno.Value = 1 Else .chkRetorno.Value = 0
                .ProcCarregaCST
            Else
                .Txt_ID_CFOP_serv = lst_NatOp.SelectedItem
                .Txt_CFOP_serv = lst_NatOp.SelectedItem.ListSubItems(1)
                .Txt_natureza_operacao_serv = lst_NatOp.SelectedItem.ListSubItems(2)
            End If
        End With
    ElseIf Faturamento = True Then
    If Formulario <> "Estoque/Ordem de faturamento" Then
            With frmFaturamento_Prod_Serv
                If Sit_REG = 1 Then
                    .txtDados_Corpo = txt_corpo_nota
                    .txtDados_DadosAdicionais = txt_dados_adicionais
                ElseIf Sit_REG = 2 Then
                        .Txt_ID_CFOP_prod = lst_NatOp.SelectedItem
                        .Txt_CFOP_prod = lst_NatOp.SelectedItem.ListSubItems(1)
                        .Txt_natureza_operacao_prod = lst_NatOp.SelectedItem.ListSubItems(2)
                        If lst_NatOp.SelectedItem.ListSubItems(14) = "SIM" Then .chkRemessa.Value = 1 Else .chkRemessa.Value = 0
                        If lst_NatOp.SelectedItem.ListSubItems(15) = "SIM" Then .chkRetorno.Value = 1 Else .chkRetorno.Value = 0
                        .ProcCarregaCSTICMS
                        If NF_enviada = False And NFe_liberada = False Then .ProcVerificaCF False
                    Else
                        .Txt_ID_CFOP_serv = lst_NatOp.SelectedItem
                        .Txt_CFOP_serv = lst_NatOp.SelectedItem.ListSubItems(1)
                        .Txt_natureza_operacao_serv = lst_NatOp.SelectedItem.ListSubItems(2)
                        .ProcCarregaCSTServ
                        If NF_enviada = False And NFe_liberada = False Then .ProcVerificaCF False
                End If
                If lst_NatOp.SelectedItem.SubItems(11) = "SIM" Then Desconto_Suframa = True Else Desconto_Suframa = False
            End With
            Else
            With frmEstoque_Ordem_Faturamento
                If Sit_REG = 1 Then
'                    .txtDados_Corpo = txt_corpo_nota
'                    .txtDados_DadosAdicionais = txt_dados_adicionais
                ElseIf Sit_REG = 2 Then
                        .Txt_ID_CFOP_prod = lst_NatOp.SelectedItem
                        .Txt_CFOP_prod = lst_NatOp.SelectedItem.ListSubItems(1)
                        .Txt_natureza_operacao_prod = lst_NatOp.SelectedItem.ListSubItems(2)
                        If lst_NatOp.SelectedItem.ListSubItems(14) = "SIM" Then .chkRemessa.Value = 1 Else .chkRemessa.Value = 0
                        If lst_NatOp.SelectedItem.ListSubItems(15) = "SIM" Then .chkRetorno.Value = 1 Else .chkRetorno.Value = 0
                        .ProcCarregaCSTICMS
                        If NF_enviada = False And NFe_liberada = False Then .ProcVerificaCF False
                    Else
                        .Txt_ID_CFOP_serv = lst_NatOp.SelectedItem
                        .Txt_CFOP_serv = lst_NatOp.SelectedItem.ListSubItems(1)
                        .Txt_natureza_operacao_serv = lst_NatOp.SelectedItem.ListSubItems(2)
                        .ProcCarregaCSTServ
                        If NF_enviada = False And NFe_liberada = False Then .ProcVerificaCF False
                End If
                If lst_NatOp.SelectedItem.SubItems(11) = "SIM" Then Desconto_Suframa = True Else Desconto_Suframa = False
            End With
            
            End If
            
        ElseIf Compras_Pedido = True Then
            With frmCompras_Pedido
                If Sit_REG = 1 Then
                    .Txt_ID_CFOP_prod = lst_NatOp.SelectedItem
                    .txtCFOP_prod = lst_NatOp.SelectedItem.ListSubItems(1)
                    .Txt_natureza_operacao_prod = lst_NatOp.SelectedItem.ListSubItems(2)
                    If lst_NatOp.SelectedItem.ListSubItems(14) = "SIM" Then .chkRemessa.Value = 1 Else .chkRemessa.Value = 0
                    .ProcCarregaCST
                    .ProcCalculaValor True
                Else
                    .Txt_ID_CFOP_serv = lst_NatOp.SelectedItem
                    .txtCFOP_serv = lst_NatOp.SelectedItem.ListSubItems(1)
                    .txtNatureza_operacao_serv = lst_NatOp.SelectedItem.ListSubItems(2)
                End If
            End With
        Else
            With frmproj_produto
                If Sit_REG = 1 Then
                    .Txt_ID_CFOP = lst_NatOp.SelectedItem
                    .Txt_CFOP = lst_NatOp.SelectedItem.ListSubItems(1)
                    .Txt_natureza_operacao = lst_NatOp.SelectedItem.ListSubItems(2)
                Else
                    .Txt_ID_CFOP1 = lst_NatOp.SelectedItem
                    .Txt_CFOP1 = lst_NatOp.SelectedItem.ListSubItems(1)
                    .Txt_natureza_operacao1 = lst_NatOp.SelectedItem.ListSubItems(2)
                End If
            End With
End If
ProcSair
      
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_NatOp_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

txt_corpo_nota = ""
txt_dados_adicionais = ""
txt_observacoes = ""
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select txt_corpo_nota, txt_dados_adicionais, Obs from tbl_NaturezaOperacao where IDCountCfop = " & lst_NatOp.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txt_corpo_nota = IIf(IsNull(TBAbrir!txt_corpo_nota), "", TBAbrir!txt_corpo_nota)
    txt_dados_adicionais = IIf(IsNull(TBAbrir!txt_dados_adicionais), "", TBAbrir!txt_dados_adicionais)
    txt_observacoes = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro
  
lst_NatOp.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro
  
lst_NatOp.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro
  
lst_NatOp.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

TextoFiltro1 = "id_CFOP is not null"
FiltroRegiao = "id_CFOP is not null"
Permitido = False
If Vendas_PI = True Or Vendas_Proposta = True Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        If .txtIDcliente <> "" And .txtUF <> "" And .txtUF <> "EX" Then
            UF = .txtUF
            Permitido = True
        End If
        TextoFiltro1 = "Proprio = 'True'"
    End With
ElseIf Faturamento = True Then
If Formulario <> "Estoque/Ordem de faturamento" Then
        With frmFaturamento_Prod_Serv
            If .txtIDcliente <> "" And .cbo_UF <> "" And .cbo_UF <> "EX" Then
                UF = .cbo_UF
                Permitido = True
            End If
        End With
Else

End If
        If Faturamento_NF_Saida = True Then TextoFiltro1 = "Proprio = 'True'" Else TextoFiltro1 = "Terceiros = 'True'"
    ElseIf Compras_Pedido = True Then
        With frmCompras_Pedido
            If .txtIDfornecedor <> "" And .txtUF <> "" And .txtUF <> "EX" Then
                UF = .txtUF
                Permitido = True
            End If
            TextoFiltro1 = "Terceiros = 'True'"
        End With
    ElseIf Compras_Fornecedores = True Then
            TextoFiltro1 = "Terceiros = 'True'"
    Else

        If Sit_REG = 1 Then TextoFiltro1 = "Terceiros = 'True'" Else TextoFiltro1 = "Proprio = 'True'"
End If
    
If Permitido = True Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from regioes where uf = '" & UF & "' and Regiao = 'DE'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        FiltroRegiao = "DE = 'True'"
    Else
        FiltroRegiao = "FE = 'True'"
    End If
    TBFI.Close
End If
If cmbfiltrarpor.Text = "Descrição" Then TextoFiltro = "txt_Descricao" Else TextoFiltro = "id_CFOP"
TextoFiltroPadrao = "DtValidacao IS NOT NULL and (" & FiltroRegiao & " or DE = 'False' and FE = 'False' or DE is null and FE is null) and (" & TextoFiltro1 & " or Proprio = 'False' and Terceiros = 'False') Order by id_CFOP"

Set TBLISTA = CreateObject("adodb.recordset")
If cmbfiltrarpor.Text = "Descrição" Then
If txtTexto.Text <> "" Then
    TBLISTA.Open "Select * FROM tbl_NaturezaOperacao where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select * FROM tbl_NaturezaOperacao where " & TextoFiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
End If
Else
If txtTexto.Text <> "_.___" Then
    TBLISTA.Open "Select * FROM tbl_NaturezaOperacao where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select * FROM tbl_NaturezaOperacao where " & TextoFiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
End If

End If

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto1_Change()
On Error GoTo tratar_erro

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
