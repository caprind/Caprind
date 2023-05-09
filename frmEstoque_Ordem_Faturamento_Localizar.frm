VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_Ordem_Faturamento_Localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque | Ordem de faturamento - Localizar"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7110
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo da ordem de faturamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   4530
      TabIndex        =   23
      Top             =   660
      Width           =   2385
      Begin VB.OptionButton optProduto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produtos"
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
         Height          =   210
         Left            =   270
         TabIndex        =   25
         Top             =   420
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptServico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serviços"
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
         Height          =   210
         Left            =   1260
         TabIndex        =   24
         Top             =   420
         Width           =   915
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   21
      Top             =   4770
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton btnFiltrar 
      Height          =   705
      Left            =   3810
      TabIndex        =   20
      Top             =   3870
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   1244
      DibPicture      =   "frmEstoque_Ordem_Faturamento_Localizar.frx":0000
      Caption         =   "Filtrar ordem(ns) de faturamento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   926
      DibPicture      =   "frmEstoque_Ordem_Faturamento_Localizar.frx":3650
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
      Icon            =   "frmEstoque_Ordem_Faturamento_Localizar.frx":6CA0
   End
   Begin VB.CheckBox Chk_data 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar por data de emissão"
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
      Left            =   1140
      TabIndex        =   9
      Top             =   3810
      Width           =   2325
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      TabIndex        =   15
      Top             =   3810
      Width           =   3615
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   2130
         TabIndex        =   11
         ToolTipText     =   "Data final."
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   199360513
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   600
         TabIndex        =   10
         ToolTipText     =   "Data inicio."
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   199360513
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
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
         Left            =   360
         TabIndex        =   17
         Top             =   390
         Width           =   195
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "até"
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
         Left            =   1845
         TabIndex        =   16
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   12
      Top             =   660
      Width           =   4395
      Begin VB.OptionButton optProduto_servico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produtos/Serviços"
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
         Height          =   210
         Left            =   6720
         TabIndex        =   22
         Top             =   960
         Width           =   1845
      End
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frmEstoque_Ordem_Faturamento_Localizar.frx":6FBA
         Left            =   180
         List            =   "frmEstoque_Ordem_Faturamento_Localizar.frx":6FBC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   3105
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   6795
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   180
         TabIndex        =   18
         Top             =   900
         Width           =   6435
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
            Left            =   5400
            TabIndex        =   8
            Top             =   240
            Width           =   705
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
            Left            =   1970
            TabIndex        =   6
            Top             =   240
            Width           =   645
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
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
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
            Left            =   3730
            TabIndex        =   7
            Top             =   240
            Width           =   555
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
         ItemData        =   "frmEstoque_Ordem_Faturamento_Localizar.frx":6FBE
         Left            =   180
         List            =   "frmEstoque_Ordem_Faturamento_Localizar.frx":6FDA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   450
         Width           =   6435
      End
      Begin VB.TextBox txtTexto 
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
         Height          =   315
         Left            =   210
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1770
         Width           =   6375
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         ToolTipText     =   "Número do CPF."
         Top             =   1770
         Visible         =   0   'False
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtcnpj 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Número do CNPJ."
         Top             =   1770
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   2662
         TabIndex        =   14
         Top             =   1560
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmEstoque_Ordem_Faturamento_Localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

 ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_data_Click()
On Error GoTo tratar_erro

If Chk_data.Value = 1 Then
    Frame2.Enabled = True
'    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "CPF" Then
    txtTexto.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = True
ElseIf cmbfiltrarpor = "CNPJ" Then
        txtTexto.Visible = False
        txtcnpj.Visible = True
        txtCpf.Visible = False
    Else
        txtTexto.Visible = True
        txtcnpj.Visible = False
        txtCpf.Visible = False
        If (cmbfiltrarpor = "Ordem de fat.") And txtTexto <> "" Then
            VerifNumero = txtTexto
            ProcVerificaNumero
            If VerifNumero = False Then
                txtTexto = ""
                txtTexto.SetFocus
                Exit Sub
            End If
        End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
Dim TextAtiva As String

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
TextAtivaCancelada = ""



With frmEstoque_Ordem_Faturamento
    .ListaNota.ListItems.Clear
    .txtIDEmpresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    DataFiltro = ""
    .Strsql_Faturamento = ""
    If Faturamento_NF_Saida = True Then CampoDataFiltro = "NF.dt_DataEmissao" Else CampoDataFiltro = "NF.dt_Saida_Entrada"
    If Chk_data.Value = 1 Then DataFiltro = " and " & CampoDataFiltro & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
'    If chkentrada.Value = Checked Then DataFiltro = DataFiltro & " And NF.Int_TipoNota = '2'"
 '   If chkSaida.Value = Checked Then DataFiltro = DataFiltro & " And NF.Int_TipoNota = '1'"
    
    CamposFiltro = "NF.Int_TipoNota, NF.ID, NF.dt_DataEmissao, NF.dt_Saida_Entrada, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, TN.dbl_Valor_Total_Nota, NF.txt_Razao_Nome, NF.txt_CNPJ_CPF, NF.Int_status, NF.Imprimir, NF.ID_empresa, NF.Aplicacao, NF.DtValidacaoOF, NF.DtValidacao, NF.RPS"
'    If Formulario = "Estoque/Ordem de faturamento" Then
        TextoFiltroVal = " and NF.int_NotaFiscal IS NULL"
        OrdenarFiltro = " order by NF.ID"
'    Else
'        TextoFiltroVal = "" '" and NF.DtValidacaoOF IS NOT NULL"
'        OrdenarFiltro = " order by NF.int_NotaFiscal desc"
'    End If
    
    Select Case cmbfiltrarpor
        Case "Ordem de fat.": TextoFiltro = "NF.ID"
        Case "Emitente": TextoFiltro = "NF.txt_Razao_Nome"
        Case "Destinatário": TextoFiltro = "NF.txt_Razao_Nome"
        Case "Código interno": TextoFiltro = "NFP.int_cod_produto"
        Case "Código de referência": TextoFiltro = "NFP.N_Referencia"
        Case "Descrição": TextoFiltro = "NFP.txt_descricao"
        Case "Pedido cliente": TextoFiltro = "NFP.pccliente"
        Case "Pedido interno/Pedido de compra": TextoFiltro = "PN.proposta"
    End Select
    
    If cmbfiltrarpor = "Código interno" Or cmbfiltrarpor = "Código de referência" Or cmbfiltrarpor = "Descrição" Or cmbfiltrarpor = "Pedido cliente" Then
        FiltroPadrao = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & TextoFiltro
        FiltroPadraoTotal = "(tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota"
    ElseIf cmbfiltrarpor = "Nosso número" Then
        FiltroPadrao = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Detalhes_Recebimento_Nboletos DR ON NF.ID = DR.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & TextoFiltro
        FiltroPadraoTotal = "(tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Detalhes_Recebimento_Nboletos DR ON NF.ID = DR.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota"
    ElseIf cmbfiltrarpor = "Pedido interno/Pedido de compra" Then
        FiltroPadrao = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & TextoFiltro
        FiltroPadraoTotal = "(tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota"
    ElseIf cmbfiltrarpor = "Chave de acesso" Then
        FiltroPadrao = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Dados_Nota_Fiscal_NFE NFE ON NF.ID = NFE.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & TextoFiltro
        FiltroPadraoTotal = "(tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Dados_Nota_Fiscal_NFE NFE ON NF.ID = NFE.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota"
    End If
    
    If txtTexto.Visible = True And txtTexto <> "" Or txtcnpj.Visible = True And txtcnpj <> "__.___.___/____-__" Or txtCpf.Visible = True And txtCpf <> "___.___.___-__" Then
        If cmbfiltrarpor = "CNPJ" Or cmbfiltrarpor = "CPF" Then
            If cmbfiltrarpor = "CNPJ" Then TextoFiltro = "NF.txt_CNPJ_CPF = '" & txtcnpj & "'" Else TextoFiltro = "NF.txt_CNPJ_CPF = '" & txtCpf & "'"
            .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
            .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
            .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
            .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
            .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
        
        If cmbSerie.Text <> "" And Left(frmEstoque_Ordem_Faturamento.Caption, 7) <> "Estoque" Then
            .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
            .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
            .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
            .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
            .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
        End If
        
        
        
        ElseIf cmbfiltrarpor = "Ordem de fat." Then
            .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " = " & txtTexto & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
            .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " = " & txtTexto & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
            .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " = " & txtTexto & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
            .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " = " & txtTexto & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
            .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & " = " & txtTexto & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
        
        
        
        ElseIf cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Then
            .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
            .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
            .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
            .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
            .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
        
'        If cmbSerie.Text <> "" And Left(frmEstoque_Ordem_Faturamento.Caption, 7) <> "Estoque" Then
'            .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
'            .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
'            .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
'            .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'            .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'        End If
        
        
        Else
            .Strsql_Faturamento = FiltroPadrao & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
            .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from " & FiltroPadraoTotal & " where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
            .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from " & FiltroPadraoTotal & " where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
            .Strsql_FaturamentoNFe = FiltroPadrao & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
            .Strsql_FaturamentoNFSe = FiltroPadrao & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
        
'        If cmbSerie.Text <> "" And Left(frmEstoque_Ordem_Faturamento.Caption, 7) <> "Estoque" Then
'            .Strsql_Faturamento = FiltroPadrao & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
'            .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from " & FiltroPadraoTotal & " where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
'            .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from " & FiltroPadraoTotal & " where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
'            .Strsql_FaturamentoNFe = FiltroPadrao & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'            .Strsql_FaturamentoNFSe = FiltroPadrao & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'        End If
        
        
        End If
        
    Else
    
' If Left(frmEstoque_Ordem_Faturamento.Caption, 7) <> "Estoque" Then
'    If cmbSerie.Text = "" Or cmbfiltrarpor.Text = "Ordem de fat." Then
'        .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
'        .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
'        .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
'        .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'        .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'    Else
'        .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
'        .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
'        .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
'        .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'        .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.serie = '" & cmbSerie.Text & "' and NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
'    End If
' Else
         .Strsql_Faturamento = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING " & TextAtivaCancelada & " NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroVal & OrdenarFiltro
        .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 1"
        .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & DataFiltro & " " & TextoFiltroVal & " and NF.Int_status = 2"
        .Strsql_FaturamentoNFe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF <> 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
        .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & ", NF.int_NotaFiscal AS NNF from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.DtValidacao IS NOT NULL AND TipoNF = 'SA' group by " & CamposFiltro & ", NF.int_NotaFiscal HAVING NF.tiponf = '" & TipoNF & "' and NF.Aplicacao = '" & .Aplicacao & "' and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & OrdenarFiltro
' End If

    
    End If
    'Debug.print .Strsql_Faturamento
    
    .ProcCarregaListaNota (1)

End With

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Faturamento_NF_Saida = True Then Chk_data.Caption = "Buscar por data de emissão" Else Chk_data.Caption = "Buscar por data de entrada"
If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Localizar"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Localizar"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            Caption = "Estoque - Ordem de faturamento - Localizar"
        Else
            Caption = "Estoque - Nota fiscal - Localizar"
End If
ProcCarregaFiltrarpor
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
TipoNF = "M1"
ProcCarregaComboEmpresa Cmb_empresa, False

If Formulario = "Estoque/Ordem de faturamento" Then
    Cmb_empresa.Text = frmEstoque_Ordem_Faturamento.txtEmpresa
Else
    Cmb_empresa.Text = frmEstoque_Ordem_Faturamento.txtEmpresa
End If

Chk_data.Value = 1

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

Private Sub optProduto_Click()
On Error GoTo tratar_erro

TipoNF = "M1"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optProduto_servico_Click()
On Error GoTo tratar_erro

TipoNF = "M1SA"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptServico_Click()
On Error GoTo tratar_erro

TipoNF = "SA"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If (cmbfiltrarpor = "Nota fiscal" Or cmbfiltrarpor = "Ordem de fat.") And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFiltrarpor()
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    If Formulario = "Faturamento/Nota fiscal/Terceiros" Or Formulario = "Estoque/Nota fiscal" Then .AddItem "Emitente" Else .AddItem "Destinatário"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Pedido cliente"
    .AddItem "Pedido interno/Pedido de compra"
    .AddItem "Ordem de fat."
    .AddItem "CNPJ"
    .AddItem "CPF"
    .AddItem "Chave de acesso"
    
    If Formulario = "Estoque/Ordem de faturamento" Then
        .Text = "Ordem de fat."
    Else
        .AddItem "Nosso número"
        .AddItem "Nota fiscal"
        .Text = "Nota fiscal"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCarregaSerie()
On Error GoTo tratar_erro

With cmbSerie
    .Clear
    .AddItem ""
    .AddItem "0"
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    If Formulario = "Estoque/Ordem de faturamento" Then
    .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCarregaFiltrarProdeServ()
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "Destinatário"
    .AddItem "Código de referência do produto"
    .AddItem "Código de referência do serviço"
    .AddItem "Código interno do produto"
    .AddItem "Código interno do serviço"
    .AddItem "Descrição do produto"
    .AddItem "Descrição do serviço"
    .AddItem "Nosso número"
    .AddItem "Nota fiscal"
    .AddItem "Pedido cliente do produto"
    .AddItem "Pedido cliente do serviço"
    .AddItem "Pedido interno/Pedido de compra"
    .Text = "Nota fiscal"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

