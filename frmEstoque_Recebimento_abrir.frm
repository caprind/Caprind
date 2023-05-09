VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_Recebimento_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estoque - Recebimento - Localizar pedidos de compra"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9390
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk_data_recebimento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data do recebimento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3090
      TabIndex        =   11
      Top             =   3210
      Width           =   2085
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   20
      Top             =   2940
      Width           =   9285
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   7800
         TabIndex        =   13
         ToolTipText     =   "Data final."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   487915521
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5910
         TabIndex        =   12
         ToolTipText     =   "Data inicio."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   487915521
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7395
         TabIndex        =   22
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5550
         TabIndex        =   21
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.CheckBox Opt_recebidos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recebidos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6390
      TabIndex        =   9
      Top             =   1133
      Width           =   1185
   End
   Begin VB.CheckBox Chk_programacao_sem_pedido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Progr. s/ pedido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7675
      TabIndex        =   10
      Top             =   1133
      Width           =   1665
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2730
      Top             =   270
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmEstoque_Recebimento_abrir.frx":0000
      Count           =   1
   End
   Begin VB.ComboBox Cmb_empresa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmEstoque_Recebimento_abrir.frx":21ED
      Left            =   1230
      List            =   "frmEstoque_Recebimento_abrir.frx":21EF
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1065
      Width           =   5055
   End
   Begin VB.Frame Frame1 
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
      Height          =   1515
      Left            =   55
      TabIndex        =   14
      Top             =   1410
      Width           =   9285
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4320
         TabIndex        =   23
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   8
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox Cmb_ordenar 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmEstoque_Recebimento_abrir.frx":21F1
         Left            =   6840
         List            =   "frmEstoque_Recebimento_abrir.frx":21FB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Ordenar por."
         Top             =   1050
         Width           =   2265
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
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   6645
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmEstoque_Recebimento_abrir.frx":221A
         Left            =   180
         List            =   "frmEstoque_Recebimento_abrir.frx":2236
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4065
      End
      Begin VB.ComboBox cmbfamilia 
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
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Familia."
         Top             =   1050
         Visible         =   0   'False
         Width           =   6645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7462
         TabIndex        =   19
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1792
         TabIndex        =   16
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2767
         TabIndex        =   15
         Top             =   840
         Width           =   1470
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   17
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   36
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   270
      TabIndex        =   18
      Top             =   1065
      Width           =   825
   End
End
Attribute VB_Name = "frmEstoque_Recebimento_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_data_recebimento_Click()
On Error GoTo tratar_erro

If Chk_data_recebimento.Value = 1 Then
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
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

Private Sub Chk_programacao_sem_pedido_Click()
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    .AddItem "Fornecedor"
    .AddItem "Nota fiscal"
    If Chk_programacao_sem_pedido.Value = 1 Then
        .AddItem "Programação de compra"
        .Text = "Programação de compra"
    Else
        .AddItem "Pedido de compra"
        .Text = "Pedido de compra"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
    If cmbfiltrarpor = "Família" Then ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True Else ProcCarregaComboGrupoFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
    txtTexto.Visible = False
    cmbfamilia.Visible = True
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With frmEstoque_Recebimento
    .Lista_movimentacao.ListItems.Clear
    .txtID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    
    StatusFiltro = ""
    StatusFiltroRel = ""
    DataFiltro = ""
    DataFiltroRel = ""
    
    If Chk_programacao_sem_pedido.Value = 1 Then
        If Opt_recebidos.Value = 1 Then
            StatusFiltro = "and (ERP.Status_item = 'RECEBIDO' or ERP.Status_item = 'PARCIAL')"
            StatusFiltroRel = "and ({Estoque_recebimento_programacao.Status_item} = 'RECEBIDO' or {Estoque_recebimento_programacao.Status_item} = 'PARCIAL')"
        Else
            StatusFiltro = "and (ERP.Status_item = 'N_RECEBIDO' or ERP.Status_item = 'PARCIAL')"
            StatusFiltroRel = "and ({Estoque_recebimento_programacao.Status_item} = 'N_RECEBIDO' or {Estoque_recebimento_programacao.Status_item} = 'PARCIAL')"
        End If
                
        If Chk_data_recebimento.Value = 1 Then
            DataFiltro = "and ERP.Data_recebimento Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            DataFiltroRel = "and {Estoque_recebimento_programacao.Data_recebimento} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Estoque_recebimento_programacao.Data_recebimento} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
        End If
        
        Select Case Cmb_ordenar
            Case "Código interno": Ordenar = "ERP.codigo"
            Case "Descrição": Ordenar = "ERP.descricao"
        End Select
        
        CamposFiltro = "ERP.Id_Item, ERP.ID_empresa, ERP.programatexto, ERP.Codigo, ERP.Descricao, ERP.Unidade, ERP.Quant_Comp, ERP.Quant_Comp_PC, ERP.Status_item"
        If txtTexto <> "" Or cmbfamilia <> "" Then
            If cmbfiltrarpor = "Código de referência" Then
                If Optinicio.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} like '" & txtTexto.Text & "*' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
                If Optmeio.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} like '*" & txtTexto.Text & "*' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
                If Optfim.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} like '*" & txtTexto.Text & "' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
                If optIgual.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia = '" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia = '" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} = '" & txtTexto.Text & "' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
            ElseIf cmbfiltrarpor = "Família" Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP where classe = '" & cmbfamilia & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP where classe = '" & cmbfamilia & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                    .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_programacao.classe} = '" & cmbfamilia & "' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                Else
                    Select Case cmbfiltrarpor
                        Case "Código interno": TextoFiltro = "codigo"
                        Case "Descrição": TextoFiltro = "descricao"
                        Case "Fornecedor": TextoFiltro = "Nome_Razao"
                        Case "Programação de compra": TextoFiltro = "Programatexto"
                        Case "Nota fiscal":
                            TextoFiltro = "nota_fiscal"
                            If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
                    End Select
                    If Optinicio.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " like '" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " like '" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_programacao." & TextoFiltro & "} like '" & txtTexto.Text & "*' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                    End If
                    If Optmeio.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_programacao." & TextoFiltro & "} like '*" & txtTexto.Text & "*' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                    End If
                    If Optfim.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_programacao." & TextoFiltro & "} like '*" & txtTexto.Text & "' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                    End If
                    If optIgual.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " = '" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP where " & TextoFiltro & " = '" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_programacao." & TextoFiltro & "} = '" & txtTexto.Text & "' and {Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                    End If
            End If
        Else
            .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_programacao ERP where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
            .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.ID_item FROM Estoque_recebimento_programacao ERP where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.ID_item, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.ID_item"
            .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_programacao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
        End If
    Else
        If Opt_recebidos.Value = 1 Then
            StatusFiltro = "and (ERP.Status_item = 'RECEBIDO' or ERP.Status_item = 'PARCIAL')"
            StatusFiltroRel = "and ({Estoque_recebimento_pedido.Status_item} = 'RECEBIDO' or {Estoque_recebimento_pedido.Status_item} = 'PARCIAL')"
        Else
            StatusFiltro = "and (ERP.Status_item = 'N_RECEBIDO' or ERP.Status_item = 'PARCIAL')"
            StatusFiltroRel = "and ({Estoque_recebimento_pedido.Status_item} = 'N_RECEBIDO' or {Estoque_recebimento_pedido.Status_item} = 'PARCIAL')"
        End If
        
        If Chk_data_recebimento.Value = 1 Then
            DataFiltro = "and ERP.Data_recebimento Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            DataFiltroRel = "and {Estoque_recebimento_pedido.Data_recebimento} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Estoque_recebimento_pedido.Data_recebimento} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
        End If
        
        Select Case Cmb_ordenar
            Case "Código interno": Ordenar = "ERP.desenho"
            Case "Descrição": Ordenar = "ERP.descricao"
        End Select
        
        CamposFiltro = "ERP.IDlista, ERP.ID_empresa, ERP.Pedido, ERP.Desenho, ERP.Descricao, ERP.UN, ERP.UNidade_com, ERP.preco_unitario, ERP.Quant_Comp, ERP.Quant_Comp_PC, ERP.Prazo, ERP.Status_item, ERP.Ordem, ERP.Qtde_estoque"
        If txtTexto <> "" Or cmbfamilia <> "" Then
            If cmbfiltrarpor = "Código de referência" Then
                If Optinicio.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} like '" & txtTexto.Text & "*' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
                If Optmeio.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "%' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} like '*" & txtTexto.Text & "*' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
                If Optfim.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia like '%" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} like '*" & txtTexto.Text & "' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
                If optIgual.Value = True Then
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia = '" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(ERP.Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP INNER JOIN item_aplicacoes IA ON ERP.Codproduto = IA.Codproduto where IA.N_referencia = '" & txtTexto.Text & "' and ERP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                    .FormulaRel_Estoque_Recebimento = "{item_aplicacoes.N_referencia} = '" & txtTexto.Text & "' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                End If
            ElseIf cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
                    If cmbfiltrarpor = "Família" Then TextoFiltro = "Familia" Else TextoFiltro = "Grupo"
                    .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " = '" & cmbfamilia & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                    .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " = '" & cmbfamilia & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                    .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido." & TextoFiltro & "} = '" & cmbfamilia & "' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                Else
                    Select Case cmbfiltrarpor
                        Case "Código interno": TextoFiltro = "desenho"
                        Case "Código de referência": TextoFiltro = "n_referencia"
                        Case "Descrição": TextoFiltro = "descricao"
                        Case "Fornecedor": TextoFiltro = "fornecedor"
                        Case "Pedido de compra": TextoFiltro = "Pedido"
                        Case "Nota fiscal":
                            TextoFiltro = "nota_fiscal"
                            If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
                    End Select
                    
                    If Optinicio.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " like '" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " like '" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido." & TextoFiltro & "} like '" & txtTexto.Text & "*' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                    End If
                    If Optmeio.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & TipoProduto & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido." & TextoFiltro & "} like '*" & txtTexto.Text & "*' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & TipoProdutoRel & " " & DataFiltroRel
                    End If
                    If Optfim.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " like '%" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & TipoProduto & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido." & TextoFiltro & "} like '*" & txtTexto.Text & "' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                    End If
                    If optIgual.Value = True Then
                        .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " = '" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
                        .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP where " & TextoFiltro & " = '" & txtTexto.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & TipoProduto & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
                        .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido." & TextoFiltro & "} = '" & txtTexto.Text & "' and {Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
                    End If
            End If
        Else
            .StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido ERP where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by " & CamposFiltro & " order by " & Ordenar
            .StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, ERP.IDlista FROM Estoque_recebimento_pedido ERP where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltro & " " & DataFiltro & " group by ERP.IDlista, ERP.Data_recebimento, ERP.Nota_fiscal order by ERP.IDlista"
            .FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & StatusFiltroRel & " " & DataFiltroRel
        End If
    End If
    If Chk_programacao_sem_pedido.Value = 1 Then Programacao = True Else Programacao = False
    .ProcCarregaLista
    ProcGravarDataFiltroRel msk_fltInicio, msk_fltFim, True, 0, ""
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
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9285, 5, True
cmbfiltrarpor = "Pedido de compra"
Cmb_ordenar = "Código interno"
ProcCarregaComboEmpresa Cmb_empresa, False
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

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

Private Sub Opt_recebidos_Click()
On Error GoTo tratar_erro

With Chk_data_recebimento
    If Opt_recebidos.Value = 1 Then
        .Enabled = True
    Else
        .Value = 0
        .Enabled = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    If cmbfiltrarpor = "Nota fiscal" Then
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

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)

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
