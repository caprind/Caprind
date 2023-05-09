VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMinuta_notas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Faturamento - Minuta de despacho - Localizar nota"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   510
      Left            =   55
      TabIndex        =   26
      Top             =   990
      Width           =   8805
      Begin VB.OptionButton optProduto_servico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produtos/Serviços"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3090
         TabIndex        =   16
         Top             =   210
         Width           =   1875
      End
      Begin VB.OptionButton optProduto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produtos"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Top             =   210
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptServico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serviços"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1650
         TabIndex        =   11
         Top             =   210
         Width           =   1185
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   22
      Top             =   5460
      Width           =   8805
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtPagIr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3930
         TabIndex        =   4
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   6150
         TabIndex        =   8
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmMinutas_notas.frx":0000
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   5610
         TabIndex        =   7
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmMinutas_notas.frx":37A7
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   4500
         TabIndex        =   5
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Ir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   5070
         TabIndex        =   6
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmMinutas_notas.frx":72B6
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   6690
         TabIndex        =   9
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmMinutas_notas.frx":B3AA
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar               reg. p/ pág."
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
         Left            =   1650
         TabIndex        =   25
         Top             =   240
         Width           =   2190
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de reg.: 0"
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
         Left            =   180
         TabIndex        =   24
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pág.: 0 de: 0"
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
         Left            =   7500
         TabIndex        =   23
         Top             =   240
         Width           =   945
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2355
      Left            =   60
      TabIndex        =   2
      Top             =   3090
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Série"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Vlr. total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Destinatário"
         Object.Width           =   5274
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   55
      TabIndex        =   17
      Top             =   1500
      Width           =   8805
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   27
         Top             =   270
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   14
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   12
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   13
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   15
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmMinutas_notas.frx":EC44
         Left            =   180
         List            =   "frmMinutas_notas.frx":EC60
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   450
         Width           =   3555
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
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1110
         Width           =   8415
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
         Left            =   3645
         TabIndex        =   19
         Top             =   900
         Width           =   1470
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
         Left            =   1537
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   6090
      Width           =   8805
      _ExtentX        =   15531
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   21
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   46
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   50
      ButtonTop3      =   2
      ButtonWidth3    =   41
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   93
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7650
         Top             =   90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmMinutas_notas.frx":ECEF
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmMinuta_notas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then
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

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If Minuta = True Then
    TextoFiltroEmpresa = " and NF.ID_empresa = " & frmMinuta.Cmb_empresa.ItemData(frmMinuta.Cmb_empresa.ListIndex)
    TextoFiltroAplic = " and NF.Aplicacao = 'P'"
Else
    TextoFiltroEmpresa = " and NF.ID_empresa = " & frmFaturamento_CartaCorrecao_NS.Cmb_empresa.ItemData(frmFaturamento_CartaCorrecao_NS.Cmb_empresa.ListIndex)
    TextoFiltroAplic = " and NF.Aplicacao = 'P'"
End If
TipoFiltro = " "
If optProduto.Value = True Then
    TipoFiltro = " and NFP.Tipo = 'P'"
ElseIf OptServico.Value = True Then
        TipoFiltro = " and NFP.Tipo = 'S'"
End If
CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, TN.dbl_Valor_Total_Nota, NF.txt_Razao_Nome, NF.Int_status"
OrdenarFiltro = " order by NF.dt_DataEmissao desc, NF.int_NotaFiscal desc"
Select Case cmbfiltrarpor
    Case "Nota fiscal":
        TextoFiltro = "NF.int_NotaFiscal"
        If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
    Case "Destinatário": TextoFiltro = "NF.txt_Razao_Nome"
    Case "Código interno":
        IMFFiltro = "NFP.int_cod_produto" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroEmpresa
        FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & IMFFiltro
    Case "Código de referência":
        IMFFiltro = "NFP.N_Referencia" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroEmpresa
        FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & IMFFiltro
    Case "Descrição"
        IMFFiltro = "NFP.txt_descricao" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroEmpresa
        FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & IMFFiltro
    Case "Pedido cliente"
        IMFFiltro = "NFP.pccliente" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroEmpresa
        FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & IMFFiltro
    Case "Nosso número"
        IMFFiltro = "DR.Nosso_numero" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroEmpresa
        FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Detalhes_Recebimento DR ON NF.ID = DR.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where " & IMFFiltro
    Case "Pedido interno/Pedido de compra":
        FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota) INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where PN.proposta" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
End Select
If txtTexto <> "" Then
    If cmbfiltrarpor = "Nota fiscal" Or cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Nosso número" Or cmbfiltrarpor = "Pedido interno/Pedido de compra" Then
        StrSqlLocProdPadrao = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Int_status = 1 and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroEmpresa & TextoFiltroAplic & " group by " & CamposFiltro & OrdenarFiltro
    Else
        StrSqlLocProdPadrao = FiltroPadrao & " and NF.tiponf = '" & TipoNF & "' and NF.Int_status = 1 " & TipoFiltro & TextoFiltroAplic & " group by " & CamposFiltro & OrdenarFiltro
    End If
Else
    StrSqlLocProdPadrao = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.tiponf = '" & TipoNF & "' and NF.Int_status = 1 " & TextoFiltroEmpresa & TextoFiltroAplic & " group by " & CamposFiltro & OrdenarFiltro
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_produto_padrao.PageCount - 1)
    Else
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_produto_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn: ListView1_DblClick
    Case vbKeyF2: ProcLocalizar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Minuta = False Then Caption = "Administrativo - Faturamento - Carta de correção - Localizar nota"
If Faturamento = True Then Caption = "Administrativo - Faturamento - Nota fiscal - Relacionamento de NFe - Localizar nota"
ProcCarregaToolBar1 Me, 8805, 5, True
TipoNF = "M1"
cmbfiltrarpor = "Nota fiscal"

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

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
If Minuta = True Then
    With frmMinuta
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Minuta_notas where id_minuta = " & .txtID & " and id_nota = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then
            TBGravar.AddNew
            TBGravar!id_minuta = .txtID
            TBGravar!ID_nota = ListView1.SelectedItem
            TBGravar.Update
            '==================================
            Modulo = "Faturamento/Minuta de despacho"
            Evento = "Nova nota"
            ID_documento = TBGravar!ID
            Documento = "Número da minuta: " & txtID
            Documento1 = "Nº nota: " & ListView1.SelectedItem.ListSubItems(2) & " - Tipo: " & ListView1.SelectedItem.ListSubItems(3) & " - Série: " & ListView1.SelectedItem.ListSubItems(4)
            ProcGravaEvento
            '==================================
            Conexao.Execute "Update tbl_Detalhes_Nota Set ID_Minuta = " & .txtID & " where ID_Nota = " & ListView1.SelectedItem
        Else
            USMsgBox ("A nota " & ListView1.SelectedItem.SubItems(2) & " já foi adicionada na minuta " & .txtID & "."), vbExclamation, "CAPRIND v5.0"
            TBGravar.Close
            Exit Sub
        End If
        TBGravar.Close
        .ProcCarregalista_nota
    End With
Else
    With frmFaturamento_CartaCorrecao_NS
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Txt_ID_nota = TBAbrir!ID
            .txtNF = IIf(IsNull(TBAbrir!int_NotaFiscal), "", TBAbrir!int_NotaFiscal)
            .Txt_serie = IIf(IsNull(TBAbrir!Serie), "", TBAbrir!Serie)
            .txtiddestinatario = IIf(IsNull(TBAbrir!Id_Int_Cliente), "", TBAbrir!Id_Int_Cliente)
            .Txt_destinatario = IIf(IsNull(TBAbrir!txt_Razao_Nome), "", TBAbrir!txt_Razao_Nome)
            Set TBNivel1 = CreateObject("adodb.recordset")
            TBNivel1.Open "Select Chave_acesso from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockReadOnly
            If TBNivel1.EOF = False Then .Txt_chave_acesso = IIf(IsNull(TBNivel1!Chave_acesso), "", TBNivel1!Chave_acesso)
            TBNivel1.Close
            .procCarregaEmpresa
        End If
'=============================================================================
' Busca o ultimo numero sequencial da carta de correção
'=============================================================================
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from NF_Carta_Correcao where ID_Nota = " & .Txt_ID_nota, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            .txtSeq.Text = IIf(IsNull(TBAbrir!Seq), "1", TBAbrir!Seq) + 1
        Else
            .txtSeq.Text = "1"
        End If
        TBAbrir.Close
        Unload Me
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_produto_padrao!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!dt_DataEmissao), "", Format(TBLocalizar_produto_padrao!dt_DataEmissao, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!int_NotaFiscal), "", TBLocalizar_produto_padrao!int_NotaFiscal)
        If IsNull(TBLocalizar_produto_padrao!TipoNF) = False Then
            If TBLocalizar_produto_padrao!TipoNF = "M1" Then Tipo = "Produto(s)"
            If TBLocalizar_produto_padrao!TipoNF = "SA" Then Tipo = "Serviço(s)"
            If TBLocalizar_produto_padrao!TipoNF = "M1SA" Then Tipo = "Prod./Serv."
        End If
        .Item(.Count).SubItems(3) = IIf(Tipo = "", "", Tipo)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!Serie), "", TBLocalizar_produto_padrao!Serie)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!dbl_Valor_Total_Nota), "", Format(TBLocalizar_produto_padrao!dbl_Valor_Total_Nota, "###,##0.00"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!txt_Razao_Nome), "", TBLocalizar_produto_padrao!txt_Razao_Nome)
        If TBLocalizar_produto_padrao!Int_status = 1 Then .Item(.Count).SubItems(7) = "Ativa" Else .Item(.Count).SubItems(7) = "Cancelada"
    End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    contador = contador + 1
    PBLista.Value = contador
Loop
lblRegistros.Caption = "Nº de reg.: " & ContadorReg
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optProduto_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TipoNF = "M1"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optProduto_servico_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TipoNF = "M1SA"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptServico_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TipoNF = "SA"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change()
On Error GoTo tratar_erro

If txtNreg <> "" Then
    VerifNumero = txtNreg
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg = ""
        txtNreg.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change()
On Error GoTo tratar_erro

If txtPagIr <> "" Then
    VerifNumero = txtPagIr
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr = ""
        txtPagIr.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcLocalizar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

