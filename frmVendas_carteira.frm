VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_carteira 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Carteira de vendas"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   1470
   ClientWidth     =   15270
   ClipControls    =   0   'False
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
   Icon            =   "frmVendas_carteira.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   60
      TabIndex        =   54
      Top             =   8790
      Width           =   15195
      Begin VB.ComboBox Cmb_opcao_lista 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmVendas_carteira.frx":000C
         Left            =   6960
         List            =   "frmVendas_carteira.frx":001C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   187
         Width           =   1965
      End
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
         Left            =   2730
         TabIndex        =   21
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
         Left            =   9540
         TabIndex        =   23
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   27
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_carteira.frx":005C
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   11220
         TabIndex        =   26
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_carteira.frx":3800
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   10110
         TabIndex        =   24
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   10680
         TabIndex        =   25
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_carteira.frx":7309
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   12300
         TabIndex        =   28
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_carteira.frx":B3F8
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
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   65
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5610
         TabIndex        =   61
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   60
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   55
         Top             =   240
         Width           =   1275
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   53
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1720
      ButtonCount     =   11
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
      ButtonCaption2  =   "Relatório"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Relatório (F5)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   51
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Faturar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Marcar como faturado (F8)"
      ButtonKey3      =   "5"
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
      ButtonLeft3     =   93
      ButtonTop3      =   2
      ButtonWidth3    =   44
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Cancelar fat."
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Cancelar faturamento (F9)"
      ButtonKey4      =   "6"
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
      ButtonLeft4     =   139
      ButtonTop4      =   2
      ButtonWidth4    =   71
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Expedir"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Marcar como expedido (F10)"
      ButtonKey5      =   "7"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   212
      ButtonTop5      =   2
      ButtonWidth5    =   44
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Cancelar exp."
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Cancelar expedição (F11)"
      ButtonKey6      =   "8"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   258
      ButtonTop6      =   2
      ButtonWidth6    =   75
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Atualizar"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Utilizado pelo administrador do sistema."
      ButtonKey7      =   "9"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   335
      ButtonTop7      =   2
      ButtonWidth7    =   50
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonAlignment8=   2
      ButtonType8     =   1
      ButtonStyle8    =   -1
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState8    =   -1
      ButtonLeft8     =   387
      ButtonTop8      =   4
      ButtonWidth8    =   2
      ButtonHeight8   =   54
      ButtonCaption9  =   "Ajuda"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Ajuda (F1)"
      ButtonKey9      =   "11"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft9     =   391
      ButtonTop9      =   2
      ButtonWidth9    =   36
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Sair"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Sair (Esc)"
      ButtonKey10     =   "12"
      ButtonAlignment10=   2
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft10    =   429
      ButtonTop10     =   2
      ButtonWidth10   =   26
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonKey11     =   "13"
      ButtonAlignment11=   2
      BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState11   =   5
      ButtonLeft11    =   457
      ButtonTop11     =   2
      ButtonWidth11   =   24
      ButtonHeight11  =   24
      ButtonUseMaskColor11=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   12210
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_carteira.frx":EC84
         Count           =   1
      End
   End
   Begin VB.CheckBox chkperiodo 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   14190
      TabIndex        =   17
      Top             =   960
      Width           =   195
   End
   Begin MSComctlLib.ListView listaitens 
      Height          =   4605
      Left            =   60
      TabIndex        =   20
      Top             =   4170
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   8123
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   27
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Ped. cliente"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Cód. referência"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Un. com."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Vlr. unitário"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Object.Tag             =   "D"
         Text            =   "Prazo final"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Tag             =   "N"
         Text            =   "Qt. lib. p/ faturar"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "N"
         Text            =   "Qtde. faturada"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Tag             =   "N"
         Text            =   "Qtde. à faturar"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "N"
         Text            =   "Saldo pedido"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Tag             =   "N"
         Text            =   "Qtde. exped."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Object.Tag             =   "D"
         Text            =   "Dt. venda"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   23
         Object.Tag             =   "D"
         Text            =   "Dt. produção"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   24
         Object.Tag             =   "D"
         Text            =   "Dt. faturam."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   25
         Object.Tag             =   "D"
         Text            =   "Dt. exped."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Object.Tag             =   "D"
         Text            =   "Dt. cancel."
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Frame framdados 
      BackColor       =   &H00E0E0E0&
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
      Height          =   945
      Left            =   55
      TabIndex        =   37
      Top             =   2670
      Width           =   15195
      Begin VB.TextBox Txt_un_com 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Unidade comercial."
         Top             =   450
         Width           =   390
      End
      Begin VB.TextBox Txt_qtde_faturar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   6045
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade à faturar."
         Top             =   450
         Width           =   870
      End
      Begin VB.TextBox Txt_saldo_pedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   6930
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Saldo do pedido."
         Top             =   450
         Width           =   750
      End
      Begin VB.TextBox Txt_grupo_cliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   8610
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Grupo do cliente."
         Top             =   450
         Width           =   2640
      End
      Begin VB.CheckBox Chk_grupo_cliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grupo do cliente"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9300
         TabIndex        =   10
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox Txtdataprod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   12255
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Data de faturamento."
         Top             =   450
         Width           =   930
      End
      Begin VB.TextBox cmbcodref 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código de referência."
         Top             =   450
         Width           =   1470
      End
      Begin VB.TextBox txtdatavendas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Data da venda."
         Top             =   450
         Width           =   955
      End
      Begin VB.TextBox txtprazofinal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3270
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Prazo final."
         Top             =   450
         Width           =   925
      End
      Begin VB.TextBox txtdatafaturado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   13200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Data de faturamento."
         Top             =   450
         Width           =   930
      End
      Begin VB.TextBox txtdataexp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   14160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Data da expedição."
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox Txtqtde_faturada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5130
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade faturada."
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox Txt_qtde_liberada_faturar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4215
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade liberada para faturar."
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox Txt_qtde_expedida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   7695
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade expedida."
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox txtun 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1635
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   450
         Width           =   390
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2445
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   450
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "com."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2115
         TabIndex        =   64
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Á faturar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6150
         TabIndex        =   63
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7065
         TabIndex        =   62
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Produz. em"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12345
         TabIndex        =   50
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Exped. em"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   14205
         TabIndex        =   49
         Top             =   255
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Faturada"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5265
         TabIndex        =   48
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Expedido"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7815
         TabIndex        =   47
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Lib. faturar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4275
         TabIndex        =   46
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Código referência"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   255
         TabIndex        =   43
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Fatur. em"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13305
         TabIndex        =   42
         Top             =   255
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendido em"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11355
         TabIndex        =   41
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo final"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   40
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Un."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1710
         TabIndex        =   39
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2640
         TabIndex        =   38
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar no período     "
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
      Height          =   1665
      Left            =   12510
      TabIndex        =   36
      Top             =   990
      Width           =   2745
      Begin VB.ComboBox cmbTipoData 
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
         ItemData        =   "frmVendas_carteira.frx":144A1
         Left            =   180
         List            =   "frmVendas_carteira.frx":144A3
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   78
         ToolTipText     =   "Filtrar pela data."
         Top             =   495
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker txtinicio 
         Height          =   315
         Left            =   180
         TabIndex        =   18
         ToolTipText     =   "Data início para pesquisa."
         Top             =   1095
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   198639619
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txtfinal 
         Height          =   315
         Left            =   1410
         TabIndex        =   19
         ToolTipText     =   "Data final para pesquisa."
         Top             =   1095
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   198639617
         CurrentDate     =   39057
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1035
         TabIndex        =   79
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1890
         TabIndex        =   45
         Top             =   900
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   44
         Top             =   900
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para filtro"
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
      Height          =   1665
      Left            =   55
      TabIndex        =   35
      Top             =   990
      Width           =   12435
      Begin VB.CheckBox Chk_retorno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar retorno"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   10950
         TabIndex        =   80
         Top             =   0
         Width           =   1305
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3540
         MaxLength       =   50
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   1110
         Width           =   1380
      End
      Begin VB.TextBox txtdescricao 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   75
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   1110
         Width           =   6945
      End
      Begin VB.TextBox txtcliente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "Cliente."
         Top             =   480
         Width           =   5505
      End
      Begin VB.CheckBox chkCodInterno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código interno"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3570
         TabIndex        =   73
         Top             =   900
         Width           =   1605
      End
      Begin VB.CheckBox chkDescricao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5520
         TabIndex        =   72
         Top             =   900
         Width           =   1065
      End
      Begin VB.CheckBox chkCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6975
         TabIndex        =   71
         Top             =   270
         Width           =   1365
      End
      Begin VB.CheckBox chkPedido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pedido do cliente"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5100
         TabIndex        =   70
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4980
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Pedido do cliente."
         Top             =   480
         Width           =   1730
      End
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmVendas_carteira.frx":144A5
         Left            =   150
         List            =   "frmVendas_carteira.frx":144A7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   68
         ToolTipText     =   "Empresa."
         Top             =   480
         Width           =   3390
      End
      Begin VB.TextBox txtPI 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   3540
         MaxLength       =   50
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   480
         Width           =   1430
      End
      Begin VB.CheckBox chkPI 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pedido interno"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3570
         TabIndex        =   66
         Top             =   270
         Width           =   1395
      End
      Begin VB.ComboBox Cmb_status 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmVendas_carteira.frx":144A9
         Left            =   150
         List            =   "frmVendas_carteira.frx":144D1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Status."
         Top             =   1110
         Width           =   3375
      End
      Begin DrawSuite2022.USButton cmdLocaliza_produto 
         Height          =   315
         Left            =   4950
         TabIndex        =   81
         ToolTipText     =   "Localizar produto"
         Top             =   1110
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmVendas_carteira.frx":1458F
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
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         Theme           =   4
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   1485
         TabIndex        =   77
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1605
         TabIndex        =   51
         Top             =   900
         Width           =   465
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00E0E0E0&
      Height          =   825
      Left            =   55
      TabIndex        =   52
      Top             =   3630
      Width           =   15195
      Begin VB.OptionButton Opt_todos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox cmbAno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmVendas_carteira.frx":32694
         Left            =   14310
         List            =   "frmVendas_carteira.frx":32696
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton OptDomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do mês"
         Height          =   195
         Left            =   930
         TabIndex        =   30
         Top             =   240
         Width           =   825
      End
      Begin VB.OptionButton OptAteomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Até o mês"
         Height          =   195
         Left            =   1815
         TabIndex        =   31
         Top             =   240
         Width           =   1035
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   315
         Left            =   3000
         TabIndex        =   33
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   556
         MultiRow        =   -1  'True
         TabMinWidth     =   1649
         TabStyle        =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   12
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jan"
               Key             =   "Jan"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fev"
               Key             =   "Fev"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mar"
               Key             =   "Mar"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Abril"
               Key             =   "Abr"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Maio"
               Key             =   "Maio"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jun"
               Key             =   "Jun"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jul"
               Key             =   "Jul"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ago"
               Key             =   "Ago"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Set"
               Key             =   "Set"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Out"
               Key             =   "Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Nov"
               Key             =   "Nov"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dez"
               Key             =   "Dez"
               ImageVarType    =   2
            EndProperty
         EndProperty
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   57
      Top             =   9390
      Width           =   15195
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   13125
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   210
         Width           =   1890
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   58
         Top             =   240
         Width           =   11685
         _ExtentX        =   20611
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total :"
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
         Left            =   12030
         TabIndex        =   59
         Top             =   210
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmVendas_carteira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Vendas_Carteira As String 'OK
Dim StrSql_Vendas_Carteira_Total As String 'OK
Public FormulaRel_Vendas_Carteira As String 'OK
Dim TBLISTA_Vendas_Carteira As ADODB.Recordset 'OK
Dim Descricao As String 'OK
Dim PedidoCliente As String 'OK
Dim status As String 'OK
Dim DataFiltro As String 'OK
Dim LimparCampo As Boolean 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=CFZ4jrjc9II&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=56&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_grupo_cliente_Click()
On Error GoTo tratar_erro

With Txt_grupo_cliente
    If Chk_grupo_cliente.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        chkCliente.Value = 0
        txtCliente = ""
        txtCliente.Locked = True
        txtCliente.TabStop = False
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With
ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chkcliente_Click()
On Error GoTo tratar_erro

With txtCliente
    If chkCliente.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_grupo_cliente.Value = 0
        Txt_grupo_cliente = ""
        Txt_grupo_cliente.Locked = True
        Txt_grupo_cliente.TabStop = False
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With
ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkCodInterno_Click()
On Error GoTo tratar_erro

With txtCodigo
    If chkCodInterno.Value = 1 Then
        chkDescricao.Value = 0
        .Locked = False
        .TabStop = True
        .SetFocus
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With
ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkDescricao_Click()
On Error GoTo tratar_erro

With txtdescricao
    If chkDescricao.Value = 1 Then
        chkCodInterno.Value = 0
        .Locked = False
        .TabStop = True
        .SetFocus
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With
ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPedido_Click()
On Error GoTo tratar_erro

With txtPedido
    If chkPedido.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With listaitens
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    Select Case Cmb_opcao_lista
        Case "Faturar":
            .ButtonState(3) = 0
            .ButtonState(4) = 5
            .ButtonState(5) = 5
            .ButtonState(6) = 5
        Case "Cancelar faturamento":
            .ButtonState(3) = 5
            .ButtonState(4) = 0
            .ButtonState(5) = 5
            .ButtonState(6) = 5
        Case "Expedir":
            .ButtonState(3) = 5
            .ButtonState(4) = 5
            .ButtonState(5) = 0
            .ButtonState(6) = 5
        Case "Cancelar expedição":
            .ButtonState(3) = 5
            .ButtonState(4) = 5
            .ButtonState(5) = 5
            .ButtonState(6) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_status_Click()
On Error GoTo tratar_erro

ProcLimparCampos
ProcCarregaComboData
Select Case Cmb_status
    Case "Emitidas": chkPedido.Enabled = True
    Case "Aberta em Analise":
        chkPedido.Value = 0
        chkPedido.Enabled = False
    Case "A Faturar": chkPedido.Enabled = True
    Case "Faturadas": chkPedido.Enabled = True
    Case "Vendidas e Faturadas": chkPedido.Enabled = True
    Case "Vendidas e Faturadas parcial": chkPedido.Enabled = True
    Case "Vendidas": chkPedido.Enabled = True
    Case "Canceladas":
        chkPedido.Value = 0
        chkPedido.Enabled = False
    Case "Perdidas por preço":
        chkPedido.Value = 0
        chkPedido.Enabled = False
    Case "Perdidas por prazo":
        chkPedido.Value = 0
        chkPedido.Enabled = False
    Case "Portal eletrônico":
        chkPedido.Value = 0
        chkPedido.Enabled = False
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTipoData_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocaliza_produto_Click()
On Error GoTo tratar_erro

If chkCodInterno.Value = 0 Then Exit Sub
cmbcodref = ""
chkDescricao.Value = 0
frmVendas_carteira_locprod.Show 1
framdados.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkperiodo_Click()
On Error GoTo tratar_erro
  
If chkPeriodo.Value = 1 Then
    Frame2.Enabled = True
    ProcCarregaComboData
    cmbTipoData.SetFocus
Else
    Frame2.Enabled = False
    txtinicio.Value = Date
    txtFinal.Value = Date
End If
ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFaturar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False

With listaitens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente faturar este(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    If USMsgBox("Algum produto/serviço selecionado será faturado com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
                    GoTo 1
                Else
                    Exit Sub
                End If
            End If
1:
            Permitido = True
            If Permitido1 = True Then
                Compras_Pedido = False
                Vendas_PI = False
                Compras_Cotacao = False
                Faturamento = False
                Sit_REG = 3
                IDlista = .ListItems.Item(InitFor)
                Permitido2 = True
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then Exit Sub
            Else
                Set TBVendas = CreateObject("adodb.recordset")
                TBVendas.Open "Select * from vendas_carteira  where Codigo = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBVendas.EOF = False Then
                
                
                '============================================================================================
                If USMsgBox("Existe nota gerada para esse item?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                
                strNF = InputBox("Informe o numero da nota fiscal", "CAPRIND v5.0")
                
                Set TBAbrir_NFe = CreateObject("adodb.recordset")
                If strNF = "" Then Exit Sub
                
                TBAbrir_NFe.Open "Select * from tbl_Dados_Nota_Fiscal  where int_NotaFiscal = '" & strNF & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir_NFe.EOF = False Then
                ID_nota = TBAbrir_NFe!ID
                Else
                USMsgBox "Nota fiscal não encontrada, por favor verifique a numeração correta e tente de novo.", vbInformation, "CAPRIND v5.0"
                Exit Sub
                End If
                TBAbrir_NFe.Close
                
                If ID_nota = 0 Then Exit Sub
                
                Set TBAbrir_NFe = CreateObject("adodb.recordset")
                TBAbrir_NFe.Open "Select * from tbl_proposta_nota  where NF = " & strNF & " and ID_Nota = '" & ID_nota & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir_NFe.EOF = False Then
                TBAbrir_NFe!Proposta = listaitens.ListItems.Item(InitFor).SubItems(2)
                TBAbrir_NFe!NF = strNF
                TBAbrir_NFe!ID_nota = ID_nota
                TBAbrir_NFe.Update
                End If
                TBAbrir_NFe.Close
                
                'Conexao.Execute "INSERT INTO tbl_proposta_nota (NF, ID_Nota) VALUES (" & strNF & ",'" & ID_nota & "')"
                'Exit Sub
                End If
                '==================================================================================================
                
                    TBVendas!Liberacao = "FATURADO"
                    TBVendas!qtdeliberada = TBVendas!quantidade
                    TBVendas!QtdeFaturada = TBVendas!quantidade
                    TBVendas!DataFaturamento = Date
                    TBVendas.Update
                                                       
                    If IsNull(TBVendas!ID_programacao) = False And TBVendas!ID_programacao <> "" Then
                        'Programação
                        Set TBProgramas = CreateObject("adodb.recordset")
                        TBProgramas.Open "Select * from Vendas_programacao where ID_prog = " & TBVendas!ID_programacao, Conexao, adOpenKeyset, adLockOptimistic
                        If TBProgramas.EOF = False Then
                            TBProgramas!QtdeFaturada = Format(TBVendas!QtdeFaturada, "###,##0.00")
                            If TBProgramas!QtdeFaturada >= TBProgramas!quantidade Then
                                TBProgramas!Status_prog = "FATURADO"
                                TBProgramas!Ordenar = 4
                            Else
                                TBProgramas!Status_prog = "PARCIAL"
                                TBProgramas!Ordenar = 1
                            End If
                            TBProgramas.Update
                        
                            Set TBItem = CreateObject("adodb.recordset")
                            TBItem.Open "Select * from vendas_programa_item where ID_item = " & TBProgramas!Id_Item, Conexao, adOpenKeyset, adLockOptimistic
                            If TBItem.EOF = False Then
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = True Then
                                    TBItem!Status_Item = "PREVISÃO FUTURA"
                                Else
                                    Set TBAbrir = CreateObject("adodb.recordset")
                                    TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBAbrir.EOF = True Then
                                        TBItem!Status_Item = "ABERTO"
                                    Else
                                        Set TBAbrir = CreateObject("adodb.recordset")
                                        TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
                                        If TBAbrir.EOF = True Then
                                            TBItem!Status_Item = "FATURADO"
                                        Else
                                            TBItem!Status_Item = "PARCIAL"
                                        End If
                                    End If
                                End If
                                TBAbrir.Close
                                TBItem.Update
                            End If
                        End If
                        TBProgramas.Close
                    End If
                    
                    FunAtualizaStatusPropPI TBVendas!Cotacao
                                                       
                    '==================================
                    Modulo = "Vendas/Follow up"
                    Evento = "Faturar produto/serviço"
                    ID_documento = .ListItems.Item(InitFor)
                    Documento = "Nº pedido: " & .ListItems.Item(InitFor).SubItems(2) & " - Rev.: " & .ListItems.Item(InitFor).SubItems(3)
                    Documento1 = "Cód. interno: " & .ListItems.Item(InitFor).SubItems(6)
                    ProcGravaEvento
                    '==================================
                End If
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de faturar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) faturado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelarFaturamento()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With listaitens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente cancelar o faturamento deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBVendas = CreateObject("adodb.recordset")
            TBVendas.Open "Select * from vendas_carteira where Codigo = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBVendas.EOF = False Then
                TBVendas!Liberacao = "VENDIDA"
                TBVendas!DataFaturamento = Null
                TBVendas!qtdeliberada = 0
                TBVendas!QtdeFaturada = 0
                TBVendas.Update
                
                If IsNull(TBVendas!ID_programacao) = False And TBVendas!ID_programacao <> "0" Then
                    'Programação
                    Set TBProgramas = CreateObject("adodb.recordset")
                    TBProgramas.Open "Select * from Vendas_programacao where ID_prog = " & TBVendas!ID_programacao, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProgramas.EOF = False Then
                        TBProgramas!QtdeFaturada = 0
                        TBProgramas!Status_prog = "ABERTO"
                        TBProgramas!Ordenar = 2
                        TBProgramas.Update
                    
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "Select * from vendas_programa_item where ID_item = " & TBProgramas!Id_Item, Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = True Then
                                TBItem!Status_Item = "PREVISÃO FUTURA"
                            Else
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = True Then
                                    TBItem!Status_Item = "ABERTO"
                                Else
                                    Set TBAbrir = CreateObject("adodb.recordset")
                                    TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBAbrir.EOF = True Then
                                        TBItem!Status_Item = "FATURADO"
                                    Else
                                        TBItem!Status_Item = "PARCIAL"
                                    End If
                                End If
                            End If
                            TBAbrir.Close
                            TBItem.Update
                        End If
                    End If
                    TBProgramas.Close
                    
                    'Programa
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select vendas_programa.ID, vendas_programa.Status from (vendas_programa INNER JOIN vendas_proposta ON vendas_programa.ID = vendas_proposta.ID_programa) INNER JOIN vendas_carteira ON vendas_carteira.Cotacao = vendas_proposta.Cotacao where vendas_carteira.Codigo = " & TBVendas!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        Do While TBItem.EOF = False
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from vendas_programa_item where id = " & TBItem!ID & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = True Then
                                TBItem!status = "PREVISÃO FUTURA"
                            Else
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "Select * from vendas_programa_item where id = " & TBItem!ID & " and Status_Item <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = True Then
                                    TBItem!status = "ABERTO"
                                Else
                                    Set TBAbrir = CreateObject("adodb.recordset")
                                    TBAbrir.Open "Select * from vendas_programa_item where id = " & TBItem!ID & " and Status_Item <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBAbrir.EOF = True Then
                                        TBItem!status = "FATURADO"
                                    Else
                                        TBItem!status = "PARCIAL"
                                    End If
                                End If
                            End If
                            TBAbrir.Close
                            TBItem.Update
                            TBItem.MoveNext
                        Loop
                    End If
                    TBItem.Close
                End If
                
                FunAtualizaStatusPropPI TBVendas!Cotacao
            End If
            '==================================
            Modulo = "Vendas/Follow up"
            Evento = "Cancelar faturamento do produto/serviço"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº pedido: " & .ListItems.Item(InitFor).SubItems(2) & " - Rev.: " & .ListItems.Item(InitFor).SubItems(3)
            Documento1 = "Cód. interno: " & .ListItems.Item(InitFor).SubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviços antes de cancelar o faturamento."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Faturamento cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExpedir()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With listaitens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente expedir este(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    If USMsgBox("Algum produto/serviço selecionado será expedido com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
                    GoTo 1
                Else
                    Exit Sub
                End If
            End If
1:
            Permitido = True
            If Permitido1 = True Then
                Compras_Pedido = False
                Vendas_PI = False
                Compras_Cotacao = False
                Faturamento = False
                Sit_REG = 4
                IDlista = .ListItems.Item(InitFor)
                Permitido2 = True
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then Exit Sub
            Else
                Set TBVendas = CreateObject("adodb.recordset")
                TBVendas.Open "Select * from vendas_carteira  where Codigo = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBVendas.EOF = False Then
                    TBVendas!qtdeexpedida = TBVendas!quantidade
                    TBVendas.Update
                    '==================================
                    Modulo = "Vendas/Follow up"
                    Evento = "Expedir produto/serviço"
                    ID_documento = .ListItems.Item(InitFor)
                    Documento = "Nº pedido: " & .ListItems.Item(InitFor).SubItems(2) & " - Rev.: " & .ListItems.Item(InitFor).SubItems(3)
                    Documento1 = "Cód. interno: " & .ListItems.Item(InitFor).SubItems(6)
                    ProcGravaEvento
                    '==================================
                End If
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de expedir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) expedido(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelarExpedicao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With listaitens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente cancelar a expedição deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "Update vendas_carteira Set qtdeexpedida = 0 where Codigo = " & .ListItems.Item(InitFor)
            '==================================
            Modulo = "Vendas/Follow up"
            Evento = "Cancelar expedição do produto/serviço"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº pedido: " & .ListItems.Item(InitFor).SubItems(2) & " - Rev.: " & .ListItems.Item(InitFor).SubItems(3)
            Documento1 = "Cód. interno: " & .ListItems.Item(InitFor).SubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviços antes de cancelar a expedição."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Expedição cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaStatusPrograma()
On Error GoTo tratar_erro

Set TBProposta = CreateObject("adodb.recordset")
TBProposta.Open "Select * from vendas_proposta where Cotacao = " & TBVendas!Cotacao, Conexao, adOpenKeyset, adLockOptimistic
If TBProposta.EOF = False Then
    Set TBCarteira = CreateObject("adodb.recordset")
    TBCarteira.Open "Select * from vendas_carteira where cotacao = " & TBProposta!Cotacao & " and Liberacao <> 'FATURAR' and Liberacao <> 'FATURAR PARCIAL' and Liberacao <> 'VENDIDA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarteira.EOF = True Then
        TBProposta!status = "VENDIDA"
    Else
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select * from vendas_carteira where cotacao = " & TBProposta!Cotacao & " and Liberacao <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = True Then
            TBProposta!status = "FATURADA"
        Else
            TBProposta!status = "FATURADA PARCIAL"
        End If
    End If
    TBCarteira.Close
    TBProposta.Update
End If
TBProposta.Close



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If chkCodInterno.Value = 1 And txtCodigo = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodigo.SetFocus
    Exit Sub
End If
If chkDescricao.Value = 1 And txtdescricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
If chkCliente.Value = 1 And txtCliente = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    txtCliente.SetFocus
    Exit Sub
End If
If chkPedido.Value = 1 And txtPedido = "" Then
    NomeCampo = "o pedido do cliente"
    ProcVerificaAcao
    txtPedido.SetFocus
    Exit Sub
End If
With txtFinal
    If FunVerificaDataFinal(txtinicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

StrSql_Vendas_Carteira = ""
StrSql_Vendas_Carteira_Total = ""

Descricao = "and FV.Desenho is not null"
DescricaoRel = "and {Follow_up_vendas.Desenho} <> 'Null'"
PedidoCliente = "and FV.Desenho is not null"
PedidoClienteRel = "and {Follow_up_vendas.Desenho} <> 'Null'"
Cliente = "and FV.Desenho is not null"
ClienteRel = "and {Follow_up_vendas.Desenho} <> 'Null'"
GrupoCliente = "and FV.Desenho is not null"
GrupoClienteRel = "and {Follow_up_vendas.Desenho} <> 'Null'"

If chkPI.Value = 1 Then
    Ncotacao = "FV.Ncotacao like '" & txtPI & "%'"
    NcotacaoRel = "{Follow_up_vendas.Ncotacao} like '" & txtPI & "*'"
Else
    Ncotacao = "FV.Ncotacao is not null"
    NcotacaoRel = "{Follow_up_vendas.Ncotacao} <> 'Null'"
End If


If chkCodInterno.Value = 1 Then
    Desenho = "FV.Desenho like '" & txtCodigo & "%'"
    DesenhoRel = "{Follow_up_vendas.Desenho} like '" & txtCodigo & "*'"
Else
    Desenho = "FV.Desenho is not null"
    DesenhoRel = "{Follow_up_vendas.Desenho} <> 'Null'"
End If
If chkDescricao.Value = 1 Then
    Descricao = "and FV.Descricao like '" & txtdescricao & "%'"
    DescricaoRel = "and {Follow_up_vendas.Descricao} like '" & txtdescricao & "*'"
End If
If chkPedido.Value = 1 Then
    PedidoCliente = "and FV.PCcliente like '" & txtPedido & "%'"
    PedidoClienteRel = "and {Follow_up_vendas.PCcliente} like '" & txtPedido & "*'"
End If
If chkCliente.Value = 1 Then
    Cliente = "and FV.cliente like '" & txtCliente & "%'"
    ClienteRel = "and {Follow_up_vendas.cliente} like '" & txtCliente & "*'"
End If
If Chk_grupo_cliente.Value = 1 Then
    GrupoCliente = "and FV.Grupo_cliente like '" & txtCliente & "%'"
    GrupoClienteRel = "and {Follow_up_vendas.Grupo_cliente} like '" & txtCliente & "*'"
End If

status = ""
StatusRel = ""
Select Case Cmb_status
    Case "":
        status = "and FV.Liberacao is not null"
        StatusRel = "and {Follow_up_vendas.Liberacao} <> 'Null'"
    Case "Emitidas":
        status = "and FV.Liberacao is not null"
        StatusRel = "and {Follow_up_vendas.Liberacao} <> 'Null'"
    Case "Aberta em Analise":
        status = "and FV.Liberacao = 'ABERTA EM ANALISE'"
        StatusRel = "and {Follow_up_vendas.Liberacao} = 'ABERTA EM ANALISE'"
    Case "A Faturar":
        status = "and (FV.liberacao = 'FATURAR' or FV.liberacao = 'FATURAR PARCIAL')"
        StatusRel = "and ({Follow_up_vendas.liberacao} = 'FATURAR' or {Follow_up_vendas.liberacao} = 'FATURAR PARCIAL')"
    Case "Faturadas":
        status = "and (FV.liberacao = 'FATURADO' or FV.liberacao = 'FATURADO PARCIAL')"
        StatusRel = "and ({Follow_up_vendas.liberacao} = 'FATURADO' or {Follow_up_vendas.liberacao} = 'FATURADO PARCIAL')"
    Case "Vendidas e Faturadas":
        status = "and (FV.liberacao = 'VENDIDA' or FV.liberacao = 'VENDIDA PARCIAL' or FV.liberacao = 'FATURAR' or FV.liberacao = 'FATURAR PARCIAL' or FV.liberacao = 'FATURADO' or FV.liberacao = 'FATURADO PARCIAL')"
        StatusRel = "and ({Follow_up_vendas.liberacao} = 'VENDIDA' or {Follow_up_vendas.liberacao} = 'VENDIDA PARCIAL' or {Follow_up_vendas.liberacao} = 'FATURAR' or {Follow_up_vendas.liberacao} = 'FATURAR PARCIAL' or {Follow_up_vendas.liberacao} = 'FATURADO' or {Follow_up_vendas.liberacao} = 'FATURADO PARCIAL')"
    Case "Vendidas e Faturadas parcial":
        status = "and (FV.liberacao = 'VENDIDA' or FV.liberacao = 'VENDIDA PARCIAL' or FV.liberacao = 'FATURAR' or FV.liberacao = 'FATURAR PARCIAL' or FV.liberacao = 'FATURADO PARCIAL')"
        StatusRel = "and ({Follow_up_vendas.liberacao} = 'VENDIDA' or {Follow_up_vendas.liberacao} = 'VENDIDA PARCIAL' or {Follow_up_vendas.liberacao} = 'FATURAR' or {Follow_up_vendas.liberacao} = 'FATURAR PARCIAL' or {Follow_up_vendas.liberacao} = 'FATURADO PARCIAL')"
    Case "Vendidas":
        status = "and (FV.liberacao = 'VENDIDA' or FV.liberacao = 'VENDIDA PARCIAL' or FV.liberacao = 'FATURAR' or FV.liberacao = 'FATURAR PARCIAL')"
        StatusRel = "and ({Follow_up_vendas.liberacao} = 'VENDIDA' or {Follow_up_vendas.liberacao} = 'VENDIDA PARCIAL' or {Follow_up_vendas.liberacao} = 'FATURAR' or {Follow_up_vendas.liberacao} = 'FATURAR PARCIAL')"
    Case "Canceladas":
        status = "and FV.liberacao = 'CANCELADA'"
        StatusRel = "and {Follow_up_vendas.Liberacao} = 'CANCELADA'"
    Case "Perdidas por preço":
        status = "and FV.liberacao = 'PERDIDO P/ PREÇO'"
        StatusRel = "and {Follow_up_vendas.Liberacao} = 'PERDIDO P/ PREÇO'"
    Case "Perdidas por prazo":
        status = "and FV.liberacao = 'PERDIDO P/ PRAZO'"
        StatusRel = "and {Follow_up_vendas.Liberacao} = 'PERDIDO P/ PRAZO'"
    Case "Portal eletrônico":
        status = "and FV.liberacao = 'PORTAL ELETRONICO'"
        StatusRel = "and {Follow_up_vendas.Liberacao} = 'PORTAL ELETRONICO'"
End Select

If Chk_retorno.Value = 1 Then
    RetornoFiltro = "and FV.Retorno = 'True'"
    RetornoFiltroRel = "and {Follow_up_vendas.Retorno} = True"
Else
    RetornoFiltro = "and FV.Retorno = 'False'"
    RetornoFiltroRel = "and {Follow_up_vendas.Retorno} = False"
End If

DataFiltro = ""
DataFiltroRel = ""
If Cmb_empresa <> "Todas" Then
    TextoFiltroEmpresa = "and FV.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TextoFiltroEmpresaRel = "and {Follow_up_vendas.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
Else
    TextoFiltroEmpresa = ""
    TextoFiltroEmpresaRel = ""
End If
If cmbTipoData = "Faturamento" Then
    DataFiltroTexto = "NF.dt_DataEmissao"
    If chkPeriodo.Value = 1 Then
        DataFiltro = "and NF.dt_DataEmissao Between '" & txtinicio.Value & "' And '" & txtFinal.Value & "'"
        DataFiltroRel = "and {tbl_Dados_Nota_Fiscal.dt_DataEmissao} >= Date(" & Year(txtinicio.Value) & "," & Month(txtinicio.Value) & "," & Day(txtinicio.Value) & ") and {tbl_Dados_Nota_Fiscal.dt_DataEmissao} <= Date(" & Year(txtFinal.Value) & "," & Month(txtFinal.Value) & "," & Day(txtFinal.Value) & ")"
    Else
        M = FunVerificaMes(TabFiltro.SelectedItem.key)
        If OptDomes.Value = True Then
            DataFiltro = "and month((FV.DataFaturamento)) = '" & M & "' and Year((FV.DataFaturamento)) = '" & cmbAno.Text & "'"
            DataFiltroRel = "and Month ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) = " & M & " and Year ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) = Year(Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & "))"
        ElseIf OptAteomes.Value = True Then
                DataFiltro = "and month((FV.DataFaturamento)) <= '" & M & "' and Year((FV.DataFaturamento)) = '" & cmbAno.Text & "'"
                DataFiltroRel = "and Month ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) <= " & M & " and Year ({tbl_Dados_Nota_Fiscal.dt_DataEmissao}) = Year(Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & "))"
        End If
    End If
Else
    Select Case cmbTipoData
        Case "": DataFiltroTexto = "FV.Data"
        Case "Emissão": DataFiltroTexto = "FV.Data"
        Case "Venda": DataFiltroTexto = "FV.Datavendas"
        Case "Prazo final": DataFiltroTexto = "FV.Prazofinal"
        Case "Cancelamento": DataFiltroTexto = "FV.Dataalteracao"
        Case "Perda": DataFiltroTexto = "FV.Dataalteracao"
    End Select
    DataFiltroTextoRel = Replace(DataFiltroTexto, "FV.", "Follow_up_vendas.")
    If chkPeriodo.Value = 1 Then
        DataFiltro = "and " & DataFiltroTexto & " Between '" & txtinicio.Value & "' And '" & txtFinal.Value & "'"
        DataFiltroRel = "and {" & DataFiltroTextoRel & "} >= Date(" & Year(txtinicio.Value) & "," & Month(txtinicio.Value) & "," & Day(txtinicio.Value) & ") and {" & DataFiltroTextoRel & "} <= Date(" & _
                            Year(txtFinal.Value) & "," & Month(txtFinal.Value) & "," & Day(txtFinal.Value) & ")"
    Else
        M = FunVerificaMes(TabFiltro.SelectedItem.key)
        If OptDomes.Value = True Then
            DataFiltro = "and month(" & DataFiltroTexto & ") = '" & M & "' and Year(" & DataFiltroTexto & ") = '" & cmbAno.Text & "'"
            DataFiltroRel = "and Month ({" & DataFiltroTextoRel & "}) = " & M & " and Year ({" & DataFiltroTextoRel & "}) = Year(Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & "))"
        ElseIf OptAteomes.Value = True Then
                DataFiltro = "and month(" & DataFiltroTexto & ") <= '" & M & "' and Year(" & DataFiltroTexto & ") = '" & cmbAno.Text & "'"
                DataFiltroRel = "and Month ({" & DataFiltroTextoRel & "}) <= " & M & " and Year ({" & DataFiltroTextoRel & "}) = Year(Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & "))"
        End If
    End If
End If

Campos = "FV.Ncotacao, FV.Codigo, FV.Empresa, FV.Cotacao, FV.Ncotacao, FV.Revisao, FV.Cliente, FV.PCcliente, FV.Desenho, FV.Rev_codinterno, FV.N_Referencia, FV.descricao_tecnica, FV.Unidade, FV.Unidade_com, FV.preco_unitario, FV.quantidade, FV.preco_lote, FV.Prazofinal, FV.Liberacao, FV.qtdeliberada, FV.qtdefaturada, FV.qtdefaturar, FV.Saldopedido, FV.qtdeexpedida, FV.Data, FV.Datavendas, FV.datapcp, FV.datafaturamento, FV.dataexpedicao, FV.dataalteracao, FV.Status"
StrSql_Vendas_Carteira = "Select " & Campos & " FROM (Follow_up_vendas FV LEFT JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = FV.Codigo and NFPP.Codinterno = FV.Desenho) LEFT JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFPP.ID_nota where " & DataFiltroTexto & " IS NOT NULL and " & Ncotacao & " AND " & Desenho & " " & Descricao & " " & PedidoCliente & " " & Cliente & " " & GrupoCliente & " " & status & " " & TextoFiltroEmpresa & " " & RetornoFiltro & " " & DataFiltro & " group by " & Campos & " order by FV.Ncotacao, FV.cotacao, FV.Codigo"
StrSql_Vendas_Carteira_Total = "Select Codigo, preco_lote FROM (Follow_up_vendas FV LEFT JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = FV.Codigo and NFPP.Codinterno = FV.Desenho) LEFT JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFPP.ID_nota where " & DataFiltroTexto & " IS NOT NULL and " & Desenho & " " & Descricao & " " & PedidoCliente & " " & Cliente & " " & GrupoCliente & " " & status & " " & TextoFiltroEmpresa & " " & RetornoFiltro & " " & DataFiltro & " group by Codigo, preco_lote order by Codigo"
FormulaRel_Vendas_Carteira = DesenhoRel & " " & DescricaoRel & " " & PedidoClienteRel & " " & ClienteRel & " " & GrupoClienteRel & " " & StatusRel & " " & TextoFiltroEmpresaRel & " " & RetornoFiltroRel & " " & DataFiltroRel
'Debug.print StrSql_Vendas_Carteira

ProcAtualizalista (1)

If Cmb_empresa = "Todas" Then IDempresa = 0 Else IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
If Cmb_status <> "" And cmbTipoData <> "" Then
    Select Case cmbTipoData
        Case "": TipoData = "Dt. de Emissão"
        Case "Emissão": TipoData = "Dt. de Emissão"
        Case "Venda": TipoData = "Dt. da Venda"
        Case "Prazo final": TipoData = "Prazo final"
        Case "Faturamento": TipoData = "Dt. de Faturamento"
        Case "Cancelamento": TipoData = "Dt. de Cancelamento"
        Case "Perda": TipoData = "Dt. da Perda"
    End Select
    Familiatext = Cmb_status & " - " & TipoData
ElseIf Cmb_status <> "" Then
        Familiatext = Cmb_status
End If
If Opt_todos.Value = True Then ProcGravarDataFiltroRel Dataini, DataFim, True, IDempresa, Familiatext Else ProcGravarDataFiltroRel txtinicio, txtFinal, True, IDempresa, Familiatext

'Grava valor total para calculo de percentual
valor = txtTotal
NovoValor = Replace(valor, ",", ".")
Conexao.Execute "UPDATE Producao_Relatorios_Total Set Numero1 = " & NovoValor & " where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

FormulaRel_Vendas_Carteira = FormulaRel_Vendas_Carteira & "and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If listaitens.ListItems.Count = 0 Then Exit Sub
frmVendas_carteira_menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

listaitens.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_Vendas_Carteira = "" Then Exit Sub
Set TBLISTA_Vendas_Carteira = CreateObject("adodb.recordset")
TBLISTA_Vendas_Carteira.Open StrSql_Vendas_Carteira, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Vendas_Carteira.EOF = False Then
    If Opt_todos.Value = True Then
        Select Case cmbTipoData
            Case "": Dataini = TBLISTA_Vendas_Carteira!Data
            Case "Emissão": Dataini = TBLISTA_Vendas_Carteira!Data
            Case "Venda": Dataini = TBLISTA_Vendas_Carteira!Datavendas
            Case "Faturamento": Dataini = TBLISTA_Vendas_Carteira!DataFaturamento
            Case "Cancelamento": Dataini = TBLISTA_Vendas_Carteira!dataalteracao
            Case "Perda": Dataini = TBLISTA_Vendas_Carteira!dataalteracao
        End Select
        TBLISTA_Vendas_Carteira.MoveLast
        Select Case cmbTipoData
            Case "": DataFim = TBLISTA_Vendas_Carteira!Data
            Case "Emissão": DataFim = TBLISTA_Vendas_Carteira!Data
            Case "Venda": DataFim = TBLISTA_Vendas_Carteira!Datavendas
            Case "Faturamento": DataFim = TBLISTA_Vendas_Carteira!DataFaturamento
            Case "Cancelamento": DataFim = TBLISTA_Vendas_Carteira!dataalteracao
            Case "Perda": DataFim = TBLISTA_Vendas_Carteira!dataalteracao
        End Select
        TBLISTA_Vendas_Carteira.MoveFirst
    End If
    ProcExibePagina (Pagina)
End If
ProcCarregaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

listaitens.ListItems.Clear
TBLISTA_Vendas_Carteira.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Vendas_Carteira.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendas_Carteira.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendas_Carteira.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendas_Carteira.PageSize * (Pagina - 1)), 0), TBLISTA_Vendas_Carteira.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Vendas_Carteira.EOF = False And (ContadorReg <= TamanhoPagina)
    With listaitens.ListItems
        .Add , , TBLISTA_Vendas_Carteira!CODIGO
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendas_Carteira!Liberacao), "", TBLISTA_Vendas_Carteira!Liberacao)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Vendas_Carteira!Ncotacao), "", TBLISTA_Vendas_Carteira!Ncotacao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Vendas_Carteira!Revisao), 0, TBLISTA_Vendas_Carteira!Revisao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendas_Carteira!Cliente), "", TBLISTA_Vendas_Carteira!Cliente)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Vendas_Carteira!PCCliente), "", TBLISTA_Vendas_Carteira!PCCliente)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Vendas_Carteira!Desenho), "", TBLISTA_Vendas_Carteira!Desenho)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Vendas_Carteira!Rev_codinterno), "", TBLISTA_Vendas_Carteira!Rev_codinterno)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Vendas_Carteira!N_referencia), "", TBLISTA_Vendas_Carteira!N_referencia)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Vendas_Carteira!descricao_tecnica), "", TBLISTA_Vendas_Carteira!descricao_tecnica)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Vendas_Carteira!Unidade), "", TBLISTA_Vendas_Carteira!Unidade)
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Vendas_Carteira!Unidade_com), "", TBLISTA_Vendas_Carteira!Unidade_com)
        .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA_Vendas_Carteira!preco_unitario), "", Format(TBLISTA_Vendas_Carteira!preco_unitario, "###,##0.0000000000"))
        .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA_Vendas_Carteira!quantidade), "", Format(TBLISTA_Vendas_Carteira!quantidade, "###,##0.0000"))
        .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_Vendas_Carteira!preco_lote), "", Format(TBLISTA_Vendas_Carteira!preco_lote, "###,##0.00"))
        .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA_Vendas_Carteira!PrazoFinal), "", Format(TBLISTA_Vendas_Carteira!PrazoFinal, "dd/mm/yy"))
        .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA_Vendas_Carteira!Liberacao), "", TBLISTA_Vendas_Carteira!Liberacao)
        .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA_Vendas_Carteira!qtdeliberada), "", Format(TBLISTA_Vendas_Carteira!qtdeliberada, "###,##0.0000"))
        .Item(.Count).SubItems(18) = IIf(IsNull(TBLISTA_Vendas_Carteira!QtdeFaturada), "", Format(TBLISTA_Vendas_Carteira!QtdeFaturada, "###,##0.0000"))
        .Item(.Count).SubItems(19) = IIf(IsNull(TBLISTA_Vendas_Carteira!QtdeFaturar), "", Format(TBLISTA_Vendas_Carteira!QtdeFaturar, "###,##0.0000"))
        .Item(.Count).SubItems(20) = IIf(IsNull(TBLISTA_Vendas_Carteira!SaldoPedido), "", Format(TBLISTA_Vendas_Carteira!SaldoPedido, "###,##0.0000"))
        .Item(.Count).SubItems(21) = IIf(IsNull(TBLISTA_Vendas_Carteira!qtdeexpedida), "", Format(TBLISTA_Vendas_Carteira!qtdeexpedida, "###,##0.0000"))
        .Item(.Count).SubItems(22) = IIf(IsNull(TBLISTA_Vendas_Carteira!Datavendas), "", Format(TBLISTA_Vendas_Carteira!Datavendas, "dd/mm/yy"))
        .Item(.Count).SubItems(23) = IIf(IsNull(TBLISTA_Vendas_Carteira!datapcp), "", Format(TBLISTA_Vendas_Carteira!datapcp, "dd/mm/yy"))
        .Item(.Count).SubItems(24) = IIf(IsNull(TBLISTA_Vendas_Carteira!DataFaturamento), "", Format(TBLISTA_Vendas_Carteira!DataFaturamento, "dd/mm/yy"))
        .Item(.Count).SubItems(25) = IIf(IsNull(TBLISTA_Vendas_Carteira!dataexpedicao), "", Format(TBLISTA_Vendas_Carteira!dataexpedicao, "dd/mm/yy"))
        .Item(.Count).SubItems(26) = IIf(IsNull(TBLISTA_Vendas_Carteira!dataalteracao), "", Format(TBLISTA_Vendas_Carteira!dataalteracao, "dd/mm/yy"))
        If TBLISTA_Vendas_Carteira!status = "CANCELADA" Or IsNull(TBLISTA_Vendas_Carteira!status) = True Then listaitens.ColumnHeaders(27).Text = "Dt. cancel."
        If TBLISTA_Vendas_Carteira!status = "PERDIDA P/ PRAZO" Then listaitens.ColumnHeaders(27).Text = "Dt. per. prazo"
        If TBLISTA_Vendas_Carteira!status = "PERDIDA P/ PREÇO" Then listaitens.ColumnHeaders(27).Text = "Dt. per. preço"
    End With
    TBLISTA_Vendas_Carteira.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Vendas_Carteira.RecordCount
If TBLISTA_Vendas_Carteira.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Vendas_Carteira.PageCount
ElseIf TBLISTA_Vendas_Carteira.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Vendas_Carteira.PageCount & " de: " & TBLISTA_Vendas_Carteira.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Vendas_Carteira.AbsolutePage - 1 & " de: " & TBLISTA_Vendas_Carteira.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

valor = 0
IDlista = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql_Vendas_Carteira_Total, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If IDlista <> TBAbrir!CODIGO Then valor = valor + IIf(IsNull(TBAbrir!preco_lote), 0, TBAbrir!preco_lote)
        IDlista = TBAbrir!CODIGO
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close
txtTotal.Text = Format(valor, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_Carteira.AbsolutePage <> 2 Then
    If TBLISTA_Vendas_Carteira.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Vendas_Carteira.PageCount - 1)
    Else
        TBLISTA_Vendas_Carteira.AbsolutePage = TBLISTA_Vendas_Carteira.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Vendas_Carteira.AbsolutePage)
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
    TBLISTA_Vendas_Carteira.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Vendas_Carteira.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_Carteira.AbsolutePage = 1
ProcExibePagina (TBLISTA_Vendas_Carteira.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_Carteira.AbsolutePage <> -3 Then
    If TBLISTA_Vendas_Carteira.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Vendas_Carteira.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Vendas_Carteira.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_Carteira.AbsolutePage = TBLISTA_Vendas_Carteira.PageCount
ProcExibePagina (TBLISTA_Vendas_Carteira.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF8: If Cmb_opcao_lista = "Faturar" Then ProcFaturar
    Case vbKeyF9: If Cmb_opcao_lista = "Cancelar faturamento" Then ProcCancelarFaturamento
    Case vbKeyF10: If Cmb_opcao_lista = "Expedir" Then ProcExpedir
    Case vbKeyF11: If Cmb_opcao_lista = "Cancelar expedição" Then ProcCancelarExpedicao
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 11, True
Formulario = "Vendas/Follow up"
Direitos
ProcLimpaVariaveisPrincipais
txtinicio.Value = Date
txtFinal.Value = Date
ProcCarregaComboEmpresa Cmb_empresa, True
ProcCarregaComboAno cmbAno, "2005", 1
Cmb_opcao_lista = "Faturar"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Follow up"
Direitos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362C" Then
    If USMsgBox("Deseja realmente atualizar os dados na tabela vendas carteira?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        ProcCorrigeDescontoProduto
        ProcCorrigeStatusProduto
        ProcCorrigeStatusProdutoVend
        ProcCorrigeStatusProdServFFP
        ProcCorrigeStatusProposta
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Vendas/Follow up"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeDescontoProduto()
On Error GoTo tratar_erro

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select * from vendas_carteira order by desconto", Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = False Then
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
        If IsNull(TBCarteira!Desconto) = True Then
            TBCarteira!Desconto = 0
            TBCarteira!ValorDesconto = 0
            TBCarteira!preco_unitario_desconto = TBCarteira!preco_unitario
            TBCarteira.Update
        End If
        Contador = Contador + 1
        PBLista.Value = Contador
        TBCarteira.MoveNext
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeStatusProduto()
On Error GoTo tratar_erro

Set TBProposta = CreateObject("adodb.recordset")
TBProposta.Open "Select * from vendas_proposta where status = 'ABERTA EM ANALISE' or status = 'CANCELADA' or status = 'REVISADA' or status = 'OUTROS' or status = 'PERDIDA P/ PRAZO' or status = 'PERDIDA P/ PREÇO' or status = 'PERDIDA P/ PREÇO' order by ncotacao, revisao", Conexao, adOpenKeyset, adLockOptimistic
If TBProposta.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBProposta.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBProposta.EOF = False
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select * from vendas_carteira where cotacao = " & TBProposta!Cotacao & " order by desenho", Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = False Then
            Do While TBCarteira.EOF = False
                If TBProposta!status = "ABERTA EM ANALISE" Then TBCarteira!Liberacao = "ABERTA EM ANALISE"
                If TBProposta!status = "CANCELADA" Then TBCarteira!Liberacao = "CANCELADO"
                If TBProposta!status = "REVISADA" Then TBCarteira!Liberacao = "REVISADA"
                If TBProposta!status = "OUTROS" Then TBCarteira!Liberacao = "OUTROS"
                If TBProposta!status = "PERDIDA P/ PRAZO" Then TBCarteira!Liberacao = "PERDIDO P/ PRAZO"
                If TBProposta!status = "PERDIDA P/ PREÇO" Then TBCarteira!Liberacao = "PERDIDO P/ PREÇO"
                TBCarteira!Datavendas = Null
                TBCarteira!dataprodsaida = Null
                TBCarteira!DataFaturamento = Null
                TBCarteira!qtdeliberada = 0
                TBCarteira!QtdeFaturada = 0
                TBCarteira.Update
                TBCarteira.MoveNext
            Loop
        End If
        TBCarteira.Close
        Contador = Contador + 1
        PBLista.Value = Contador
        TBProposta.MoveNext
    Loop
End If
TBProposta.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeStatusProdutoVend()
On Error GoTo tratar_erro

Set TBProposta = CreateObject("adodb.recordset")
TBProposta.Open "Select * from vendas_proposta where status = 'VENDIDA' or status = 'VENDIDA PARCIAL' or status = 'FATURAR' or status = 'FATURAR PARCIAL' or status = 'FATURADA' or status = 'FATURADA PARCIAL' order by ncotacao, revisao", Conexao, adOpenKeyset, adLockOptimistic
If TBProposta.EOF = False Then
    TBProposta.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBProposta.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBProposta.MoveFirst
    Do While TBProposta.EOF = False
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select * from vendas_carteira where cotacao = " & TBProposta!Cotacao & " order by desenho", Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = False Then
            Do While TBCarteira.EOF = False
                TBCarteira!Liberacao = "VENDIDA"
                If TBCarteira!Datavendas = "" Then
                    If IsNull(TBProposta!Datavendas) = False Then TBCarteira!Datavendas = TBProposta!Datavendas
                End If
                TBCarteira!DataFaturamento = Null
                TBCarteira.Update
                TBCarteira.MoveNext
            Loop
        End If
        TBCarteira.Close
        Contador = Contador + 1
        PBLista.Value = Contador
        TBProposta.MoveNext
    Loop
End If
TBProposta.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeStatusProdServFFP()
On Error GoTo tratar_erro

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select * from vendas_carteira where liberacao = 'VENDIDA' or liberacao = 'FATURADO' or liberacao = 'FATURADO PARCIAL' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = False Then
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
        QtdeFaturada = 0
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_Detalhes_Nota where Codigo = " & TBCarteira!CODIGO & " and int_Cod_Produto = '" & TBCarteira!Desenho & "' order by Int_notafiscal", Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            Do While TBLISTA.EOF = False
                QtdeFaturada = QtdeFaturada + TBLISTA!int_Qtd
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where id = " & TBLISTA!ID_nota, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    DataFaturamento = TBAbrir!dt_DataEmissao
                End If
                TBAbrir.Close
                TBLISTA.MoveNext
            Loop
        End If
        TBCarteira!DataFaturamento = DataFaturamento
        TBCarteira!QtdeFaturada = QtdeFaturada
        If QtdeFaturada <> 0 Then
            If QtdeFaturada >= TBCarteira!quantidade Then TBCarteira!Liberacao = "FATURADO" Else TBCarteira!Liberacao = "FATURADO PARCIAL"
        Else
            TBCarteira!Liberacao = "VENDIDA"
        End If
        TBCarteira.Update
        TBLISTA.Close
        Contador = Contador + 1
        PBLista.Value = Contador
        TBCarteira.MoveNext
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeStatusProposta()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select proposta,revisao from tbl_proposta_nota order by proposta", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBAbrir.MoveFirst
    Do While TBAbrir.EOF = False
        Faturado = False
        Faturado_Parcial = False
        Faturar = False
        Cotacao = 0
        Set TBProposta = CreateObject("adodb.recordset")
        TBProposta.Open "Select * from vendas_proposta where Ncotacao = '" & TBAbrir!Proposta & "' and revisao = " & TBAbrir!Revisao, Conexao, adOpenKeyset, adLockOptimistic
        If TBProposta.EOF = False Then
            Set TBCarteira = CreateObject("adodb.recordset")
            TBCarteira.Open "Select * from vendas_carteira where cotacao = " & TBProposta!Cotacao & " order by desenho", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                Cotacao = TBCarteira!Cotacao
                Do While TBCarteira.EOF = False
                    If TBCarteira!Liberacao = "FATURADO" Then Faturado = True
                    If TBCarteira!Liberacao = "FATURADO PARCIAL" Then Faturado_Parcial = True
                    If TBCarteira!Liberacao = "FATURAR" Or TBCarteira!Liberacao = "FATURAR PARCIAL" Or TBCarteira!Liberacao = "VENDIDA" Then Faturar = True
                    TBCarteira.MoveNext
                Loop
            End If
            TBCarteira.Close
            If Faturado = True And Faturado_Parcial = False And Faturar = False Then TBProposta!status = "FATURADA"
            If Faturado = True And Faturar = True Or Faturado_Parcial = True Then TBProposta!status = "FATURADA PARCIAL"
            If Faturado = False And Faturado_Parcial = False And Faturar = True Then TBProposta!status = "VENDIDA"
            TBProposta.Update
        End If
        TBProposta.Close
        Contador = Contador + 1
        PBLista.Value = Contador
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listaitens_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With listaitens
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                StatusTexto = .ListItems.Item(InitFor).ListSubItems.Item(16).Text
                If Cmb_opcao_lista = "Faturar" Then
                    If StatusTexto <> "VENDIDA" And Left(StatusTexto, 7) <> "FATURAR" And Left(StatusTexto, 8) <> "FATURADO" Then
                        .ListItems(InitFor).Checked = False
                        GoTo Proximo
                    End If
                ElseIf Cmb_opcao_lista = "Cancelar faturamento" Then
                        If Left(StatusTexto, 8) <> "FATURADO" Then
                            .ListItems(InitFor).Checked = False
                            GoTo Proximo
                        End If
                    ElseIf Cmb_opcao_lista = "Expedir" Or Cmb_opcao_lista = "Cancelar expedição" Then
                            If StatusTexto <> "VENDIDA" And Left(StatusTexto, 7) <> "FATURAR" And Left(StatusTexto, 8) <> "FATURADO" Then
                                .ListItems(InitFor).Checked = False
                                GoTo Proximo
                            End If
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView listaitens, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listaitens_DblClick()
On Error GoTo tratar_erro

With listaitens
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturadaProdServ .SelectedItem, .SelectedItem.ListSubItems(6), False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listaitens_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With listaitens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            StatusTexto = .ListItems.Item(InitFor).ListSubItems.Item(16).Text
            If Cmb_opcao_lista = "Faturar" Then
                If StatusTexto <> "VENDIDA" And Left(StatusTexto, 7) <> "FATURAR" And Left(StatusTexto, 8) <> "FATURADO" Then
                    USMsgBox ("Só é permitido faturar produto/serviço com o status vendida, faturar, faturar parcial, faturado ou faturado parcial."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            ElseIf Cmb_opcao_lista = "Cancelar faturamento" Then
                    If Left(StatusTexto, 8) <> "FATURADO" Then
                        USMsgBox ("Só é permitido cancelar o faturamento de produto/serviço com o status faturado ou faturado parcial."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                ElseIf Cmb_opcao_lista = "Expedir" Or Cmb_opcao_lista = "Cancelar expedição" Then
                        If StatusTexto <> "VENDIDA" And Left(StatusTexto, 7) <> "FATURAR" And Left(StatusTexto, 8) <> "FATURADO" Then
                            If Cmb_opcao_lista = "Expedir" Then TextoMsg = "expedir" Else TextoMsg = "cancelar expedição de"
                            USMsgBox ("Só é permitido " & TextoMsg & " produto/serviço com o status vendida, faturar, faturar parcial, faturado ou faturado parcial."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
            End If
        End If
      Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listaitens_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listaitens.SelectedItem.ListSubItems.Count = 0 Then Exit Sub
Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select * from Follow_up_vendas where Codigo = " & listaitens.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = False Then
    LimparCampo = False
'    Cmb_status.Text = IIf(IsNull(TBCarteira!Liberacao), "", TBCarteira!Liberacao)
    txtPI.Text = IIf(IsNull(TBCarteira!Ncotacao), "", TBCarteira!Ncotacao)
    txtCodigo = IIf(IsNull(TBCarteira!Desenho), "", TBCarteira!Desenho)
    cmbcodref = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    txtUN = IIf(IsNull(TBCarteira!Unidade), "", TBCarteira!Unidade)
    Txt_un_com = IIf(IsNull(TBCarteira!Unidade_com), "", TBCarteira!Unidade_com)
    txtQuantidade = IIf(IsNull(TBCarteira!quantidade), "", Format(TBCarteira!quantidade, "###,##0.000"))
    txtPedido = IIf(IsNull(TBCarteira!PCCliente), "", TBCarteira!PCCliente)
    txtPrazoFinal = IIf(IsNull(TBCarteira!PrazoFinal), "", Format(TBCarteira!PrazoFinal, "dd/mm/yy"))
    Txt_qtde_liberada_faturar = IIf(IsNull(TBCarteira!qtdeliberada), "", Format(TBCarteira!qtdeliberada, "###,##0.000"))
    Txtqtde_faturada = IIf(IsNull(TBCarteira!QtdeFaturada), "", Format(TBCarteira!QtdeFaturada, "###,##0.000"))
    Txt_qtde_faturar = IIf(IsNull(TBCarteira!QtdeFaturar), "", Format(TBCarteira!QtdeFaturar, "###,##0.000"))
    Txt_saldo_pedido = IIf(IsNull(TBCarteira!SaldoPedido), "", Format(TBCarteira!SaldoPedido, "###,##0.000"))
    Txt_qtde_expedida = IIf(IsNull(TBCarteira!qtdeexpedida), "", Format(TBCarteira!qtdeexpedida, "###,##0.000"))
    txtdescricao = IIf(IsNull(TBCarteira!descricao_tecnica), "", TBCarteira!descricao_tecnica)
    txtCliente = IIf(IsNull(TBCarteira!Cliente), "", TBCarteira!Cliente)
    Txt_grupo_cliente = IIf(IsNull(TBCarteira!Grupo_Cliente), "", TBCarteira!Grupo_Cliente)
    txtDatavendas = IIf(IsNull(TBCarteira!Datavendas), "", Format(TBCarteira!Datavendas, "dd/mm/yy"))
    Txtdataprod = IIf(IsNull(TBCarteira!datapcp), "", Format(TBCarteira!datapcp, "dd/mm/yy"))
    txtdatafaturado = IIf(IsNull(TBCarteira!DataFaturamento), "", Format(TBCarteira!DataFaturamento, "dd/mm/yy"))
    txtdataexp = IIf(IsNull(TBCarteira!dataexpedicao), "", Format(TBCarteira!dataexpedicao, "dd/mm/yy"))
    LimparCampo = True
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

cmbcodref = ""
txtUN = ""
Txt_un_com = ""
txtQuantidade = ""
txtPrazoFinal = ""
Txt_qtde_liberada_faturar = ""
Txtqtde_faturada = ""
Txt_qtde_faturar = ""
Txt_saldo_pedido = ""
Txt_qtde_expedida = ""
txtDatavendas = ""
Txtdataprod = ""
txtdatafaturado = ""
txtdataexp = ""
'listaitens.ListItems.Clear
txtTotal = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_todos_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAteomes_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro_Click()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_grupo_cliente_Change()
On Error GoTo tratar_erro

If LimparCampo = True Then ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcliente_Change()
On Error GoTo tratar_erro

If LimparCampo = True Then ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

If LimparCampo = True Then ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDescricao_Change()
On Error GoTo tratar_erro

If chkDescricao.Value = 1 Then chkCodInterno.Value = 0
If LimparCampo = True Then ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtfinal_Change()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtinicio_Change()
On Error GoTo tratar_erro

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboData()
On Error GoTo tratar_erro

TipoDataAntigo = cmbTipoData
With cmbTipoData
    .Clear
    .AddItem "Emissão"
    .Text = "Emissão"
    If Cmb_status = "Faturadas" Or Cmb_status = "A Faturar" Or Cmb_status = "Vendidas e Faturadas" Or Cmb_status = "Vendidas e Faturadas parcial" Or Cmb_status = "Vendidas" Then
        .AddItem "Venda"
        .AddItem "Prazo final"
        .AddItem "Faturamento"
        If TipoDataAntigo <> "" And (TipoDataAntigo = "Emissão" Or TipoDataAntigo = "Venda" Or TipoDataAntigo = "Prazo final" Or TipoDataAntigo = "Faturamento") Then
            .Text = TipoDataAntigo
        Else
            If Cmb_status = "Faturadas" Or Cmb_status = "Vendidas e Faturadas" Or Cmb_status = "Vendidas e Faturadas parcial" Then .Text = "Faturamento" Else .Text = "Venda"
        End If
    End If
    If Cmb_status = "Canceladas" Then
        .AddItem "Cancelamento"
        .Text = "Cancelamento"
    End If
    If Cmb_status = "Perdidas por prazo" Or Cmb_status = "Perdidas por preço" Then
        .AddItem "Perda"
        .Text = "Perda"
    End If
End With

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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 3: ProcFaturar
    Case 4: ProcCancelarFaturamento
    Case 5: ProcExpedir
    Case 6: ProcCancelarExpedicao
    Case 7: procAtualiza
    Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
