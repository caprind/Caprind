VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_atualizacao_valores 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Vendas - Atualização de valores"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   45
      Top             =   8250
      Width           =   15195
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3780
         TabIndex        =   16
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtPagIr 
         Height          =   315
         Left            =   9540
         TabIndex        =   17
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   10110
         TabIndex        =   18
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
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   21
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_atualizacao_valores.frx":0000
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
         TabIndex        =   20
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_atualizacao_valores.frx":37A4
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
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   10680
         TabIndex        =   19
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_atualizacao_valores.frx":72AD
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
         TabIndex        =   22
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "Frm_atualizacao_valores.frx":B39C
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
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
         Left            =   4410
         TabIndex        =   50
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
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
         Left            =   3090
         TabIndex        =   48
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
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
         Left            =   13050
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
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
         TabIndex        =   46
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.CheckBox optPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data do cadastro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11460
      TabIndex        =   12
      Top             =   1020
      Width           =   1755
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8310
      TabIndex        =   38
      Top             =   990
      Width           =   3015
      Begin VB.ComboBox cmbStatus 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "Frm_atualizacao_valores.frx":EC28
         Left            =   180
         List            =   "Frm_atualizacao_valores.frx":EC35
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Status."
         Top             =   270
         Width           =   2670
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   33
      Top             =   990
      Width           =   8235
      Begin VB.CheckBox Chk_valores_diferentes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clientes com valores diferentes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   330
         Width           =   4725
      End
      Begin VB.CheckBox chkProdutos 
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
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   330
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.CheckBox chkServicos 
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
         Height          =   195
         Left            =   1260
         TabIndex        =   9
         Top             =   330
         Value           =   1  'Checked
         Width           =   915
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7140
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "Frm_atualizacao_valores.frx":EC50
      Count           =   1
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   60
      TabIndex        =   30
      Top             =   9150
      Width           =   15195
      Begin VB.ComboBox cmbCasasDecimais 
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
         ItemData        =   "Frm_atualizacao_valores.frx":11579
         Left            =   12705
         List            =   "Frm_atualizacao_valores.frx":1157B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Casas decimais após a virgula."
         Top             =   390
         Width           =   1035
      End
      Begin VB.CheckBox Chk_atualizar_vr 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atualizar valor de revenda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5820
         TabIndex        =   25
         Top             =   375
         Value           =   1  'Checked
         Width           =   2205
      End
      Begin VB.CheckBox Chk_atualizar_vc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atualizar valor de consumo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3330
         TabIndex        =   24
         Top             =   375
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.TextBox Txt_valor 
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
         Left            =   13755
         TabIndex        =   29
         ToolTipText     =   "VValor."
         Top             =   390
         Width           =   1245
      End
      Begin VB.ComboBox Cmb_atualizar_para 
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
         ItemData        =   "Frm_atualizacao_valores.frx":1157D
         Left            =   10620
         List            =   "Frm_atualizacao_valores.frx":1157F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Atualizar valores para."
         Top             =   390
         Width           =   2055
      End
      Begin VB.ComboBox Cmb_atualizar_por 
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
         ItemData        =   "Frm_atualizacao_valores.frx":11581
         Left            =   8550
         List            =   "Frm_atualizacao_valores.frx":1158B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Atualizar valores por."
         Top             =   390
         Width           =   2055
      End
      Begin VB.CheckBox Chk_atualizar_com_mesmo_valor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atualizar clientes com o mesmo valor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   375
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Casas dec."
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
         Left            =   12780
         TabIndex        =   49
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
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
         Left            =   14160
         TabIndex        =   44
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atualizar valores para"
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
         Left            =   10710
         TabIndex        =   43
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atualizar valores por"
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
         Left            =   8692
         TabIndex        =   42
         Top             =   180
         Width           =   1770
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5610
      Left            =   60
      TabIndex        =   15
      Top             =   2640
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   9895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Cód. de ref."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   10675
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Un. com."
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Vlr. consumo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Vlr. revenda"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "ID_cli_for"
         Object.Width           =   0
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   32
      Top             =   8880
      Width           =   15195
      _ExtentX        =   26802
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
      Left            =   60
      TabIndex        =   31
      Top             =   0
      Width           =   15200
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   6
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
      ButtonCaption2  =   "Atualizar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Atualizar (F3)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   46
      ButtonTop2      =   2
      ButtonWidth2    =   59
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   107
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
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
      ButtonLeft4     =   111
      ButtonTop4      =   2
      ButtonWidth4    =   41
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   154
      ButtonTop5      =   2
      ButtonWidth5    =   30
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   186
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
   Begin VB.Frame FrameData 
      BackColor       =   &H00E0E0E0&
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
      Height          =   735
      Left            =   11340
      TabIndex        =   39
      Top             =   990
      Width           =   3915
      Begin MSComCtl2.DTPicker Msk_final 
         Height          =   315
         Left            =   2430
         TabIndex        =   14
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   300
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
         Format          =   179896321
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker Msk_inicio 
         Height          =   315
         Left            =   570
         TabIndex        =   13
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   300
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
         Format          =   179896321
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         Left            =   1980
         TabIndex        =   41
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
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
         TabIndex        =   40
         Top             =   300
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   60
      TabIndex        =   34
      Top             =   1740
      Width           =   15195
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   2940
         TabIndex        =   51
         Top             =   210
         Width           =   4785
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
            TabIndex        =   3
            Top             =   180
            Width           =   1155
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
            TabIndex        =   1
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
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
            TabIndex        =   2
            Top             =   180
            Width           =   1275
         End
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
            TabIndex        =   4
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.TextBox txtTexto 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7800
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   5025
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
         ItemData        =   "Frm_atualizacao_valores.frx":115A2
         Left            =   180
         List            =   "Frm_atualizacao_valores.frx":115A4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2685
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
         ItemData        =   "Frm_atualizacao_valores.frx":115A6
         Left            =   12840
         List            =   "Frm_atualizacao_valores.frx":115B0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Ordenar por."
         Top             =   390
         Width           =   2175
      End
      Begin VB.ComboBox cmbfamilia 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Visible         =   0   'False
         Width           =   5025
      End
      Begin VB.Label Label9 
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
         Left            =   1102
         TabIndex        =   37
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label8 
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
         Left            =   9577
         TabIndex        =   36
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label4 
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
         Left            =   13417
         TabIndex        =   35
         Top             =   180
         Width           =   1020
      End
   End
End
Attribute VB_Name = "Frm_atualizacao_valores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sql_Atualizacao_Valores  As String 'OK
Dim TBLISTA_Atualizacao_Valores As ADODB.Recordset 'OK

Private Sub Chk_valores_diferentes_Click()
On Error GoTo tratar_erro

With Lista
    .ListItems.Clear
    If Chk_valores_diferentes.Value = 1 Then
        If Vendas_Atualização_Valores = True Then .ColumnHeaders(4).Width = 2976 Else .ColumnHeaders(4).Width = 4176
        .ColumnHeaders(7).Width = 3076
    Else
        If Vendas_Atualização_Valores = True Then .ColumnHeaders(4).Width = 6052 Else .ColumnHeaders(4).Width = 7252
        .ColumnHeaders(7).Width = 0
    End If
End With
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkProdutos_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkServicos_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_atualizar_por_Click()
On Error GoTo tratar_erro

With Cmb_atualizar_para
    .Clear
    .AddItem "Mais"
    .AddItem "Menos"
    If Cmb_atualizar_por = "Valor" Then .AddItem "Igual"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ordenar_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Grupo do cliente" Then
    txtTexto.Visible = False
    With cmbfamilia
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            If Vendas_Atualização_Valores = True Then
                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", True
            Else
                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
            End If
        ElseIf cmbfiltrarpor = "Grupo" Then
                If Vendas_Atualização_Valores = True Then
                    ProcCarregaComboGrupoFamilia cmbfamilia, "Grupo <> 'Null' and vendas = 'True'", True
                Else
                    ProcCarregaComboGrupoFamilia cmbfamilia, "Grupo <> 'Null' and compras = 'True'", True
                End If
            ElseIf cmbfiltrarpor = "Cliente" Then
                    Set TBClientes = CreateObject("adodb.recordset")
                    TBClientes.Open "Select IDCliente, NomeRazao from Clientes where NomeRazao <> 'Null' group by IDCliente, NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
                    If TBClientes.EOF = False Then
                        Do While TBClientes.EOF = False
                            .AddItem Trim(TBClientes!NomeRazao)
                            .ItemData(.NewIndex) = TBClientes!IDCliente
                            TBClientes.MoveNext
                        Loop
                    End If
                    TBClientes.Close
                ElseIf cmbfiltrarpor = "Grupo do cliente" Then
                        Set TBFamilia = CreateObject("adodb.recordset")
                        TBFamilia.Open "Select ID, Texto from Clientes_grupos where Texto <> 'Null' group by ID, Texto", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFamilia.EOF = False Then
                            Do While TBFamilia.EOF = False
                                .AddItem TBFamilia!Texto
                                .ItemData(.NewIndex) = TBFamilia!ID
                                TBFamilia.MoveNext
                            Loop
                        End If
                        TBFamilia.Close
                    Else
                        Set TBFornecedor = CreateObject("adodb.recordset")
                        TBFornecedor.Open "Select IDCliente, Nome_Razao from Compras_fornecedores where Nome_Razao <> 'Null' group by IDCliente, Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFornecedor.EOF = False Then
                            Do While TBFornecedor.EOF = False
                                .AddItem Trim(TBFornecedor!Nome_Razao)
                                .ItemData(.NewIndex) = TBFornecedor!IDCliente
                                TBFornecedor.MoveNext
                            Loop
                        End If
                        TBFornecedor.Close
        End If
    End With
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
End If


With Lista
    If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Grupo do cliente" Then
        If Vendas_Atualização_Valores = True Then .ColumnHeaders(4).Width = 2976 Else .ColumnHeaders(4).Width = 4176
        .ColumnHeaders(7).Width = 3076
    Else
        If Vendas_Atualização_Valores = True Then .ColumnHeaders(4).Width = 6052 Else .ColumnHeaders(4).Width = 7252
        .ColumnHeaders(7).Width = 0
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With Msk_final
    If FunVerificaDataFinal(Msk_inicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
Lista.ListItems.Clear
If Vendas_Atualização_Valores = True Then AplicacaoFiltro = "vendas = 'True'" Else AplicacaoFiltro = "compras = 'True'"
TipoProduto = "Tipo <> 'Null'"
If chkProdutos.Value = 1 And chkServicos.Value = 0 Then
    TipoProduto = "Tipo <> 'S'"
ElseIf chkProdutos.Value = 0 And chkServicos.Value = 1 Then
        TipoProduto = "Tipo = 'S'"
End If
StatusFiltro = " and desenho <> 'Null'"
If cmbStatus <> "" Then
    If cmbStatus = "Liberado" Then StatusFiltro = " and bloqueado = 'False'" Else StatusFiltro = " and bloqueado = 'True'"
End If
If Cmb_ordenar = "Código interno" Then Ordenar = "Desenho" Else Ordenar = "Descricao"

If optPeriodo.Value = 1 Then DataFiltro = "(data) Between '" & Msk_inicio.Value & "' And '" & Msk_final.Value & "'" Else DataFiltro = "desenho <> 'Null'"

FiltroValorDif = ""
If Chk_valores_diferentes.Value = 1 Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Or cmbfiltrarpor = "Fornecedor" Then
    If Vendas_Atualização_Valores = True Then
        CamposFiltro = "codProduto, Desenho, Descricao, Unidade_com, Classe, Idcliente, Cliente, PConsumo_cliente, PRevenda_cliente"
        If Chk_valores_diferentes.Value = 1 Then FiltroValorDif = "and PConsumo_cliente <> 0 and PRevenda_cliente <> 0 and (PConsumo <> PConsumo_cliente or PRevenda <> PRevenda_cliente)"
    Else
        CamposFiltro = "codProduto, Desenho, Descricao, Unidade_com, Classe, Idfornecedor, Fornecedor, PCusto_fornecedor"
        If Chk_valores_diferentes.Value = 1 Then FiltroValorDif = "and PCusto <> PCusto_fornecedor"
    End If
Else
    CamposFiltro = "codProduto, Desenho, Descricao, Unidade_com, Classe, PCusto, PConsumo, PRevenda"
End If

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
        If cmbfiltrarpor = "Família" Then TextoFiltro = "Classe" Else TextoFiltro = "Grupo_familia"
        Sql_Atualizacao_Valores = "Select " & CamposFiltro & " from Vendas_compras_atualizacao_valores where " & TextoFiltro & " = '" & cmbfamilia & "' and " & AplicacaoFiltro & " and " & TipoProduto & " and " & DataFiltro & " " & StatusFiltro & " " & FiltroValorDif & " group by " & CamposFiltro & " order by " & Ordenar
    ElseIf cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Or cmbfiltrarpor = "Fornecedor" Then
            Select Case cmbfiltrarpor
                Case "Cliente": TextoFiltro = "IDCliente"
                Case "Grupo do cliente": TextoFiltro = "IDGrupo"
                Case "Fornecedor": TextoFiltro = "IDFornecedor"
            End Select
            Sql_Atualizacao_Valores = "Select " & CamposFiltro & " from Vendas_compras_atualizacao_valores where " & TextoFiltro & " = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & AplicacaoFiltro & " and " & TipoProduto & " and " & DataFiltro & " " & StatusFiltro & " " & FiltroValorDif & " group by " & CamposFiltro & " order by " & Ordenar
    Else
        Select Case cmbfiltrarpor
            Case "Código interno": TextoFiltro = "desenho"
            Case "Código de referência": TextoFiltro = "N_referencia"
            Case "Descrição": TextoFiltro = "descricao"
            Case "Descrição comercial": TextoFiltro = "Descricaotecnica"
        End Select
        If Optinicio.Value = True Then Sql_Atualizacao_Valores = "Select " & CamposFiltro & " from Vendas_compras_atualizacao_valores where " & TextoFiltro & " like '" & txtTexto & "%' and " & AplicacaoFiltro & " and " & TipoProduto & " and " & DataFiltro & " " & StatusFiltro & " " & FiltroValorDif & " group by " & CamposFiltro & " order by " & Ordenar
        If Optmeio.Value = True Then Sql_Atualizacao_Valores = "Select " & CamposFiltro & " from Vendas_compras_atualizacao_valores where " & TextoFiltro & " like '%" & txtTexto & "%' and " & AplicacaoFiltro & " and " & TipoProduto & " and " & DataFiltro & " " & StatusFiltro & " " & FiltroValorDif & " group by " & CamposFiltro & " order by " & Ordenar
        If Optfim.Value = True Then Sql_Atualizacao_Valores = "Select " & CamposFiltro & " from Vendas_compras_atualizacao_valores where " & TextoFiltro & " like '%" & txtTexto & "' and " & AplicacaoFiltro & " and " & TipoProduto & " and " & DataFiltro & " " & StatusFiltro & " " & FiltroValorDif & " group by " & CamposFiltro & " order by " & Ordenar
        If optIgual.Value = True Then Sql_Atualizacao_Valores = "Select " & CamposFiltro & " from Vendas_compras_atualizacao_valores where " & TextoFiltro & " = '" & txtTexto & "' and " & AplicacaoFiltro & " and " & TipoProduto & " and " & DataFiltro & " " & StatusFiltro & " " & FiltroValorDif & " group by " & CamposFiltro & " order by " & Ordenar
    End If
Else
    Sql_Atualizacao_Valores = "Select " & CamposFiltro & " from Vendas_compras_atualizacao_valores where " & AplicacaoFiltro & " and " & TipoProduto & " and " & DataFiltro & " " & StatusFiltro & " " & FiltroValorDif & " group by " & CamposFiltro & " order by " & Ordenar
End If
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If Sql_Atualizacao_Valores = "" Then Exit Sub
Set TBLISTA_Atualizacao_Valores = CreateObject("adodb.recordset")
TBLISTA_Atualizacao_Valores.Open Sql_Atualizacao_Valores, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Atualizacao_Valores.EOF = False Then ProcExibePagina (Pagina)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Atualizacao_Valores.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Atualizacao_Valores.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Atualizacao_Valores.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Atualizacao_Valores.RecordCount - IIf(Pagina > 1, (TBLISTA_Atualizacao_Valores.PageSize * (Pagina - 1)), 0), TBLISTA_Atualizacao_Valores.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Atualizacao_Valores.EOF = False And (ContadorReg <= TamanhoPagina)
    nReferencia = ""
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select N_Referencia from item_aplicacoes where codproduto = " & TBLISTA_Atualizacao_Valores!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        nReferencia = IIf(IsNull(TBItem!N_referencia), "", (TBItem!N_referencia))
    End If
    TBItem.Close
    
    With Lista.ListItems
        .Add , , TBLISTA_Atualizacao_Valores!Codproduto
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Desenho), "", TBLISTA_Atualizacao_Valores!Desenho)
        .Item(.Count).SubItems(2) = nReferencia
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Descricao), "", TBLISTA_Atualizacao_Valores!Descricao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Unidade_com), "", TBLISTA_Atualizacao_Valores!Unidade_com)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Classe), "", TBLISTA_Atualizacao_Valores!Classe)
        If Chk_valores_diferentes.Value = 1 Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Or cmbfiltrarpor = "Fornecedor" Then
            If Vendas_Atualização_Valores = True Then
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Cliente), "", TBLISTA_Atualizacao_Valores!Cliente)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Pconsumo_cliente), "0,00000", Format(TBLISTA_Atualizacao_Valores!Pconsumo_cliente, "###,##0.0000000000"))
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Atualizacao_Valores!PRevenda_cliente), "0,00000", Format(TBLISTA_Atualizacao_Valores!PRevenda_cliente, "###,##0.0000000000"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Atualizacao_Valores!IDCliente), "", TBLISTA_Atualizacao_Valores!IDCliente)
            Else
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Fornecedor), "", TBLISTA_Atualizacao_Valores!Fornecedor)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Atualizacao_Valores!Pcusto_fornecedor), "0,00000", Format(TBLISTA_Atualizacao_Valores!Pcusto_fornecedor, "###,##0.0000000000"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Atualizacao_Valores!IDFornecedor), "", TBLISTA_Atualizacao_Valores!IDFornecedor)
            End If
        Else
            If Vendas_Atualização_Valores = True Then
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Atualizacao_Valores!PConsumo), "0,00000", Format(TBLISTA_Atualizacao_Valores!PConsumo, "###,##0.0000000000"))
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Atualizacao_Valores!PRevenda), "0,00000", Format(TBLISTA_Atualizacao_Valores!PRevenda, "###,##0.0000000000"))
            Else
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Atualizacao_Valores!PCusto), "0,00000", Format(TBLISTA_Atualizacao_Valores!PCusto, "###,##0.0000000000"))
            End If
        End If
    End With
    TBLISTA_Atualizacao_Valores.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Atualizacao_Valores.RecordCount
If TBLISTA_Atualizacao_Valores.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Atualizacao_Valores.PageCount
ElseIf TBLISTA_Atualizacao_Valores.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Atualizacao_Valores.PageCount & " de: " & TBLISTA_Atualizacao_Valores.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Atualizacao_Valores.AbsolutePage - 1 & " de: " & TBLISTA_Atualizacao_Valores.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbstatus_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Atualizacao_Valores.AbsolutePage <> 2 Then
    If TBLISTA_Atualizacao_Valores.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Atualizacao_Valores.PageCount - 1)
    Else
        TBLISTA_Atualizacao_Valores.AbsolutePage = TBLISTA_Atualizacao_Valores.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Atualizacao_Valores.AbsolutePage)
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
    TBLISTA_Atualizacao_Valores.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Atualizacao_Valores.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Atualizacao_Valores.AbsolutePage = 1
ProcExibePagina (TBLISTA_Atualizacao_Valores.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Atualizacao_Valores.AbsolutePage <> -3 Then
    If TBLISTA_Atualizacao_Valores.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Atualizacao_Valores.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Atualizacao_Valores.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Atualizacao_Valores.AbsolutePage = TBLISTA_Atualizacao_Valores.PageCount
ProcExibePagina (TBLISTA_Atualizacao_Valores.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcAtualizar
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

ProcCarregaToolBar1 Me, 15195, 6, True
With cmbfiltrarpor
    .Clear
    .AddItem "Código interno"
    .AddItem "Código de referência"
    .AddItem "Descrição"
    .AddItem "Descrição Comercial"
    .AddItem "Família"
    .AddItem "Grupo"
    If Vendas_Atualização_Valores = True Then
        Formulario = "Vendas/Atualização de valores"
        Caption = "Vendas - Atualização de valores"
        .AddItem "Cliente"
        .AddItem "Grupo do cliente"
        Familiatext = "V"
    Else
        Formulario = "Compras/Atualização de valores"
        Caption = "Compras - Atualização de valores"
        .AddItem "Fornecedor"
        Chk_valores_diferentes.Caption = "Fornecedores com valores diferentes"
        
        Lista.ColumnHeaders(4).Width = 7252
        Lista.ColumnHeaders(7).Text = "Fornecedor"
        Lista.ColumnHeaders(8).Text = "Pcusto"
        Lista.ColumnHeaders(9).Width = 0
        
        Chk_atualizar_com_mesmo_valor.Caption = "Atualizar fornecedores com o mesmo valor"
        Chk_atualizar_vc.Visible = False
        Chk_atualizar_vr.Visible = False
        Familiatext = "C"
    End If
End With

With cmbCasasDecimais
    .Clear
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
End With

Direitos
ProcLimpaVariaveisPrincipais
cmbStatus = "Liberado"

ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, 0, "Produtos/Serviços", Familiatext, False
If Permitido = False Then cmbfiltrarpor = "Código interno"

Cmb_ordenar = "Código interno"
Msk_final.Value = Date
Msk_inicio.Value = Date

ProcRemoveObjetosResize Me

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

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Msk_inicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If optPeriodo.Value = 1 Then
    FrameData.Enabled = True
    Msk_inicio.SetFocus
Else
    FrameData.Enabled = False
    Msk_inicio.Value = Date
    Msk_final.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_Change()
On Error GoTo tratar_erro

If Txt_valor <> "" Then
    VerifNumero = Txt_valor
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor = ""
        Txt_valor.SetFocus
        Exit Sub
    End If
End If
    
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

Lista.ListItems.Clear
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcAtualizar
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

'If Alterar = False Then
'    usMsgbox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
Acao = "atualizar os valores"
If Cmb_atualizar_por = "" Then
    NomeCampo = "por qual opção será atualizado"
    ProcVerificaAcao
    Cmb_atualizar_por.SetFocus
    Exit Sub
End If
If Cmb_atualizar_para = "" Then
    NomeCampo = "para qual opção será atualizado"
    ProcVerificaAcao
    Cmb_atualizar_para.SetFocus
    Exit Sub
End If
If cmbCasasDecimais = "" Then
    NomeCampo = "quantas casas decimais após a virgula"
    ProcVerificaAcao
    cmbCasasDecimais.SetFocus
    Exit Sub
End If
If Txt_valor = "" Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Txt_valor.SetFocus
    Exit Sub
End If
valor = Txt_valor
If valor < 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Txt_valor.SetFocus
    Exit Sub
End If

If Cmb_atualizar_por = "Valor" Then
    If USMsgBox("Os valores que estão zerados serão atualizados?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True Else Permitido1 = False
End If

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente atualizar o valor desse(s) registro(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            If Chk_valores_diferentes.Value = 1 Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Or cmbfiltrarpor = "Fornecedor" Then
                If Vendas_Atualização_Valores = True Then
                    INNERJOINTEXTO = "PC.* from Projproduto_clientes PC INNER JOIN Projproduto P ON P.Codproduto = PC.Codproduto"
                    TextoFiltro = "PC.Codproduto = " & .ListItems(InitFor) & " and PC.Idcliente = " & .ListItems(InitFor).ListSubItems(9)
                Else
                    INNERJOINTEXTO = "PF.* from Projproduto_fornecedor PF INNER JOIN Projproduto P ON P.Codproduto = PF.Codproduto"
                    TextoFiltro = "PF.Codproduto = " & .ListItems(InitFor) & " and PF.Idfornecedor = " & .ListItems(InitFor).ListSubItems(9)
                End If
                TBFI.Open "Select " & INNERJOINTEXTO & " where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            Else
                TBFI.Open "Select * from projproduto where Codproduto = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            End If
            
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If Vendas_Atualização_Valores = True Then
                        If Chk_atualizar_com_mesmo_valor.Value = 1 And Chk_valores_diferentes.Value = 0 And cmbfiltrarpor <> "Cliente" And cmbfiltrarpor <> "Grupo do cliente" And cmbfiltrarpor <> "Fornecedor" Then
                            NovoValor = Replace(TBFI!PConsumo, ",", ".")
                            NovoValor1 = Replace(TBFI!PRevenda, ",", ".")
                            Set TBFIltro = CreateObject("adodb.recordset")
                            TBFIltro.Open "Select * from Projproduto_clientes where Codproduto = " & .ListItems(InitFor) & " and (PConsumo = " & NovoValor & " or PRevenda = " & NovoValor1 & ")", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFIltro.EOF = False Then
                                Do While TBFIltro.EOF = False
                                    If Cmb_atualizar_por = "Percentual" Then
                                        If Cmb_atualizar_para = "Mais" Then
                                            If Chk_atualizar_vc.Value = 1 And TBFI!PConsumo = TBFIltro!PConsumo Then TBFIltro!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PConsumo + ((TBFIltro!PConsumo * valor) / 100))
                                            If Chk_atualizar_vr.Value = 1 And TBFI!PRevenda = TBFIltro!PRevenda Then TBFIltro!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PRevenda + ((TBFIltro!PRevenda * valor) / 100))
                                        Else
                                            If Chk_atualizar_vc.Value = 1 And TBFI!PConsumo = TBFIltro!PConsumo Then TBFIltro!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PConsumo - ((TBFIltro!PConsumo * valor) / 100))
                                            If Chk_atualizar_vr.Value = 1 And TBFI!PRevenda = TBFIltro!PRevenda Then TBFIltro!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PRevenda - ((TBFIltro!PRevenda * valor) / 100))
                                        End If
                                    Else
                                        If Cmb_atualizar_para = "Mais" Then
                                            If Chk_atualizar_vc.Value = 1 And TBFI!PConsumo = TBFIltro!PConsumo Then TBFIltro!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PConsumo + valor)
                                            If Chk_atualizar_vr.Value = 1 And TBFI!PRevenda = TBFIltro!PRevenda Then TBFIltro!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PRevenda + valor)
                                        ElseIf Cmb_atualizar_para = "Menos" Then
                                                If Chk_atualizar_vc.Value = 1 And TBFI!PConsumo = TBFIltro!PConsumo Then TBFIltro!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PConsumo - valor)
                                                If Chk_atualizar_vr.Value = 1 And TBFI!PRevenda = TBFIltro!PRevenda Then TBFIltro!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PRevenda - valor)
                                            Else
                                                If Chk_atualizar_vc.Value = 1 And TBFI!PConsumo = TBFIltro!PConsumo Then TBFIltro!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, valor)
                                                If Chk_atualizar_vr.Value = 1 And TBFI!PRevenda = TBFIltro!PRevenda Then TBFIltro!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, valor)
                                        End If
                                    End If
                                    TBFIltro.Update
                                    TBFIltro.MoveNext
                                Loop
                            End If
                            TBFIltro.Close
                        End If
                        
                        If Cmb_atualizar_por = "Percentual" Then
                            If Cmb_atualizar_para = "Mais" Then
                                If Chk_atualizar_vc.Value = 1 Then TBFI!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PConsumo + ((TBFI!PConsumo * valor) / 100))
                                If Chk_atualizar_vr.Value = 1 Then TBFI!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PRevenda + ((TBFI!PRevenda * valor) / 100))
                            Else
                                If Chk_atualizar_vc.Value = 1 Then TBFI!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PConsumo - ((TBFI!PConsumo * valor) / 100))
                                If Chk_atualizar_vr.Value = 1 Then TBFI!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PRevenda - ((TBFI!PRevenda * valor) / 100))
                            End If
                        Else
                            If Cmb_atualizar_para = "Mais" Then
                                If Chk_atualizar_vc.Value = 1 Then
                                    If TBFI!PConsumo <> 0 Or TBFI!PConsumo = 0 And Permitido1 = True Then TBFI!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PConsumo + valor)
                                End If
                                If Chk_atualizar_vr.Value = 1 Then
                                    If TBFI!PRevenda <> 0 Or TBFI!PRevenda = 0 And Permitido1 = True Then TBFI!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PRevenda + valor)
                                End If
                            ElseIf Cmb_atualizar_para = "Menos" Then
                                    If Chk_atualizar_vc.Value = 1 Then
                                        If TBFI!PConsumo <> 0 Or TBFI!PConsumo = 0 And Permitido1 = True Then TBFI!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PConsumo - valor)
                                    End If
                                    If Chk_atualizar_vr.Value = 1 Then
                                        If TBFI!PRevenda <> 0 Or TBFI!PRevenda = 0 And Permitido1 = True Then TBFI!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PRevenda - valor)
                                    End If
                                Else
                                    If Chk_atualizar_vc.Value = 1 Then
                                        If TBFI!PConsumo <> 0 Or TBFI!PConsumo = 0 And Permitido1 = True Then TBFI!PConsumo = FunFormataCasasDecimais(cmbCasasDecimais, valor)
                                    End If
                                    If Chk_atualizar_vr.Value = 1 Then
                                        If TBFI!PRevenda <> 0 Or TBFI!PRevenda = 0 And Permitido1 = True Then TBFI!PRevenda = FunFormataCasasDecimais(cmbCasasDecimais, valor)
                                    End If
                            End If
                        End If
                    Else
                        If Chk_atualizar_com_mesmo_valor.Value = 1 And Chk_valores_diferentes.Value = 0 And cmbfiltrarpor <> "Fornecedor" Then
                            NovoValor = Replace(TBFI!PCusto, ",", ".")
                            Set TBFIltro = CreateObject("adodb.recordset")
                            TBFIltro.Open "Select * from Projproduto_fornecedor where Codproduto = " & .ListItems(InitFor) & " and PCusto = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFIltro.EOF = False Then
                                Do While TBFIltro.EOF = False
                                    If Cmb_atualizar_por = "Percentual" Then
                                        If Cmb_atualizar_para = "Mais" Then TBFIltro!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PCusto + ((TBFIltro!PCusto * valor) / 100)) Else TBFIltro!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PCusto - ((TBFIltro!PCusto * valor) / 100))
                                    Else
                                        If Cmb_atualizar_para = "Mais" Then
                                            TBFIltro!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PCusto + valor)
                                        ElseIf Cmb_atualizar_para = "Menos" Then
                                                TBFIltro!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFIltro!PCusto - valor)
                                            Else
                                                TBFIltro!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, valor)
                                        End If
                                    End If
                                    TBFIltro.Update
                                    TBFIltro.MoveNext
                                Loop
                            End If
                            TBFIltro.Close
                        End If
                        
                        If Cmb_atualizar_por = "Percentual" Then
                            If Cmb_atualizar_para = "Mais" Then TBFI!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PCusto + ((TBFI!PCusto * valor) / 100)) Else TBFI!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PCusto - ((TBFI!PCusto * valor) / 100))
                        Else
                            If TBFI!PCusto <> 0 Or TBFI!PCusto = 0 And Permitido1 = True Then
                                If Cmb_atualizar_para = "Mais" Then
                                    TBFI!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PCusto + valor)
                                ElseIf Cmb_atualizar_para = "Menos" Then
                                        TBFI!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, TBFI!PCusto - valor)
                                    Else
                                        TBFI!PCusto = FunFormataCasasDecimais(cmbCasasDecimais, valor)
                                End If
                            End If
                        End If
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
            
            '==================================
            Modulo = Formulario
            Evento = "Atualizar valor"
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & .ListItems(InitFor).ListSubItems(1)
            If Chk_valores_diferentes.Value = 1 Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Or cmbfiltrarpor = "Fornecedor" Then
                If Vendas_Atualização_Valores = True Then Documento1 = "Cliente : " & .ListItems(InitFor).ListSubItems(6) Else Documento1 = "Fornecedor : " & .ListItems(InitFor).ListSubItems(6)
            End If
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) registro(s) antes de atualizar os valores."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


