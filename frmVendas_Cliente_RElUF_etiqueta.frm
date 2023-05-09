VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_Cliente_RElUF_etiqueta 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Vendas - Clientes - Gerar etiquetas"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7650
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
   Icon            =   "frmVendas_Cliente_RElUF_etiqueta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   6030
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_Cliente_RElUF_etiqueta.frx":030A
      Count           =   1
   End
   Begin VB.TextBox Txt_ID_nome 
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
      Left            =   1290
      Locked          =   -1  'True
      MouseIcon       =   "frmVendas_Cliente_RElUF_etiqueta.frx":431C
      MousePointer    =   99  'Custom
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "ID nome."
      Top             =   5130
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtid 
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
      Left            =   930
      Locked          =   -1  'True
      MouseIcon       =   "frmVendas_Cliente_RElUF_etiqueta.frx":4626
      MousePointer    =   99  'Custom
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "ID etiqueta."
      Top             =   5130
      Visible         =   0   'False
      Width           =   345
   End
   Begin MSComctlLib.ListView lst_clientes 
      Height          =   2630
      Left            =   60
      TabIndex        =   13
      Top             =   4455
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   4630
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Pos."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   4128
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Contato"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Departamento"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   55
      TabIndex        =   27
      Top             =   1020
      Width           =   7530
      Begin VB.OptionButton Opt_cobranca 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cobrança"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6240
         TabIndex        =   2
         ToolTipText     =   "Física"
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton Opt_principal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Jurídica"
         Top             =   270
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton Opt_entrega 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrega"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3240
         TabIndex        =   1
         ToolTipText     =   "Física"
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   55
      TabIndex        =   14
      Top             =   1620
      Width           =   7530
      Begin VB.ComboBox txtendereco 
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
         ItemData        =   "frmVendas_Cliente_RElUF_etiqueta.frx":4930
         Left            =   1320
         List            =   "frmVendas_Cliente_RElUF_etiqueta.frx":4932
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Endereço."
         Top             =   930
         Width           =   5985
      End
      Begin VB.TextBox Txt_documento 
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
         Left            =   4380
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Documento."
         Top             =   180
         Width           =   2925
      End
      Begin VB.ComboBox cmbposicao 
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
         ItemData        =   "frmVendas_Cliente_RElUF_etiqueta.frx":4934
         Left            =   1320
         List            =   "frmVendas_Cliente_RElUF_etiqueta.frx":4974
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Posição inicial para impressão."
         Top             =   180
         Width           =   1950
      End
      Begin VB.ComboBox cmbcliente 
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
         ItemData        =   "frmVendas_Cliente_RElUF_etiqueta.frx":49C8
         Left            =   1320
         List            =   "frmVendas_Cliente_RElUF_etiqueta.frx":49CA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Cliente."
         Top             =   552
         Width           =   5985
      End
      Begin VB.ComboBox cmbcontato 
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
         ItemData        =   "frmVendas_Cliente_RElUF_etiqueta.frx":49CC
         Left            =   1320
         List            =   "frmVendas_Cliente_RElUF_etiqueta.frx":49CE
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Contato do cliente."
         Top             =   1995
         Width           =   5985
      End
      Begin VB.TextBox txtdepartamento 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Departamento."
         Top             =   2370
         Width           =   5985
      End
      Begin VB.TextBox txtcidade 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Cidade."
         Top             =   1638
         Width           =   3405
      End
      Begin VB.TextBox txtuf 
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
         Left            =   6945
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "UF."
         Top             =   1635
         Width           =   360
      End
      Begin VB.TextBox txtcep 
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
         Left            =   5325
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Cep."
         Top             =   1635
         Width           =   1080
      End
      Begin VB.TextBox txtbairro 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Bairro."
         Top             =   1281
         Width           =   5985
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3375
         TabIndex        =   26
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Posição inicial :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   23
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   645
         TabIndex        =   22
         Top             =   555
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contato :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   555
         TabIndex        =   21
         Top             =   1995
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Departam. :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   20
         Top             =   2370
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   465
         TabIndex        =   19
         Top             =   930
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   645
         TabIndex        =   18
         Top             =   1635
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "UF :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   6570
         TabIndex        =   17
         Top             =   1635
         Width           =   300
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   4845
         TabIndex        =   16
         Top             =   1635
         Width           =   390
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   15
         Top             =   1281
         Width           =   525
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   1720
      ButtonCount     =   8
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Novo"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Novo (Insert)"
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
      ButtonWidth1    =   36
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Salvar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Salvar (F3)"
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
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   44
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Excluir"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Excluir (F4)"
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
      ButtonLeft3     =   86
      ButtonTop3      =   2
      ButtonWidth3    =   45
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Relatório"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Relatório (F5)"
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
      ButtonLeft4     =   133
      ButtonTop4      =   2
      ButtonWidth4    =   60
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonAlignment5=   2
      ButtonType5     =   1
      ButtonStyle5    =   -1
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   -1
      ButtonLeft5     =   195
      ButtonTop5      =   4
      ButtonWidth5    =   2
      ButtonHeight5   =   54
      ButtonCaption6  =   "Ajuda"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Ajuda (F1)"
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   199
      ButtonTop6      =   2
      ButtonWidth6    =   41
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Sair"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Sair (Esc)"
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   242
      ButtonTop7      =   2
      ButtonWidth7    =   30
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState8    =   5
      ButtonLeft8     =   274
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   29
      Top             =   7080
      Width           =   7530
      _ExtentX        =   13282
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
End
Attribute VB_Name = "frmVendas_Cliente_RElUF_etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Cliente_Etiqueta As Boolean 'OK

Private Sub cmbcliente_Click()
On Error GoTo tratar_erro

ProcCarregaDadosCliente
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosCliente()
On Error GoTo tratar_erro

If cmbcliente = "" Then Exit Sub
txtBairro.Text = ""
txtCidade.Text = ""
txtCEP.Text = ""
txtuf.Text = ""
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes where IDCliente = " & cmbcliente.ItemData(cmbcliente.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    Txt_ID_nome = IIf(IsNull(TBClientes!IDCliente), 0, TBClientes!IDCliente)
    If Opt_principal.Value = True Then
        ProcCarregaEndPrincipal
    ElseIf Opt_entrega.Value = True Then
            ProcCarregaEndEntrega
        Else
            ProcCarregaEndCobranca
    End If
    Nome = ""
    cmbcontato.Clear
    txtdepartamento.Text = ""
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Clientes_Contatos where idcliente = " & TBClientes!IDCliente & " and nomecontato <> 'Null' order by nomecontato", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If Nome <> TBAbrir!NomeContato Then
                cmbcontato.AddItem TBAbrir!NomeContato
                cmbcontato.ItemData(cmbcontato.NewIndex) = TBAbrir!idcontato
            End If
            Nome = TBAbrir!NomeContato
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEndPrincipal()
On Error GoTo tratar_erro

txtendereco.Clear
If IsNull(TBClientes!Endereco) = False And TBClientes!Endereco <> "" Then
    If IsNull(TBClientes!complemento) = False And TBClientes!complemento <> "" Then ComplementoTexto = TBClientes!complemento Else ComplementoTexto = ""
    txtendereco.AddItem TBClientes!Endereco & "," & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & IIf(ComplementoTexto <> "", "," & ComplementoTexto, "")
    txtendereco.Text = TBClientes!Endereco & "," & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & IIf(ComplementoTexto <> "", "," & ComplementoTexto, "")
End If
txtBairro.Text = IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
txtCidade.Text = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
txtCEP.Text = IIf(IsNull(TBClientes!CEP), "", TBClientes!CEP)
txtuf.Text = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEndEntrega()
On Error GoTo tratar_erro

txtendereco.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from clientes_entrega where idcliente = " & TBClientes!IDCliente & " and Tipo = 'C' order by endereco_entrega", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If IsNull(TBAbrir!endereco_entrega) = False And TBAbrir!endereco_entrega <> "" Then
            If IsNull(TBAbrir!complemento) = False And TBAbrir!complemento <> "" Then ComplementoTexto = TBAbrir!complemento Else ComplementoTexto = ""
            txtendereco.AddItem TBAbrir!endereco_entrega & "," & IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero) & IIf(ComplementoTexto <> "", "," & ComplementoTexto, "")
            txtendereco.ItemData(txtendereco.NewIndex) = TBAbrir!identrega
        End If
        TBAbrir.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEndCobranca()
On Error GoTo tratar_erro

txtendereco.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from clientes_cobranca where idcliente = " & TBClientes!IDCliente & " and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If IsNull(TBAbrir!endereco_Cobranca) = False And TBAbrir!endereco_Cobranca <> "" Then
            If IsNull(TBAbrir!complemento) = False And TBAbrir!complemento <> "" Then ComplementoTexto = TBAbrir!complemento Else ComplementoTexto = ""
            txtendereco.AddItem TBAbrir!endereco_Cobranca & "," & IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero) & IIf(ComplementoTexto <> "", "," & ComplementoTexto, "")
            txtendereco.ItemData(txtendereco.NewIndex) = TBAbrir!idCobranca
        End If
        TBAbrir.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbcontato_Click()
On Error GoTo tratar_erro

If cmbcontato = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Clientes_Contatos where IDContato = " & cmbcontato.ItemData(cmbcontato.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtdepartamento.Text = IIf(IsNull(TBAbrir!Departamento), "", TBAbrir!Departamento)
End If
TBAbrir.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRelatorio()
On Error GoTo tratar_erro

If lst_clientes.ListItems.Count = 0 Then Exit Sub
frmVendas_Cliente_RElUF_etiqueta_menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcRelatorio
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7530, 8, True
cmbcliente.Clear
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes where DtValidacao IS NOT NULL order by nomerazao", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    Do While TBClientes.EOF = False
        cmbcliente.AddItem TBClientes!NomeRazao
        cmbcliente.ItemData(cmbcliente.NewIndex) = TBClientes!IDCliente
        TBClientes.MoveNext
    Loop
End If
TBClientes.Close
ProcCarregaLista
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With lst_clientes
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) etiqueta(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from etiqueta where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Clientes/Etiqueta"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Cliente: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) etiqueta(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Etiqueta(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparCampos
    ProcCarregaLista
    Frame1.Enabled = False
    Novo_Cliente_Etiqueta = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimparCampos
Novo_Cliente_Etiqueta = True
Frame1.Enabled = True
cmbposicao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Cliente_Etiqueta = True Then
    If USMsgBox("A etiqueta ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Cliente_Etiqueta = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Cliente_Etiqueta = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbposicao.Text = "" Then
    USMsgBox ("Informe a posição inicial antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmbposicao.SetFocus
    Exit Sub
End If
If cmbcliente.Text = "" Then
    USMsgBox ("Informe o cliente antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmbcliente.SetFocus
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from etiqueta where Tipo = 'C' and Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and posicao = " & cmbposicao.Text & " and id <> " & txtId & " and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Informe outro número para posição inicial, pois esta posição está sendo utilizada."), vbExclamation, "CAPRIND v5.0"
    cmbposicao.SetFocus
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close
ProcGravar
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravar()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from etiqueta where id = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
If Novo_Cliente_Etiqueta = True Then
    Posicao = cmbposicao.Text
    Cont = cmbposicao.Text
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select posicao from etiqueta where posicao <> 0 and posicao <> " & Cont & " and Tipo = 'C' and Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' order by posicao", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
       TBFI.MoveLast
        If TBFI!Posicao < Cont Then
            i = TBFI!Posicao + 1
            If i >= 20 Then i = 1
        Else
            i = 1
        End If
    Else
        i = 1
    End If
    Do While i <> Cont
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from etiqueta where posicao = " & i & " and Tipo = 'C' and Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Posicao = i
        TBGravar!Tipo = "C"
        TBGravar!Modulo = Formulario
        TBGravar!Responsavel = pubUsuario
        TBGravar.Update
        TBGravar.Close
        i = i + 1
    Loop
    TBFI.Close
End If

TBAbrir!ID_nome = Txt_ID_nome
TBAbrir!Tipo = "C"
TBAbrir!Documento = Txt_documento
TBAbrir!Nome = cmbcliente.Text
TBAbrir!Endereco = txtendereco
TBAbrir!Bairro = txtBairro
TBAbrir!Cidade = txtCidade
TBAbrir!CEP = txtCEP
TBAbrir!UF = txtuf
TBAbrir!contato = IIf(cmbcontato = "", Null, cmbcontato)
TBAbrir!Departamento = IIf(txtdepartamento = "", Null, txtdepartamento)
TBAbrir!Posicao = cmbposicao
TBAbrir!Modulo = Formulario
TBAbrir!Responsavel = pubUsuario
TBAbrir.Update
txtId = TBAbrir!ID
TBAbrir.Close
ProcCarregaLista
If Novo_Cliente_Etiqueta = True Then
    USMsgBox ("Nova etiqueta cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And lst_clientes.ListItems.Count <> 0 Then
        lst_clientes.SelectedItem = lst_clientes.ListItems(CodigoLista)
        lst_clientes.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Clientes/Etiqueta"
Evento = "Salvar"
ID_documento = txtId
Documento = "Cliente: " & cmbcliente
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Cliente_Etiqueta = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

txtId.Text = 0
Txt_ID_nome = 0
With cmbposicao
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With
Txt_documento = ""
cmbcliente.ListIndex = -1
txtendereco.ListIndex = -1
txtBairro.Text = ""
txtCidade.Text = ""
txtCEP.Text = ""
txtuf.Text = ""
cmbcontato.ListIndex = -1
txtdepartamento.Text = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lst_clientes.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from etiqueta where Tipo = 'C' and Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' order by posicao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With lst_clientes.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Posicao), "", FunTamanhoTextoZeroEsq(TBLISTA!Posicao, 2))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Nome), "", TBLISTA!Nome)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!contato), "", TBLISTA!contato)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Departamento), "", TBLISTA!Departamento)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do error : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lst_clientes
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView lst_clientes, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_clientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lst_clientes.ListItems.Count = 0 Then Exit Sub
ProcLimparCampos
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from etiqueta where id = " & lst_clientes.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtId.Text = lst_clientes.SelectedItem
    If lst_clientes.SelectedItem.ListSubItems(1) <> 0 Then
        Txt_ID_nome = IIf(IsNull(TBLISTA!ID_nome), "", TBLISTA!ID_nome)
        With cmbposicao
            .Text = lst_clientes.SelectedItem.ListSubItems(1)
            .Locked = True
            .TabStop = False
        End With
        Txt_documento = IIf(IsNull(TBLISTA!Documento), "", TBLISTA!Documento)
        If IsNull(TBLISTA!Nome) = False And TBLISTA!Nome <> "" Then cmbcliente.Text = TBLISTA!Nome
        If IsNull(TBLISTA!Endereco) = False And TBLISTA!Endereco <> "" Then txtendereco = TBLISTA!Endereco
        txtBairro.Text = IIf(IsNull(TBLISTA!Bairro), "", TBLISTA!Bairro)
        txtCidade.Text = IIf(IsNull(TBLISTA!Cidade), "", TBLISTA!Cidade)
        txtCEP.Text = IIf(IsNull(TBLISTA!CEP), "", TBLISTA!CEP)
        txtuf.Text = IIf(IsNull(TBLISTA!UF), "", TBLISTA!UF)
        If IsNull(TBLISTA!contato) = False And TBLISTA!contato <> "" Then cmbcontato.Text = TBLISTA!contato
        txtdepartamento.Text = IIf(IsNull(TBLISTA!Departamento), "", TBLISTA!Departamento)
    End If
    Novo_Cliente_Etiqueta = False
    CodigoLista = lst_clientes.SelectedItem.index
End If
TBLISTA.Close
Frame1.Refresh
Frame1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do error : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_cobranca_Click()
On Error GoTo tratar_erro

ProcCarregaDadosCliente

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do error : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_entrega_Click()
On Error GoTo tratar_erro

ProcCarregaDadosCliente

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do error : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_principal_Click()
On Error GoTo tratar_erro

ProcCarregaDadosCliente

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do error : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtendereco_Click()
On Error GoTo tratar_erro

If txtendereco <> "" Then
    If Opt_entrega.Value = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from clientes_entrega where identrega = " & txtendereco.ItemData(txtendereco.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtBairro.Text = IIf(IsNull(TBAbrir!bairro_entrega), "", TBAbrir!bairro_entrega)
            txtCidade.Text = IIf(IsNull(TBAbrir!cidade_entrega), "", TBAbrir!cidade_entrega)
            txtCEP.Text = IIf(IsNull(TBAbrir!cep_entrega), "", TBAbrir!cep_entrega)
            txtuf.Text = IIf(IsNull(TBAbrir!uf_entrega), "", TBAbrir!uf_entrega)
        End If
        TBAbrir.Close
    ElseIf Opt_cobranca.Value = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from clientes_cobranca where idcobranca = " & txtendereco.ItemData(txtendereco.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                txtBairro.Text = IIf(IsNull(TBAbrir!bairro_Cobranca), "", TBAbrir!bairro_Cobranca)
                txtCidade.Text = IIf(IsNull(TBAbrir!cidade_Cobranca), "", TBAbrir!cidade_Cobranca)
                txtCEP.Text = IIf(IsNull(TBAbrir!cep_Cobranca), "", TBAbrir!cep_Cobranca)
                txtuf.Text = IIf(IsNull(TBAbrir!uf_Cobranca), "", TBAbrir!uf_Cobranca)
            End If
            TBAbrir.Close
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do error : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: ProcRelatorio
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

