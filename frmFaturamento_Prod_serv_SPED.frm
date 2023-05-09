VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Prod_serv_SPED 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Faturamento - SPED Fiscal Bloco K"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   16875
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   16875
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escolha a empresa"
      Height          =   825
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   4815
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
         ItemData        =   "frmFaturamento_Prod_serv_SPED.frx":0000
         Left            =   180
         List            =   "frmFaturamento_Prod_serv_SPED.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Empresa."
         Top             =   330
         Width           =   4425
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções SPED Bloco K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7035
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   4815
      Begin DrawSuite2022.USCheckBox chk200 
         Height          =   375
         Left            =   210
         TabIndex        =   8
         Top             =   1350
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 200 - Estoque escriturado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         Value           =   1
      End
      Begin DrawSuite2022.USCheckBox chk210 
         Height          =   375
         Left            =   210
         TabIndex        =   9
         Top             =   1785
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 210 - Desmontagem (Origem)"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk215 
         Height          =   375
         Left            =   210
         TabIndex        =   10
         Top             =   2235
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 215 - Desmontagem (Destino)"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk220 
         Height          =   375
         Left            =   210
         TabIndex        =   11
         Top             =   2670
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 220 - Outras movimentações internas"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk230 
         Height          =   375
         Left            =   210
         TabIndex        =   12
         Top             =   3105
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 230 - Itens produzidos"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk235 
         Height          =   375
         Left            =   210
         TabIndex        =   13
         Top             =   3540
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 235 - Insumos consumidos"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk250 
         Height          =   375
         Left            =   210
         TabIndex        =   14
         Top             =   3975
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 250 - Industrialização por terceiros"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox CHK275 
         Height          =   375
         Left            =   210
         TabIndex        =   15
         Top             =   6165
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   661
         Caption         =   "Bloco K 275 - Correções apontamentos ret. insumos"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk270 
         Height          =   375
         Left            =   210
         TabIndex        =   16
         Top             =   5730
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   661
         Caption         =   "Bloco K 270 - Correções (K210,K220,K230,K250,K260)"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk265 
         Height          =   375
         Left            =   210
         TabIndex        =   17
         Top             =   5280
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 265 - Reparo (Retrabalho - Insumos)"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk260 
         Height          =   375
         Left            =   210
         TabIndex        =   18
         Top             =   4845
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 260 - Reparo (Retrabalho)"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk255 
         Height          =   375
         Left            =   210
         TabIndex        =   19
         Top             =   4410
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         Caption         =   "Bloco K 255 - Industrialização por terceiros (Insumos)"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox chk01 
         Height          =   375
         Left            =   210
         TabIndex        =   20
         Top             =   480
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 001 - Abertura do Bloco K"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         Value           =   1
      End
      Begin DrawSuite2022.USCheckBox chk100 
         Height          =   375
         Left            =   210
         TabIndex        =   21
         Top             =   915
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   661
         Caption         =   "Bloco K 100 - Período de apuração do ICMS e IPI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         Value           =   1
      End
      Begin DrawSuite2022.USCheckBox chk990 
         Height          =   375
         Left            =   210
         TabIndex        =   22
         Top             =   6600
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         Caption         =   "Bloco K 990 - Encerramento do Bloco K"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         Value           =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gerar SPED Bloco K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1905
      Left            =   0
      TabIndex        =   0
      Top             =   7890
      Width           =   4815
      Begin VB.ComboBox cmbAno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmFaturamento_Prod_serv_SPED.frx":0004
         Left            =   3615
         List            =   "frmFaturamento_Prod_serv_SPED.frx":0014
         TabIndex        =   25
         Top             =   420
         Width           =   975
      End
      Begin DrawSuite2022.USButton cmdGerar 
         Height          =   825
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Gerar SPED Bloco K"
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1455
         DibPicture      =   "frmFaturamento_Prod_serv_SPED.frx":0030
         Caption         =   "Gerar SPED Bloco K"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
         ToolTipTitle    =   "Caprind V5.0"
      End
      Begin VB.ComboBox cmbMes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmFaturamento_Prod_serv_SPED.frx":9ADD
         Left            =   1830
         List            =   "frmFaturamento_Prod_serv_SPED.frx":9B05
         TabIndex        =   5
         Top             =   420
         Width           =   1335
      End
      Begin DrawSuite2022.USButton cmdSair 
         Height          =   825
         Left            =   2400
         TabIndex        =   23
         ToolTipText     =   "Sair do módulo de geração arquivo Sped Bloco K"
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1455
         DibPicture      =   "frmFaturamento_Prod_serv_SPED.frx":9B6E
         Caption         =   "Sair"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   3
         ToolTipTitle    =   "Caprind V5.0"
      End
      Begin DrawSuite2022.USCheckBox chkBloco200 
         Height          =   375
         Left            =   180
         TabIndex        =   24
         Top             =   390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Bloco 0200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   -1  'True
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
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
         Index           =   0
         Left            =   3255
         TabIndex        =   26
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Index           =   2
         Left            =   11730
         TabIndex        =   2
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
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
         Index           =   3
         Left            =   1500
         TabIndex        =   1
         Top             =   480
         Width           =   285
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   9810
      Width           =   16845
      _ExtentX        =   29713
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
   Begin MSComctlLib.ListView Lista 
      Height          =   9795
      Left            =   4830
      TabIndex        =   4
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   17277
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "REG"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "DT_EST"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "COD_ITEM"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DESC_ITEM"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "QTD"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "IND_EST"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "COD_PART"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "COD"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "DESC_COD"
         Object.Width           =   3528
      EndProperty
   End
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
      FormHeightDT    =   10530
      FormWidthDT     =   16995
      FormScaleHeightDT=   10065
      FormScaleWidthDT=   16875
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frmFaturamento_Prod_serv_SPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mmes As Integer

Private Sub ProcSPED()
On Error GoTo tratar_erro

If cmbMes.Text = "" Then
  USMsgBox "Favor escolher o mês para filtrar o Bloco K", vbCritical, "CAPRIND v5.0"
  cmbMes.SetFocus
  Exit Sub
End If

If cmbAno.Text = "" Then
  USMsgBox "Favor escolher o ano para filtrar o Bloco K", vbCritical, "CAPRIND v5.0"
  cmbAno.SetFocus
  Exit Sub
End If


If USMsgBox("Deseja realmente gerar o SPED bloco K?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    
'Verifica se existe a pasta para salvar o arquivo
    If FileOrDirExists(Localrel & "\Arquivos exportados\SPED enviar") = False Then
        MkDir (Localrel & "\Arquivos exportados\SPED enviar")
        'USMsgBox ("Não é permitido gerar o SPED, pois não foi encontrado o caminho " & Localrel & "\Arquivos exportados\SPED enviar, onde será armazenado os aquivos."), vbExclamation, "CAPRIND v5.0"
        'Exit Sub
    End If

    ProcGerarSPEDBlocoK

    USMsgBox ("SPED Bloco K gerado com sucesso."), vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub PrimeiroUltimoDiaMes(DataMes As Date)
On Error GoTo tratar_erro
      
    Dim primeiro   As Date
    Dim Ultimo   As Date
      
    'Usamos a função DAteSerial para obter o primeiro e o último dia
    primeiro = DateSerial(Year(DataMes), Month(DataMes) + 0, 1)
    Ultimo = DateSerial(Year(DataMes), Month(DataMes) + 1, 0)
      
    USMsgBox "Primeiro dia:" & Primer & vbNewLine & _
           "Último dia:" & Ultimo, vbInformation, "CAPRIND v5.0"
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGerarSPEDBlocoK()
On Error GoTo tratar_erro
Lista.ListItems.Clear
Dim countLinha As Double
countLinha = 0

PBLista.Min = 0
PBLista.Max = 4
PBLista.Value = 0
DataMes = "01/" & Mmes & "/" & cmbAno.Text
'=========================================================
' Acertar ID Tipo na movimentação do estoque
'=========================================================
Conexao.Execute ("update Estoque_movimentacao Set ID_Tipo = PP.ID_Tipo from Estoque_movimentacao as EM inner join projproduto as PP on EM.Desenho= PP.desenho")
'=========================================================
'============================================
' Criar nome do arquivo do SPED Bloco K
'============================================
Set ArqTXT = GerArqPastas.CreateTextFile(Localrel & "\Arquivos exportados\SPED enviar\SPEDF_MES_" & cmbMes & ".txt", True)
With ArqTXT
'============================================
' Criar cabeçalho com nome da informante
'============================================
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Razao, CNPJ, UF, cidade, IE, IM, Codigo_SUFRAMA from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        FamiliaAntiga = RemoveAccents(TBAbrir!Cidade)
        FamiliaAntiga = FunVerificaCodMunicipio(FamiliaAntiga, TBAbrir!UF)
        .WriteLine "|0000|012|0|" & ReturnNumbersOnly(DateSerial(Year(DataMes), Month(DataMes) + 0, 1)) & "|" & ReturnNumbersOnly(DateSerial(Year(DataMes), Month(DataMes) + 1, 0)) & "|" & TBAbrir!Razao & "|" & DS.ReturnNumbersOnly(TBAbrir!CNPJ) & "||" & TBAbrir!UF & "|" & IIf(IsNull(TBAbrir!IE), "", DS.ReturnNumbersOnly(TBAbrir!IE)) & "|" & FamiliaAntiga & "|" & IIf(IsNull(TBAbrir!IM), "", TBAbrir!IM) & "|" & IIf(IsNull(TBAbrir!Codigo_SUFRAMA), "", TBAbrir!Codigo_SUFRAMA) & "|A|0|"
        countLinha = countLinha + 1
    End If
    TBAbrir.Close
    
'============================================
'Bloco 0190 com as unidades do bloco 200
'============================================
    Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "SELECT UM.Unidade, UM.Descricao from Unidade_Medida UM inner join projproduto PP on PP.Unidade = UM.Unidade GROUP BY UM.UNIDADE, UM.Descricao"
     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
     If TBAbrir.EOF = False Then
     .WriteLine "|0001|0|"
     countLinha = countLinha + 1
        Do While TBAbrir.EOF = False
            .WriteLine "|0190|" & TBAbrir!Unidade & "|" & TBAbrir!Descricao & "|"
            countLinha = countLinha + 1
            TBAbrir.MoveNext

        Loop
     End If
     TBAbrir.Close

'============================================
' Bloco 0200 de acordo com bloco K200
'============================================
If chkBloco200.Value = Checked Then
'=======================================
DataResultado = DateSerial(Year(DataMes), Month(DataMes) + 1, 0)
Var = ""
Var1 = "07"
Set TBCodigoDesc = CreateObject("adodb.recordset")
 TBCodigoDesc.Open "Select * from Projproduto_Tipo where Codigo = " & Var1 & "", Conexao, adOpenKeyset, adLockOptimistic
   If TBCodigoDesc.EOF = False Then
     Var = TBCodigoDesc!ID
   End If
   TBCodigoDesc.Close

Vdata = Format(DataResultado, "YYYYMMDD")

StrSql = "select PPT.Codigo as Cod_item ,PPT.Descricao as Desc_cod,EM.Desenho,EM.Descricao,sum(Entrada) as TTEntrada, sum(Saida) as TTSaida,Sum(Entrada) - Sum(Saida) as Saldo from Estoque_movimentacao AS EM INNER JOIN Projproduto_Tipo AS PPT ON EM.ID_Tipo = PPT.ID where EM.ID_Tipo < '" & Int(Var) & "' AND  DATA <= '" & Vdata & "' GROUP BY EM.Desenho,EM.Descricao, EM.ID_Tipo,PPT.Codigo,PPT.Descricao"
'Debug.print StrSql

Contador = 2
Dim Saldo As Integer
Dim Total200 As Long

Saldo = 0
Desenho = ""
Total200 = 0

Set TBAbrir_NFe = CreateObject("adodb.recordset")
TBAbrir_NFe.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir_NFe.EOF = False Then
'Total200 = TBAbrir_NFe.RecordCount
  Do While TBAbrir_NFe.EOF = False
   Set TBAbrir = CreateObject("adodb.recordset")
   StrSql = "SELECT projproduto.Desenho AS COD_ITEM, projproduto.descricao AS DESCR_ITEM, projproduto.Unidade AS UNID_INV, Projproduto_Tipo.Codigo AS Tipo_ITEM, tbl_ClassificacaoFiscal.IDIntClasse AS NCM, tbl_ClassificacaoFiscal.dbl_ICMS_de AS ALIQ_ICMS, Projproduto_Genero.Codigo AS COD_GEN, tbl_ClassificacaoFiscal.CEST AS TIPI, projproduto.GTIN, projproduto.Cod_servico_NFSE AS COD_LST FROM projproduto INNER JOIN Projproduto_Tipo ON projproduto.ID_Tipo = Projproduto_Tipo.ID INNER JOIN tbl_ClassificacaoFiscal ON projproduto.ID_CF = tbl_ClassificacaoFiscal.Idclass LEFT JOIN Projproduto_Genero ON projproduto.ID_Genero = Projproduto_Genero.ID WHERE projproduto.Desenho = '" & TBAbrir_NFe!Desenho & "'" 'Projproduto_tipo.Codigo <> '07' and Projproduto_tipo.Codigo <> '08' and Projproduto_tipo.Codigo <> '09' and Projproduto_tipo.Codigo <>'10' and Projproduto_tipo.Codigo <>'99' order by projproduto.desenho"
'Debug.print StrSql

     TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
     If TBAbrir.EOF = False And TBAbrir_NFe!Saldo > Saldo Then
             If TBAbrir!COD_ITEM <> Desenho Then
                 .WriteLine "|0200|" & Trim(TBAbrir!COD_ITEM) & "|" & Trim(TBAbrir!DESCR_ITEM) & "|" & TBAbrir!GTIN & "||" & TBAbrir!UNID_INV & "|" & TBAbrir!Tipo_ITEM & "|" & ReturnNumbersOnly(TBAbrir!NCM) & "||" & Left(TBAbrir!NCM, 2) & "|" & TBAbrir!COD_LST & "|" & Format(TBAbrir!ALIQ_ICMS, "0.00") & "|" & TBAbrir!TIPI & "|"
                countLinha = countLinha + 1
             End If
             
             Desenho = TBAbrir!COD_ITEM
     End If
     TBAbrir.Close
   TBAbrir_NFe.MoveNext
 Loop
End If
.WriteLine "|0990|" & countLinha + 1 & "|"
End If
countLinha = 0
'============================================
' Inicio do Bloco K
'============================================
' Identificação do bloco K
'============================================
            With Lista.ListItems
                .Add , , "K001|0|"
            End With
                .WriteLine "|K001|0|"
                countLinha = countLinha + 1
            With Lista.ListItems
                .Add , , "K100"
                .Item(.Count).SubItems(1) = DateSerial(Year(DataMes), Month(DataMes) + 0, 1)
                .Item(.Count).SubItems(2) = DateSerial(Year(DataMes), Month(DataMes) + 1, 0)
            End With
                .WriteLine "|K100|" & ReturnNumbersOnly(DateSerial(Year(DataMes), Month(DataMes) + 0, 1)) & "|" & ReturnNumbersOnly(DateSerial(Year(DataMes), Month(DataMes) + 1, 0)) & "|"
                countLinha = countLinha + 1
'===========================================
' Inicio do K200 Estoque escriturado
'===========================================
' Pega ultimo dia do mês pesquisado
'===========================================
    DataResultado = DateSerial(Year(DataMes), Month(DataMes) + 1, 0)
'===========================================
' Filtra Mes na movimentação do estoque
'===========================================
' Localiza
'===========================================
Var = ""
Var1 = "07"
 Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select * from Projproduto_Tipo where Codigo = " & Var1 & "", Conexao, adOpenKeyset, adLockOptimistic
      If TBCodigoDesc.EOF = False Then
        Var = TBCodigoDesc!ID
      End If
      TBCodigoDesc.Close
      
'Dim Vdata

Vdata = Format(DataResultado, "YYYYMMDD")

' StrSql = "select PPT.Codigo as Cod_item ,PPT.Descricao as Desc_cod,EM.Desenho,EM.Descricao,sum(Entrada) as TTEntrada, sum(Saida) as TTSaida,Sum(Entrada) - Sum(Saida) as Saldo from Estoque_movimentacao AS EM INNER JOIN Projproduto_Tipo AS PPT ON EM.ID_Tipo = PPT.ID where EM.ID_Tipo < '" & Int(Var) & "' AND  DATA <= '" & Vdata & "' GROUP BY EM.Desenho,EM.Descricao, EM.ID_Tipo,PPT.Codigo,PPT.Descricao"
 StrSql = "select PPT.Codigo as Cod_item ,PPT.Descricao as Desc_cod,EM.Desenho,sum(Entrada) as TTEntrada, sum(Saida) as TTSaida,Sum(Entrada) - Sum(Saida) as Saldo from Estoque_movimentacao AS EM INNER JOIN Projproduto_Tipo AS PPT ON EM.ID_Tipo = PPT.ID where EM.ID_Tipo < '" & Int(Var) & "' AND  DATA <= '" & Vdata & "' GROUP BY EM.Desenho, EM.ID_Tipo,PPT.Codigo,PPT.Descricao"
 
 'Debug.print StrSql
 
 Contador = 2
 Set TBAbrir = CreateObject("adodb.recordset")
 TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
 If TBAbrir.EOF = False Then
 'Debug.print TBAbrir.RecordCount
     Do While TBAbrir.EOF = False
    ' Saldo = "0,001"
         With Lista.ListItems
         If Format(TBAbrir!Saldo, "0.000") > Saldo Then
             .Add , , "K200"
             .Item(.Count).SubItems(1) = DateSerial(Year(DataMes), Month(DataMes) + 1, 0)
             .Item(.Count).SubItems(2) = Trim(TBAbrir!Desenho)
             '.Item(.Count).SubItems(3) = Trim(TBAbrir!Descricao)
             .Item(.Count).SubItems(4) = Format(TBAbrir!Saldo, "0.000")
             .Item(.Count).SubItems(5) = "0"
             .Item(.Count).SubItems(7) = Trim(TBAbrir!COD_ITEM)
             .Item(.Count).SubItems(8) = Trim(TBAbrir!Desc_cod)
         End If
         End With
         If Format(TBAbrir!Saldo, "0.000") > Saldo Then
         .WriteLine "|K200|" & ReturnNumbersOnly(DateSerial(Year(DataMes), Month(DataMes) + 1, 0)) & "|" & Trim(TBAbrir!Desenho) & "|" & Format(TBAbrir!Saldo, "0.000") & "|0||"
         countLinha = countLinha + 1
         End If
         Contador = Contador + 1
         TBAbrir.MoveNext
         QtdeSaida = 0
         Qtde = 0
     Loop
 End If
'    'BK_RK200_Estoque escriturado (Antigo)
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from SPEDF_BK_RK200", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        Do While TBAbrir.EOF = False
'            .WriteLine "|K200|" & TBAbrir!DT_EST & "|" & Trim(TBAbrir!COD_ITEM) & "|" & Format(TBAbrir!Qtd, "0.000") & "|" & TBAbrir!IND_EST & "|" & TBAbrir!COD_PART & "|"
'            TBAbrir.MoveNext
'        Loop
'    End If
    
'    'BK_RK220_Outras movimentações internas entre mercadorias
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from SPEDF_BK_RK220", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        Do While TBAbrir.EOF = False
'            .WriteLine "|K220|" & TBAbrir!DT_MOV & "|" & Trim(TBAbrir!COD_ITEM_ORI) & "|" & Trim(TBAbrir!COD_ITEM_DEST) & "|" & Format(TBAbrir!Qtd, "0.000") & "|"
'            TBAbrir.MoveNext
'        Loop
'    End If
'    'BK_RK230_Itens produzidos
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from SPEDF_BK_RK230", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        Do While TBAbrir.EOF = False
'            .WriteLine "|K230|" & TBAbrir!DT_INI_OP & "|" & TBAbrir!DT_FIN_OP & "|" & TBAbrir!COD_DOC_OP & "|" & Trim(TBAbrir!COD_ITEM) & "|" & TBAbrir!QTD_ENC & "|"
'            TBAbrir.MoveNext
'        Loop
'    End If
'    'BK_RK235_Insumos consumidos
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from SPEDF_BK_RK235", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        Do While TBAbrir.EOF = False
'            .WriteLine "|K235|" & TBAbrir!DT_SAIDA & "|" & TBAbrir!COD_ITEM & "|" & Format(TBAbrir!Qtd, "0.000") & "|" & TBAbrir!COD_INS_SUBST & "|"
'            TBAbrir.MoveNext
'        Loop
'    End If
'    'BK_RK250_Industrialização efetuada por terceiros - itens produzidos
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from SPEDF_BK_RK250", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        Do While TBAbrir.EOF = False
'            .WriteLine "|K250|" & TBAbrir!DT_PROD & "|" & Trim(TBAbrir!COD_ITEM) & "|" & Format(TBAbrir!Qtd, "0.000") & "|"
'            TBAbrir.MoveNext
'        Loop
'    End If
'    'BK_RK255_Industrialização em terceiros - insumos consumidos
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select * from SPEDF_BK_RK255", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        Do While TBAbrir.EOF = False
'            .WriteLine "|K255|" & TBAbrir!DT_CONS & "|" & Trim(TBAbrir!COD_ITEM) & "|" & Format(TBAbrir!Qtd, "0.000") & "|" & TBAbrir!COD_INS_SUBST & "|"
'            TBAbrir.MoveNext
'        Loop
'    End If
    PBLista.Value = 4
.WriteLine "|K990|" & countLinha + 1 & "|"
.Close
End With
'=========================================
' Fecha bloco K
'=========================================
With Lista.ListItems
    .Add , , "K990"
    .Item(.Count).SubItems(1) = countLinha + 1
End With

Contador = 0

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


Private Sub cmbMes_Change()
On Error GoTo tratar_erro

Select Case cmbMes.Text
    Case "Janeiro": Mmes = 1
    Case "Fevereiro": Mmes = 2
    Case "Março": Mmes = 3
    Case "Abril": Mmes = 4
    Case "Maio": Mmes = 5
    Case "Junho": Mmes = 6
    Case "Julho": Mmes = 7
    Case "Agosto": Mmes = 8
    Case "Setembro": Mmes = 9
    Case "Outubro": Mmes = 10
    Case "Novembro": Mmes = 11
    Case "Dezembro": Mmes = 12
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbMes_Click()
On Error GoTo tratar_erro

Select Case cmbMes.Text
    Case "Janeiro": Mmes = 1
    Case "Fevereiro": Mmes = 2
    Case "Março": Mmes = 3
    Case "Abril": Mmes = 4
    Case "Maio": Mmes = 5
    Case "Junho": Mmes = 6
    Case "Julho": Mmes = 7
    Case "Agosto": Mmes = 8
    Case "Setembro": Mmes = 9
    Case "Outubro": Mmes = 10
    Case "Novembro": Mmes = 11
    Case "Dezembro": Mmes = 12
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdGerar_Click()
On Error GoTo tratar_erro

 ProcSPED

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSPED
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboEmpresa Cmb_empresa, False
ProcRemoveObjetosResize Me
cmbAno.Clear
Contador = 0

Do While Contador <= 5
cmbAno.AddItem Year(Date) - Contador
Contador = Contador + 1
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


