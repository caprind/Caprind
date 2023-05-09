VERSION 5.00
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Impostos_Substituicao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Fiscal | Classificação fiscal | Substituição tributária"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   10770
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7815
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   7410
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   688
      DibPicture      =   "frmFaturamento_Impostos_Substituicao.frx":0000
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
      Icon            =   "frmFaturamento_Impostos_Substituicao.frx":7180
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   2
      Top             =   390
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   1720
      ButtonCount     =   7
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   33
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   38
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   39
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   118
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "9"
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
      ButtonLeft5     =   122
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "10"
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
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "11"
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
      ButtonState7    =   5
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5340
         Top             =   270
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_Impostos_Substituicao.frx":749A
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6990
      Width           =   10515
      _ExtentX        =   18547
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
   Begin MSComctlLib.ListView Lista1 
      Height          =   4035
      Left            =   120
      TabIndex        =   4
      Top             =   2940
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "NCM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Grupo"
         Object.Width           =   5586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "CST"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Margem"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Alíquota"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame4 
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
      Height          =   1485
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   10515
      Begin VB.CommandButton Cmd_CF 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6930
         Picture         =   "frmFaturamento_Impostos_Substituicao.frx":A885
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
         Top             =   420
         Width           =   315
      End
      Begin VB.TextBox Txt_ID_CF 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "ID da NCM."
         Top             =   420
         Width           =   525
      End
      Begin VB.ComboBox cmbTributaria 
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
         ItemData        =   "frmFaturamento_Impostos_Substituicao.frx":A987
         Left            =   180
         List            =   "frmFaturamento_Impostos_Substituicao.frx":A9A0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Tipo antecipação tributária."
         Top             =   1020
         Width           =   10170
      End
      Begin VB.TextBox txtResponsavel1 
         Alignment       =   2  'Center
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
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   420
         Width           =   3330
      End
      Begin VB.TextBox txtData1 
         Alignment       =   2  'Center
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txtAliquota 
         Alignment       =   2  'Center
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
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   9
         ToolTipText     =   "Alíquota interna do ICMS."
         Top             =   420
         Width           =   1080
      End
      Begin VB.TextBox txtMargem 
         Alignment       =   2  'Center
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
         Left            =   8205
         MaxLength       =   10
         TabIndex        =   8
         ToolTipText     =   "Margem do valor agregado (MVA)."
         Top             =   420
         Width           =   1020
      End
      Begin VB.ComboBox cmbCST 
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
         ItemData        =   "frmFaturamento_Impostos_Substituicao.frx":AC61
         Left            =   7275
         List            =   "frmFaturamento_Impostos_Substituicao.frx":AD22
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Código de situação tributária"
         Top             =   420
         Width           =   930
      End
      Begin VB.TextBox Txt_CF 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Classificação fiscal."
         Top             =   420
         Width           =   1155
      End
      Begin ControlesUteis.txt txtUF1 
         Height          =   360
         Left            =   4635
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "UF."
         Top             =   420
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   635
         Tamanho         =   570
         Text            =   ""
         FocusColor      =   16777215
         CaptionColor    =   0
         ShowCaption     =   0   'False
         Caption         =   ""
         Locked          =   -1  'True
         MaxLength       =   2
         BackColor       =   14737632
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin VB.TextBox txtId1 
         Alignment       =   2  'Center
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
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Data do cadastro."
         Top             =   9240
         Width           =   375
      End
      Begin VB.TextBox txtId 
         Alignment       =   2  'Center
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
         Left            =   390
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Data do cadastro."
         Top             =   9240
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   5400
         TabIndex        =   24
         Top             =   210
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6172
         TabIndex        =   23
         Top             =   210
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo antecipação tributária"
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
         Index           =   1
         Left            =   4290
         TabIndex        =   22
         Top             =   810
         Width           =   1920
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
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
         Left            =   2498
         TabIndex        =   21
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         Left            =   600
         TabIndex        =   20
         Top             =   210
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alíq. int. (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   9308
         TabIndex        =   19
         Top             =   210
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4830
         TabIndex        =   18
         Top             =   210
         Width           =   195
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "MVA (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   8445
         TabIndex        =   17
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "CST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   7530
         TabIndex        =   16
         Top             =   210
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmFaturamento_Impostos_Substituicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_CF_Regiao  As Boolean 'OK
Dim Novo_CF_Regiao1 As Boolean 'OK

Private Sub cmbCST_Click()
On Error GoTo tratar_erro

txtMargem = ""
txtAliquota = ""
cmbTributaria.ListIndex = -1
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from cst where id_uf = " & txtId & " and ID_CF = " & IIf(Txt_ID_CF = "", 0, Txt_ID_CF) & " and cst = '" & cmbCST & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    txtMargem = IIf(IsNull(TBFI!Margem), "", TBFI!Margem)
    txtAliquota = IIf(IsNull(TBFI!Aliquota), "", TBFI!Aliquota)
    NomeCampo = "Não foi encontrado o campo tipo antecipação tributária"
    If IsNull(TBFI!Tipo) = False And TBFI!Tipo <> "" Then cmbTributaria = TBFI!Tipo
End If
TBFI.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox (NomeCampo & ", favor alterar."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End Sub

Private Sub ProcExcluirSub()
On Error GoTo tratar_erro

Permitido = False
With Lista1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) substituição(ões) tributária?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CST WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Fiscal/Classificação fiscal/Cadastro de regiões"
                Evento = "Excluir substituição tributária"
                ID_documento = .ListItems(InitFor)
                Documento = "UF: " & Lista.SelectedItem.ListSubItems(1)
                Documento1 = "CST: " & TBFI!CST
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from CST where ID = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) substituição(ões) tributária antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Substituição(ões) tributária excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos1
    ProcCarregaLista1
    Frame4.Enabled = False
    Novo_CF_Regiao1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcNovoSub()
On Error GoTo tratar_erro

ProcLimpaCampos1
Frame4.Enabled = True
'Cmd_CF_Click
Novo_CF_Regiao1 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_CF_Regiao1 = True Then
    If USMsgBox("A substituição tributária ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarSub
        If Novo_CF_Regiao1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_CF_Regiao = False
Novo_CF_Regiao1 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcGravarSub()
On Error GoTo tratar_erro

If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_CF = "" Then
    NomeCampo = "a classificação fiscal"
    ProcVerificaAcao
    Txt_CF.SetFocus
    Exit Sub
End If
If cmbCST = "" Then
    NomeCampo = "a CST"
    ProcVerificaAcao
    cmbCST.SetFocus
    Exit Sub
End If
If IsNumeric(txtMargem) = False And txtMargem <> "" Then
    NomeCampo = "a margem"
    ProcVerificaAcao
    txtMargem.SetFocus
    Exit Sub
End If
If IsNumeric(txtAliquota) = False And txtAliquota <> "" Then
    NomeCampo = "a alíquota"
    ProcVerificaAcao
    txtAliquota.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CST where id = " & txtId1, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
If txtData1 <> "" Then TBGravar!Data = txtData1 Else TBGravar!Data = Date
If txtResponsavel1 <> "" Then TBGravar!Responsavel = txtResponsavel1 Else TBGravar!Responsavel = pubUsuario
TBGravar!ID_UF = txtId
TBGravar!ID_CF = Txt_ID_CF
TBGravar!CST = cmbCST
TBGravar!Margem = IIf(txtMargem = "", "0", txtMargem)
TBGravar!Aliquota = IIf(txtAliquota = "", "0", txtAliquota)
TBGravar!Tipo = IIf(cmbTributaria = "", Null, cmbTributaria)
TBGravar.Update
txtId1 = TBGravar!ID
TBGravar.Close
ProcCarregaLista1
If Novo_CF_Regiao1 = True Then
    USMsgBox ("Nova CST cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova substituição tributária"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar substituição tributária"
    If CodigoLista1 <> 0 And Lista1.ListItems.Count <> 0 Then
        Lista1.SelectedItem = Lista1.ListItems(CodigoLista1)
        Lista1.SetFocus
    End If
End If
'==================================
Modulo = "Fiscal/Classificação fiscal/Cadastro de regiões"
ID_documento = txtId1
Documento = "UF: " & txtUF1.Text
Documento1 = "CST: " & cmbCST
ProcGravaEvento
'==================================
Novo_CF_Regiao1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_CF_Click()
On Error GoTo tratar_erro

Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Clientes = False
Compras_Pedido = False
Familia_NCM = False
ClassFiscal = True
frmProj_Classificacao_Fiscal.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovoSub
    Case vbKeyF3: ProcGravarSub
    Case vbKeyF4: ProcExcluirSub
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10515, 7, True
ProcCarregaLista1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista1()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select CST.*, CF.IDIntClasse, CF.txt_Grupo from CST LEFT JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = CST.ID_CF " 'where CST.ID_UF = " & txtId
'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    With Lista1.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDIntClasse), "", TBLISTA!IDIntClasse)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Txt_grupo), "", TBLISTA!Txt_grupo)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!CST), "", TBLISTA!CST)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Margem), "", TBLISTA!Margem)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Aliquota), "", TBLISTA!Aliquota)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End With
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista1.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos1
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CST where id = " & Lista1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcCarregaDadosST
    CodigoLista1 = Lista1.SelectedItem.index
    Frame4.Enabled = True
End If
TBLISTA.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosST()
On Error GoTo tratar_erro

txtId1.Text = TBLISTA!ID
txtData1 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel1 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Regioes where Id = " & IIf(IsNull(TBLISTA!ID_UF), 0, TBLISTA!ID_UF), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtUF1.Text = IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF)
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBLISTA!ID_CF), 0, TBLISTA!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_CF = TBAbrir!Idclass
    Txt_CF = IIf(IsNull(TBAbrir!IDIntClasse), "", TBAbrir!IDIntClasse)
End If
TBAbrir.Close

NomeCampo = "A CST " & TBLISTA!CST & " não é com substituição tributária"
If IsNull(TBLISTA!CST) = False And TBLISTA!CST <> "" Then cmbCST = TBLISTA!CST
1:
    Novo_CF_Regiao1 = False

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox (NomeCampo & ", favor alterar."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Sub ProcLimpaCampos1()
On Error GoTo tratar_erro

txtId1 = 0
txtData1 = Format(Date, "dd/mm/yy")
txtResponsavel1 = pubUsuario
txtUF1.Text = frmFaturamento_Prod_Serv.cbo_UF.Text
Txt_ID_CF = frmFaturamento_Prod_Serv.Txt_ID_CF.Text
Txt_CF = frmFaturamento_Prod_Serv.Txt_CF.Text
cmbCST.ListIndex = -1
txtMargem = ""
txtAliquota = ""
cmbTributaria.ListIndex = -1
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtAliquota_LostFocus()
On Error GoTo tratar_erro

If txtAliquota.Text <> "" Then
    VerifNumero = txtAliquota.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtAliquota.Text = ""
        txtAliquota.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMargem_LostFocus()
    On Error GoTo tratar_erro

If txtMargem.Text <> "" Then
    VerifNumero = txtMargem.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtMargem.Text = ""
        txtMargem.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtUF1_Change()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select ID from regioes where UF = '" & txtUF1.Text & "'"
'Debug.print StrSql

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
txtId = TBAbrir!ID
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoSub
    Case 2: ProcGravarSub
    Case 3: ProcExcluirSub
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
