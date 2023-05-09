VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Classificacao_Fiscal 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Fiscal - Classificação fiscal"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Classificacao_Fiscal.frx":0000
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
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
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
      Caption         =   "Alíquotas"
      Enabled         =   0   'False
      Height          =   1485
      Left            =   13470
      TabIndex        =   74
      Top             =   1830
      Width           =   1725
      Begin MSMask.MaskEdBox Txt_aliq_importacao 
         Height          =   315
         Left            =   690
         TabIndex        =   31
         ToolTipText     =   "Alíquota importação. ""LEI 12741"""
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_aliq_nacional 
         Height          =   315
         Left            =   690
         TabIndex        =   30
         ToolTipText     =   "Alíquota nacional. ""LEI 12741"""
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Imp. :"
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
         TabIndex        =   78
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   4
         Left            =   1365
         TabIndex        =   77
         Top             =   900
         Width           =   165
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nac. :"
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
         TabIndex        =   76
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1365
         TabIndex        =   75
         Top             =   540
         Width           =   165
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   70
      Top             =   9120
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
         ItemData        =   "frm_Classificacao_Fiscal.frx":1042
         Left            =   6990
         List            =   "frm_Classificacao_Fiscal.frx":104C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   81
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
         TabIndex        =   33
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
         TabIndex        =   34
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   38
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Classificacao_Fiscal.frx":1064
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
         TabIndex        =   37
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Classificacao_Fiscal.frx":480B
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
         TabIndex        =   35
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
         TabIndex        =   36
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Classificacao_Fiscal.frx":8319
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
         TabIndex        =   39
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Classificacao_Fiscal.frx":C40B
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
      Begin VB.Label Label30 
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
         Left            =   3360
         TabIndex        =   89
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
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
         Index           =   29
         Left            =   5610
         TabIndex        =   82
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label21 
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
         Left            =   2040
         TabIndex        =   73
         Top             =   240
         Width           =   645
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
         TabIndex        =   72
         Top             =   240
         Width           =   1275
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
         TabIndex        =   71
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox Chk_desoneracao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desoneração da foha de pagam."
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
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   2340
      Width           =   2685
   End
   Begin VB.CheckBox chkReducao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Com redução na BC do ICMS"
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
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   2115
      Width           =   2415
   End
   Begin VB.TextBox TxtID 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1710
      TabIndex        =   47
      Text            =   "0"
      ToolTipText     =   "Ref."
      Top             =   6420
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5775
      Left            =   60
      TabIndex        =   32
      Top             =   3330
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   10186
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Grupo"
         Object.Width           =   11633
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Classificação"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "IPI"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "ICMS (DE)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "ICMS (SS)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "ICMS (NN)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "ICMS (CO)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "ICMS (EX)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Validada"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   825
      Left            =   55
      TabIndex        =   40
      Top             =   990
      Width           =   15135
      Begin VB.TextBox txtRespValidacao 
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
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pela validação."
         Top             =   375
         Width           =   2745
      End
      Begin VB.TextBox txtDtValidacao 
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   375
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   375
         Width           =   945
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   1135
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   375
         Width           =   2805
      End
      Begin VB.TextBox txtGrupo 
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
         Left            =   8370
         MaxLength       =   255
         TabIndex        =   4
         ToolTipText     =   "Grupo."
         Top             =   375
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mskClassificacao 
         Height          =   315
         Left            =   12240
         TabIndex        =   5
         ToolTipText     =   "Número da classificação fiscal."
         Top             =   375
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####.##.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtIPI 
         Height          =   315
         Left            =   13290
         TabIndex        =   6
         ToolTipText     =   "Alíquota de IPI."
         Top             =   375
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_CEST 
         Height          =   315
         Left            =   14010
         TabIndex        =   7
         ToolTipText     =   "Código especificador da substituição tributária."
         Top             =   375
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CEST"
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
         Left            =   14295
         TabIndex        =   88
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável pela validação"
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
         Index           =   2
         Left            =   5992
         TabIndex        =   80
         Top             =   180
         Width           =   1980
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data/hora validação"
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
         Left            =   4050
         TabIndex        =   79
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label19 
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
         Left            =   480
         TabIndex        =   49
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
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
         Index           =   9
         Left            =   2080
         TabIndex        =   48
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo"
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
         Left            =   10080
         TabIndex        =   43
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Classificação"
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
         Left            =   12300
         TabIndex        =   42
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IPI (%)"
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
         Left            =   13365
         TabIndex        =   41
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.CheckBox Chk_retem_PIS_Cofins 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   11190
      TabIndex        =   25
      Top             =   1860
      Width           =   225
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retem PIS/Cofins "
      Enabled         =   0   'False
      Height          =   1485
      Left            =   9510
      TabIndex        =   50
      Top             =   1830
      Width           =   1965
      Begin MSMask.MaskEdBox Txt_Cofins 
         Height          =   315
         Left            =   780
         TabIndex        =   27
         ToolTipText     =   "Alíquota cofins."
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_PIS 
         Height          =   315
         Left            =   780
         TabIndex        =   26
         ToolTipText     =   "Alíquota PIS."
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cofins :"
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
         TabIndex        =   54
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   9
         Left            =   1455
         TabIndex        =   53
         Top             =   900
         Width           =   165
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PIS :"
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
         Left            =   390
         TabIndex        =   52
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1455
         TabIndex        =   51
         Top             =   540
         Width           =   165
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alíquota PIS/Cofins "
      Enabled         =   0   'False
      Height          =   1485
      Left            =   11490
      TabIndex        =   61
      Top             =   1830
      Width           =   1965
      Begin MSMask.MaskEdBox Txt_Cofins1 
         Height          =   315
         Left            =   870
         TabIndex        =   29
         ToolTipText     =   "Alíquota cofins."
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_PIS1 
         Height          =   315
         Left            =   870
         TabIndex        =   28
         ToolTipText     =   "Alíquota PIS."
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   2
         Left            =   1545
         TabIndex        =   65
         Top             =   540
         Width           =   165
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PIS :"
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
         Left            =   480
         TabIndex        =   64
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1545
         TabIndex        =   63
         Top             =   900
         Width           =   165
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cofins :"
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
         Left            =   270
         TabIndex        =   62
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alíquota interestadual do ICMS"
      Enabled         =   0   'False
      Height          =   1485
      Left            =   2970
      TabIndex        =   44
      Top             =   1830
      Width           =   6525
      Begin MSMask.MaskEdBox txtCTDE 
         Height          =   315
         Left            =   600
         TabIndex        =   15
         ToolTipText     =   "Carga tributaria dentro do estado."
         Top             =   660
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtICMSde 
         Height          =   315
         Left            =   600
         TabIndex        =   10
         ToolTipText     =   "ICMS dentro do estado."
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtICMSss 
         Height          =   315
         Left            =   1860
         TabIndex        =   11
         ToolTipText     =   "ICMS sul e sudeste."
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCTSS 
         Height          =   315
         Left            =   1860
         TabIndex        =   16
         ToolTipText     =   "Carga tributaria sul e sudeste."
         Top             =   660
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtICMSnn 
         Height          =   315
         Left            =   3120
         TabIndex        =   12
         ToolTipText     =   "ICMS norte e nordeste."
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCTNN 
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         ToolTipText     =   "Carga tributaria norte e nordeste."
         Top             =   660
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtICMSco 
         Height          =   315
         Left            =   4395
         TabIndex        =   13
         ToolTipText     =   "ICMS centro oeste."
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCTCO 
         Height          =   315
         Left            =   4395
         TabIndex        =   18
         ToolTipText     =   "Carga tributaria centro oeste."
         Top             =   660
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtICMSex 
         Height          =   315
         Left            =   5625
         TabIndex        =   14
         ToolTipText     =   "ICMS exterior."
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCTEX 
         Height          =   315
         Left            =   5625
         TabIndex        =   19
         ToolTipText     =   "Carga tributaria exterior."
         Top             =   660
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_DIF_DE 
         Height          =   315
         Left            =   600
         TabIndex        =   20
         ToolTipText     =   "Percentual de diferimento do ICMS dentro do estado."
         Top             =   1020
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_DIF_SS 
         Height          =   315
         Left            =   1860
         TabIndex        =   21
         ToolTipText     =   "Percentual de diferimento do ICMS sul e sudeste."
         Top             =   1020
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_DIF_NN 
         Height          =   315
         Left            =   3120
         TabIndex        =   22
         ToolTipText     =   "Percentual de diferimento do ICMS norte e nordeste."
         Top             =   1020
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_DIF_CO 
         Height          =   315
         Left            =   4395
         TabIndex        =   23
         ToolTipText     =   "Percentual de diferimento do ICMS centro oeste."
         Top             =   1020
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txt_DIF_EX 
         Height          =   315
         Left            =   5625
         TabIndex        =   24
         ToolTipText     =   "Carga tributaria exterior."
         Top             =   1020
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DIF :"
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
         TabIndex        =   87
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DIF :"
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
         Left            =   1410
         TabIndex        =   86
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DIF :"
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
         Left            =   2685
         TabIndex        =   85
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DIF :"
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
         Left            =   3945
         TabIndex        =   84
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DIF :"
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
         Left            =   5190
         TabIndex        =   83
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CT :"
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
         Left            =   5250
         TabIndex        =   67
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EX :"
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
         Left            =   5265
         TabIndex        =   66
         Top             =   300
         Width           =   285
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CO :"
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
         Left            =   3975
         TabIndex        =   60
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CT :"
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
         Left            =   4005
         TabIndex        =   59
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CT :"
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
         Left            =   2745
         TabIndex        =   58
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NN :"
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
         Left            =   2730
         TabIndex        =   57
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CT :"
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
         Left            =   1470
         TabIndex        =   56
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SS :"
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
         Left            =   1485
         TabIndex        =   55
         Top             =   300
         Width           =   285
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DE :"
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
         Left            =   240
         TabIndex        =   46
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CT :"
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
         Left            =   240
         TabIndex        =   45
         Top             =   660
         Width           =   300
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   68
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   1720
      ButtonCount     =   13
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
      ButtonCaption2  =   "Filtrar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Filtrar (F2)"
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
      ButtonWidth2    =   36
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Salvar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Salvar (F3)"
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
      ButtonLeft3     =   75
      ButtonTop3      =   2
      ButtonWidth3    =   38
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Excluir"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Excluir (F4)"
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
      ButtonLeft4     =   115
      ButtonTop4      =   2
      ButtonWidth4    =   39
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Relatório"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Relatório (F5)"
      ButtonKey5      =   "5"
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
      ButtonLeft5     =   156
      ButtonTop5      =   2
      ButtonWidth5    =   51
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Anterior"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Registro anterior."
      ButtonKey6      =   "6"
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
      ButtonLeft6     =   209
      ButtonTop6      =   2
      ButtonWidth6    =   47
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Próximo"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Próximo registro."
      ButtonKey7      =   "7"
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
      ButtonLeft7     =   258
      ButtonTop7      =   2
      ButtonWidth7    =   46
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Região"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Abrir formulário para cadastro de UF por região (F7)"
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   306
      ButtonTop8      =   2
      ButtonWidth8    =   41
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Validação"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Validar/Cancelar validação (F8)"
      ButtonKey9      =   "9"
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
      ButtonLeft9     =   349
      ButtonTop9      =   2
      ButtonWidth9    =   53
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonAlignment10=   2
      ButtonType10    =   1
      ButtonStyle10   =   -1
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState10   =   -1
      ButtonLeft10    =   404
      ButtonTop10     =   4
      ButtonWidth10   =   2
      ButtonHeight10  =   54
      ButtonCaption11 =   "Ajuda"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Ajuda (F1)"
      ButtonKey11     =   "11"
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
      ButtonLeft11    =   408
      ButtonTop11     =   2
      ButtonWidth11   =   36
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      ButtonCaption12 =   "Sair"
      ButtonEnabled12 =   0   'False
      ButtonIconSize12=   32
      ButtonToolTipText12=   "Sair (Esc)"
      ButtonKey12     =   "12"
      ButtonAlignment12=   2
      BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft12    =   446
      ButtonTop12     =   2
      ButtonWidth12   =   26
      ButtonHeight12  =   21
      ButtonUseMaskColor12=   0   'False
      ButtonEnabled13 =   0   'False
      ButtonIconSize13=   32
      ButtonKey13     =   "13"
      ButtonAlignment13=   2
      BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState13   =   5
      ButtonLeft13    =   474
      ButtonTop13     =   2
      ButtonWidth13   =   24
      ButtonHeight13  =   24
      ButtonUseMaskColor13=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   12150
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Classificacao_Fiscal.frx":FC98
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   69
      Top             =   9750
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
      SearchText      =   ""
      Value           =   0
   End
End
Attribute VB_Name = "frm_Classificacao_Fiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_CF As Boolean 'OK
Public StrSql_CF As String 'OK
Public FormulaRel_CF As String 'OK
Dim TBLISTA_CF As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=hPbnn3LJPAY&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=55&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_retem_PIS_Cofins_Click()
On Error GoTo tratar_erro

If Chk_retem_PIS_Cofins.Value = 1 Then
    Frame8.Enabled = True
    Txt_PIS.SetFocus
Else
    Frame8.Enabled = False
    Txt_PIS = ""
    Txt_Cofins = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
NomeRel = "Fiscal_Classificacao fiscal.rpt"
ProcImprimirRel FormulaRel_CF, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkReducao_Click()
On Error GoTo tratar_erro

If chkReducao.Value = 1 Then
    txtCTDE.Enabled = True
    txtCTSS.Enabled = True
    txtCTNN.Enabled = True
    txtCTCO.Enabled = True
    txtCTEX.Enabled = True
Else
    txtCTDE = ""
    txtCTDE.Enabled = False
    txtCTSS = ""
    txtCTSS.Enabled = False
    txtCTNN = ""
    txtCTNN.Enabled = False
    txtCTCO = ""
    txtCTCO.Enabled = False
    txtCTEX = ""
    txtCTEX.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_ClassificacaoFiscal order by IDIntClasse", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("Idclass = " & txtId)
    TBAbrir.MovePrevious
    If TBAbrir.BOF = False Then
        txtId = TBAbrir!Idclass
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        USMsgBox ("Fim dos cadastros de classificação fiscal."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_ClassificacaoFiscal order by IDIntClasse", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("Idclass = " & txtId)
    TBAbrir.MoveNext
    If TBAbrir.EOF = False Then
        txtId = TBAbrir!Idclass
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        USMsgBox ("Fim dos cadastros de classificação fiscal."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procRegiao()
On Error GoTo tratar_erro
  
frm_Classificacao_Fiscal_regiao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(9) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(9) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CF.AbsolutePage <> 2 Then
    If TBLISTA_CF.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_CF.PageCount - 1)
    Else
        TBLISTA_CF.AbsolutePage = TBLISTA_CF.AbsolutePage - 2
        ProcExibePagina (TBLISTA_CF.AbsolutePage)
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
    TBLISTA_CF.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_CF.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CF.AbsolutePage = 1
ProcExibePagina (TBLISTA_CF.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CF.AbsolutePage <> -3 Then
    If TBLISTA_CF.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_CF.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_CF.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CF.AbsolutePage = TBLISTA_CF.PageCount
ProcExibePagina (TBLISTA_CF.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcLocalizar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
    Case vbKeyF7: procRegiao
    Case vbKeyF8: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, "Faturamento/Fiscal/Classificação fiscal"
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
txtGrupo.Text = ""
mskClassificacao = "____.__.__"
txtIPI = ""
Txt_CEST = "__.___.__"
chkReducao.Value = 0
Chk_desoneracao.Value = 0
txtICMSde = ""
txtCTDE = ""
txtICMSss = ""
txtCTSS = ""
txtICMSnn = ""
txtCTNN = ""
TxtICMSco = ""
txtCTCO = ""
TxtICMSex = ""
txtCTEX = ""
Txt_DIF_DE = ""
Txt_DIF_SS = ""
Txt_DIF_NN = ""
Txt_DIF_CO = ""
Txt_DIF_EX = ""
Chk_retem_PIS_Cofins.Value = 0
Txt_PIS = ""
Txt_Cofins = ""
Txt_PIS1 = ""
Txt_Cofins1 = ""
Txt_aliq_nacional = ""
Txt_aliq_importacao = ""
CodigoLista = 0
Caption = "Administrativo - Faturamento - Fiscal - Classificação fiscal"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) classificação(ões) fiscal(ais)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_ClassificacaoFiscal WHERE Idclass = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Faturamento/Fiscal/Classificação fiscal"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Classificação: " & TBFI!IDIntClasse
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE FROM tbl_ClassificacaoFiscal WHERE IDClass = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) classificação(ões) fiscal(ais) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Classificação(ões) fiscal(ais) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    ProcDesabilitaCampos
    Novo_CF = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos
ProcHabilitaCampos
Novo_CF = True
txtGrupo.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaCampos()
On Error GoTo tratar_erro

Chk_retem_PIS_Cofins.Enabled = True
Frame2.Enabled = True
Frame1.Enabled = True
Frame4.Enabled = True
Frame3.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesabilitaCampos()
On Error GoTo tratar_erro

Chk_retem_PIS_Cofins.Enabled = False
Frame2.Enabled = False
Frame1.Enabled = False
Frame4.Enabled = False
Frame3.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
If Novo_CF = True Then
    If USMsgBox("A classificação fiscal ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_CF = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_CF = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

If txtData <> "" Then TBAliquota!Data = txtData Else TBAliquota!Data = Date
If txtResponsavel <> "" Then TBAliquota!Responsavel = txtResponsavel Else TBAliquota!Responsavel = pubUsuario
TBAliquota!Txt_grupo = txtGrupo.Text
TBAliquota!IDIntClasse = mskClassificacao
TBAliquota!dbl_IPI = IIf(txtIPI = "", Null, txtIPI)
TBAliquota!CEST = IIf(Txt_CEST = "__.___.__", Null, Txt_CEST)
TBAliquota!dbl_ICMS_de = IIf(txtICMSde = "", Null, txtICMSde)
TBAliquota!dbl_ICMS_ss = IIf(txtICMSss = "", Null, txtICMSss)
TBAliquota!dbl_ICMS_nn = IIf(txtICMSnn = "", Null, txtICMSnn)
TBAliquota!dbl_ICMS_co = IIf(TxtICMSco = "", Null, TxtICMSco)
TBAliquota!dbl_ICMS_ex = IIf(TxtICMSex = "", Null, TxtICMSex)
If chkReducao.Value = 1 Then
    TBAliquota!basereduzida = True
    TBAliquota!CTDE = IIf(txtCTDE = "", Null, txtCTDE)
    TBAliquota!CTNN = IIf(txtCTNN = "", Null, txtCTNN)
    TBAliquota!CTSS = IIf(txtCTSS = "", Null, txtCTSS)
    TBAliquota!CTCO = IIf(txtCTCO = "", Null, txtCTCO)
    TBAliquota!CTEX = IIf(txtCTEX = "", Null, txtCTEX)
Else
    TBAliquota!basereduzida = False
End If
TBAliquota!DIFDE = IIf(Txt_DIF_DE = "", Null, Txt_DIF_DE)
TBAliquota!DIFNN = IIf(Txt_DIF_NN = "", Null, Txt_DIF_NN)
TBAliquota!DIFSS = IIf(Txt_DIF_SS = "", Null, Txt_DIF_SS)
TBAliquota!DIFCO = IIf(Txt_DIF_CO = "", Null, Txt_DIF_CO)
TBAliquota!DIFEX = IIf(Txt_DIF_EX = "", Null, Txt_DIF_EX)

If Chk_desoneracao.Value = 1 Then TBAliquota!Desoneracao = True Else TBAliquota!Desoneracao = False
If Chk_retem_PIS_Cofins.Value = 1 Then
    TBAliquota!Retem_PIS_Cofins = True
    TBAliquota!PIS = IIf(Txt_PIS = "", Null, Txt_PIS)
    TBAliquota!Cofins = IIf(Txt_Cofins = "", Null, Txt_Cofins)
Else
    TBAliquota!Retem_PIS_Cofins = False
End If
TBAliquota!PIS_destaca = IIf(Txt_PIS1 = "", Null, Txt_PIS1)
TBAliquota!Cofins_destaca = IIf(Txt_Cofins1 = "", Null, Txt_Cofins1)
TBAliquota!Aliq_nacional = IIf(Txt_aliq_nacional = "", Null, Txt_aliq_nacional)
TBAliquota!Aliq_importacao = IIf(Txt_aliq_importacao = "", Null, Txt_aliq_importacao)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtGrupo.Text = "" Then
    NomeCampo = "o grupo"
    ProcVerificaAcao
    txtGrupo.SetFocus
    Exit Sub
End If
If mskClassificacao.Text = "____.__.__" Then
    NomeCampo = "a classificação"
    ProcVerificaAcao
    mskClassificacao.SetFocus
    Exit Sub
End If
If txtIPI.Text = "" Then
    NomeCampo = "a alíquota do IPI"
    ProcVerificaAcao
    txtIPI.SetFocus
    Exit Sub
End If
If txtICMSde.Text = "" Then
    NomeCampo = "a alíquota de ICMS"
    ProcVerificaAcao
    txtICMSde.SetFocus
    Exit Sub
End If
If txtICMSss.Text = "" Then
    NomeCampo = "a alíquota de ICMS"
    ProcVerificaAcao
    txtICMSss.SetFocus
    Exit Sub
End If
If txtICMSnn.Text = "" Then
    NomeCampo = "a alíquota de ICMS"
    ProcVerificaAcao
    txtICMSnn.SetFocus
    Exit Sub
End If
If TxtICMSco.Text = "" Then
    NomeCampo = "a alíquota de ICMS"
    ProcVerificaAcao
    TxtICMSco.SetFocus
    Exit Sub
End If
If chkReducao.Value = 1 Then
    If txtCTDE = "" Or txtCTSS = "" Or txtCTNN = "" Or txtCTCO = "" Or txtCTEX = "" Then
        USMsgBox ("É necessário preencher todos os campos antes de salvar."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
End If
If Chk_retem_PIS_Cofins.Value = 1 Then
    If Txt_PIS = "" Then
        NomeCampo = "a alíquota de PIS"
        ProcVerificaAcao
        Txt_PIS.SetFocus
        Exit Sub
    End If
    If Txt_Cofins = "" Then
        NomeCampo = "a alíquota de Cofins"
        ProcVerificaAcao
        Txt_Cofins.SetFocus
        Exit Sub
    End If
End If
If Txt_PIS1 = "" Then
    NomeCampo = "a alíquota de PIS"
    ProcVerificaAcao
    Txt_PIS1.SetFocus
    Exit Sub
End If
If Txt_Cofins1 = "" Then
    NomeCampo = "a alíquota de Cofins"
    ProcVerificaAcao
    Txt_Cofins1.SetFocus
    Exit Sub
End If
If Txt_aliq_nacional = "" Then
    NomeCampo = "a alíquota nacional"
    ProcVerificaAcao
    Txt_aliq_nacional.SetFocus
    Exit Sub
End If
If Txt_aliq_importacao = "" Then
    NomeCampo = "a alíquota de importação"
    ProcVerificaAcao
    Txt_aliq_importacao.SetFocus
    Exit Sub
End If

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from Tbl_ClassificacaoFiscal where IDIntClasse = '" & mskClassificacao.Text & "' and Idclass <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    If USMsgBox("Este código de classificação está sendo utilizado, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
End If

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from Tbl_ClassificacaoFiscal where IDClass = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = True Then
    TBAliquota.AddNew
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesma", "classificação fiscal", False) = False Then Exit Sub
End If
ProcEnviaDados
TBAliquota.Update
txtId = TBAliquota!Idclass
TBAliquota.Close
If Novo_CF = True Then
    USMsgBox ("Nova classificação fiscal cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    'StrSql_CF = "Select * from Tbl_ClassificacaoFiscal where IDClass = " & TxtID
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Faturamento/Fiscal/Classificação fiscal"
ID_documento = txtId
Documento = "Classificação: " & mskClassificacao
Documento1 = ""
ProcGravaEvento
'==================================
Novo_CF = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15135, 13, True

Cmb_opcao_lista = "Validação"
Formulario = "Faturamento/Fiscal/Classificação fiscal"
Direitos
ProcLimpaVariaveisPrincipais
ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If StrSql_CF = "" Then Exit Sub
Set TBLISTA_CF = CreateObject("adodb.recordset")
TBLISTA_CF.Open StrSql_CF, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_CF.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_CF.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_CF.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_CF.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_CF.RecordCount - IIf(Pagina > 1, (TBLISTA_CF.PageSize * (Pagina - 1)), 0), TBLISTA_CF.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_CF.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_CF!Idclass
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_CF!Txt_grupo), "", TBLISTA_CF!Txt_grupo)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_CF!IDIntClasse), "", TBLISTA_CF!IDIntClasse)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_CF!dbl_IPI), "", TBLISTA_CF!dbl_IPI)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_CF!dbl_ICMS_de), "", TBLISTA_CF!dbl_ICMS_de)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_CF!dbl_ICMS_ss), "", TBLISTA_CF!dbl_ICMS_ss)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_CF!dbl_ICMS_nn), "", TBLISTA_CF!dbl_ICMS_nn)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_CF!dbl_ICMS_co), "", TBLISTA_CF!dbl_ICMS_co)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_CF!dbl_ICMS_ex), "", TBLISTA_CF!dbl_ICMS_ex)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_CF!DtValidacao) = False, "Sim", "Não")
    End With
    TBLISTA_CF.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_CF.RecordCount
If TBLISTA_CF.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_CF.PageCount
ElseIf TBLISTA_CF.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_CF.PageCount & " de: " & TBLISTA_CF.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_CF.AbsolutePage - 1 & " de: " & TBLISTA_CF.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frm_Classificacao_Fiscal_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Faturamento/Fiscal/Classificação fiscal"
Direitos
ProcLimpaVariaveisPrincipais

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
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidadoSemMsg("tbl_ClassificacaoFiscal", "Idclass = " & .ListItems(InitFor), True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    
                    ProcVerificaRegistroUtilizadoSemMsg "projproduto", "ID_CF = " & .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Projproduto_clientes", "ID_CF = " & .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "vendas_Carteira", "ID_CF = " & .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "ID_CF = " & .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "CST", "ID_CF = " & .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
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

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Cmb_opcao_lista = "Excluir" Then
                If FunVerificaRegistroValidado("tbl_ClassificacaoFiscal", "Idclass = " & .ListItems(InitFor), "a mesma", "esta classificação fiscal", "excluir", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                Mensagem = "Não é permitido excluir esta classificação fiscal, pois a mesma está sendo utilizada no módulo"
                ProcVerificaRegistroUtilizado "projproduto", "ID_CF = " & .ListItems(InitFor), "Engenharia/Produtos e serviços"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Projproduto_clientes", "ID_CF = " & .ListItems(InitFor), "Vendas/Produtos e serviços"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "vendas_Carteira", "ID_CF = " & .ListItems(InitFor), "Vendas"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_Detalhes_Nota", "ID_CF = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "CST", "ID_CF = " & .ListItems(InitFor), "Faturamento/Classificação fiscal/Cadastro de regiões"
                If Permitido = False Then
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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_ClassificacaoFiscal where idClass = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtId.Text = IIf(IsNull(TBLISTA!Idclass) = False, TBLISTA!Idclass, "")
txtData = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
txtDtValidacao = IIf(IsNull(TBLISTA!DtValidacao), "", TBLISTA!DtValidacao)
txtRespValidacao = IIf(IsNull(TBLISTA!RespValidacao), "", TBLISTA!RespValidacao)
Caption = "Administrativo - Faturamento - Fiscal - Classificação fiscal (Classificação : " & TBLISTA!IDIntClasse & ")"
txtGrupo.Text = IIf(IsNull(TBLISTA!Txt_grupo) = False, TBLISTA!Txt_grupo, "")
mskClassificacao.Text = IIf(IsNull(TBLISTA!IDIntClasse) = False, TBLISTA!IDIntClasse, "____.__.__")
txtIPI = IIf(IsNull(TBLISTA!dbl_IPI) = False, TBLISTA!dbl_IPI, "")
Txt_CEST = IIf(IsNull(TBLISTA!CEST) = False, TBLISTA!CEST, "__.___.__")
txtICMSde = IIf(IsNull(TBLISTA!dbl_ICMS_de), "", TBLISTA!dbl_ICMS_de)
txtICMSss = IIf(IsNull(TBLISTA!dbl_ICMS_ss), "", TBLISTA!dbl_ICMS_ss)
txtICMSnn = IIf(IsNull(TBLISTA!dbl_ICMS_nn), "", TBLISTA!dbl_ICMS_nn)
TxtICMSco = IIf(IsNull(TBLISTA!dbl_ICMS_co), "", TBLISTA!dbl_ICMS_co)
TxtICMSex = IIf(IsNull(TBLISTA!dbl_ICMS_ex), "", TBLISTA!dbl_ICMS_ex)
If TBLISTA!basereduzida = True Then
    txtCTCO.Text = IIf(IsNull(TBLISTA!CTCO), "", TBLISTA!CTCO)
    txtCTNN.Text = IIf(IsNull(TBLISTA!CTNN), "", TBLISTA!CTNN)
    txtCTSS.Text = IIf(IsNull(TBLISTA!CTSS), "", TBLISTA!CTSS)
    txtCTDE.Text = IIf(IsNull(TBLISTA!CTDE), "", TBLISTA!CTDE)
    txtCTEX.Text = IIf(IsNull(TBLISTA!CTEX), "", TBLISTA!CTEX)
    chkReducao.Value = 1
Else
    chkReducao.Value = 0
End If
Txt_DIF_DE = IIf(IsNull(TBLISTA!DIFDE), "", TBLISTA!DIFDE)
Txt_DIF_SS = IIf(IsNull(TBLISTA!DIFSS), "", TBLISTA!DIFSS)
Txt_DIF_NN = IIf(IsNull(TBLISTA!DIFNN), "", TBLISTA!DIFNN)
Txt_DIF_CO = IIf(IsNull(TBLISTA!DIFCO), "", TBLISTA!DIFCO)
Txt_DIF_EX = IIf(IsNull(TBLISTA!DIFEX), "", TBLISTA!DIFEX)
If TBLISTA!Desoneracao = True Then Chk_desoneracao.Value = 1 Else Chk_desoneracao.Value = 0
If TBLISTA!Retem_PIS_Cofins = True Then
    Chk_retem_PIS_Cofins.Value = 1
    Txt_PIS = IIf(IsNull(TBLISTA!PIS), "", TBLISTA!PIS)
    Txt_Cofins = IIf(IsNull(TBLISTA!Cofins), "", TBLISTA!Cofins)
Else
    Chk_retem_PIS_Cofins.Value = 0
End If
Txt_PIS1 = IIf(IsNull(TBLISTA!PIS_destaca), "", TBLISTA!PIS_destaca)
Txt_Cofins1 = IIf(IsNull(TBLISTA!Cofins_destaca), "", TBLISTA!Cofins_destaca)
Txt_aliq_nacional = IIf(IsNull(TBLISTA!Aliq_nacional), "", TBLISTA!Aliq_nacional)
Txt_aliq_importacao = IIf(IsNull(TBLISTA!Aliq_importacao), "", TBLISTA!Aliq_importacao)
Novo_CF = False
ProcHabilitaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_aliq_importacao_Change()
On Error GoTo tratar_erro

If Txt_aliq_importacao <> "" Then
    VerifNumero = Txt_aliq_importacao
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_aliq_importacao = ""
        Txt_aliq_importacao.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_aliq_nacional_Change()
On Error GoTo tratar_erro

If Txt_aliq_nacional <> "" Then
    VerifNumero = Txt_aliq_nacional
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_aliq_nacional = ""
        Txt_aliq_nacional.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Cofins_Change()
On Error GoTo tratar_erro

If Txt_Cofins.Text <> "" Then
    VerifNumero = Txt_Cofins.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_Cofins.Text = ""
        Txt_Cofins.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Cofins1_Change()
On Error GoTo tratar_erro

If Txt_Cofins1.Text <> "" Then
    VerifNumero = Txt_Cofins1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_Cofins1.Text = ""
        Txt_Cofins1.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Txt_DIF_CO_Change()
On Error GoTo tratar_erro

If Txt_DIF_CO <> "" Then
    VerifNumero = Txt_DIF_CO
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_DIF_CO = ""
        Txt_DIF_CO.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_CO_LostFocus()
On Error GoTo tratar_erro

Txt_DIF_CO = Format(Txt_DIF_CO, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_DE_Change()
On Error GoTo tratar_erro

If Txt_DIF_DE <> "" Then
    VerifNumero = Txt_DIF_DE
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_DIF_DE = ""
        Txt_DIF_DE.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_DE_LostFocus()
On Error GoTo tratar_erro

Txt_DIF_DE = Format(Txt_DIF_DE, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_EX_Change()
On Error GoTo tratar_erro

If Txt_DIF_EX <> "" Then
    VerifNumero = Txt_DIF_EX
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_DIF_EX = ""
        Txt_DIF_EX.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_EX_LostFocus()
On Error GoTo tratar_erro

Txt_DIF_EX = Format(Txt_DIF_EX, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_NN_Change()
On Error GoTo tratar_erro

If Txt_DIF_NN <> "" Then
    VerifNumero = Txt_DIF_NN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_DIF_NN = ""
        Txt_DIF_NN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_NN_LostFocus()
On Error GoTo tratar_erro

Txt_DIF_NN = Format(Txt_DIF_NN, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_SS_Change()
On Error GoTo tratar_erro

If Txt_DIF_SS <> "" Then
    VerifNumero = Txt_DIF_SS
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_DIF_SS = ""
        Txt_DIF_SS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_DIF_SS_LostFocus()
On Error GoTo tratar_erro

Txt_DIF_SS = Format(Txt_DIF_SS, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_PIS_Change()
On Error GoTo tratar_erro

If Txt_PIS.Text <> "" Then
    VerifNumero = Txt_PIS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_PIS.Text = ""
        Txt_PIS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_PIS1_Change()
On Error GoTo tratar_erro

If Txt_PIS1.Text <> "" Then
    VerifNumero = Txt_PIS1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_PIS1.Text = ""
        Txt_PIS1.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCTCO_Change()
On Error GoTo tratar_erro

If txtCTCO.Text <> "" Then
    VerifNumero = txtCTCO.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCTCO.Text = ""
        txtCTCO.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCTDE_Change()
On Error GoTo tratar_erro

If txtCTDE.Text <> "" Then
    VerifNumero = txtCTDE.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCTDE.Text = ""
        txtCTDE.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCTEX_Change()
On Error GoTo tratar_erro

If txtCTEX <> "" Then
    VerifNumero = txtCTEX
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCTEX = ""
        txtCTEX.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCTNN_Change()
On Error GoTo tratar_erro

If txtCTNN.Text <> "" Then
    VerifNumero = txtCTNN.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCTNN.Text = ""
        txtCTNN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCTSS_Change()
On Error GoTo tratar_erro

If txtCTSS.Text <> "" Then
    VerifNumero = txtCTSS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCTSS.Text = ""
        txtCTSS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtICMSco_Change()
On Error GoTo tratar_erro

If TxtICMSco.Text <> "" Then
    VerifNumero = TxtICMSco.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        TxtICMSco.Text = ""
        TxtICMSco.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtICMSde_Change()
On Error GoTo tratar_erro

If txtICMSde.Text <> "" Then
    VerifNumero = txtICMSde.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtICMSde.Text = ""
        txtICMSde.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtICMSnn_Change()
On Error GoTo tratar_erro

If txtICMSnn.Text <> "" Then
    VerifNumero = txtICMSnn.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtICMSnn.Text = ""
        txtICMSnn.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtICMSss_Change()
On Error GoTo tratar_erro

If txtICMSss.Text <> "" Then
    VerifNumero = txtICMSss.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtICMSss.Text = ""
        txtICMSss.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIPI_Change()
On Error GoTo tratar_erro

If txtIPI.Text <> "" Then
    VerifNumero = txtIPI.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIPI.Text = ""
        txtIPI.SetFocus
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: procRegiao
    Case 9: ProcValidarRegistros Lista, "Faturamento/Fiscal/Classificação fiscal"
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

