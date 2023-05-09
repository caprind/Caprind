VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmRelatorios_Custos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Custos - Relatórios - Resumido"
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
      ScreenHeight    =   768
      ScreenWidth     =   1366
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   9870
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmRelatorios_Custos.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   40
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   46
      ButtonTop2      =   2
      ButtonWidth2    =   60
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
      ButtonLeft3     =   108
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
      ButtonLeft4     =   112
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
      ButtonLeft5     =   155
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
      ButtonLeft6     =   187
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   55
      TabIndex        =   25
      Top             =   8760
      Width           =   15195
      Begin VB.TextBox Txt_outras 
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
         Left            =   8475
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de outras despesas."
         Top             =   480
         Width           =   1605
      End
      Begin VB.TextBox txtTerceiros 
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
         Left            =   6795
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de terceiros."
         Top             =   480
         Width           =   1665
      End
      Begin VB.TextBox txtValorTotal 
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Valor total."
         Top             =   480
         Width           =   1635
      End
      Begin VB.TextBox txtPercentual 
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
         Left            =   13380
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Percentual."
         Top             =   480
         Width           =   1635
      End
      Begin VB.TextBox txtLucro 
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
         Left            =   11745
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Lucro total."
         Top             =   480
         Width           =   1625
      End
      Begin VB.TextBox txtImpostos 
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
         Left            =   10095
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de impostos."
         Top             =   480
         Width           =   1635
      End
      Begin VB.TextBox txtObra 
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
         Left            =   3474
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de mão de obra."
         Top             =   480
         Width           =   1625
      End
      Begin VB.TextBox txtPrima 
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
         Left            =   5121
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de material."
         Top             =   480
         Width           =   1665
      End
      Begin VB.TextBox txtQtdePeca 
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
         Left            =   1827
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total."
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outras despesas"
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
         Left            =   8587
         TabIndex        =   44
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiros"
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
         Left            =   7230
         TabIndex        =   38
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentual"
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
         Left            =   13740
         TabIndex        =   33
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total"
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
         Left            =   555
         TabIndex        =   32
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   2157
         TabIndex        =   31
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
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
         Left            =   5601
         TabIndex        =   30
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mão de obra"
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
         Left            =   3766
         TabIndex        =   29
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impostos"
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
         Left            =   10507
         TabIndex        =   28
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lucro"
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
         Left            =   12330
         TabIndex        =   27
         Top             =   270
         Width           =   465
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6165
      Left            =   60
      TabIndex        =   8
      Top             =   2580
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   10874
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código interno"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "D"
         Text            =   "Dt. conclusão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. unit."
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Ct. mat."
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Ct. m. obra"
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Ct. terc."
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Ct. out."
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Impostos"
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "Lucro"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Object.Tag             =   "N"
         Text            =   "Percentual"
         Object.Width           =   2337
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Object.Tag             =   "T"
         Text            =   "Res. validado"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1015
      Left            =   7230
      TabIndex        =   26
      Top             =   1545
      Width           =   3315
      Begin VB.CheckBox Chk_qtde_prod 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Qtde. prod."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   690
         Width           =   1275
      End
      Begin VB.CheckBox chkLucro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lucro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1800
         TabIndex        =   4
         Top             =   690
         Width           =   1275
      End
      Begin VB.CheckBox chkValor 
         BackColor       =   &H00E0E0E0&
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
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkQuantidade 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1800
         TabIndex        =   2
         Top             =   390
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   55
      TabIndex        =   35
      Top             =   960
      Width           =   4485
      Begin VB.OptionButton Opt_valor_unitario 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor unitário"
         DisabledPicture =   "frmRelatorios_Custos.frx":2DF4
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
         Left            =   210
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton Opt_valor_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor total"
         DisabledPicture =   "frmRelatorios_Custos.frx":2F46
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
         Left            =   1860
         TabIndex        =   10
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4560
      TabIndex        =   36
      Top             =   960
      Width           =   10695
      Begin VB.OptionButton Opt_NF 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nota fiscal"
         DisabledPicture =   "frmRelatorios_Custos.frx":3098
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
         Left            =   1860
         TabIndex        =   42
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton Opt_PI 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pedido interno"
         DisabledPicture =   "frmRelatorios_Custos.frx":31EA
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
         Left            =   210
         TabIndex        =   41
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1035
      Left            =   10560
      TabIndex        =   20
      Top             =   1530
      Width           =   4695
      Begin VB.ComboBox cmbTipoData 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmRelatorios_Custos.frx":333C
         Left            =   150
         List            =   "frmRelatorios_Custos.frx":3343
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Tipo da data."
         Top             =   480
         Width           =   1560
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         ToolTipText     =   "Data final."
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   489816065
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   1710
         TabIndex        =   6
         ToolTipText     =   "Data inicio."
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   489816065
         CurrentDate     =   39057
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   645
         TabIndex        =   34
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   2325
         TabIndex        =   22
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
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
         Left            =   3705
         TabIndex        =   21
         Top             =   270
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   55
      TabIndex        =   23
      Top             =   1530
      Width           =   7155
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
         ItemData        =   "frmRelatorios_Custos.frx":334F
         Left            =   180
         List            =   "frmRelatorios_Custos.frx":3365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   480
         Width           =   6825
      End
      Begin VB.Label Label8 
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
         Left            =   3172
         TabIndex        =   24
         Top             =   270
         Width           =   840
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   37
      Top             =   9750
      Width           =   11775
      _ExtentX        =   20770
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
   Begin MSComctlLib.ListView Lista_ordem 
      Height          =   6165
      Left            =   13830
      TabIndex        =   43
      Top             =   2580
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   10874
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1826
      EndProperty
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11910
      TabIndex        =   39
      Top             =   9765
      Width           =   3315
   End
End
Attribute VB_Name = "frmRelatorios_Custos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataBanco As String 'OK
Dim DataTexto As String 'OK

Private Sub Chk_qtde_prod_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
If Chk_qtde_prod.Value = 1 Then
    chkValor.Value = 0
    chkQuantidade.Value = 0
    chkLucro.Value = 0
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

With Lista
    .ListItems.Clear
    ProcLimpaCamposTotais
    .ColumnHeaders(2).Text = cmbfiltrarpor
    If cmbfiltrarpor = "Ordem" Then
        .Width = 15195
        .ColumnHeaders(2).Tag = "N"
        .ColumnHeaders(2).Width = 800
        .ColumnHeaders(3).Width = 1400
        .ColumnHeaders(4).Width = 1200
        .ColumnHeaders(5).Width = 2800
        .ColumnHeaders(6).Width = 4000
        .ColumnHeaders(7).Width = 1300
        .ColumnHeaders(8).Width = 1000
        .ColumnHeaders(9).Width = 1200
        .ColumnHeaders(10).Width = 1275
        .ColumnHeaders(11).Width = 1275
        .ColumnHeaders(12).Width = 1275
        .ColumnHeaders(13).Width = 1275
        .ColumnHeaders(14).Width = 1275
        .ColumnHeaders(15).Width = 1275
        .ColumnHeaders(16).Width = 1275
        .ColumnHeaders(17).Width = 1325
        .ColumnHeaders(18).Width = 1400
    Else
        .Width = 13755
        .ColumnHeaders(2).Tag = "T"
        .ColumnHeaders(2).Width = 3200
        .ColumnHeaders(3).Width = 0
        .ColumnHeaders(4).Width = 0
        .ColumnHeaders(5).Width = 0
        .ColumnHeaders(6).Width = 0
        .ColumnHeaders(7).Width = 0
        .ColumnHeaders(8).Width = 1091
        .ColumnHeaders(9).Width = 1200
        .ColumnHeaders(10).Width = 0
        .ColumnHeaders(11).Width = 1091
        .ColumnHeaders(12).Width = 1091
        .ColumnHeaders(13).Width = 1091
        .ColumnHeaders(14).Width = 1091
        .ColumnHeaders(15).Width = 1091
        .ColumnHeaders(16).Width = 1091
        .ColumnHeaders(17).Width = 1325
        .ColumnHeaders(18).Width = 0
    End If
End With
Lista_ordem.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
frmRelatorios_Custos_menuimpressao.Show 1

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcAbrir
    Case vbKeyF5: ProcImprimir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If TBLISTA.EOF = False Then
    Posicao = TBLISTA.RecordCount
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            If cmbfiltrarpor = "Ordem" Then
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select Desenho, N_Referencia, Produto, dataentrega, DtValidacao_custo from Producao where Ordem = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBOrdem!N_referencia), "", TBOrdem!N_referencia)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Execucaoprev), "", TBLISTA!Execucaoprev)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBOrdem!DataEntrega), "", Format(TBOrdem!DataEntrega, "dd/mm/yy"))
                    .Item(.Count).SubItems(17) = IIf(IsNull(TBOrdem!DtValidacao_custo), "Não", "Sim")
                End If
                TBOrdem.Close
            Else
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
            End If
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!QtdePrev), "0,00", Format(TBLISTA!QtdePrev, "###,##0.00"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!qtdeOK), "0,0000", Format(TBLISTA!qtdeOK, "###,##0.0000"))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Data1), "", TBLISTA!Data1)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Qtdetotalprod), "0,00", Format(TBLISTA!Qtdetotalprod, "###,##0.00"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Eficiencia), "0,00", Format(TBLISTA!Eficiencia, "###,##0.00"))
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Terceiros), "0,00", Format(TBLISTA!Terceiros, "###,##0.00"))
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Numero1), "0,00", Format(TBLISTA!Numero1, "###,##0.00"))
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!impostos), "0,00", Format(TBLISTA!impostos, "###,##0.00"))
            .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Lucro), "0,00", Format(TBLISTA!Lucro, "###,##0.00"))
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!Valor1), "0,00", Format(TBLISTA!Valor1, "###,##0.00") & "%")
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
    
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select QtdePrevista, QtdeProduzida, QtdeNC, QtdeOrdem, CustoMat, CustoObra, Terceros, Numero1, Lucro, Valor1 from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtQtdePeca = IIf(IsNull(TBLISTA!QtdePrevista), "0,0000", Format(TBLISTA!QtdePrevista, "###,##0.0000"))
    txtimpostos = IIf(IsNull(TBLISTA!QtdeProduzida), "0,00", Format(TBLISTA!QtdeProduzida, "###,##0.00"))
    txtValorTotal = IIf(IsNull(TBLISTA!qtdeNC), "0,00", Format(TBLISTA!qtdeNC, "###,##0.00"))
    txtQtdeOrdem = IIf(IsNull(TBLISTA!QtdeOrdem), "0,0000", Format(TBLISTA!QtdeOrdem, "###,##0.0000"))
    txtPrima = IIf(IsNull(TBLISTA!CustoMat), "0,00", Format(TBLISTA!CustoMat, "###,##0.00"))
    txtObra = IIf(IsNull(TBLISTA!CustoObra), "0,00", Format(TBLISTA!CustoObra, "###,##0.00"))
    txtTerceiros = IIf(IsNull(TBLISTA!Terceros), "0,00", Format(TBLISTA!Terceros, "###,##0.00"))
    Txt_outras = IIf(IsNull(TBLISTA!Numero1), "0,00", Format(TBLISTA!Numero1, "###,##0.00"))
    txtLucro = IIf(IsNull(TBLISTA!Lucro), "0,00", Format(TBLISTA!Lucro, "###,##0.00"))
    txtPercentual = IIf(IsNull(TBLISTA!Valor1), "0,00", Format(TBLISTA!Valor1, "###,##0.00") & "%")
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
txtimpostos = ""
txtLucro = ""
txtObra = ""
txtPrima = ""
txtQtdeOrdem = ""
txtQtdePeca = ""
txtTerceiros = ""
Txt_outras = ""
txtValorTotal = ""
txtPercentual = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
Formulario = "Custos/Relatórios/Resumido"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
cmbfiltrarpor.Text = "Código interno"
cmbTipoData = "Vendas"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Custos/Relatórios/Resumido"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Or cmbfiltrarpor = "Ordem" Then Exit Sub
Lista_ordem.ListItems.Clear
Set TBCarteira = CreateObject("adodb.recordset")
If Opt_PI.Value = True Then
    Select Case cmbfiltrarpor
        Case "Código interno": TextoFiltro = "VC.Desenho"
        Case "Código de referência": TextoFiltro = "VC.n_referencia"
        Case "Descrição": TextoFiltro = "VC.Descricao_tecnica"
        Case "Família": TextoFiltro = "VC.familia"
        Case "Cliente": TextoFiltro = "VP.Cliente"
    End Select
    TBCarteira.Open "Select P.Ordem As Ordem, P.Dtvalidacao_custo from ((Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem) INNER JOIN vendas_carteira VC ON VC.codigo = PP.IDCarteira) INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao where " & TextoFiltro & " = '" & Lista.SelectedItem.ListSubItems(1) & "' and VC.Datavendas Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and VC.Retorno = 'False' and (VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'VENDIDA' or VC.liberacao = 'VENDIDA PARCIAL') group by P.Ordem, P.Dtvalidacao_custo", Conexao, adOpenKeyset, adLockReadOnly
Else
    Select Case cmbfiltrarpor
        Case "Código interno": TextoFiltro = "NFP.int_Cod_Produto"
        Case "Código de referência": TextoFiltro = "NFP.n_referencia"
        Case "Descrição": TextoFiltro = "NFP.Descricao_tecnica"
        Case "Família": TextoFiltro = "NFP.familia"
        Case "Cliente": TextoFiltro = "NFP.Cliente"
    End Select
    TBCarteira.Open "Select EM.Lote As Ordem, P.Dtvalidacao_custo from (((((Estoque_movimentacao EM INNER JOIN Ordens_texto_SA OT ON OT.Ordem = EM.Lote) INNER JOIN Producao P ON P.ordem = OT.Ordem) INNER JOIN tbl_Detalhes_Nota NFP ON EM.ID_prod_NF = NFP.Int_codigo) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFP.ID_nota) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP) INNER JOIN tbl_Detalhes_Nota_CST_ICMS CST ON CST.ID_item = NFP.Int_codigo where " & TextoFiltro & " = '" & Lista.SelectedItem.ListSubItems(1) & "' and NF.int_TipoNota = 1 and NF.int_status = 1 and NF.dt_DataEmissao Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and NFP.Retorno = 'False' and NFP.Remessa = 'False' and IsNumeric(EM.Lote) = 1 and (CFOP.Vendas = 'True' or CFOP.MaoObra = 'True') and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL') group by EM.Lote, P.Dtvalidacao_custo", Conexao, adOpenKeyset, adLockReadOnly
End If
Lista_ordem.ListItems.Clear
If TBCarteira.EOF = False Then
    Do While TBCarteira.EOF = False
        With Lista_ordem.ListItems
            .Add , , TBCarteira!Ordem & " - " & IIf(IsNull(TBCarteira!DtValidacao_custo), "RNV", "RV")
        End With
        TBCarteira.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

Acao = "filtrar"
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
Inicio = Time
ProcLimpaCamposTotais
ProcAbrirTabelas
ProcGravarTotalizacoes

If chkValor.Value = 1 Then
    Texto = "Qtdeprev desc"
ElseIf chkQuantidade.Value = 1 Then
        Texto = "QtdeOK desc"
    ElseIf Chk_qtde_prod.Value = 1 Then
        Texto = "QtdeNC desc"
    ElseIf chkLucro.Value = 1 Then
            Texto = "Lucro desc"
        Else
            If cmbfiltrarpor = "Ordem" Then Texto = "Ordem" Else Texto = "maquina"
End If
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by " & Texto, Conexao, adOpenKeyset, adLockOptimistic
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

IDlista = 0
Ordem = 0

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

If Opt_PI.Value = True Then
    CamposFiltro = "VC.Codigo As ID_carteira, VC.Desenho As Desenho, VC.Tipo, VC.Quantidade As Qtde, VC.dbl_Valor_ICMS As VlrICMS, VC.preco_unitario_desconto As VlrUnit, VC.Total_PIS_prod As VlrPISProd, VC.Total_Cofins_prod As VlrCofinsProd, VC.Total_CSLL_prod As VlrCSLLProd, VC.Total_IRPJ_prod As VlrIRPJProd, VC.vlriss As VlrISSServ, VC.Total_IRRF_serv As VlrIRRFServ, VC.Total_IRPJ_serv As VlrIRPJServ, VC.Total_PIS_serv As VlrPISServ, VC.Total_Cofins_serv As VlrCofinsServ, VC.Total_CSLL_serv As VlrCSLLServ, VC.Total_INSS_serv As VlrINSSServ"
    Grupar = "VC.Codigo, VC.Desenho, VC.Tipo, VC.Quantidade, VC.dbl_Valor_ICMS, VC.preco_unitario_desconto, VC.Total_PIS_prod, VC.Total_Cofins_prod, VC.Total_CSLL_prod, VC.Total_IRPJ_prod, VC.vlriss, VC.Total_IRRF_serv, VC.Total_IRPJ_serv, VC.Total_PIS_serv, VC.Total_Cofins_serv, VC.Total_CSLL_serv, VC.Total_INSS_serv"
    Set TBCarteira = CreateObject("adodb.recordset")
    If cmbfiltrarpor = "Ordem" Then
        TBCarteira.Open "Select P.Ordem As CampoFiltro, " & CamposFiltro & ", VP.Cliente from ((Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem) INNER JOIN vendas_carteira VC ON VC.codigo = PP.IDCarteira) INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao where VC.Datavendas Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and VC.Retorno = 'False' and (VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'VENDIDA' or VC.liberacao = 'VENDIDA PARCIAL') group by P.Ordem, " & Grupar & ", VP.Cliente order by P.Ordem", Conexao, adOpenKeyset, adLockReadOnly
    Else
        Select Case cmbfiltrarpor
            Case "Código interno": Ordenar = "VC.Desenho"
            Case "Código de referência": Ordenar = "VC.n_referencia"
            Case "Descrição": Ordenar = "VC.Descricao_tecnica"
            Case "Família": Ordenar = "VC.familia"
            Case "Cliente": Ordenar = "VP.Cliente"
        End Select
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select " & Ordenar & " As CampoFiltro, " & CamposFiltro & " from vendas_carteira VC INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao where VC.Datavendas Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and VC.Retorno = 'False' and (VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'VENDIDA' or VC.liberacao = 'VENDIDA PARCIAL') order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
    End If
Else
    Set TBCarteira = CreateObject("adodb.recordset")
    If cmbfiltrarpor = "Ordem" Then
        CamposFiltro = "NFP.Int_codigo, NFPP.ID_carteira As ID_carteira, NFP.int_Cod_Produto As Desenho, NFP.Tipo, NFPP.Quantidade As Qtde, NF.int_NotaFiscal, CST.Valor_ICMS As VlrICMS, NFP.dbl_ValorUnitario As VlrUnit, NFP.Total_PIS_prod As VlrPISProd, NFP.Total_Cofins_prod As VlrCofinsProd, NFP.Total_CSLL_prod As VlrCSLLProd, NFP.Total_IRPJ_prod As VlrIRPJProd, NFP.VlrISS As VlrISSServ, NFP.Total_IRRF_serv As VlrIRRFServ, NFP.Total_IRPJ_serv As VlrIRPJServ, NFP.Total_PIS_serv As VlrPISServ, NFP.Total_Cofins_serv As VlrCofinsServ, NFP.Total_CSLL_serv As VlrCSLLServ, NFP.Total_INSS_serv As VlrINSSServ"
        Grupar = "NFP.Int_codigo,  NFPP.ID_carteira, NFP.int_Cod_Produto, NFP.Tipo, NFPP.Quantidade, NF.int_NotaFiscal, CST.Valor_ICMS, NFP.dbl_ValorUnitario, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, NFP.Total_CSLL_prod, NFP.Total_IRPJ_prod, NFP.VlrISS, NFP.Total_IRRF_serv, NFP.Total_IRPJ_serv, NFP.Total_PIS_serv, NFP.Total_Cofins_serv, NFP.Total_CSLL_serv, NFP.Total_INSS_serv"
        TBCarteira.Open "Select EM.Lote As CampoFiltro, EM.Saida, " & CamposFiltro & ", VP.Cliente from ((((((Estoque_movimentacao EM INNER JOIN tbl_Detalhes_Nota NFP ON EM.ID_prod_NF = NFP.Int_codigo) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFP.ID_nota) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo) INNER JOIN vendas_carteira VC ON VC.codigo = NFPP.ID_carteira and VC.Desenho = NFPP.Codinterno) INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao) INNER JOIN tbl_Detalhes_Nota_CST_ICMS CST ON CST.ID_item = NFP.Int_codigo where NF.int_TipoNota = 1 and NF.int_status = 1 and NF.dt_DataEmissao Between '" & _
            Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and NFP.Retorno = 'False' and NFP.Remessa = 'False' and IsNumeric(EM.Lote) = 1 and (CFOP.Vendas = 'True' or CFOP.MaoObra = 'True') and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL') group by EM.Lote, EM.Saida, " & Grupar & ", VP.Cliente order by EM.Lote", Conexao, adOpenKeyset, adLockReadOnly
    Else
        CamposFiltro = "NFP.Int_codigo, NFPP.ID_carteira As ID_carteira, NFP.int_Cod_Produto As Desenho, NFP.Tipo, NFPP.Quantidade As Qtde, NF.int_NotaFiscal, CST.Valor_ICMS As VlrICMS, NFP.dbl_ValorUnitario As VlrUnit, NFP.Total_PIS_prod As VlrPISProd, NFP.Total_Cofins_prod As VlrCofinsProd, NFP.Total_CSLL_prod As VlrCSLLProd, NFP.Total_IRPJ_prod As VlrIRPJProd, NFP.VlrISS As VlrISSServ, NFP.Total_IRRF_serv As VlrIRRFServ, NFP.Total_IRPJ_serv As VlrIRPJServ, NFP.Total_PIS_serv As VlrPISServ, NFP.Total_Cofins_serv As VlrCofinsServ, NFP.Total_CSLL_serv As VlrCSLLServ, NFP.Total_INSS_serv As VlrINSSServ"
        Grupar = "NFP.Int_codigo, NFPP.ID_carteira, NFP.int_Cod_Produto, NFP.Tipo, NFPP.Quantidade, NF.int_NotaFiscal, CST.Valor_ICMS, NFP.dbl_ValorUnitario, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, NFP.Total_CSLL_prod, NFP.Total_IRPJ_prod, NFP.VlrISS, NFP.Total_IRRF_serv, NFP.Total_IRPJ_serv, NFP.Total_PIS_serv, NFP.Total_Cofins_serv, NFP.Total_CSLL_serv, NFP.Total_INSS_serv"
        Select Case cmbfiltrarpor
            Case "Código interno": Ordenar = "NFP.int_Cod_Produto"
            Case "Código de referência": Ordenar = "NFP.n_referencia"
            Case "Descrição": Ordenar = "NFP.Descricao_tecnica"
            Case "Família": Ordenar = "NFP.familia"
            Case "Cliente": Ordenar = "NFP.Cliente"
        End Select
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select " & Ordenar & " As CampoFiltro, " & CamposFiltro & " from (((tbl_Detalhes_Nota NFP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NFP.ID_nota = NF.ID) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP) INNER JOIN tbl_Detalhes_Nota_CST_ICMS CST ON CST.ID_item = NFP.Int_codigo where NF.int_TipoNota = 1 and NF.int_status = 1 and NF.dt_DataEmissao Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and NFP.Retorno = 'False' and NFP.Remessa = 'False' and (CFOP.Vendas = 'True' or CFOP.MaoObra = 'True') order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
    End If
End If
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If TBCarteira.EOF = False Then
    Permitido = True
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
        Set TBProdutividade = CreateObject("adodb.recordset")
        If cmbfiltrarpor = "Ordem" Then TextoFiltro = "and Execucaoprev = '" & TBCarteira!Cliente & "' and Total = " & TBCarteira!ID_carteira Else TextoFiltro = ""
        TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!CampoFiltro & "' " & TextoFiltro & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcEnviaDadosResumido
        TBCarteira.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

Qtde = 0
qt = 0
Qtd = 0
quantidade = 0
qtdeliberada = 0
ValorICMS = 0

If TBProdutividade.EOF = True Then TBProdutividade.AddNew
Texto = TBCarteira!CampoFiltro
If cmbfiltrarpor = "Ordem" Then
    TBProdutividade!Ordem = TBCarteira!CampoFiltro
    TBProdutividade!Execucaoprev = TBCarteira!Cliente
    TBProdutividade!Total = TBCarteira!ID_carteira
    
    If Opt_NF.Value = True Then TBProdutividade!Data1 = TBCarteira!int_NotaFiscal
End If
TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario

'Busca valor do ICMS
ValorICMS = IIf(IsNull(TBCarteira!vlrICMS), 0, TBCarteira!vlrICMS)

If TBCarteira!ID_carteira <> 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    CamposFiltro = "P.Ordem, P.Quant, P.QuantProd, P.QuantNC, P.CTTReal, P.CPR, P.CTMaterial, P.CTOutras, P.CTServico, P.consignacao"
    If Opt_PI.Value = True Then
        INNERJOINTEXTO = "Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem where PP.IDCarteira"
    Else
        INNERJOINTEXTO = "((Producao P INNER JOIN Estoque_movimentacao EM ON CAST(P.Ordem AS varchar(10)) = EM.Lote) INNER JOIN tbl_Detalhes_Nota NFP ON EM.ID_prod_NF = NFP.Int_codigo) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo where NFP.Retorno = 'False' and NFP.Remessa = 'False' and EM.ID_prod_NF = " & TBCarteira!Int_codigo & " and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL') and NFPP.ID_carteira"
        CamposFiltro = CamposFiltro & ", EM.Saida"
    End If
    If cmbfiltrarpor = "Ordem" Then TextoFiltro = "and P.Ordem = " & TBCarteira!CampoFiltro Else TextoFiltro = ""
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & CamposFiltro & " from " & INNERJOINTEXTO & " = " & TBCarteira!ID_carteira & " and P.Desenho = '" & TBCarteira!Desenho & "' and P.Tipo = 'E' and P.QuantProd > 0 " & TextoFiltro & " group by " & CamposFiltro & " order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select " & CamposFiltro & " from " & INNERJOINTEXTO & " = " & TBCarteira!ID_carteira & " and P.Tipo = 'E' and P.QuantProd > 0 " & TextoFiltro & " group by " & CamposFiltro & " order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select " & CamposFiltro & " from " & INNERJOINTEXTO & " = " & TBCarteira!ID_carteira & " and P.QuantProd > 0 " & TextoFiltro & " group by " & CamposFiltro & " order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
        End If
    End If
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Texto = TBAbrir!Ordem
            If Opt_PI.Value = True Then QuantsolicitadoN2 = TBCarteira!Qtde Else QuantsolicitadoN2 = TBAbrir!Saida
            
                                                'ORDEM         QTDE. PREVISTA                                QTDE. OK                                              QT. PROD.(OK+NC)                                                                                         CUSTO LOTE                                        CUSTO PEÇA                                CUSTO TERCEIROS                                       CUSTO MATERIAL                                          CUSTO OUTRAS                                        ORDEM CONSIGNADA
            ValorunitPC = FunCalculaValorUnitOrdem(TBAbrir!Ordem, IIf(IsNull(TBAbrir!Quant), 0, TBAbrir!Quant), IIf(IsNull(TBAbrir!QuantProd), 0, TBAbrir!QuantProd), IIf(IsNull(TBAbrir!QuantProd), 0, TBAbrir!QuantProd) + IIf(IsNull(TBAbrir!QuantNC), 0, TBAbrir!QuantNC), IIf(IsNull(TBAbrir!CTTReal), 0, TBAbrir!CTTReal), IIf(IsNull(TBAbrir!CPR), 0, TBAbrir!CPR), IIf(IsNull(TBAbrir!CTServico), 0, TBAbrir!CTServico), IIf(IsNull(TBAbrir!CTMaterial), 0, TBAbrir!CTMaterial), IIf(IsNull(TBAbrir!CTOutras), 0, TBAbrir!CTOutras), TBAbrir!consignacao)
            
            'Verifica quantidade vendida e calcula custo total
            ValorTotal = Valor3 * QuantsolicitadoN2
            Valor1 = Valor_Produto * QuantsolicitadoN2
            Valor2 = ValorPagar * QuantsolicitadoN2
            Valor_ICMS_SN = Valor_DAS * QuantsolicitadoN2
                       
            TBProdutividade!qtdeOK = IIf(IsNull(TBProdutividade!qtdeOK), 0, TBProdutividade!qtdeOK) + QuantsolicitadoN2 'Qtde. peça
            
            If Opt_valor_total.Value = True Then
                TBProdutividade!Eficiencia = IIf(IsNull(TBProdutividade!Eficiencia), 0, TBProdutividade!Eficiencia) + ValorTotal 'Custo mão de obra
                TBProdutividade!Qtdetotalprod = IIf(IsNull(TBProdutividade!Qtdetotalprod), 0, TBProdutividade!Qtdetotalprod) + Valor1 'Custo materia prima
                TBProdutividade!Terceiros = IIf(IsNull(TBProdutividade!Terceiros), 0, TBProdutividade!Terceiros) + Valor2 'Custo terceiros
                TBProdutividade!Numero1 = IIf(IsNull(TBProdutividade!Numero1), 0, TBProdutividade!Numero1) + Valor_ICMS_SN 'Custo outras
                
                TBProdutividade!QtdePrev = IIf(IsNull(TBProdutividade!QtdePrev), 0, TBProdutividade!QtdePrev) + (IIf(IsNull(TBCarteira!VlrUnit), 0, TBCarteira!VlrUnit) * QuantsolicitadoN2) 'Valor total
                
                'Impostos
                If TBCarteira!Tipo = "P" Then
                    TBProdutividade!impostos = ((IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + ValorICMS + IIf(IsNull(TBCarteira!VlrPISProd), 0, TBCarteira!VlrPISProd) + IIf(IsNull(TBCarteira!VlrCofinsProd), 0, TBCarteira!VlrCofinsProd) + IIf(IsNull(TBCarteira!VlrCSLLProd), 0, TBCarteira!VlrCSLLProd) + IIf(IsNull(TBCarteira!VlrIRPJProd), 0, TBCarteira!VlrIRPJProd)) / TBCarteira!Qtde) * QuantsolicitadoN2
                Else
                    TBProdutividade!impostos = ((IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + IIf(IsNull(TBCarteira!VlrISSServ), 0, TBCarteira!VlrISSServ) + IIf(IsNull(TBCarteira!VlrIRRFServ), 0, TBCarteira!VlrIRRFServ) + IIf(IsNull(TBCarteira!VlrIRPJServ), 0, TBCarteira!VlrIRPJServ) + IIf(IsNull(TBCarteira!VlrPISServ), 0, TBCarteira!VlrPISServ) + IIf(IsNull(TBCarteira!VlrCofinsServ), 0, TBCarteira!VlrCofinsServ) + IIf(IsNull(TBCarteira!VlrCSLLServ), 0, TBCarteira!VlrCSLLServ) + IIf(IsNull(TBCarteira!VlrINSSServ), 0, TBCarteira!VlrINSSServ)) / TBCarteira!Qtde) * QuantsolicitadoN2
                End If
            Else
                TBProdutividade!Eficiencia = IIf(IsNull(TBProdutividade!Eficiencia), 0, TBProdutividade!Eficiencia) + Valor3 'Custo mão de obra
                TBProdutividade!Qtdetotalprod = IIf(IsNull(TBProdutividade!Qtdetotalprod), 0, TBProdutividade!Qtdetotalprod) + Valor_Produto 'Custo materia prima
                TBProdutividade!Terceiros = IIf(IsNull(TBProdutividade!Terceiros), 0, TBProdutividade!Terceiros) + ValorPagar 'Custo terceiros
                TBProdutividade!Numero1 = IIf(IsNull(TBProdutividade!Numero1), 0, TBProdutividade!Numero1) + Valor_DAS 'Custo outras
                
                TBProdutividade!QtdePrev = IIf(IsNull(TBProdutividade!QtdePrev), 0, TBProdutividade!QtdePrev) + IIf(IsNull(TBCarteira!VlrUnit), 0, TBCarteira!VlrUnit) 'Valor unitário
                
                'Impostos
                If TBCarteira!Tipo = "P" Then
                    TBProdutividade!impostos = (IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + ValorICMS + IIf(IsNull(TBCarteira!VlrPISProd), 0, TBCarteira!VlrPISProd) + IIf(IsNull(TBCarteira!VlrCofinsProd), 0, TBCarteira!VlrCofinsProd) + IIf(IsNull(TBCarteira!VlrCSLLProd), 0, TBCarteira!VlrCSLLProd) + IIf(IsNull(TBCarteira!VlrIRPJProd), 0, TBCarteira!VlrIRPJProd)) / TBCarteira!Qtde
                Else
                    TBProdutividade!impostos = (IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + IIf(IsNull(TBCarteira!VlrISSServ), 0, TBCarteira!VlrISSServ) + IIf(IsNull(TBCarteira!VlrIRRFServ), 0, TBCarteira!VlrIRRFServ) + IIf(IsNull(TBCarteira!VlrIRPJServ), 0, TBCarteira!VlrIRPJServ) + IIf(IsNull(TBCarteira!VlrPISServ), 0, TBCarteira!VlrPISServ) + IIf(IsNull(TBCarteira!VlrCofinsServ), 0, TBCarteira!VlrCofinsServ) + IIf(IsNull(TBCarteira!VlrCSLLServ), 0, TBCarteira!VlrCSLLServ) + IIf(IsNull(TBCarteira!VlrINSSServ), 0, TBCarteira!VlrINSSServ)) / TBCarteira!Qtde
                End If
            End If
            
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
Else
    TBProdutividade!qtdeOK = IIf(IsNull(TBProdutividade!qtdeOK), 0, TBProdutividade!qtdeOK) + TBCarteira!Qtde 'Qtde. peça
    QuantsolicitadoN1 = TBCarteira!Qtde
    If Opt_valor_total.Value = True Then
        TBProdutividade!QtdePrev = IIf(IsNull(TBProdutividade!QtdePrev), 0, TBProdutividade!QtdePrev) + (IIf(IsNull(TBCarteira!VlrUnit), 0, TBCarteira!VlrUnit) * QuantsolicitadoN1) 'Valor total
        'Impostos
        If TBCarteira!Tipo = "P" Then
            TBProdutividade!impostos = IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + ValorICMS + IIf(IsNull(TBCarteira!VlrPISProd), 0, TBCarteira!VlrPISProd) + IIf(IsNull(TBCarteira!VlrCofinsProd), 0, TBCarteira!VlrCofinsProd) + IIf(IsNull(TBCarteira!VlrCSLLProd), 0, TBCarteira!VlrCSLLProd) + IIf(IsNull(TBCarteira!VlrIRPJProd), 0, TBCarteira!VlrIRPJProd)
        Else
            TBProdutividade!impostos = IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + IIf(IsNull(TBCarteira!VlrISSServ), 0, TBCarteira!VlrISSServ) + IIf(IsNull(TBCarteira!VlrIRRFServ), 0, TBCarteira!VlrIRRFServ) + IIf(IsNull(TBCarteira!VlrIRPJServ), 0, TBCarteira!VlrIRPJServ) + IIf(IsNull(TBCarteira!VlrPISServ), 0, TBCarteira!VlrPISServ) + IIf(IsNull(TBCarteira!VlrCofinsServ), 0, TBCarteira!VlrCofinsServ) + IIf(IsNull(TBCarteira!VlrCSLLServ), 0, TBCarteira!VlrCSLLServ) + IIf(IsNull(TBCarteira!VlrINSSServ), 0, TBCarteira!VlrINSSServ)
        End If
    Else
        TBProdutividade!QtdePrev = IIf(IsNull(TBProdutividade!QtdePrev), 0, TBProdutividade!QtdePrev) + IIf(IsNull(TBCarteira!VlrUnit), 0, TBCarteira!VlrUnit) 'Valor unitário
        'Impostos
        If TBCarteira!Tipo = "P" Then
            TBProdutividade!impostos = IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + ((ValorICMS + IIf(IsNull(TBCarteira!VlrPISProd), 0, TBCarteira!VlrPISProd) + IIf(IsNull(TBCarteira!VlrCofinsProd), 0, TBCarteira!VlrCofinsProd) + IIf(IsNull(TBCarteira!VlrCSLLProd), 0, TBCarteira!VlrCSLLProd) + IIf(IsNull(TBCarteira!VlrIRPJProd), 0, TBCarteira!VlrIRPJProd)) / QuantsolicitadoN1)
        Else
            TBProdutividade!impostos = IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) + ((IIf(IsNull(TBCarteira!VlrISSServ), 0, TBCarteira!VlrISSServ) + IIf(IsNull(TBCarteira!VlrIRRFServ), 0, TBCarteira!VlrIRRFServ) + IIf(IsNull(TBCarteira!VlrIRPJServ), 0, TBCarteira!VlrIRPJServ) + IIf(IsNull(TBCarteira!VlrPISServ), 0, TBCarteira!VlrPISServ) + IIf(IsNull(TBCarteira!VlrCofinsServ), 0, TBCarteira!VlrCofinsServ) + IIf(IsNull(TBCarteira!VlrCSLLServ), 0, TBCarteira!VlrCSLLServ) + IIf(IsNull(TBCarteira!VlrINSSServ), 0, TBCarteira!VlrINSSServ)) / QuantsolicitadoN1)
        End If
    End If
End If

Qtde = IIf(IsNull(TBProdutividade!QtdePrev), 0, TBProdutividade!QtdePrev) 'Valor de venda/fat
qt = IIf(IsNull(TBProdutividade!Eficiencia), 0, TBProdutividade!Eficiencia) 'Custo mão de obra
Qtd = IIf(IsNull(TBProdutividade!Qtdetotalprod), 0, TBProdutividade!Qtdetotalprod) 'Custo materia prima
quantidade = IIf(IsNull(TBProdutividade!Terceiros), 0, TBProdutividade!Terceiros) 'Custo terceiros
qtdeliberada = IIf(IsNull(TBProdutividade!Numero1), 0, TBProdutividade!Numero1) 'Custo outras
ValorICMS = IIf(IsNull(TBProdutividade!impostos), 0, TBProdutividade!impostos) 'Impostos

VltUnit = Qtde - (Qtd + qt + quantidade + ValorICMS)
TBProdutividade!Lucro = VltUnit 'Lucro
If VltUnit <> 0 And Qtde <> 0 Then TBProdutividade!Valor1 = (VltUnit / Qtde) * 100

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

Qtde = 0
Qtd = 0
qt = 0
quantidade = 0
ValorTotal = 0
ValorICMS = 0
VltUnit = 0
qtdeliberada = 0
VlttTotal = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
TBAbrir!Texto = cmbfiltrarpor
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
If Lista.ColumnHeaders(8).Text = "Vlr. unit." Then TBAbrir!Turno = True Else TBAbrir!Turno = False
If Opt_PI.Value = True Then TBAbrir!Totalutilizada = 0 Else TBAbrir!Totalutilizada = 1
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(QtdeOK) as Qtde, Sum(Impostos) as Qtd, Sum(QtdePrev) as qt, Sum(QtdeNC) as Quantidade, Sum(Qtdetotalprod) as Valortotal, Sum(Eficiencia) as ValorICMS, Sum(Terceiros) as VltUnit, Sum(Numero1) as qtdeliberada, Sum(Lucro) as VlttTotal from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Qtde = IIf(IsNull(TBproducao!Qtde), 0, TBproducao!Qtde)
    Qtd = IIf(IsNull(TBproducao!Qtd), 0, TBproducao!Qtd)
    qt = IIf(IsNull(TBproducao!qt), 0, TBproducao!qt)
    quantidade = IIf(IsNull(TBproducao!quantidade), 0, TBproducao!quantidade)
    ValorTotal = IIf(IsNull(TBproducao!ValorTotal), 0, TBproducao!ValorTotal)
    ValorICMS = IIf(IsNull(TBproducao!ValorICMS), 0, TBproducao!ValorICMS)
    VltUnit = IIf(IsNull(TBproducao!VltUnit), 0, TBproducao!VltUnit)
    qtdeliberada = IIf(IsNull(TBproducao!qtdeliberada), 0, TBproducao!qtdeliberada)
    VlttTotal = IIf(IsNull(TBproducao!VlttTotal), 0, TBproducao!VlttTotal)
End If
TBproducao.Close

TBAbrir!QtdePrevista = Qtde 'Qtde peça
TBAbrir!QtdeProduzida = Qtd 'Impostos
TBAbrir!qtdeNC = qt 'Valor total
TBAbrir!QtdeOrdem = quantidade 'Qtde ordem
TBAbrir!CustoMat = ValorTotal 'Custo materia prima
TBAbrir!CustoObra = ValorICMS 'Custo mão de obra
TBAbrir!Terceros = VltUnit 'Custo terceiros
TBAbrir!Numero1 = qtdeliberada 'Custo outras
TBAbrir!Lucro = VlttTotal 'Lucro
If qt <> 0 Then TBAbrir!Valor1 = (VlttTotal / qt) * 100

TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkQuantidade_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
If chkQuantidade.Value = 1 Then
    chkValor.Value = 0
    Chk_qtde_prod.Value = 0
    chkLucro.Value = 0
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkValor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
If chkValor.Value = 1 Then
    chkQuantidade.Value = 0
    Chk_qtde_prod.Value = 0
    chkLucro.Value = 0
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkLucro_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
If chkLucro.Value = 1 Then
    chkValor.Value = 0
    chkQuantidade.Value = 0
    Chk_qtde_prod.Value = 0
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_NF_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
If Opt_NF.Value = True Then
    With cmbTipoData
        .Clear
        .AddItem "Faturamento"
        .Text = "Faturamento"
    End With
    If cmbfiltrarpor = "Ordem" Then Lista.ColumnHeaders(10).Width = 1275 Else Lista.ColumnHeaders(10).Width = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_PI_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
If Opt_PI.Value = True Then
    With cmbTipoData
        .Clear
        .AddItem "Vendas"
        .Text = "Vendas"
    End With
    Lista.ColumnHeaders(10).Width = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_valor_total_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
Lista.ColumnHeaders(8).Text = "Vlr. total"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_valor_unitario_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
ProcLimpaCamposTotais
Lista.ColumnHeaders(8).Text = "Vlr. unit."

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcAbrir
    Case 2: ProcImprimir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

