VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_Relatorios_Historico 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Relatórios - Histórico"
   ClientHeight    =   10035
   ClientLeft      =   10350
   ClientTop       =   450
   ClientWidth     =   15360
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
   ForeColor       =   &H8000000D&
   Icon            =   "frmVendas_Relatorios_Historico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
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
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
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
      Height          =   825
      Left            =   60
      TabIndex        =   36
      Top             =   9180
      Width           =   15195
      Begin VB.TextBox txtTotal_geral 
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
         Left            =   10317
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Total vendido."
         Top             =   375
         Width           =   2205
      End
      Begin VB.TextBox txtTotal_ipi 
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
         Left            =   6938
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Total de IPI."
         Top             =   375
         Width           =   2205
      End
      Begin VB.TextBox txtTotal_produtos 
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
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Total de produtos."
         Top             =   375
         Width           =   2205
      End
      Begin VB.TextBox Txt_total_servicos 
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
         Left            =   3559
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Total de serviços."
         Top             =   375
         Width           =   2205
      End
      Begin VB.TextBox Txt_qtde_total_vendido 
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
         Left            =   13050
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total vendida."
         Top             =   375
         Width           =   1935
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   2912
         TabIndex        =   44
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   195
         Left            =   9670
         TabIndex        =   43
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total vendido"
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
         Left            =   10842
         TabIndex        =   42
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total produtos"
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
         Left            =   660
         TabIndex        =   41
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total IPI"
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
         Left            =   7673
         TabIndex        =   40
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   6291
         TabIndex        =   39
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total serviços"
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
         Left            =   4069
         TabIndex        =   38
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total vendida"
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
         Left            =   13207
         TabIndex        =   37
         Top             =   180
         Width           =   1620
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   12720
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_Relatorios_Historico.frx":0442
      Count           =   1
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6585
      Left            =   60
      TabIndex        =   18
      Top             =   2310
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11615
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Pedido int."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Valor IPI"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Tag             =   "D"
         Text            =   "Dt. venda"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6585
      Left            =   60
      TabIndex        =   19
      Top             =   2310
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11615
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
      NumItems        =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   35
      Top             =   0
      Width           =   15195
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
      ButtonLeft3     =   93
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
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
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
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
      Height          =   1335
      Left            =   60
      TabIndex        =   25
      Top             =   960
      Width           =   1695
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
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
         Left            =   180
         TabIndex        =   0
         Top             =   450
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
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
         Left            =   180
         TabIndex        =   1
         Top             =   720
         Width           =   1425
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
      Height          =   1335
      Left            =   1770
      TabIndex        =   26
      Top             =   960
      Width           =   1455
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
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
         Left            =   180
         TabIndex        =   2
         Top             =   450
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
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
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1335
      Left            =   3240
      TabIndex        =   27
      Top             =   960
      Width           =   9885
      Begin VB.TextBox Txt_limite 
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Limite de registros para carregar na lista."
         Top             =   900
         Width           =   555
      End
      Begin VB.OptionButton Opt_valor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor"
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
         Height          =   195
         Left            =   7470
         TabIndex        =   7
         Top             =   990
         Width           =   915
      End
      Begin VB.OptionButton Opt_quantidade 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quantidade"
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
         Height          =   195
         Left            =   8430
         TabIndex        =   8
         Top             =   990
         Width           =   1425
      End
      Begin VB.ComboBox cmbTexto 
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":323B
         Left            =   3480
         List            =   "frmVendas_Relatorios_Historico.frx":323D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   450
         Width           =   6225
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":323F
         Left            =   180
         List            =   "frmVendas_Relatorios_Historico.frx":3255
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   450
         Width           =   3285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limitar em                registros"
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
         Left            =   180
         TabIndex        =   34
         Top             =   900
         Width           =   2400
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
         Left            =   1402
         TabIndex        =   29
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5857
         TabIndex        =   28
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1325
      Left            =   13135
      TabIndex        =   30
      Top             =   975
      Width           =   2115
      Begin VB.ComboBox cmbPor 
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":32A6
         Left            =   630
         List            =   "frmVendas_Relatorios_Historico.frx":32AD
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Por."
         Top             =   210
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   11
         ToolTipText     =   "Data final."
         Top             =   930
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
         Format          =   103284737
         CurrentDate     =   39799
      End
      Begin VB.ComboBox Cmb_ano_ate1 
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":32B6
         Left            =   630
         List            =   "frmVendas_Relatorios_Historico.frx":32B8
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   10
         ToolTipText     =   "Data inicio."
         Top             =   570
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
         Format          =   103284737
         CurrentDate     =   39799
      End
      Begin VB.ComboBox Cmb_mes_de 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":32BA
         Left            =   630
         List            =   "frmVendas_Relatorios_Historico.frx":32E2
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Mês de."
         Top             =   570
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_de 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":3323
         Left            =   1260
         List            =   "frmVendas_Relatorios_Historico.frx":3325
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox Cmb_ano_de1 
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":3327
         Left            =   630
         List            =   "frmVendas_Relatorios_Historico.frx":3329
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_mes_ate 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":332B
         Left            =   630
         List            =   "frmVendas_Relatorios_Historico.frx":3353
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Mês até."
         Top             =   930
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_ate 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmVendas_Relatorios_Historico.frx":3394
         Left            =   1260
         List            =   "frmVendas_Relatorios_Historico.frx":3396
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   33
         Top             =   278
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   630
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   990
         Width           =   360
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   45
      Top             =   8910
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
      Left            =   11940
      TabIndex        =   46
      Top             =   8940
      Width           =   3315
   End
End
Attribute VB_Name = "frmVendas_Relatorios_Historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=FOVIOhJT6Dw&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=31&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ano_ate1_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_ano_de1 <> "" And Cmb_ano_ate1 <> "" Then
    qt = Cmb_ano_ate1
    Qtd = Cmb_ano_de1
    If qt < Qtd Then
        USMsgBox ("O ano final não pode ser menor que o ano inicial."), vbExclamation, "CAPRIND v5.0"
        Cmb_ano_ate1 = Cmb_ano_de1
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ano_de_Click()
On Error GoTo tratar_erro

Cmb_ano_ate = Cmb_ano_de
Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ano_de1_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_ano_de1 <> "" And Cmb_ano_ate1 <> "" Then
    qt = Cmb_ano_de1
    Qtd = Cmb_ano_ate1
    If qt > Qtd Then
        USMsgBox ("O ano inicial não pode ser maior que o ano final."), vbExclamation, "CAPRIND v5.0"
        Cmb_ano_de1 = Cmb_ano_ate1
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_ate_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_mes_de <> "" And Cmb_mes_ate <> "" Then
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    If Qtd < qt Then
        USMsgBox ("O mês final não pode ser menor que o mês inicial."), vbExclamation, "CAPRIND v5.0"
        Cmb_mes_ate = Cmb_mes_de
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_de_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_mes_de <> "" And Cmb_mes_ate <> "" Then
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    If qt > Qtd Then
        USMsgBox ("O mês inicial não pode ser maior que o mês final."), vbExclamation, "CAPRIND v5.0"
        Cmb_mes_de = Cmb_mes_ate
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
ProcMostrarEsconderCombosData

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then
        If Lista.ListItems.Count = 0 Then Exit Sub
    Else
        If Lista1.ListItems.Count = 0 Then Exit Sub
    End If
Else
    If Lista1.ListItems.Count = 0 Then Exit Sub
End If
Vendas_Relatorio_Historico = True
Vendas_Relatorio_IndiceAtraso = False
Vendas_Relatorio_Comissao = False
Compras_Relatorio_IndiceAtraso = False
PCP_relatorios_indice_atraso = False
Manutencao_Relatorio_Historico = False
FrmMenu_impressao_padrao.Show 1

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
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Familiatext = ""
Contador1 = 1
Posicao = 0
Lista.ListItems.Clear
Lista1.ListItems.Clear
If TBLISTA.EOF = False Then
    If optDetalhado.Value = True Then Posicao = TBLISTA.RecordCount
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Vendas_relatorios_historico_detalhado where Codigo = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Add , , TBAbrir!CODIGO
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Ncotacao), "", TBAbrir!Ncotacao)
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Revisao), "", TBAbrir!Revisao)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!descricao_tecnica), "", TBAbrir!descricao_tecnica)
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!quantidade), "", Format(TBAbrir!quantidade, "###,##0.0000"))
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!preco_unitario_desconto), "", (Format(TBAbrir!preco_unitario_desconto, "###,##0.0000000000")))
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBAbrir!dbl_valoripi), "", (Format(TBAbrir!dbl_valoripi, "###,##0.00")))
                    .Item(.Count).SubItems(11) = Format(IIf(IsNull(TBAbrir!preco_lote), 0, TBAbrir!preco_lote) + IIf(IsNull(TBAbrir!dbl_valoripi), 0, TBAbrir!dbl_valoripi), "###,##0.00")
                    .Item(.Count).SubItems(12) = IIf(IsNull(TBAbrir!Datavendas), "", Format(TBAbrir!Datavendas, "dd/mm/yy"))
                    .Item(.Count).SubItems(13) = IIf(IsNull(TBAbrir!Liberacao), "", TBAbrir!Liberacao)
                End If
                TBAbrir.Close
            End With
        Else
            If TBLISTA!maquina <> "" Then
                With Lista1.ListItems
                    Contador1 = 1
                    If cmbPor = "Dia" Then
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> Format(TBLISTA!Execucaoprev, "dd/mm/yy")
                            Contador1 = Contador1 + 1
                        Loop
                    ElseIf cmbPor = "Mês" Then
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                                Contador1 = Contador1 + 1
                            Loop
                        Else
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                                Contador1 = Contador1 + 1
                            Loop
                    End If
                    
                    If TBLISTA!maquina <> Familiatext Then
                        .Add , , TBLISTA!maquina
                        Posicao = Posicao + 1
                    End If
                    .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!qtdeOK), "", Format(TBLISTA!qtdeOK, "###,##0.00"))
                    
                    'Carrega valor ou quantidade total
                    Contador1 = 1
                    If Opt_valor.Value = True Then
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Valor total"
                            Contador1 = Contador1 + 1
                        Loop
                    Else
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Qtde. total"
                            Contador1 = Contador1 + 1
                        Loop
                    End If
                    .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!OS), "", Format(TBLISTA!OS, "###,##0.00"))
                    
                    If cmbfiltrarpor = "Código interno" And Opt_quantidade.Value = True Then
                        'Carrega dados do material
                        Contador1 = 1
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Cód. material"
                            Contador1 = Contador1 + 1
                        Loop
                        .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!Totalhsprev), "", TBLISTA!Totalhsprev)
                        
                        Contador1 = 1
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Descrição material"
                            Contador1 = Contador1 + 1
                        Loop
                        .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil)
                        
                        Contador1 = 1
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Qtde. material"
                            Contador1 = Contador1 + 1
                        Loop
                        .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!Qtdetotalprod), "", Format(TBLISTA!Qtdetotalprod, "###,##0.0000"))
                        
                        Contador1 = 1
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Un."
                            Contador1 = Contador1 + 1
                        Loop
                        .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento)
                                                
                        Contador1 = 1
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Qtde. total mat."
                            Contador1 = Contador1 + 1
                        Loop
                        .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!impostos), "", Format(TBLISTA!impostos, "###,##0.0000"))
                    End If
                End With
            End If
        End If
        Familiatext = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    If optDetalhado.Value = True Then Else
End If
TBLISTA.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtTotal_produtos = Format(TBLISTA!qtdeNC, "###,##0.00")
    Txt_total_servicos = Format(TBLISTA!Totalutilizada, "###,##0.00")
    txtTotal_ipi = Format(TBLISTA!Totalprevista, "###,##0.00")
    txtTotal_geral = Format(TBLISTA!qtdeNC + TBLISTA!Totalutilizada + TBLISTA!Totalprevista, "###,##0.00")
    Txt_qtde_total_vendido = Format(TBLISTA!QtdePrevista, "###,##0.0000")
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparListaeCampos()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
txtTotal_produtos = ""
Txt_total_servicos = ""
txtTotal_ipi = ""
txtTotal_geral = ""
Txt_qtde_total_vendido = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True

Formulario = "Vendas/Relatórios/Histórico"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcCarregaComboAno Cmb_ano_ate, "2005", 1
ProcCarregaComboAno Cmb_ano_ate1, "2005", 1
ProcCarregaComboAno Cmb_ano_de, "2005", 1
ProcCarregaComboAno Cmb_ano_de1, "2005", 1
cmbfiltrarpor.Text = "Cliente"
cmbPor = "Dia"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Relatórios/Histórico"
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

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTexto()
On Error GoTo tratar_erro

cmbTexto.Clear
Ordenar = ""
Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_individual.Value = True Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Código interno x Cliente" Or cmbfiltrarpor = "Código de referência x Cliente" Then
    If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Código interno x Cliente" Or cmbfiltrarpor = "Código de referência x Cliente" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Cliente from Vendas_relatorios_historico_detalhado where cliente <> 'Null' Group by cliente", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                cmbTexto.AddItem TBAbrir!Cliente
                TBAbrir.MoveNext
            Loop
        End If
    Else
        Select Case cmbfiltrarpor
            Case "Código de referência":
            Case "Código de referência": Ordenar = "n_referencia"
            Case "Família": Ordenar = "familia"
            Case "Código interno": Ordenar = "desenho"
            Case "Descrição": Ordenar = "descricao_tecnica"
        End Select
        If Ordenar <> "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select " & Ordenar & " as NomeCampo1 from Vendas_relatorios_historico_detalhado where " & Ordenar & " <> 'Null' Group by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TBAbrir!NomeCampo1 <> "" Then cmbTexto.AddItem TBAbrir!NomeCampo1
                TBAbrir.MoveNext
            Loop
        End If
        Else
        Select Case cmbfiltrarpor
            Case "Vendedor": Ordenar = "Vendedor"
        End Select
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select " & Ordenar & " from Vendas_Vendedores where " & Ordenar & " <> 'Null' Group by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TBAbrir!vendedor <> "" Then cmbTexto.AddItem TBAbrir!vendedor
                TBAbrir.MoveNext
            Loop
        End If
        
    End If
    TBAbrir.Close
    End If
    
End If
If Opt_comparativo = True And optResumido.Value = True Then
    If cmbfiltrarpor = "Código interno x Cliente" Or cmbfiltrarpor = "Código de referência x Cliente" Then cmbTexto.Enabled = True Else cmbTexto.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Opt_comparativo.Value = True And cmbfiltrarpor = "Código interno x Cliente" And cmbTexto = "" Or Opt_comparativo.Value = True And cmbfiltrarpor = "Código de referência x Cliente" And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If
If optResumido.Value = True Then
    ProcVerificaPeriodoMax
    If Permitido = False Then
        USMsgBox ("Só é permitido colocar um período de " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
        msk_fltInicio.SetFocus
        Exit Sub
    End If
    If cmbPor = "Mês" Then
        If Cmb_mes_de = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_de.SetFocus
            Exit Sub
        End If
        If Cmb_mes_ate = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_ate.SetFocus
            Exit Sub
        End If
        If Cmb_ano_de = "" Then
            NomeCampo = "o ano"
            ProcVerificaAcao
            Cmb_ano_de.SetFocus
            Exit Sub
        End If
    ElseIf cmbPor = "Ano" Then
            If Cmb_ano_de1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_de1.SetFocus
                Exit Sub
            End If
            If Cmb_ano_ate1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_ate1.SetFocus
                Exit Sub
            End If
    End If
End If
If Txt_limite <> "" Then
    If Txt_limite < 10 Then
        USMsgBox ("O campo (Limitar em) não pode ser menor que 10."), vbExclamation, "CAPRIND v5.0"
        Txt_limite.SetFocus
        Exit Sub
    End If
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

Inicio = Time
ProcLimpaCamposTotais
ProcAbrirTabelas
If optResumido.Value = True Then
    ProcCriaColunas
    
    'Soma e grava o total geral
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select maquina, Sum(QtdeOK) as QtdeSaida from Producao_relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' Group by Maquina", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            quantidade = IIf(IsNull(TBLISTA!QtdeSaida), 0, TBLISTA!QtdeSaida) 'Qtde. vendida
            NovoValor = Replace(quantidade, ",", ".")
            Conexao.Execute "Update Producao_relatorios Set OS = " & NovoValor & " where Maquina = '" & TBLISTA!maquina & "'"
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
End If
If Txt_limite <> "" Then ProcVerificaLimiteRegistros
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True And optDetalhado.Value = True Then
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data, Maquina", Conexao, adOpenKeyset, adLockReadOnly
Else
    'TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Maquina <> 'Null' order by OS desc, Maquina, Ordem", Conexao, adOpenKeyset, adLockReadOnly
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Maquina <> 'Null' order by Maquina, Ordem", Conexao, adOpenKeyset, adLockReadOnly
End If
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

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

FamiliaAntiga = ""
Select Case cmbfiltrarpor
    Case "Código interno":
        Grupo = "desenho, Descricao_tecnica"
        If cmbTexto <> "" Then FamiliaAntiga = "desenho = '" & cmbTexto & "' and "
    Case "Código de referência":
        Grupo = "n_referencia, Descricao_tecnica"
        If cmbTexto <> "" Then FamiliaAntiga = "n_referencia = '" & cmbTexto & "' and "
    Case "Descrição":
        Grupo = "Descricao_tecnica"
        If cmbTexto <> "" Then FamiliaAntiga = "Descricao_tecnica = '" & cmbTexto & "' and "
    Case "Família":
        Grupo = "familia"
        If cmbTexto <> "" Then FamiliaAntiga = "familia = '" & cmbTexto & "' and "
    Case "Cliente":
        Grupo = "Cliente"
        If cmbTexto <> "" Then FamiliaAntiga = "Cliente = '" & cmbTexto & "' and "
    Case "Código interno x Cliente":
        Grupo = "desenho, Descricao_tecnica"
        If cmbTexto <> "" Then FamiliaAntiga = "desenho = '" & cmbTexto & "' and "
    Case "Código de referência x Cliente":
        Grupo = "n_referencia, Descricao_tecnica"
        If cmbTexto <> "" Then FamiliaAntiga = "n_referencia = '" & cmbTexto & "' and "
     Case "Vendedor":
        Grupo = "Vend_Int"
        If cmbTexto <> "" Then FamiliaAntiga = "Vend_Int = '" & cmbTexto & "' and "
     
End Select
If optDetalhado.Value = True Then
    StrSql = "Select * from Vendas_relatorios_historico_detalhado where " & FamiliaAntiga & " (datavendas) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    'Debug.print StrSql
    
Else
    If Opt_quantidade.Value = True Then TextoFiltro = "Quantidade" Else TextoFiltro = "Preco_lote"
    
    Par1 = ""
    Permitido = False
    Select Case cmbPor
        Case "Dia":
            Dataini = msk_fltInicio
            DataFim = msk_fltFim
            Do While Dataini <= DataFim
                If Permitido = False Then Par1 = "[" & Dataini & "]" Else Par1 = Par1 & " , [" & Dataini & "]"
                Permitido = True
                Dataini = Dataini + 1
            Loop
            Pesquisa = "(datavendas) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for datavendas In (" & Par1 & "))"
            Pesquisa2 = "datavendas"
        Case "Mês":
            qt = FunVerificaMes(Cmb_mes_de)
            Qtd = FunVerificaMes(Cmb_mes_ate)
            MesX = qt
            MesX1 = Qtd
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Month(datavendas) >= '" & MesX & "' and Year(datavendas) = '" & Cmb_ano_de & "' and Month(datavendas) <= '" & MesX1 & "' and Year(datavendas) = '" & Cmb_ano_ate & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Mes In (" & Par1 & "))"
            Pesquisa2 = "Mes"
        Case "Ano":
            qt = Cmb_ano_de1
            Qtd = Cmb_ano_ate1
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Year(datavendas) >= '" & Cmb_ano_de1 & "' and Year(datavendas) <= '" & Cmb_ano_ate1 & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Ano In (" & Par1 & "))"
            Pesquisa2 = "Ano"
    End Select
    If Opt_individual.Value = True Then
        If cmbfiltrarpor = "Vendedor" Then
        FamiliaAntiga = "Vend_int"
        End If
        
        StrSql = "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Vendas_relatorios_historico_detalhado Where " & FamiliaAntiga & " = '" & cmbTexto & "' and " & Pesquisa & ") p " & Pesquisa1 & " pvt"

    Else
        If cmbfiltrarpor = "Código interno x Cliente" Or cmbfiltrarpor = "Código de referência x Cliente" Then
            StrSql = "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Vendas_relatorios_historico_detalhado Where Cliente = '" & cmbTexto & "' and " & Pesquisa & ") p " & Pesquisa1 & " pvt"
        Else
            StrSql = "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Vendas_relatorios_historico_detalhado Where " & Pesquisa & ") p " & Pesquisa1 & " as pvt"
        End If
    End If
End If
'Debug.print StrSql

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
ProcFiltrar1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

Produto = ""
Familiatext = ""
quantidade = 0
QTLOTE = 0
Valor_Produto = 0
Valor_Cofins_Serv = 0
ValorIPI = 0
If TBCarteira.EOF = False Then
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
        If optDetalhado.Value = True Then
            Set TBProdutividade = CreateObject("adodb.recordset")
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            ProcCriarResumido
        End If
        TBCarteira.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

Permitido = True
TBProdutividade.AddNew
TBProdutividade!Ordem = TBCarteira!CODIGO
TBProdutividade!Data = TBCarteira!Datavendas
TBProdutividade!QtdePrev = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade) 'Qtde. vendida
TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote) + IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) 'Valor total
If TBCarteira!Tipo = "P" Then
    TBProdutividade!qtdeNC = IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote) 'Valor total produtos
Else
    TBProdutividade!Qtdetotalprod = IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote) 'Valor total serviços
End If
TBProdutividade!Eficiencia = IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) 'Valor total IPI
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = Familiatext

quantidade = quantidade + TBProdutividade!QtdePrev 'Qtde. vendida
QTLOTE = QTLOTE + TBProdutividade!qtdeOK 'Valor total
If TBCarteira!Tipo = "P" Then
    Valor_Produto = Valor_Produto + TBProdutividade!qtdeNC 'Valor total produtos
Else
    Valor_Cofins_Serv = Valor_Cofins_Serv + TBProdutividade!Qtdetotalprod 'Valor total serviços
End If
ValorIPI = ValorIPI + TBProdutividade!Eficiencia 'Valor total IPI

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarResumido()
On Error GoTo tratar_erro

Permitido = True
Select Case cmbPor
    Case "Dia":
        qt = 0
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            qt = qt + 1
            ProcEnviaDadosResumido
            Dataini = Dataini + 1
        Loop
    Case "Mês":
        qt = MesX
        Qtd = MesX1
        Do While qt <= Qtd
            ProcEnviaDadosResumido
            qt = qt + 1
        Loop
    Case "Ano":
        qt = Cmb_ano_de1
        Qtd = Cmb_ano_ate1
        Do While qt <= Qtd
            ProcEnviaDadosResumido
            qt = qt + 1
        Loop
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Código interno": Familiatext = TBCarteira!Desenho
    Case "Código de referência": Familiatext = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    Case "Descrição": Familiatext = TBCarteira!descricao_tecnica
    Case "Família": Familiatext = TBCarteira!Familia
    Case "Cliente": Familiatext = TBCarteira!Cliente
    Case "Código interno x Cliente": Familiatext = TBCarteira!Desenho
    Case "Código de referência x Cliente": Familiatext = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    Case "Vendedor": Familiatext = IIf(IsNull(TBCarteira!vend_int), "S/CADASTRO", TBCarteira!vend_int)
    
End Select
Select Case cmbPor
    Case "Dia": DataTexto = Dataini
    Case "Mês": DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano": DataTexto = "01" & "/01/" & qt
End Select
Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
TBProdutividade.AddNew
TBProdutividade!Data = Format(DataTexto, "dd/mm/yy")
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!qtdeOK = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00"))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![1]), 0, Format(TBCarteira![1], "###,##0.00"))
            Case 2: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![2]), 0, Format(TBCarteira![2], "###,##0.00"))
            Case 3: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![3]), 0, Format(TBCarteira![3], "###,##0.00"))
            Case 4: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![4]), 0, Format(TBCarteira![4], "###,##0.00"))
            Case 5: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![5]), 0, Format(TBCarteira![5], "###,##0.00"))
            Case 6: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![6]), 0, Format(TBCarteira![6], "###,##0.00"))
            Case 7: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![7]), 0, Format(TBCarteira![7], "###,##0.00"))
            Case 8: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![8]), 0, Format(TBCarteira![8], "###,##0.00"))
            Case 9: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![9]), 0, Format(TBCarteira![9], "###,##0.00"))
            Case 10: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![10]), 0, Format(TBCarteira![10], "###,##0.00"))
            Case 11: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![11]), 0, Format(TBCarteira![11], "###,##0.00"))
            Case 12: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![12]), 0, Format(TBCarteira![12], "###,##0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!qtdeOK = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00"))
        TBProdutividade!Execucaoprev = qt
End Select

TBProdutividade!Ordem = qt

If cmbfiltrarpor = "Código interno" Or cmbfiltrarpor = "Código de referência" Or cmbfiltrarpor = "Código interno x Cliente" Or cmbfiltrarpor = "Código de referência x Cliente" Then
    TBProdutividade!maquina = Left(Familiatext & " " & TBCarteira!descricao_tecnica, 25)
    If cmbfiltrarpor = "Código interno" And Opt_quantidade.Value = True Then ProcGravaDadosEstrutura
Else
    TBProdutividade!maquina = Left(Familiatext, 25)
End If

TBProdutividade.Update
TBProdutividade.Close

If cmbfiltrarpor = "Código interno" And Opt_quantidade.Value = True Then
    'Quantidade total de material
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(Terceiros) as QtdeSaida from Producao_Relatorios where Maquina = '" & Left(Familiatext & " " & TBCarteira!descricao_tecnica, 25) & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Qtd_Prog = IIf(IsNull(TBAbrir!QtdeSaida), 0, TBAbrir!QtdeSaida) 'Qtde. vendida
    End If
    TBAbrir.Close
    NovoValor = Replace(Qtd_Prog, ",", ".")
    Conexao.Execute "Update Producao_Relatorios Set Impostos = " & NovoValor & " where Maquina = '" & Left(Familiatext & " " & TBCarteira!descricao_tecnica, 25) & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"
End If

Select Case cmbfiltrarpor
    Case "Código interno": Produto = TBCarteira!Desenho
    Case "Código de referência": Produto = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    Case "Descrição": Produto = TBCarteira!descricao_tecnica
    Case "Família": Produto = TBCarteira!Familia
    Case "Cliente": Produto = TBCarteira!Cliente
    Case "Código interno x Cliente": Produto = TBCarteira!Desenho
    Case "Código de referência x Cliente": Produto = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    Case "Vendedor": Produto = IIf(IsNull(TBCarteira!vend_int), 0, TBCarteira!vend_int)
End Select
'Debug.print cmbfiltrarpor.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravaDadosEstrutura()
On Error GoTo tratar_erro

If Familiatext <> Produto Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select codproduto, Desenho from projproduto where desenho = '" & Familiatext & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Codproduto = TBAbrir!Codproduto
        Desenho = TBAbrir!Desenho
    End If
    TBAbrir.Close
    ProcNivel1
End If

TBProdutividade!Totalhsprev = Desenho 'Cód. do material
TBProdutividade!Totalhsutil = IIf(DesenhoProduto <> "", DesenhoProduto, 0) 'Descrição do material
TBProdutividade!Qtdetotalprod = QuantSolicitado 'Quantidade de material
TBProdutividade!DescEvento = IIf(Par2 <> "", Par2, 0) 'Unidade do material
TBProdutividade!Terceiros = QuantSolicitado * TBProdutividade!qtdeOK 'Qtde. total de material

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel1()
On Error GoTo tratar_erro

Set TBNivel1 = CreateObject("adodb.recordset")
TBNivel1.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel1.EOF = False Then
    Do While TBNivel1.EOF = False
        Desenho = TBNivel1!Desenho
        DesenhoProduto = TBNivel1!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel1!Unidade = "KG" Then QuantSolicitado = TBNivel1!PesoTotal
            If TBNivel1!Unidade = "MT" Then QuantSolicitado = (TBNivel1!Dimensoes * TBNivel1!quantidade) / 1000
            If TBNivel1!Unidade = "MM" Then QuantSolicitado = TBNivel1!Dimensoes * TBNivel1!quantidade
            If TBNivel1!Unidade <> "KG" And TBNivel1!Unidade <> "MT" And TBNivel1!Unidade <> "MM" Then QuantSolicitado = TBNivel1!quantidade
            Par2 = TBNivel1!Unidade
            ProcNivel2
        End If
        TBNivel1.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel2()
On Error GoTo tratar_erro

Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
        Desenho = TBNivel2!Desenho
        DesenhoProduto = TBNivel2!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel2!Unidade = "KG" Then QuantSolicitado = TBNivel2!PesoTotal
            If TBNivel2!Unidade = "MT" Then QuantSolicitado = (TBNivel2!Dimensoes * TBNivel2!quantidade) / 1000
            If TBNivel2!Unidade = "MM" Then QuantSolicitado = TBNivel2!Dimensoes * TBNivel2!quantidade
            If TBNivel2!Unidade <> "KG" And TBNivel2!Unidade <> "MT" And TBNivel2!Unidade <> "MM" Then QuantSolicitado = TBNivel2!quantidade
            Par2 = TBNivel2!Unidade
            ProcNivel3
        End If
        TBNivel2.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel3()
On Error GoTo tratar_erro

Set TBNivel3 = CreateObject("adodb.recordset")
TBNivel3.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel3.EOF = False Then
    Do While TBNivel3.EOF = False
        Desenho = TBNivel3!Desenho
        DesenhoProduto = TBNivel3!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel3!Unidade = "KG" Then QuantSolicitado = TBNivel3!PesoTotal
            If TBNivel3!Unidade = "MT" Then QuantSolicitado = (TBNivel3!Dimensoes * TBNivel3!quantidade) / 1000
            If TBNivel3!Unidade = "MM" Then QuantSolicitado = TBNivel3!Dimensoes * TBNivel3!quantidade
            If TBNivel3!Unidade <> "KG" And TBNivel3!Unidade <> "MT" And TBNivel3!Unidade <> "MM" Then QuantSolicitado = TBNivel3!quantidade
            Par2 = TBNivel3!Unidade
            ProcNivel4
        End If
        TBNivel3.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel4()
On Error GoTo tratar_erro

Set TBNivel4 = CreateObject("adodb.recordset")
TBNivel4.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel4.EOF = False Then
    Do While TBNivel4.EOF = False
        Desenho = TBNivel4!Desenho
        DesenhoProduto = TBNivel4!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel4!Unidade = "KG" Then QuantSolicitado = TBNivel4!PesoTotal
            If TBNivel4!Unidade = "MT" Then QuantSolicitado = (TBNivel4!Dimensoes * TBNivel4!quantidade) / 1000
            If TBNivel4!Unidade = "MM" Then QuantSolicitado = TBNivel4!Dimensoes * TBNivel4!quantidade
            If TBNivel4!Unidade <> "KG" And TBNivel4!Unidade <> "MT" And TBNivel4!Unidade <> "MM" Then QuantSolicitado = TBNivel4!quantidade
            Par2 = TBNivel4!Unidade
            ProcNivel5
        End If
        TBNivel4.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel5()
On Error GoTo tratar_erro

Set TBNivel5 = CreateObject("adodb.recordset")
TBNivel5.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel5.EOF = False Then
    Do While TBNivel5.EOF = False
        Desenho = TBNivel5!Desenho
        DesenhoProduto = TBNivel5!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel5!Unidade = "KG" Then QuantSolicitado = TBNivel5!PesoTotal
            If TBNivel5!Unidade = "MT" Then QuantSolicitado = (TBNivel5!Dimensoes * TBNivel5!quantidade) / 1000
            If TBNivel5!Unidade = "MM" Then QuantSolicitado = TBNivel5!Dimensoes * TBNivel5!quantidade
            If TBNivel5!Unidade <> "KG" And TBNivel5!Unidade <> "MT" And TBNivel5!Unidade <> "MM" Then QuantSolicitado = TBNivel5!quantidade
            Par2 = TBNivel5!Unidade
            ProcNivel6
        End If
        TBNivel5.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel6()
On Error GoTo tratar_erro

Set TBNivel6 = CreateObject("adodb.recordset")
TBNivel6.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel6.EOF = False Then
    Do While TBNivel6.EOF = False
        Desenho = TBNivel6!Desenho
        DesenhoProduto = TBNivel6!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel6!Unidade = "KG" Then QuantSolicitado = TBNivel6!PesoTotal
            If TBNivel6!Unidade = "MT" Then QuantSolicitado = (TBNivel6!Dimensoes * TBNivel6!quantidade) / 1000
            If TBNivel6!Unidade = "MM" Then QuantSolicitado = TBNivel6!Dimensoes * TBNivel6!quantidade
            If TBNivel6!Unidade <> "KG" And TBNivel6!Unidade <> "MT" And TBNivel6!Unidade <> "MM" Then QuantSolicitado = TBNivel6!quantidade
            Par2 = TBNivel6!Unidade
            ProcNivel7
        End If
        TBNivel6.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel7()
On Error GoTo tratar_erro

Set TBNivel7 = CreateObject("adodb.recordset")
TBNivel7.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel7.EOF = False Then
    Do While TBNivel7.EOF = False
        Desenho = TBNivel7!Desenho
        DesenhoProduto = TBNivel7!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel7!Unidade = "KG" Then QuantSolicitado = TBNivel7!PesoTotal
            If TBNivel7!Unidade = "MT" Then QuantSolicitado = (TBNivel7!Dimensoes * TBNivel7!quantidade) / 1000
            If TBNivel7!Unidade = "MM" Then QuantSolicitado = TBNivel7!Dimensoes * TBNivel7!quantidade
            If TBNivel7!Unidade <> "KG" And TBNivel7!Unidade <> "MT" And TBNivel7!Unidade <> "MM" Then QuantSolicitado = TBNivel7!quantidade
            Par2 = TBNivel7!Unidade
            ProcNivel8
        End If
        TBNivel7.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel8()
On Error GoTo tratar_erro

Set TBNivel8 = CreateObject("adodb.recordset")
TBNivel8.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel8.EOF = False Then
    Do While TBNivel8.EOF = False
        Desenho = TBNivel8!Desenho
        DesenhoProduto = TBNivel8!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel8!Unidade = "KG" Then QuantSolicitado = TBNivel8!PesoTotal
            If TBNivel8!Unidade = "MT" Then QuantSolicitado = (TBNivel8!Dimensoes * TBNivel8!quantidade) / 1000
            If TBNivel8!Unidade = "MM" Then QuantSolicitado = TBNivel8!Dimensoes * TBNivel8!quantidade
            If TBNivel8!Unidade <> "KG" And TBNivel8!Unidade <> "MT" And TBNivel8!Unidade <> "MM" Then QuantSolicitado = TBNivel8!quantidade
            Par2 = TBNivel8!Unidade
            ProcNivel9
        End If
        TBNivel8.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel9()
On Error GoTo tratar_erro

Set TBNivel9 = CreateObject("adodb.recordset")
TBNivel9.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel9.EOF = False Then
    Do While TBNivel9.EOF = False
        Desenho = TBNivel9!Desenho
        DesenhoProduto = TBNivel9!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel9!Unidade = "KG" Then QuantSolicitado = TBNivel9!PesoTotal
            If TBNivel9!Unidade = "MT" Then QuantSolicitado = (TBNivel9!Dimensoes * TBNivel9!quantidade) / 1000
            If TBNivel9!Unidade = "MM" Then QuantSolicitado = TBNivel9!Dimensoes * TBNivel9!quantidade
            If TBNivel9!Unidade <> "KG" And TBNivel9!Unidade <> "MT" And TBNivel9!Unidade <> "MM" Then QuantSolicitado = TBNivel9!quantidade
            Par2 = TBNivel9!Unidade
            ProcNivel10
        End If
        TBNivel9.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel10()
On Error GoTo tratar_erro

Set TBNivel10 = CreateObject("adodb.recordset")
TBNivel10.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel10.EOF = False Then
    Do While TBNivel10.EOF = False
        Desenho = TBNivel10!Desenho
        DesenhoProduto = TBNivel10!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel10!Unidade = "KG" Then QuantSolicitado = TBNivel10!PesoTotal
            If TBNivel10!Unidade = "MT" Then QuantSolicitado = (TBNivel10!Dimensoes * TBNivel10!quantidade) / 1000
            If TBNivel10!Unidade = "MM" Then QuantSolicitado = TBNivel10!Dimensoes * TBNivel10!quantidade
            If TBNivel10!Unidade <> "KG" And TBNivel10!Unidade <> "MT" And TBNivel10!Unidade <> "MM" Then QuantSolicitado = TBNivel10!quantidade
            Par2 = TBNivel10!Unidade
            ProcNivel11
        End If
        TBNivel10.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel11()
On Error GoTo tratar_erro

Set TBNivel11 = CreateObject("adodb.recordset")
TBNivel11.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel11.EOF = False Then
    Do While TBNivel11.EOF = False
        Desenho = TBNivel11!Desenho
        DesenhoProduto = TBNivel11!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel11!Unidade = "KG" Then QuantSolicitado = TBNivel11!PesoTotal
            If TBNivel11!Unidade = "MT" Then QuantSolicitado = (TBNivel11!Dimensoes * TBNivel11!quantidade) / 1000
            If TBNivel11!Unidade = "MM" Then QuantSolicitado = TBNivel11!Dimensoes * TBNivel11!quantidade
            If TBNivel11!Unidade <> "KG" And TBNivel11!Unidade <> "MT" And TBNivel11!Unidade <> "MM" Then QuantSolicitado = TBNivel11!quantidade
            Par2 = TBNivel11!Unidade
            ProcNivel12
        End If
        TBNivel11.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel12()
On Error GoTo tratar_erro

Set TBNivel12 = CreateObject("adodb.recordset")
TBNivel12.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel12.EOF = False Then
    Do While TBNivel12.EOF = False
        Desenho = TBNivel12!Desenho
        DesenhoProduto = TBNivel12!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel12!Unidade = "KG" Then QuantSolicitado = TBNivel12!PesoTotal
            If TBNivel12!Unidade = "MT" Then QuantSolicitado = (TBNivel12!Dimensoes * TBNivel12!quantidade) / 1000
            If TBNivel12!Unidade = "MM" Then QuantSolicitado = TBNivel12!Dimensoes * TBNivel12!quantidade
            If TBNivel12!Unidade <> "KG" And TBNivel12!Unidade <> "MT" And TBNivel12!Unidade <> "MM" Then QuantSolicitado = TBNivel12!quantidade
            Par2 = TBNivel12!Unidade
            ProcNivel13
        End If
        TBNivel12.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel13()
On Error GoTo tratar_erro

Set TBNivel13 = CreateObject("adodb.recordset")
TBNivel13.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel13.EOF = False Then
    Do While TBNivel13.EOF = False
        Desenho = TBNivel13!Desenho
        DesenhoProduto = TBNivel13!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel13!Unidade = "KG" Then QuantSolicitado = TBNivel13!PesoTotal
            If TBNivel13!Unidade = "MT" Then QuantSolicitado = (TBNivel13!Dimensoes * TBNivel13!quantidade) / 1000
            If TBNivel13!Unidade = "MM" Then QuantSolicitado = TBNivel13!Dimensoes * TBNivel13!quantidade
            If TBNivel13!Unidade <> "KG" And TBNivel13!Unidade <> "MT" And TBNivel13!Unidade <> "MM" Then QuantSolicitado = TBNivel13!quantidade
            Par2 = TBNivel13!Unidade
            ProcNivel14
        End If
        TBNivel13.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel14()
On Error GoTo tratar_erro

Set TBNivel14 = CreateObject("adodb.recordset")
TBNivel14.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel14.EOF = False Then
    Do While TBNivel14.EOF = False
        Desenho = TBNivel14!Desenho
        DesenhoProduto = TBNivel14!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel14!Unidade = "KG" Then QuantSolicitado = TBNivel14!PesoTotal
            If TBNivel14!Unidade = "MT" Then QuantSolicitado = (TBNivel14!Dimensoes * TBNivel14!quantidade) / 1000
            If TBNivel14!Unidade = "MM" Then QuantSolicitado = TBNivel14!Dimensoes * TBNivel14!quantidade
            If TBNivel14!Unidade <> "KG" And TBNivel14!Unidade <> "MT" And TBNivel14!Unidade <> "MM" Then QuantSolicitado = TBNivel14!quantidade
            Par2 = TBNivel14!Unidade
            ProcNivel15
        End If
        TBNivel14.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNivel15()
On Error GoTo tratar_erro

Set TBNivel15 = CreateObject("adodb.recordset")
TBNivel15.Open "Select * from projconjunto where codproduto = " & Codproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel15.EOF = False Then
    Do While TBNivel15.EOF = False
        Desenho = TBNivel15!Desenho
        DesenhoProduto = TBNivel15!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Codproduto = TBAbrir!Codproduto
            If TBNivel15!Unidade = "KG" Then QuantSolicitado = TBNivel15!PesoTotal
            If TBNivel15!Unidade = "MT" Then QuantSolicitado = (TBNivel15!Dimensoes * TBNivel15!quantidade) / 1000
            If TBNivel15!Unidade = "MM" Then QuantSolicitado = TBNivel15!Dimensoes * TBNivel15!quantidade
            If TBNivel15!Unidade <> "KG" And TBNivel15!Unidade <> "MT" And TBNivel15!Unidade <> "MM" Then QuantSolicitado = TBNivel15!quantidade
            Par2 = TBNivel15!Unidade
        End If
        TBNivel15.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaColunas()
On Error GoTo tratar_erro

Lista1.ColumnHeaders.Clear
Contador = 1
With Lista1.ColumnHeaders
    .Add
    If cmbfiltrarpor <> "Código interno x Cliente" And cmbfiltrarpor <> "Código de referência x Cliente" And cmbfiltrarpor <> "Vendedor" Then
        .Item(Contador).Text = cmbfiltrarpor.Text
    Else
        If cmbfiltrarpor = "Código interno x Cliente" Then
            .Item(Contador).Text = "Código interno"
        End If
        If cmbfiltrarpor = "Código de referência" Then
            .Item(Contador).Text = "Código de referência"
        End If
        If cmbfiltrarpor = "Vendedor" Then
            .Item(Contador).Text = "Vendedor"
        End If
    End If
If .Count = 1 Then
    .Item(Contador).Width = 2500
Else
    .Item(Contador).Width = 1500
End If
    If cmbPor.Text = "Dia" Then
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = Format(Dataini, "dd/mm/yy")
            .Item(Contador).Alignment = lvwColumnRight
            Dataini = Dataini + 1
        Loop
    End If
    If cmbPor.Text = "Mês" Then
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)
        Do While qt <= Qtd
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = qt & "/" & Cmb_ano_de
            .Item(Contador).Alignment = lvwColumnRight
            .Item(Contador).Width = 1000
            qt = qt + 1
        Loop
    End If
    If cmbPor.Text = "Ano" Then
        qt = Cmb_ano_de1
        Do While qt <= Cmb_ano_ate1
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = qt
            .Item(Contador).Alignment = lvwColumnRight
            qt = qt + 1
        Loop
    End If
    .Add
    Contador = Contador + 1
    If Opt_valor.Value = True Then
        .Item(Contador).Text = "Valor total"
        .Item(Contador).Alignment = lvwColumnRight
    Else
        .Item(Contador).Text = "Qtde. total"
        .Item(Contador).Alignment = lvwColumnRight
        If cmbfiltrarpor = "Código interno" Then
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = "Cód. material"
            .Item(Contador).Alignment = lvwColumnLeft
            .Item(Contador).Width = 1200
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = "Descrição material"
            .Item(Contador).Alignment = lvwColumnLeft
            .Item(Contador).Width = 2500
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = "Qtde. material"
            .Item(Contador).Alignment = lvwColumnRight
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = "Un."
            .Item(Contador).Alignment = lvwColumnCenter
            .Item(Contador).Width = 500
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = "Qtde. total mat."
            .Item(Contador).Alignment = lvwColumnRight
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaLimiteRegistros()
On Error GoTo tratar_erro

Contador1 = Txt_limite
Valor_total = 0
Cliente = ""
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by OS desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    If TBLISTA.RecordCount < 10 Then Exit Sub
    Do While Contador1 <> 0
        If Cliente <> TBLISTA!maquina Then
            Valor_total = TBLISTA!OS
            Contador1 = Contador1 - 1
        End If
        Cliente = TBLISTA!maquina
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close
NovoValor = Replace(Valor_total, ",", ".")
Conexao.Execute "DELETE from Producao_Relatorios where OS < " & NovoValor & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    If Opt_individual.Value = True Then
        'Produtos e IPI
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select Sum(preco_lote) as Valor, Sum(dbl_valoripi) as ValorIPI from Vendas_relatorios_historico_detalhado where " & FamiliaAntiga & " = '" & cmbTexto & "' and " & Pesquisa & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            Valor_Produto = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total produtos
            ValorIPI = IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total IPI
            QTLOTE = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) + IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total
        End If
            
        'Serviços
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select Sum(preco_lote) as Valor from Vendas_relatorios_historico_detalhado where " & FamiliaAntiga & " = '" & cmbTexto & "' and " & Pesquisa & " and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            Valor_Cofins_Serv = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total serviços
        End If
        
        'Quantidade
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select Sum(quantidade) as QtdeSaida from Vendas_relatorios_historico_detalhado where " & FamiliaAntiga & " = '" & cmbTexto & "' and " & Pesquisa, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            quantidade = IIf(IsNull(TBLISTA!QtdeSaida), 0, TBLISTA!QtdeSaida) 'Qtde. vendida
        End If
    Else
        If cmbfiltrarpor = "Código interno x Cliente" Or cmbfiltrarpor = "Código de referência x Cliente" Then
            'Produtos e IPI
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select Sum(preco_lote) as Valor, Sum(dbl_valoripi) as ValorIPI from Vendas_relatorios_historico_detalhado where Cliente = '" & cmbTexto & "' and " & Pesquisa & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Valor_Produto = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total produtos
                ValorIPI = IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total IPI
                QTLOTE = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) + IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total
            End If
                
            'Serviços
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select Sum(preco_lote) as Valor from Vendas_relatorios_historico_detalhado where Cliente = '" & cmbTexto & "' and " & Pesquisa & " and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Valor_Cofins_Serv = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total serviços
            End If
            
            'Quantidade
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select Sum(quantidade) as QtdeSaida from Vendas_relatorios_historico_detalhado where Cliente = '" & cmbTexto & "' and " & Pesquisa, Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                quantidade = IIf(IsNull(TBLISTA!QtdeSaida), 0, TBLISTA!QtdeSaida) 'Qtde. vendida
            End If
            TBLISTA.Close
        Else
            'Produtos e IPI
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select Sum(preco_lote) as Valor, Sum(dbl_valoripi) as ValorIPI from Vendas_relatorios_historico_detalhado where " & Pesquisa & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Valor_Produto = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total produtos
                ValorIPI = IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total IPI
                QTLOTE = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) + IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total
            End If
                
            'Serviços
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select Sum(preco_lote) as Valor from Vendas_relatorios_historico_detalhado where " & Pesquisa & " and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Valor_Cofins_Serv = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total serviços
            End If
            
            'Quantidade
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select Sum(quantidade) as QtdeSaida from Vendas_relatorios_historico_detalhado where " & Pesquisa, Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                quantidade = IIf(IsNull(TBLISTA!QtdeSaida), 0, TBLISTA!QtdeSaida) 'Qtde. vendida
            End If
        End If
    End If
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew

Select Case cmbPor
    Case "Dia":
        Tipo = "D"
        TBAbrir!Data1 = Format(msk_fltInicio.Value, "dd/mm/yy")
        TBAbrir!Data2 = Format(msk_fltFim.Value, "dd/mm/yy")
    Case "Mês":
        Tipo = "M"
        TBAbrir!Data1 = Cmb_mes_de & "/" & Cmb_ano_de
        TBAbrir!Data2 = Cmb_mes_ate & "/" & Cmb_ano_ate
    Case "Ano":
        Tipo = "A"
        TBAbrir!Data1 = Cmb_ano_de1
        TBAbrir!Data2 = Cmb_ano_ate1
End Select

TBAbrir!Data3 = Tipo
If cmbfiltrarpor <> "Código interno x Cliente" And cmbfiltrarpor <> "Código de referência x Cliente" Then
    If Opt_individual.Value = True Then TBAbrir!Texto = cmbfiltrarpor & " : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor
    TBAbrir!QtdeOrdem = "1"
Else
    If cmbfiltrarpor = "Código interno x Cliente" Then
        If Opt_individual.Value = True Then
            TBAbrir!Texto = "Código interno" & " : " & cmbTexto
        Else
            TBAbrir!Texto = "Código interno"
        End If
        TBAbrir!QtdeOrdem = "2"
    Else
        If Opt_individual.Value = True Then
            TBAbrir!Texto = "Código de referência" & " : " & cmbTexto
        Else
            TBAbrir!Texto = "Código de referência"
        End If
        TBAbrir!QtdeOrdem = "2"
    End If
    TBAbrir!Texto1 = cmbTexto
End If

TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
If Opt_quantidade.Value = True Then TBAbrir!Turno = True Else TBAbrir!Turno = False
TBAbrir!QtdePrevista = quantidade 'Qtde. vendida
TBAbrir!QtdeProduzida = QTLOTE 'Valor total
TBAbrir!qtdeNC = Valor_Produto 'Valor total produtos
TBAbrir!Totalutilizada = Format(Valor_Cofins_Serv, "###,##0.00") 'Valor serviços
TBAbrir!Totalprevista = Format(ValorIPI, "###,##0.00") 'Valor total IPI
If Opt_valor = True Then TBAbrir!TotalEficiencia = 1 Else TBAbrir!TotalEficiencia = 2
TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaPeriodoMax()
On Error GoTo tratar_erro

Permitido = True
If cmbPor = "Dia" Then
    Dataini = msk_fltInicio
    DataFim = msk_fltFim
    If DataFim - Dataini > 10 Then
        Permitido = False
        NomeCampo = "10 dias"
    End If
ElseIf Cmb_ano_ate1 - Cmb_ano_de1 > 5 Then
        Permitido = False
        NomeCampo = "5 anos"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
    Txt_limite.Locked = False
    Txt_limite.TabStop = True
End If
With cmbfiltrarpor
    .Clear
    .AddItem "Cliente"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    .AddItem "Código interno x Cliente"
    .AddItem "Código de referência x Cliente"
    .AddItem "Vendedor"
    .Text = "Cliente"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
    Txt_limite = ""
    Txt_limite.Locked = True
    Txt_limite.TabStop = False
End If
With cmbfiltrarpor
    .Clear
    .AddItem "Cliente"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    .AddItem "Vendedor"
     .Text = "Cliente"
End With
ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_quantidade_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_valor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    Lista.ListItems.Clear
    Lista.Visible = True
    Lista1.ListItems.Clear
    Lista1.Visible = False
    Opt_valor.Value = False
    Opt_valor.Enabled = False
    Opt_quantidade.Value = False
    Opt_quantidade.Enabled = False
    With cmbPor
        .Clear
        .AddItem "Dia"
        .Text = "Dia"
    End With
    ProcMostrarEsconderCombosData
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    Lista.ListItems.Clear
    Lista.Visible = False
    Lista1.ListItems.Clear
    Lista1.Visible = True
    Opt_valor.Value = True
    Opt_valor.Enabled = True
    Opt_quantidade.Enabled = True
    With cmbPor
        .Clear
        .AddItem "Dia"
        .AddItem "Mês"
        .AddItem "Ano"
        .Text = "Dia"
    End With
    ProcMostrarEsconderCombosData
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarEsconderCombosData()
On Error GoTo tratar_erro

If cmbPor = "Dia" Then
    msk_fltInicio.Visible = True
    msk_fltFim.Visible = True
    Cmb_mes_de.Visible = False
    Cmb_mes_ate.Visible = False
    Cmb_ano_de.Visible = False
    Cmb_ano_ate.Visible = False
    Cmb_ano_de1.Visible = False
    Cmb_ano_ate1.Visible = False
ElseIf cmbPor = "Mês" Then
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = True
        Cmb_mes_ate.Visible = True
        Cmb_ano_de.Visible = True
        Cmb_ano_ate.Visible = True
        Cmb_ano_de1.Visible = False
        Cmb_ano_ate1.Visible = False
    Else
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = False
        Cmb_mes_ate.Visible = False
        Cmb_ano_de.Visible = False
        Cmb_ano_ate.Visible = False
        Cmb_ano_de1.Visible = True
        Cmb_ano_ate1.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_limite_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Txt_limite <> "" Then
    VerifNumero = Txt_limite
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_limite = ""
        txt_ValorPago.SetFocus
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
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
