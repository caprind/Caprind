VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCompras_Relatorios_Historico 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Relatórios - Histórico"
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
   Icon            =   "frmCompras_Relatorios_Historico.frx":0000
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
      Height          =   825
      Left            =   60
      TabIndex        =   38
      Top             =   9180
      Width           =   15195
      Begin VB.TextBox txtICMS_ST 
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
         Left            =   9795
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do ICMS substituto."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox txtSubtotal 
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
         Left            =   6333
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Subtotal."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox txtDesconto_percentual 
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
         Left            =   5010
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Percentual do desconto."
         Top             =   375
         Width           =   1005
      End
      Begin VB.TextBox txtDesconto 
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
         Left            =   3642
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do(s) serviço(s)."
         Top             =   375
         Width           =   1335
      End
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
         Left            =   11526
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Valor total comprado."
         Top             =   375
         Width           =   1395
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
         Left            =   8064
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do IPI."
         Top             =   375
         Width           =   1395
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
         ToolTipText     =   "Valor total do(s) produto(s)."
         Top             =   375
         Width           =   1395
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
         Left            =   1911
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do(s) serviço(s)."
         Top             =   375
         Width           =   1395
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
         Left            =   13260
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total comprado."
         Top             =   375
         Width           =   1755
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS ST"
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
         Left            =   9870
         TabIndex        =   57
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal"
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
         Left            =   6670
         TabIndex        =   56
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   9570
         TabIndex        =   55
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   195
         Left            =   6120
         TabIndex        =   54
         Top             =   435
         Width           =   120
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
         Height          =   195
         Left            =   5055
         TabIndex        =   53
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total desconto"
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
         Left            =   3679
         TabIndex        =   52
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   3450
         TabIndex        =   51
         Top             =   435
         Width           =   60
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   1680
         TabIndex        =   46
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   195
         Left            =   11295
         TabIndex        =   45
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total comprado"
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
         Left            =   11556
         TabIndex        =   44
         Top             =   180
         Width           =   1335
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
         Left            =   255
         TabIndex        =   43
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
         Left            =   8394
         TabIndex        =   42
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   7830
         TabIndex        =   41
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
         Left            =   2016
         TabIndex        =   40
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. total comprado"
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
         Left            =   13290
         TabIndex        =   39
         Top             =   180
         Width           =   1695
      End
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Pedido"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   6085
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Valor desconto"
         Object.Width           =   2205
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
         Text            =   "Valor ICMS ST"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "D"
         Text            =   "Dt. compra"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Tag             =   "T"
         Text            =   "Posto de trabalho"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6590
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   48
      Top             =   8910
      Width           =   11325
      _ExtentX        =   19976
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13920
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_Relatorios_Historico.frx":0442
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   49
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
      TabIndex        =   29
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
      TabIndex        =   50
      Top             =   960
      Width           =   1455
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
      TabIndex        =   30
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
         Left            =   7500
         TabIndex        =   7
         Top             =   990
         Width           =   735
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
         Width           =   1275
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":3236
         Left            =   180
         List            =   "frmCompras_Relatorios_Historico.frx":3255
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   450
         Width           =   3165
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":32DB
         Left            =   3360
         List            =   "frmCompras_Relatorios_Historico.frx":32DD
         MouseIcon       =   "frmCompras_Relatorios_Historico.frx":32DF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   450
         Width           =   6345
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
         TabIndex        =   37
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
         Left            =   1342
         TabIndex        =   32
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5797
         TabIndex        =   31
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   13140
      TabIndex        =   33
      Top             =   960
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":35E9
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico.frx":35F0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Por."
         Top             =   210
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
         Format          =   131727361
         CurrentDate     =   39799
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   14
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
         Format          =   131727361
         CurrentDate     =   39799
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":35F9
         Left            =   1260
         List            =   "frmCompras_Relatorios_Historico.frx":35FB
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   675
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":35FD
         Left            =   1260
         List            =   "frmCompras_Relatorios_Historico.frx":35FF
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   675
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":3601
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico.frx":3629
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Mês até."
         Top             =   930
         Visible         =   0   'False
         Width           =   645
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":366A
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico.frx":3692
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Mês de."
         Top             =   570
         Visible         =   0   'False
         Width           =   645
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":36D3
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico.frx":36D5
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   1305
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
         ItemData        =   "frmCompras_Relatorios_Historico.frx":36D7
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico.frx":36D9
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   990
         Width           =   360
      End
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
      Left            =   11580
      TabIndex        =   47
      Top             =   8940
      Width           =   3315
   End
End
Attribute VB_Name = "frmCompras_Relatorios_Historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=UNB5MhQdTA0&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=34&feature=plcp")

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
frmCompras_Relatorios_Historico_MenuImpressao.Show 1

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
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Compras_relatorios_historico_detalhado where IDLista = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Add , , TBAbrir!IDlista
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Pedido), "", TBAbrir!Pedido)
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Fornecedor), "", TBAbrir!Fornecedor)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!Quant_Comp), "", Format(TBAbrir!Quant_Comp, "###,##0.0000"))
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!preco_unitario), "", (Format(TBAbrir!preco_unitario, "###,##0.0000000000")))
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!ValorDesconto), "", (Format(TBAbrir!ValorDesconto, "###,##0.0000000000")))
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBAbrir!VlrIPI), "", (Format(TBAbrir!VlrIPI, "###,##0.00")))
                    .Item(.Count).SubItems(11) = IIf(IsNull(TBAbrir!Valor_ICMS_ST), "", (Format(TBAbrir!Valor_ICMS_ST, "###,##0.00")))
                    .Item(.Count).SubItems(12) = Format(IIf(IsNull(TBAbrir!preco_total), 0, TBAbrir!preco_total) + IIf(IsNull(TBAbrir!VlrIPI), 0, TBAbrir!VlrIPI), "###,##0.00")
                    .Item(.Count).SubItems(13) = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
                    If IsNull(TBAbrir!Status_Item) = False And TBAbrir!Status_Item <> "" Then
                        If TBAbrir!Status_Item = "N_RECEBIDO" Then .Item(.Count).SubItems(14) = "COMPRADO" Else .Item(.Count).SubItems(14) = TBAbrir!Status_Item
                    End If
                    .Item(.Count).SubItems(15) = IIf(IsNull(TBAbrir!maquina), "", TBAbrir!maquina)
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
                End With
            End If
        End If
        Familiatext = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtTotal_produtos = Format(TBLISTA!qtdeNC, "###,##0.00")
    Txt_total_servicos = Format(TBLISTA!Totalutilizada, "###,##0.00")
    txtDesconto = Format(TBLISTA!Valor2, "###,##0.00")
    txtSubtotal = Format(TBLISTA!qtdeNC + TBLISTA!Totalutilizada - TBLISTA!Valor2, "###,##0.00")
    txtTotal_ipi = Format(TBLISTA!Totalprevista, "###,##0.00")
    txtICMS_ST = Format(TBLISTA!CustoMat, "###,##0.00")
    txtDesconto_percentual = Format(TBLISTA!Valor1, "###,##0.00") & "%"
    txtTotal_geral = Format((TBLISTA!qtdeNC + TBLISTA!Totalutilizada + TBLISTA!Totalprevista + TBLISTA!CustoMat) - TBLISTA!Valor2, "###,##0.00")
    Txt_qtde_total_vendido = Format(TBLISTA!QtdePrevista, "###,##0.0000")
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
txtTotal_produtos = ""
Txt_total_servicos = ""
txtDesconto = ""
txtDesconto_percentual = ""
txtSubtotal = ""
txtTotal_ipi = ""
txtICMS_ST = ""
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
Formulario = "Compras/Relatórios/Histórico"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcCarregaComboAno Cmb_ano_ate, "2005", 1
ProcCarregaComboAno Cmb_ano_ate1, "2005", 1
ProcCarregaComboAno Cmb_ano_de, "2005", 1
ProcCarregaComboAno Cmb_ano_de1, "2005", 1
cmbfiltrarpor.Text = "Fornecedor"
cmbPor = "Dia"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Relatórios/Histórico"
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
If optDetalhado.Value = True And cmbfiltrarpor = "Posto de trabalho" Then Lista.ColumnHeaders(14).Width = 1500 Else Lista.ColumnHeaders(14).Width = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTexto()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
With cmbTexto
    .Clear
    .AddItem ""
    If Opt_individual.Value = True Or cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Or cmbfiltrarpor = "Família x Grupo" Then
        If cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Fornecedor from Compras_pedido where Data_aprovado IS NOT NULL group by Fornecedor", Conexao, adOpenKeyset, adLockReadOnly
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    .AddItem TBAbrir!Fornecedor
                    TBAbrir.MoveNext
                Loop
            End If
        Else
            If cmbfiltrarpor = "Código de referência" Then
                Ordenar = "n_referencia"
            ElseIf cmbfiltrarpor = "Família" Then
                    Ordenar = "familia"
                ElseIf cmbfiltrarpor = "Grupo" Then
                        Ordenar = "Grupo"
                    ElseIf cmbfiltrarpor = "Família x Grupo" Then
                            Ordenar = "Grupo"
                        ElseIf cmbfiltrarpor = "Descrição" Then
                                Ordenar = "desenho, descricao"
                                TextoFiltro = "descricao"
                            ElseIf cmbfiltrarpor = "Posto de trabalho" Then
                                    Ordenar = "maquina"
                                Else
                                    Ordenar = "desenho"
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            If cmbfiltrarpor = "Descrição" Then
                TBAbrir.Open "Select " & Ordenar & " as NomeCampo1 from Compras_relatorios_historico_detalhado where " & TextoFiltro & " <> 'Null' Group by " & Ordenar, Conexao, adOpenKeyset, adLockReadOnly
            Else
                TBAbrir.Open "Select " & Ordenar & " as NomeCampo1 from Compras_relatorios_historico_detalhado where " & Ordenar & " <> 'Null' Group by " & Ordenar, Conexao, adOpenKeyset, adLockReadOnly
            End If
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    .AddItem TBAbrir!NomeCampo1
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
        End If
    End If
    If Opt_comparativo = True And optResumido.Value = True Then
        If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Or cmbfiltrarpor = "Família x Grupo" Then .Enabled = True Else .Enabled = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Opt_comparativo.Value = True And cmbfiltrarpor = "Código interno x Fornecedor" And cmbTexto = "" Or Opt_comparativo.Value = True And cmbfiltrarpor = "Código de referência x Fornecedor" And cmbTexto = "" Or cmbfiltrarpor = "Família x Grupo" And cmbTexto = "" Then
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
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
If Txt_limite <> "" Then
    If Txt_limite < 10 Then
        USMsgBox ("O campo (Limitar em) não pode ser menor que 10."), vbExclamation, "CAPRIND v5.0"
        Txt_limite.SetFocus
        Exit Sub
    End If
End If

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
            quantidade = IIf(IsNull(TBLISTA!QtdeSaida), 0, TBLISTA!QtdeSaida) 'Qtde. comprada
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
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data, Maquina", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Maquina <> 'Null' order by Maquina, Ordem", Conexao, adOpenKeyset, adLockOptimistic
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
        Grupo = "desenho, Descricao"
        If cmbTexto <> "" Then FamiliaAntiga = "desenho = '" & cmbTexto & "' and "
    Case "Código de referência":
        Grupo = "n_referencia, Descricao"
        If cmbTexto <> "" Then FamiliaAntiga = "n_referencia = '" & cmbTexto & "' and "
    Case "Descrição":
        Grupo = "Descricao"
        If cmbTexto <> "" Then FamiliaAntiga = "Descricao = '" & cmbTexto & "' and "
    Case "Família x Grupo":
        Grupo = "Grupo, Familia"
        FamiliaAntiga = "Grupo = '" & cmbTexto & "' and "
    Case "Família":
        Grupo = "familia"
        If cmbTexto <> "" Then FamiliaAntiga = "familia = '" & cmbTexto & "' and "
    Case "Grupo"
        Grupo = "Grupo"
        If cmbTexto <> "" Then FamiliaAntiga = "Grupo = '" & cmbTexto & "' and "
    Case "Fornecedor":
        Grupo = "Fornecedor"
        If cmbTexto <> "" Then FamiliaAntiga = "Fornecedor = '" & cmbTexto & "' and "
    Case "Posto de trabalho":
        Grupo = "maquina"
        If cmbTexto <> "" Then FamiliaAntiga = "maquina = '" & cmbTexto & "' and " Else FamiliaAntiga = "maquina IS NOT NULL and "
    Case "Código interno x Fornecedor":
        Grupo = "desenho, Descricao"
    Case "Código de referência x Fornecedor":
        Grupo = "n_referencia, Descricao"
End Select

Set TBCarteira = CreateObject("adodb.recordset")
If optDetalhado.Value = True Then
    TBCarteira.Open "Select * from Compras_relatorios_historico_detalhado where " & FamiliaAntiga & " (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' order by Data, IDLista", Conexao, adOpenKeyset, adLockReadOnly
Else
    If Opt_quantidade.Value = True Then TextoFiltro = "Quant_Comp" Else TextoFiltro = "preco_total"
    
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
            Pesquisa = "(Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Data In (" & Par1 & "))"
            Pesquisa2 = "Data"
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
            Pesquisa = "Month(Data) >= '" & MesX & "' and Year(Data) = '" & Cmb_ano_de & "' and Month(Data) <= '" & MesX1 & "' and Year(Data) = '" & Cmb_ano_ate & "'"
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
            Pesquisa = "Year(Data) >= '" & Cmb_ano_de1 & "' and Year(Data) <= '" & Cmb_ano_ate1 & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Ano In (" & Par1 & "))"
            Pesquisa2 = "Ano"
    End Select
    If Opt_individual.Value = True Then
        TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where " & FamiliaAntiga & Pesquisa & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
    Else
        Pesquisa3 = "Desenho is not null"
        If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then
            TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where Fornecedor = '" & cmbTexto & "' and " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
        ElseIf cmbfiltrarpor = "Família x Grupo" Then
                TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where Grupo = '" & cmbTexto & "' and " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
            Else
                TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
        End If
    End If
End If
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
Valor_Cofins_Prod = 0
Valor_ICMS_SN = 0
Desconto = 0
If TBCarteira.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBCarteira.EOF = False
        If optDetalhado.Value = True Then
            Set TBProdutividade = CreateObject("adodb.recordset")
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            ProcCriarResumido
        End If
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

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

Permitido = True
TBProdutividade.AddNew
TBProdutividade!Ordem = TBCarteira!IDlista
TBProdutividade!Data = TBCarteira!Data
TBProdutividade!QtdePrev = IIf(IsNull(TBCarteira!Quant_Comp), 0, TBCarteira!Quant_Comp) 'Qtde. comprada
TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira!preco_total), 0, TBCarteira!preco_total) + IIf(IsNull(TBCarteira!VlrIPI), 0, TBCarteira!VlrIPI)   'Valor total
TBProdutividade!Terceiros = IIf(IsNull(TBCarteira!ValorDesconto), 0, TBCarteira!ValorDesconto) * IIf(IsNull(TBCarteira!Quant_Comp), 0, TBCarteira!Quant_Comp) 'Valor desconto

If TBCarteira!Tipo = "P" Then
    TBProdutividade!qtdeNC = IIf(IsNull(TBCarteira!preco_total), 0, TBCarteira!preco_total) 'Valor total produtos
Else
    TBProdutividade!Qtdetotalprod = IIf(IsNull(TBCarteira!preco_total), 0, TBCarteira!preco_total) 'Valor total serviços
End If
TBProdutividade!Eficiencia = IIf(IsNull(TBCarteira!VlrIPI), 0, TBCarteira!VlrIPI) 'Valor total IPI
TBProdutividade!impostos = IIf(IsNull(TBCarteira!Valor_ICMS_ST), 0, TBCarteira!Valor_ICMS_ST) 'Valor_ICMS_ST
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = Familiatext
TBProdutividade.Update

quantidade = quantidade + TBProdutividade!QtdePrev 'Qtde. comprada
QTLOTE = QTLOTE + TBProdutividade!qtdeOK 'Valor total
If TBCarteira!Tipo = "P" Then
    Valor_Produto = Valor_Produto + (IIf(IsNull(TBCarteira!Quant_Comp), 0, TBCarteira!Quant_Comp) * IIf(IsNull(TBCarteira!preco_unitario), 0, TBCarteira!preco_unitario)) 'Valor total produtos
Else
    Valor_Cofins_Serv = Valor_Cofins_Serv + (IIf(IsNull(TBCarteira!Quant_Comp), 0, TBCarteira!Quant_Comp) * IIf(IsNull(TBCarteira!preco_unitario), 0, TBCarteira!preco_unitario)) 'Valor total serviços
End If
ValorIPI = ValorIPI + TBProdutividade!Eficiencia 'Valor total IPI
Valor_Cofins_Prod = Valor_Cofins_Prod + TBProdutividade!Terceiros 'Valor desconto
Valor_ICMS_SN = Valor_ICMS_SN + IIf(IsNull(TBCarteira!Valor_ICMS_ST), 0, TBCarteira!Valor_ICMS_ST) 'Valor ICMS ST

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
    Case "Descrição": Familiatext = TBCarteira!Descricao
    Case "Família x Grupo": Familiatext = TBCarteira!Familia
    Case "Família": Familiatext = TBCarteira!Familia
    Case "Grupo": Familiatext = TBCarteira!Grupo
    Case "Fornecedor": Familiatext = TBCarteira!Fornecedor
    Case "Posto de trabalho": Familiatext = IIf(IsNull(TBCarteira!maquina), "", TBCarteira!maquina)
    Case "Código interno x Fornecedor": Familiatext = TBCarteira!Desenho
    Case "Código de referência x Fornecedor": Familiatext = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
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

If cmbfiltrarpor = "Código interno" Or cmbfiltrarpor = "Código de referência" Or cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then
    TBProdutividade!maquina = Left(Familiatext & " " & TBCarteira!Descricao, 25)
Else
    TBProdutividade!maquina = Left(Familiatext, 25)
End If

TBProdutividade.Update
TBProdutividade.Close

Select Case cmbfiltrarpor
    Case "Código interno": Produto = TBCarteira!Desenho
    Case "Código de referência": Produto = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    Case "Descrição": Produto = TBCarteira!Descricao
    Case "Família x Grupo": Produto = TBCarteira!Familia
    Case "Família": Produto = TBCarteira!Familia
    Case "Grupo": Produto = TBCarteira!Grupo
    Case "Fornecedor": Produto = TBCarteira!Fornecedor
    Case "Posto de trabalho": Produto = IIf(IsNull(TBCarteira!maquina), "", TBCarteira!maquina)
    Case "Código interno x Fornecedor": Produto = TBCarteira!Desenho
    Case "Código de referência x Fornecedor": Produto = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaColunas()
On Error GoTo tratar_erro

Lista1.ColumnHeaders.Clear
contador = 1
With Lista1.ColumnHeaders
    .Add
    If cmbfiltrarpor <> "Código interno x Fornecedor" And cmbfiltrarpor <> "Código de referência x Fornecedor" And cmbfiltrarpor <> "Família x Grupo" Then
        .Item(contador).Text = cmbfiltrarpor.Text
    Else
        If cmbfiltrarpor = "Código interno x Fornecedor" Then
            .Item(contador).Text = "Código interno"
        ElseIf cmbfiltrarpor = "Código de referência x Fornecedor" Then
                .Item(contador).Text = "Código de referência"
            Else
                .Item(contador).Text = "Família"
        End If
    End If
    .Item(contador).Width = 3500
    If cmbPor.Text = "Dia" Then
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            .Add
            contador = contador + 1
            .Item(contador).Text = Format(Dataini, "dd/mm/yy")
            .Item(contador).Alignment = lvwColumnRight
            Dataini = Dataini + 1
        Loop
    End If
    If cmbPor.Text = "Mês" Then
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)
        Do While qt <= Qtd
            .Add
            contador = contador + 1
            .Item(contador).Text = qt & "/" & Cmb_ano_de
            .Item(contador).Alignment = lvwColumnRight
            qt = qt + 1
        Loop
    End If
    If cmbPor.Text = "Ano" Then
        qt = Cmb_ano_de1
        Do While qt <= Cmb_ano_ate1
            .Add
            contador = contador + 1
            .Item(contador).Text = qt
            .Item(contador).Alignment = lvwColumnRight
            qt = qt + 1
        Loop
    End If
    .Add
    contador = contador + 1
    If Opt_valor.Value = True Then
        .Item(contador).Text = "Valor total"
        .Item(contador).Alignment = lvwColumnRight
    Else
        .Item(contador).Text = "Qtde. total"
        .Item(contador).Alignment = lvwColumnRight
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
Fornecedor = ""
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select maquina, OS from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' Group by Maquina, OS order by OS desc", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    If TBLISTA.RecordCount < 10 Then Exit Sub
    Do While Contador1 <> 0
        If Fornecedor <> TBLISTA!maquina Then
            Valor_total = TBLISTA!OS
            Contador1 = Contador1 - 1
        End If
        Fornecedor = TBLISTA!maquina
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
        TextoFiltroPadrao = FamiliaAntiga & Pesquisa
    Else
        If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Or cmbfiltrarpor = "Família x Grupo" Then
            If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then TextoFiltro = "Fornecedor" Else TextoFiltro = "Grupo"
            TextoFiltroPadrao = TextoFiltro & " = '" & cmbTexto & "' and " & Pesquisa
        Else
            TextoFiltroPadrao = IIf(cmbfiltrarpor = "Posto de trabalho", FamiliaAntiga, "") & Pesquisa
        End If
    End If
    'Produtos e IPI
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Sum(Quant_Comp * preco_unitario) as Valor, Sum(vlripi) as ValorIPI, Sum(valordesconto * Quant_Comp) as Desconto, Sum(Valor_ICMS_ST) as ICMS_ST from Compras_relatorios_historico_detalhado where " & TextoFiltroPadrao & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Valor_Produto = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total produto
        ValorIPI = IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total IPI
        Valor_Cofins_Prod = IIf(IsNull(TBLISTA!Desconto), 0, TBLISTA!Desconto)
        Valor_ICMS_SN = IIf(IsNull(TBLISTA!ICMS_ST), 0, TBLISTA!ICMS_ST) 'Valor total ICMS ST
        QTLOTE = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) + IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI) 'Valor total
    End If
        
    'Serviços
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Sum(Quant_Comp * preco_unitario) as Valor, Sum(valordesconto * Quant_Comp) as Desconto from Compras_relatorios_historico_detalhado where " & TextoFiltroPadrao & " and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Valor_Cofins_Serv = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor total serviços
        Desconto = IIf(IsNull(TBLISTA!Desconto), 0, TBLISTA!Desconto)
    End If
    
    'Quantidade
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Sum(Quant_Comp) as QtdeSaida from Compras_relatorios_historico_detalhado where " & TextoFiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        quantidade = IIf(IsNull(TBLISTA!QtdeSaida), 0, TBLISTA!QtdeSaida) 'Qtde. comprada
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
If cmbfiltrarpor <> "Código interno x Fornecedor" And cmbfiltrarpor <> "Código de referência x Fornecedor" And cmbfiltrarpor <> "Família x Grupo" Then
    If Opt_individual.Value = True And cmbTexto <> "" Then TBAbrir!Texto = cmbfiltrarpor & " : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor
    TBAbrir!QtdeOrdem = "1"
Else
    If cmbfiltrarpor = "Código interno x Fornecedor" Then
        If Opt_individual.Value = True And cmbTexto <> "" Then
            TBAbrir!Texto = "Código interno" & " : " & cmbTexto
        Else
            TBAbrir!Texto = "Código interno"
        End If
        TBAbrir!QtdeOrdem = "2"
    ElseIf cmbfiltrarpor = "Código de referência x Fornecedor" Then
            If Opt_individual.Value = True And cmbTexto <> "" Then
                TBAbrir!Texto = "Código de referência" & " : " & cmbTexto
            Else
                TBAbrir!Texto = "Código de referência"
            End If
            TBAbrir!QtdeOrdem = "2"
        ElseIf cmbfiltrarpor = "Família x Grupo" Then
            If Opt_individual.Value = True And cmbTexto <> "" Then
                TBAbrir!Texto = "Fámilia" & " : " & cmbTexto
            Else
                TBAbrir!Texto = "Família"
            End If
            TBAbrir!QtdeOrdem = "3"
    End If
    TBAbrir!Texto1 = cmbTexto
End If

TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
If Opt_quantidade.Value = True Then TBAbrir!Turno = True Else TBAbrir!Turno = False
TBAbrir!QtdePrevista = quantidade 'Qtde. comprada
TBAbrir!QtdeProduzida = QTLOTE 'Valor total
TBAbrir!qtdeNC = Valor_Produto 'Valor total produtos
TBAbrir!Totalutilizada = Format(Valor_Cofins_Serv, "###,##0.00") 'Valor serviços
TBAbrir!Totalprevista = Format(ValorIPI, "###,##0.00") 'Valor total IPI
TBAbrir!Valor2 = Format(Valor_Cofins_Prod + Desconto, "###,##0.00") 'Valor total desconto
If TBAbrir!qtdeNC + TBAbrir!Totalutilizada = 0 Then TBAbrir!Valor1 = 0 Else TBAbrir!Valor1 = (TBAbrir!Valor2 * 100) / (TBAbrir!qtdeNC + TBAbrir!Totalutilizada)  'Percentual desconto
TBAbrir!CustoMat = Format(Valor_ICMS_SN, "###,##0.00") 'Valor total ICMS ST

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
    .AddItem "Fornecedor"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Posto de trabalho"
    .AddItem "Família x Grupo"
    .AddItem "Família"
    .AddItem "Grupo"
    .AddItem "Código interno x Fornecedor"
    .AddItem "Código de referência x Fornecedor"
    .Text = "Fornecedor"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

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
    .AddItem "Fornecedor"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Posto de trabalho"
    .AddItem "Família x Grupo"
    .AddItem "Grupo"
    .AddItem "Família"
    .Text = "Fornecedor"
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
