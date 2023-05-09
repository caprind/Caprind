VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoque_Recebimento 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Recebimento"
   ClientHeight    =   10035
   ClientLeft      =   1050
   ClientTop       =   1665
   ClientWidth     =   15360
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
   ForeColor       =   &H00000000&
   Icon            =   "frmEstoque_Recebimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximizado
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1360
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
   Begin VB.TextBox txtID_empresa 
      Alignment       =   2  'Centralizar
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
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   255
      MouseIcon       =   "frmEstoque_Recebimento.frx":014A
      MousePointer    =   99  'Custom
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   3030
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   55
      TabIndex        =   48
      Top             =   1020
      Width           =   11745
      Begin VB.TextBox Txt_ID_pedido 
         Alignment       =   2  'Centralizar
         BackColor       =   &H80000014&
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
         Height          =   315
         Left            =   120
         TabIndex        =   103
         ToolTipText     =   "ID do pedido"
         Top             =   390
         Width           =   735
      End
      Begin VB.TextBox txtuf 
         Alignment       =   2  'Centralizar
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
         Left            =   6380
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         ToolTipText     =   "UF."
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Txt_ID_forn 
         Alignment       =   2  'Centralizar
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Código do fornecedor."
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtEmpresa 
         Alignment       =   2  'Centralizar
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
         Left            =   6990
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Empresa."
         Top             =   960
         Width           =   4605
      End
      Begin VB.TextBox txtcondpagamento 
         Alignment       =   2  'Centralizar
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
         Left            =   5925
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Condições de pagamento."
         Top             =   390
         Width           =   5295
      End
      Begin VB.CommandButton cmdPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2490
         Picture         =   "frmEstoque_Recebimento.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Filtrar."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtProg_pedido 
         Alignment       =   2  'Centralizar
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
         Left            =   870
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Programação de compra."
         Top             =   390
         Width           =   1605
      End
      Begin VB.CommandButton cmdPagamento 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   11250
         Picture         =   "frmEstoque_Recebimento.frx":086F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Localizar condições de pagamento."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtvalortotal 
         Alignment       =   1  'Alinhar à Direita
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
         Left            =   4170
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Valor total."
         Top             =   390
         Width           =   1740
      End
      Begin VB.TextBox txtfornecedor 
         Alignment       =   2  'Centralizar
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
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Fornecedor."
         Top             =   960
         Width           =   5525
      End
      Begin VB.TextBox txtdata 
         Alignment       =   2  'Centralizar
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
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data."
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   104
         Top             =   180
         Width           =   165
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparente
         Caption         =   "UF"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6583
         TabIndex        =   98
         Top             =   750
         Width           =   195
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparente
         Caption         =   "ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   94
         Top             =   750
         Width           =   165
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   8925
         TabIndex        =   75
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Valor total"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4673
         TabIndex        =   54
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblPedido 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Pedido de compra"
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
         Left            =   930
         TabIndex        =   53
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparente
         Caption         =   "Fornecedor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3190
         TabIndex        =   52
         Top             =   750
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Condições de pagamento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7665
         TabIndex        =   51
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3352
         TabIndex        =   50
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   11820
      TabIndex        =   56
      Top             =   1020
      Width           =   2295
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Centralizar
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
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   16
         ToolTipText     =   "Série."
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Cmd_salvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1800
         Picture         =   "frmEstoque_Recebimento.frx":0971
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salvar dados da nota fiscal no produto/serviço recebido."
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtnotafiscal 
         Alignment       =   2  'Centralizar
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
         MaxLength       =   9
         TabIndex        =   15
         ToolTipText     =   "Número da nota fiscal."
         Top             =   960
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtdataemissao 
         Height          =   315
         Left            =   180
         TabIndex        =   13
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   390
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Image imgCalendario 
         Height          =   360
         Left            =   1785
         Picture         =   "frmEstoque_Recebimento.frx":09C4
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Série"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1357
         TabIndex        =   97
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Data emissão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   495
         TabIndex        =   58
         Top             =   180
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "N° nota fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   765
         Width           =   960
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2925
      Left            =   60
      TabIndex        =   60
      Top             =   7230
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   5159
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Recebimento"
      TabPicture(0)   =   "frmEstoque_Recebimento.frx":0E47
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdReceber"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(3)=   "Frame11"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Movimentação"
      TabPicture(1)   =   "frmEstoque_Recebimento.frx":0E63
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdCancelar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin DrawSuite2014.USButton cmdCancelar 
         Height          =   2430
         Left            =   14100
         TabIndex        =   47
         ToolTipText     =   "Excluir recebimento (F4)"
         Top             =   330
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   4286
         DibPicture      =   "frmEstoque_Recebimento.frx":0E7F
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Excluir recebimento (F4)"
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
         PicAlign        =   8
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
      End
      Begin DrawSuite2014.USButton cmdReceber 
         Height          =   1500
         Left            =   -60990
         TabIndex        =   45
         ToolTipText     =   "Receber (F3)"
         Top             =   1260
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   2646
         DibPicture      =   "frmEstoque_Recebimento.frx":6D46
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Receber (F3)"
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
         HandPointer     =   0   'False
         PicAlign        =   8
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Height          =   2450
         Left            =   30
         TabIndex        =   69
         Top             =   330
         Width           =   14055
         Begin MSComctlLib.ListView Lista_movimentacao 
            Height          =   1860
            Left            =   180
            TabIndex        =   46
            Top             =   195
            Width           =   13665
            _ExtentX        =   24104
            _ExtentY        =   3281
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "RE"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Operação"
               Object.Width           =   4912
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Entrada"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Entrada PÇ"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Documento"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "N. de série"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Insp."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "IDestCR"
               Object.Width           =   0
            EndProperty
         End
         Begin DrawSuite2014.USProgressBar PBlista1 
            Height          =   255
            Left            =   180
            TabIndex        =   79
            Top             =   2070
            Width           =   13665
            _ExtentX        =   24104
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
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   915
         Left            =   -74970
         TabIndex        =   61
         Top             =   330
         Width           =   5865
         Begin VB.ComboBox Cmb_codigo_ref 
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   27
            ToolTipText     =   "Codigo de referência."
            Top             =   390
            Width           =   1890
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   2  'Centralizar
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
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1605
         End
         Begin VB.TextBox txtUN 
            Alignment       =   2  'Centralizar
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
            Left            =   3690
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Unidade."
            Top             =   390
            Width           =   390
         End
         Begin VB.TextBox txtstatus 
            Alignment       =   2  'Centralizar
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
            Left            =   4095
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1650
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparente
            Caption         =   "Código de ref."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2223
            TabIndex        =   95
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparente
            Caption         =   "Código interno"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   457
            TabIndex        =   64
            Top             =   180
            Width           =   1050
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Un"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3795
            TabIndex        =   63
            Top             =   180
            Width           =   195
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Status"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4688
            TabIndex        =   62
            Top             =   180
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1515
         Left            =   -74970
         TabIndex        =   65
         Top             =   1260
         Width           =   13965
         Begin VB.CheckBox Chk_LA 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4770
            TabIndex        =   101
            Top             =   180
            Width           =   195
         End
         Begin VB.CheckBox Chk_Dt_rcbto 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7830
            TabIndex        =   100
            Top             =   180
            Width           =   195
         End
         Begin VB.TextBox Txt_numero_serie 
            Alignment       =   2  'Centralizar
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
            Height          =   335
            Left            =   12030
            MaxLength       =   50
            TabIndex        =   43
            ToolTipText     =   "Número de série."
            Top             =   390
            Width           =   1725
         End
         Begin VB.TextBox txtQuantidade_PC 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   10830
            MaxLength       =   50
            TabIndex        =   42
            ToolTipText     =   "Quantidade de peças à receber."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txtOBS 
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
            Height          =   405
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            ToolTipText     =   "Observações."
            Top             =   990
            Width           =   13605
         End
         Begin VB.TextBox txtQuantidade 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   9210
            MaxLength       =   50
            TabIndex        =   40
            ToolTipText     =   "Quantidade à receber."
            Top             =   390
            Width           =   1185
         End
         Begin VB.CommandButton cmdcalc_peso 
            BackColor       =   &H00C0C0C0&
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
            Left            =   10410
            Picture         =   "frmEstoque_Recebimento.frx":94ED
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Abrir calculadora para cálculo de peso."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtcertificado 
            Alignment       =   2  'Centralizar
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
            Height          =   335
            Left            =   1960
            MaxLength       =   50
            TabIndex        =   37
            ToolTipText     =   "Número do certificado."
            Top             =   390
            Width           =   1790
         End
         Begin VB.TextBox txtcorrida 
            Alignment       =   2  'Centralizar
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
            Height          =   335
            Left            =   150
            MaxLength       =   50
            TabIndex        =   36
            ToolTipText     =   "Número da corrida."
            Top             =   390
            Width           =   1800
         End
         Begin VB.ComboBox cmbLocal_armaz 
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
            Left            =   3765
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   38
            ToolTipText     =   "Local de armazenamento."
            Top             =   390
            Width           =   4050
         End
         Begin MSMask.MaskEdBox Txt_data_recebimento 
            Height          =   315
            Left            =   7830
            TabIndex        =   39
            ToolTipText     =   "Data do recebimento."
            Top             =   390
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Image imgCalendario_receb 
            Height          =   360
            Left            =   8820
            Picture         =   "frmEstoque_Recebimento.frx":9756
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Número de série"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12307
            TabIndex        =   96
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Qtd. receber PÇ"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10837
            TabIndex        =   90
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6435
            TabIndex        =   89
            Top             =   780
            Width           =   945
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Dt. receb."
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
            Left            =   8085
            TabIndex        =   88
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Qtd. receber"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9330
            TabIndex        =   73
            Top             =   180
            Width           =   930
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Certificado"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2465
            TabIndex        =   68
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Corrida"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   788
            TabIndex        =   67
            Top             =   180
            Width           =   525
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Local armazenamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5010
            TabIndex        =   66
            Top             =   180
            Width           =   1560
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   915
         Left            =   -69090
         TabIndex        =   70
         Top             =   330
         Width           =   9225
         Begin VB.TextBox txtrecebida_PC 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   6090
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade recebida em peça."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox txtrequisitado_PC 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade comprada em peça."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox txtSaldo_PC 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   7560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Saldo em peça."
            Top             =   390
            Width           =   1485
         End
         Begin VB.TextBox txtSaldo 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   3150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Saldo."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox txtrequisitado 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   210
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade comprada."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox txtrecebida 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade recebida."
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Comprada PÇ"
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
            Left            =   4785
            TabIndex        =   93
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Recebida PÇ"
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
            Left            =   6300
            TabIndex        =   92
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Saldo PÇ"
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
            Left            =   7942
            TabIndex        =   91
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Saldo"
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
            Left            =   3645
            TabIndex        =   86
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Recebida"
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
            Left            =   2017
            TabIndex        =   72
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Comprada"
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
            Left            =   502
            TabIndex        =   71
            Top             =   180
            Width           =   870
         End
      End
   End
   Begin VB.TextBox txtidlista 
      Alignment       =   2  'Centralizar
      BackColor       =   &H80000014&
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
      Height          =   315
      Left            =   8460
      TabIndex        =   55
      ToolTipText     =   "ID da lista"
      Top             =   420
      Visible         =   0   'False
      Width           =   1335
   End
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   77
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
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
      ButtonCaption3  =   "Status"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Alterar status do(s) produto(s)/serviço(s) (F7)"
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
      ButtonLeft3     =   108
      ButtonTop3      =   2
      ButtonWidth3    =   45
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Atualizar"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Utilizado pelo administrador do sistema."
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
      ButtonLeft4     =   155
      ButtonTop4      =   2
      ButtonWidth4    =   59
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
      ButtonLeft5     =   216
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
      ButtonLeft6     =   220
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
      ButtonLeft7     =   263
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
      ButtonLeft8     =   295
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
      Begin DrawSuite2014.USImageList USImageList1 
         Left            =   5820
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmEstoque_Recebimento.frx":9BD9
         Count           =   1
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista de produtos/serviços "
      ForeColor       =   &H00000000&
      Height          =   3525
      Left            =   55
      TabIndex        =   59
      Top             =   2460
      Width           =   15195
      Begin VB.TextBox txtUnCom 
         Alignment       =   2  'Centralizar
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
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         ToolTipText     =   "Unidade comercial."
         Top             =   900
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtEspecificacoes 
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
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   255
         MouseIcon       =   "frmEstoque_Recebimento.frx":DD17
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   570
         Visible         =   0   'False
         Width           =   4395
      End
      Begin MSComctlLib.ListView listprod 
         Height          =   3135
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   5530
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
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Pedido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   6227
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Vlr. unitário"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Qtde. PÇ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "D"
            Text            =   "Prazo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11100
         TabIndex        =   76
         Top             =   0
         Width           =   3900
         Begin VB.OptionButton OptTodos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Todos"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2970
            TabIndex        =   12
            Top             =   30
            Width           =   885
         End
         Begin VB.OptionButton optRecebidos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Recebidos"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1770
            TabIndex        =   11
            Top             =   30
            Width           =   1215
         End
         Begin VB.OptionButton OptAreceber 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Á receber"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   660
            TabIndex        =   10
            Top             =   30
            Value           =   -1  'True
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   83
      Top             =   6000
      Width           =   15195
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Centralizar
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
         Left            =   3780
         TabIndex        =   18
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
         TabIndex        =   19
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2014.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   23
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":E021
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
      Begin DrawSuite2014.USButton cmdPagAnt 
         Height          =   315
         Left            =   11220
         TabIndex        =   22
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":117C5
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
      Begin DrawSuite2014.USButton cmdPagIr 
         Height          =   315
         Left            =   10110
         TabIndex        =   20
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
      Begin DrawSuite2014.USButton cmdPagPrim 
         Height          =   315
         Left            =   10680
         TabIndex        =   21
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":152CE
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
      Begin DrawSuite2014.USButton cmdPagUlt 
         Height          =   315
         Left            =   12300
         TabIndex        =   24
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":193BD
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
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4410
         TabIndex        =   99
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3090
         TabIndex        =   87
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   85
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   84
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   55
      TabIndex        =   80
      Top             =   6630
      Width           =   15195
      Begin VB.TextBox txtQtde_total 
         Alignment       =   1  'Alinhar à Direita
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
         Height          =   315
         Left            =   13440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total."
         Top             =   180
         Width           =   1560
      End
      Begin DrawSuite2014.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   81
         Top             =   210
         Width           =   12105
         _ExtentX        =   21352
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
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Qtde. total :"
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
         Left            =   12390
         TabIndex        =   82
         Top             =   180
         Width           =   2415
         WordWrap        =   -1  'True
      End
   End
   Begin DrawSuite2014.USButton cmdNota 
      Height          =   1425
      Left            =   14140
      TabIndex        =   49
      ToolTipText     =   "Emitir nota fiscal."
      Top             =   1020
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   2514
      DibPicture      =   "frmEstoque_Recebimento.frx":1CC49
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      Caption         =   "Emitir NF"
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
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
End
Attribute VB_Name = "frmEstoque_Recebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLISTA_Estoque_RecebimentoPedido As ADODB.Recordset 'OK
Public StrSql_Estoque_Recebimento_Localizar As String 'OK
Public StrSql_Estoque_Recebimento_LocalizarTotal As String 'OK
Public FormulaRel_Estoque_Recebimento As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=BZk-gwHpncU&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=47&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaMovimentacao()
On Error GoTo tratar_erro

Lista_movimentacao.ListItems.Clear
If txtcodigo = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select EM.*, ECR.recebido, ECR.ID, EC.Numero_serie from (estoque_controle_recebimento ECR INNER JOIN estoque_movimentacao EM on ECR.ID = EM.idestoque_recebimento) INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where ECR.IDLista = " & IIf(txtidlista = "", 0, txtidlista) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " order by EM.Idoperacao", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    PBlista1.Min = 0
    PBlista1.Max = TBProduto.RecordCount
    PBlista1.Value = 1
    Contador = 0
    Do While TBProduto.EOF = False
        With Lista_movimentacao.ListItems
            .Add , , TBProduto!IDoperacao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProduto!IDestoque), 0, TBProduto!IDestoque)
            .Item(.Count).SubItems(2) = TBProduto!Operacao
            .Item(.Count).SubItems(3) = Format(TBProduto!data, "dd/mm/yy")
            .Item(.Count).SubItems(4) = Format(TBProduto!Entrada, "###,##0.0000")
            .Item(.Count).SubItems(5) = IIf(IsNull(TBProduto!Entrada_PC), "", Format(TBProduto!Entrada_PC, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBProduto!Documento), "", TBProduto!Documento)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBProduto!Numero_serie), "", TBProduto!Numero_serie)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select ID from Compras_recebimento where IDestoque = " & TBProduto!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then .Item(.Count).SubItems(9) = "Sim" Else .Item(.Count).SubItems(9) = "Não"
            
            .Item(.Count).SubItems(10) = IIf(IsNull(TBProduto!ID), 0, TBProduto!ID)
        End With
        TBProduto.MoveNext
        Contador = Contador + 1
        PBlista1.Value = Contador
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_Click()
On Error GoTo tratar_erro

Acao = "salvar"
If txtidlista = "" Or txtidlista = "0" Then
    NomeCampo = "o produto/serviço"
    ProcVerificaAcao
    Exit Sub
End If
If Lista_movimentacao.ListItems.Count = 0 Then
    NomeCampo = "a movimentação na lista"
    ProcVerificaAcao
    Exit Sub
End If
If Lista_movimentacao.SelectedItem = False Then
    NomeCampo = "a movimentação na lista"
    ProcVerificaAcao
    Exit Sub
End If
If txtstatus = "NÃO_RECEBIDO" Then Exit Sub
If IsDate(txtdataemissao) = False Then
    NomeCampo = "a data de emissão da nota fiscal"
    ProcVerificaAcao
    txtdataemissao.SetFocus
    Exit Sub
End If
If txtnotafiscal = "" Then
    NomeCampo = "o número da nota fiscal"
    ProcVerificaAcao
    txtnotafiscal.SetFocus
    Exit Sub
End If
If txtSerie = "" Then
    NomeCampo = "a série"
    ProcVerificaAcao
    txtSerie.SetFocus
    Exit Sub
End If
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_movimentacao where Idoperacao = " & Lista_movimentacao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    If IsNull(TBEstoque!IDEstoque_recebimento) = False And TBEstoque!IDEstoque_recebimento <> "" Then
        'Atualiza dados da nota na tabela de estoque_controle_recebimento
        Conexao.Execute "Update estoque_controle_recebimento Set Nota_fiscal = '" & txtnotafiscal & "', Serie = '" & txtSerie & "', Data_emissao =  '" & Format(txtdataemissao, "Short Date") & "' where Id = " & TBEstoque!IDEstoque_recebimento
        
        TBEstoque!Documento = txtnotafiscal
        TBEstoque.Update
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Recebimento/Pedido de compra"
        Evento = "Salvar nota fiscal"
        ID_documento = txtidlista
        Documento = "Cód. interno: " & txtcodigo & " - Nº lote: " & txtProg_pedido & " - Nº corrida: " & txtcorrida & " - Nº certificado: " & txtcertificado & " - Local armaz.: " & cmbLocal_armaz
        Documento1 = "Operação: " & Lista_movimentacao.SelectedItem.SubItems(2) & " - Documento: " & Lista_movimentacao.SelectedItem.SubItems(6)
        ProcGravaEvento
        '==================================
        ProcAtualizaVlrEntradaEstoque True
        ProcCarregaListaMovimentacao
    Else
        USMsgBox ("Não é possivel salvar os dados da nota fiscal, pois não foi encontrato o id do recebimento na movimentação"), vbExclamation, "CAPRIND v5.0"
    End If
End If
TBEstoque.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNota_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

ID_nota = 0
Acao = "emitir a nota"
Servicos = False
Prod = False

If txtProg_pedido = "" Then
    NomeCampo = "o pedido"
    ProcVerificaAcao
    txtProg_pedido.SetFocus
    Exit Sub
End If
If txtuf <> "EX" Then
    If IsDate(txtdataemissao) = False Then
        NomeCampo = "a data de emissão da nota fiscal"
        ProcVerificaAcao
        txtdataemissao.SetFocus
        Exit Sub
    End If
    If txtnotafiscal = "" Then
        NomeCampo = "a nota fiscal"
        ProcVerificaAcao
        txtnotafiscal.SetFocus
        Exit Sub
    End If
    If txtSerie = "" Then
        NomeCampo = "a série"
        ProcVerificaAcao
        txtSerie.SetFocus
        Exit Sub
    End If
End If

TextoFiltro = ""
If txtuf <> "EX" Then TextoFiltro = " and ECR.Nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Data_emissao = '" & Format(txtdataemissao, "Short Date") & "'"

'Verifica se tem algum produto/serviço recebido para o pedido
Prodpedido = False
ServPedido = False
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select CPL.Tipo from (compras_pedido_lista CPL INNER JOIN Estoque_controle_recebimento ECR ON CPL.IDPedido = ECR.IDPedido and CPL.IDLista = ECR.IDLista and CPL.Desenho = ECR.Desenho) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDpedido where CP.Pedido = '" & txtProg_pedido & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = True Then
    USMsgBox ("É necessário receber o(s) produto(s)/serviço(s) deste pedido antes de emitir a nota."), vbExclamation, "CAPRIND v5.0"
    TBPedido.Close
    Exit Sub
Else
    'Verifica tipo da nota
    Do While TBPedido.EOF = False
        If TBPedido!Tipo = "P" Then Prodpedido = True Else ServPedido = True
        TBPedido.MoveNext
    Loop
End If
TBPedido.Close
If Prodpedido = True And ServPedido = True Then
    TipoNF = "M1SA"
ElseIf Prodpedido = True Then
        TipoNF = "M1"
    Else
        TipoNF = "SA"
End If

'Cria a nota fiscal
If txtuf = "EX" Then
    If txtnotafiscal = "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select CAST(int_NotaFiscal AS int) AS NF, Serie FROM tbl_Dados_Nota_Fiscal where tipoNF = '" & TipoNF & "' and Aplicacao = 'P' and ID_empresa = " & txtID_empresa & " and int_NotaFiscal IS NOT NULL order by dt_DataEmissao desc, NF desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            QuantsolicitadoN1 = TBAbrir!NF + 1
            FamiliaAntiga = QuantsolicitadoN1
            Familiatext = FunTamanhoTextoZeroEsq(FamiliaAntiga, 9)
            SerieNF = IIf(IsNull(TBAbrir!Serie), 1, TBAbrir!Serie)
        Else
            Familiatext = "000000001"
            SerieNF = 1
        End If
        txtdataemissao = Format(Date, "dd/mm/yyyy")
        txtnotafiscal = FunVerifExisteNumNF(TipoNF, txtID_empresa, Familiatext, SerieNF)
        txtSerie = SerieNF
        
        Conexao.Execute "Update ECR set ECR.Nota_fiscal = '" & txtnotafiscal & "', ECR.Serie = '" & txtSerie & "', ECR.Data_emissao =  '" & Format(txtdataemissao, "Short Date") & "' from Estoque_controle_recebimento ECR INNER JOIN Compras_pedido_lista CPL ON CPL.IDPedido = ECR.IDPedido and CPL.IDLista = ECR.IDLista and CPL.Desenho = ECR.Desenho where ECR.IDPedido = " & IIf(Txt_ID_pedido = "", 0, Txt_ID_pedido) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " and (ECR.Nota_fiscal IS NULL or ECR.Nota_fiscal = N'')"
        Conexao.Execute "Update EM set EM.Documento = '" & txtnotafiscal & "' from Estoque_movimentacao EM INNER JOIN Estoque_controle_recebimento ECR ON ECR.Id = EM.IDEstoque_recebimento INNER JOIN Compras_pedido_lista CPL ON CPL.IDPedido = ECR.IDPedido and CPL.IDLista = ECR.IDLista and CPL.Desenho = ECR.Desenho where ECR.IDPedido = " & IIf(Txt_ID_pedido = "", 0, Txt_ID_pedido) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " and ECR.Nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Data_emissao = '" & Format(txtdataemissao, "Short Date") & "'"
    End If
End If
   
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal where ID_empresa = " & txtID_empresa & " and Id_Int_Cliente = " & Txt_ID_forn & " and int_NotaFiscal = '" & txtnotafiscal & "' and Serie = '" & txtSerie & "' and int_TipoNota = 2 and TipoNF = '" & TipoNF & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select CF.*, CP.ID_empresa, CC.Moeda FROM (Compras_fornecedores CF INNER JOIN Compras_pedido CP ON CF.IDCliente = CP.idfornecedor) LEFT JOIN Compras_comercial CC ON CC.IDpedido = CP.IDpedido where CP.Pedido = '" & txtProg_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        
        Moeda = "REAL"
        ValorMoeda = 1
        If IsNull(TBFornecedor!Moeda) = False And TBFornecedor!Moeda <> "" And TBFornecedor!Moeda <> "REAL" Then
            Moeda = TBFornecedor!Moeda
Mensagem:
            Dolar = InputBox("Favor informar o valor do " & Moeda & ".")
            If Dolar = "" Then Exit Sub
            If IsNumeric(Dolar) = False Then
                USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            ValorMoeda = Dolar
        End If
        
        TBGravar.AddNew
        TBGravar!TabelaSN = 0
        TBGravar!Regime = FunVerifRegimeEmpresa(TBFornecedor!ID_empresa)
        TBGravar!pedido_interno = False
        TBGravar!DtValidacaoOF = Now
        TBGravar!RespValidacaoOF = pubUsuario
        TBGravar!int_NotaFiscal = txtnotafiscal
        TBGravar!Serie = txtSerie
        If txtuf = "EX" Then TBGravar!Aplicacao = "P" Else TBGravar!Aplicacao = "T"
        TBGravar!int_TipoNota = "2"
        TBGravar!dt_DataEmissao = txtdataemissao
        TBGravar!txt_Hora_Saida = Format(Now, "hh:mm:ss")
        TBGravar!TipoNF = TipoNF
        TBGravar!txt_Razao_Nome = txtfornecedor
        TBGravar!Moeda = Moeda
        TBGravar!ValorMoeda = ValorMoeda
        TBGravar!ID_empresa = TBFornecedor!ID_empresa
        TBGravar!Id_Int_Cliente = TBFornecedor!IDCliente
        TBGravar!txt_Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
        TBGravar!Numero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
        TBGravar!txt_Bairro = IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
        TBGravar!Txt_CEP = IIf(IsNull(TBFornecedor!CEP), "", TBFornecedor!CEP)
        TBGravar!txt_Municipio = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
        TBGravar!txt_Fone_Fax = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
        
        TBGravar!txt_UF = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
        UF = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
        
        If TBFornecedor!idTipoEmpresa = 1 Then TBGravar!txt_CNPJ_CPF = IIf(IsNull(TBFornecedor!CPF_CNPJ), "", TBFornecedor!CPF_CNPJ)
        TBGravar!txt_IE_Cliente = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
        If TBFornecedor!Pessoa = "JURÍDICA" Then TBGravar!txt_tipocliente = "J" Else TBGravar!txt_tipocliente = "F"
    End If
    TBGravar!Int_status = "1"
    TBGravar.Update
    ID_nota = TBGravar!ID
Else
    ID_nota = TBGravar!ID
    ValorMoeda = TBGravar!ValorMoeda
    
    'Verifica se a NF já foi validada e não permite alteração
    If IsNull(TBGravar!DtValidacao) = False Then
        USMsgBox ("Esta nota fiscal não será alterada, pois a mesma já foi validada."), vbInformation, "CAPRIND v5.0"
        TBGravar.Close
        GoTo Validada
    End If
End If
TBGravar.Close

'Cria ou altera os produtos
Desenho = ""
OrdemTexto = ""
valor = 0
ValorTotal = 0
OF = 0
NovoValor = ""
Set TBReceber = CreateObject("adodb.recordset")
TBReceber.Open "Select ECR.*, CP.idfornecedor, CP.pedido from (Estoque_Controle_recebimento ECR INNER JOIN Compras_pedido CP ON ECR.idpedido = CP.idpedido) INNER JOIN Compras_fornecedores CF ON CF.IDCliente = CP.idfornecedor where CP.Pedido = '" & txtProg_pedido & "' and  CP.idfornecedor = " & Txt_ID_forn & " and ECR.nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " order by ECR.Desenho", Conexao, adOpenKeyset, adLockOptimistic
If TBReceber.EOF = False Then
    Do While TBReceber.EOF = False
        If TBReceber!Desenho <> Desenho Then OF = 0

        Set TBPedido = CreateObject("adodb.recordset")
        TBPedido.Open "Select * from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBPedido.EOF = False Then
            If TBPedido!Tipo = "P" Then
                Prod = True
                Prodpedido = True
                ServPedido = False
            Else
                Prodpedido = False
                ServPedido = True
            End If
            If TBPedido!Tipo = "S" Then Servicos = True
            If Desenho <> TBReceber!Desenho Or Desenho = TBReceber!Desenho And valor <> IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto) Or OrdemTexto <> IIf(IsNull(TBPedido!Ordem), 0, TBPedido!Ordem) Then
                ValorTotal = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
                OF = IIf(IsNull(TBPedido!Ordem), 0, TBPedido!Ordem)
                GoTo Prosseguir
            Else
                GoTo Proximo
            End If
        End If
        TBPedido.Close
Prosseguir:
        If OF = 0 Then TextoFiltro = "(Ordem = 0 or Ordem is null)" Else TextoFiltro = "Ordem = '" & OF & "'"
        
        qt = 0
        Set TBPedido = CreateObject("adodb.recordset")
        TBPedido.Open "Select * from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBPedido.EOF = False Then
            NovoValor1 = Replace(ValorTotal, ",", ".")
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Sum(ECR.recebido) as qt from compras_pedido_lista CPL inner join estoque_controle_recebimento ECR on CPL.idlista = ECR.idlista where CPL.preco_unitario_desconto = " & IIf(NovoValor1 = "", 0, NovoValor1) & " and " & TextoFiltro & " and ECR.nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Desenho = '" & TBReceber!Desenho & "' and ECR.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                qt = Format(IIf(IsNull(TBFI!qt), 0, TBFI!qt), "###,##0.0000")
                If TBPedido!Un <> TBPedido!Unidade_com Then
                    If FunVerifUNConversao(TBPedido!Un, TBPedido!Unidade_com) = True Then
                        qt = FunConverteUN(TBPedido!Unidade_com, TBPedido!Un, qt, TBReceber!Desenho)
                    Else
                        qt = qt * FunVerificaTabelaConversaoUnidade(TBPedido!Un, TBPedido!Unidade_com)
                    End If
                End If
            End If
            TBFI.Close
        End If
        TBPedido.Close
        
        ValorTotal = Format(ValorTotal * ValorMoeda, "###,##0.0000000000")
        NovoValor = Replace(ValorTotal, ",", ".")
        If Prodpedido = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Detalhes_Nota where int_Cod_Produto = '" & TBReceber!Desenho & "' and id_nota = " & ID_nota & " and dbl_valorunitario = " & IIf(NovoValor = "", 0, NovoValor) & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then TBAbrir.AddNew
            TBAbrir!Tipo = "P"
            TBAbrir!int_Cod_Produto = TBReceber!Desenho
            TBAbrir!int_Qtd = qt
            TBAbrir!Saldo = qt
            TBAbrir!int_NotaFiscal = txtnotafiscal
            TBAbrir!ID_nota = ID_nota
            
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select * from compras_pedido_lista where idpedido = " & TBReceber!IDpedido & " and IDLista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                TBAbrir!Txt_descricao = IIf(IsNull(TBPedido!Descricao), "", TBPedido!Descricao)
                TBAbrir!Codproduto = IIf(IsNull(TBPedido!Codproduto), "", TBPedido!Codproduto)
                IDlista = IIf(IsNull(TBPedido!IDlista), "", TBPedido!IDlista)
                TBAbrir!txt_Unid = IIf(IsNull(TBPedido!Un), "", TBPedido!Un)
                TBAbrir!Unidade_com = IIf(IsNull(TBPedido!Unidade_com), "", TBPedido!Unidade_com)
                TBAbrir!Familia = IIf(IsNull(TBPedido!Familia), "", TBPedido!Familia)
                TBAbrir!N_referencia = IIf(IsNull(TBPedido!N_referencia), "", TBPedido!N_referencia)
                TBAbrir!Ordem = TBPedido!Ordem
                If TBPedido!Remessa = True Then TBAbrir!retorno = True
                
                If IsNull(TBPedido!ID_CFOP) = False And TBPedido!ID_CFOP <> "" Then TBAbrir!ID_CFOP = TBPedido!ID_CFOP
                If IsNull(TBPedido!ID_CF) = False And TBPedido!ID_CF <> "" Then TBAbrir!ID_CF = TBPedido!ID_CF
                If IsNull(TBPedido!CST) = False And TBPedido!CST <> "" Then TBAbrir!txt_CST = TBPedido!CST
            End If
            TBPedido.Close
            
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select ID_CFOP, ID_CF from projproduto where desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                If IsNull(TBAbrir!ID_CFOP) = True Or TBAbrir!ID_CFOP = "" Then TBAbrir!ID_CFOP = IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP)
                If IsNull(TBAbrir!ID_CF) = True Or TBAbrir!ID_CF = "" Then TBAbrir!ID_CF = IIf(IsNull(TBItem!ID_CF), 0, TBItem!ID_CF)
            End If
            
            If IsNull(TBAbrir!ID_CFOP) = False And TBAbrir!ID_CFOP <> "" Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from tbl_NaturezaOperacao_CST where ID_CFOP = " & TBAbrir!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    If TBItem.RecordCount = 1 Then
                        If IsNull(TBAbrir!txt_CST) = True Or TBAbrir!txt_CST = "" Then TBAbrir!txt_CST = TBItem!CST_ICMS
                        TBAbrir!CST_IPI = TBItem!CST_IPI
                        TBAbrir!CST_PIS = TBItem!CST_PIS
                        TBAbrir!CST_Cofins = TBItem!CST_Cofins
                    End If
                End If
            End If
            TBItem.Close
            
            Set TBPI_Lista_produto = CreateObject("adodb.recordset")
            TBPI_Lista_produto.Open "Select CPL.* from compras_pedido_lista CPL INNER JOIN compras_pedido CP ON CPL.idpedido = CP.idpedido where CPL.idlista = " & TBReceber!IDlista & " and CP.idpedido = " & TBReceber!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBPI_Lista_produto.EOF = False Then
                TBAbrir!dbl_ValorUnitario = Format(TBPI_Lista_produto!preco_unitario_desconto * ValorMoeda, "###,##0.0000000000")
                TBAbrir!int_ICMS = IIf(IsNull(TBPI_Lista_produto!ICMS), 0, TBPI_Lista_produto!ICMS)
                TBAbrir!int_IPI = IIf(IsNull(TBPI_Lista_produto!IPI), 0, TBPI_Lista_produto!IPI)
                TBAbrir!dbl_valoripi = Format(((TBAbrir!dbl_ValorUnitario * qt) * IIf(IsNull(TBPI_Lista_produto!IPI), 0, TBPI_Lista_produto!IPI)) / 100, "###,##0.00")
                TBAbrir!dbl_ValorTotal = Format(TBAbrir!dbl_ValorUnitario * qt, "###,##0.00")
                TBAbrir!Valor_frete = Format(IIf(IsNull(TBPI_Lista_produto!Frete), 0, TBPI_Lista_produto!Frete) * ValorMoeda, "###,##0.00")
                TBAbrir!Valor_seguro = Format(IIf(IsNull(TBPI_Lista_produto!Seguro), 0, TBPI_Lista_produto!Seguro) * ValorMoeda, "###,##0.00")
                TBAbrir!Valor_acessorias = Format(IIf(IsNull(TBPI_Lista_produto!Acessorias), 0, TBPI_Lista_produto!Acessorias) * ValorMoeda, "###,##0.00")
                TBAbrir!Tem_IPI_frete = TBPI_Lista_produto!Frete_IPI
                If IsNull(TBPI_Lista_produto!OS) = False And TBPI_Lista_produto!OS <> "" Then ProcAtualizaCTTEROrdem TBPI_Lista_produto!OS
            End If
            TBPI_Lista_produto.Close
            TBAbrir.Update
            'Salvar CST
            'ProcSalvarCSTLista
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Detalhes_Nota where int_Cod_Produto = '" & TBReceber!Desenho & "' and id_nota = " & ID_nota & " and dbl_ValorUnitario = " & IIf(NovoValor = "", 0, NovoValor) & " and Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then TBAbrir.AddNew
            TBAbrir!Tipo = "S"
            TBAbrir!int_Cod_Produto = TBReceber!Desenho
            TBAbrir!int_Qtd = qt
            TBAbrir!int_NotaFiscal = txtnotafiscal
            TBAbrir!ID_nota = ID_nota
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select * from compras_pedido_lista where idpedido = " & TBReceber!IDpedido & " and IDLista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                TBAbrir!Txt_descricao = IIf(IsNull(TBPedido!Descricao), "", TBPedido!Descricao)
                TBAbrir!Codproduto = IIf(IsNull(TBPedido!Codproduto), "", TBPedido!Codproduto)
                IDlista = IIf(IsNull(TBPedido!IDlista), "", TBPedido!IDlista)
                TBAbrir!txt_Unid = IIf(IsNull(TBPedido!Un), "", TBPedido!Un)
                TBAbrir!Unidade_com = IIf(IsNull(TBPedido!Unidade_com), "", TBPedido!Unidade_com)
                TBAbrir!Familia = IIf(IsNull(TBPedido!Familia), "", TBPedido!Familia)
                TBAbrir!N_referencia = IIf(IsNull(TBPedido!N_referencia), "", TBPedido!N_referencia)
                TBAbrir!Ordem = TBPedido!Ordem
                
                If IsNull(TBPedido!ID_CFOP) = False And TBPedido!ID_CFOP <> "" Then TBAbrir!ID_CFOP = TBPedido!ID_CFOP
            End If
            TBPedido.Close
            
            If IsNull(TBAbrir!ID_CFOP) = True Or TBAbrir!ID_CFOP = "" Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from projproduto where desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    TBAbrir!ID_CFOP = IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP)
                End If
            End If
            
            Set TBPI_Lista_produto = CreateObject("adodb.recordset")
            TBPI_Lista_produto.Open "Select CPL.* from compras_pedido_lista CPL INNER JOIN compras_pedido CP ON CPL.idpedido = CP.idpedido where CPL.idlista = " & TBReceber!IDlista & " and CP.idpedido = " & TBReceber!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBPI_Lista_produto.EOF = False Then
                TBAbrir!dbl_ValorUnitario = Format(TBPI_Lista_produto!preco_unitario_desconto * ValorMoeda, "###,##0.0000000000")
                TBAbrir!ISS = IIf(IsNull(TBPI_Lista_produto!ISSQN), 0, TBPI_Lista_produto!ISSQN)
                TBAbrir!VlrISS = Format(((TBAbrir!dbl_ValorUnitario * qt) * IIf(IsNull(TBPI_Lista_produto!ISSQN), 0, TBPI_Lista_produto!ISSQN)) / 100, "###,##0.00")
                TBAbrir!dbl_ValorTotal = Format(TBAbrir!dbl_ValorUnitario * qt, "###,##0.00")
                
                If IsNull(TBPI_Lista_produto!OS) = False And TBPI_Lista_produto!OS <> "" Then ProcAtualizaCTTEROrdem TBPI_Lista_produto!OS
            End If
            TBPI_Lista_produto.Close
            TBAbrir.Update
        End If
                    
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select ECR.IDlista, ECR.Recebido from Estoque_Controle_recebimento ECR inner join Compras_pedido CP on ECR.idpedido = CP.idpedido where CP.idfornecedor = " & TBReceber!IDfornecedor & " and ECR.nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " and ECR.Desenho = '" & TBAbrir!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Do While TBFIltro.EOF = False
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "Select idlista from compras_pedido_lista where idlista = " & TBFIltro!IDlista & " and preco_unitario_desconto = " & IIf(NovoValor1 = "", 0, NovoValor1) & " and (Ordem = " & OF & " or Ordem IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
                If TBPedido.EOF = False Then
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * FROM tbl_Detalhes_Nota_pedidos where ID_nota = " & ID_nota & " and ID_prod_NF = " & TBAbrir!Int_codigo & " and ID_carteira = " & TBFIltro!IDlista & " and Codinterno = '" & TBAbrir!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then TBGravar.AddNew
                    TBGravar!ID_nota = ID_nota
                    TBGravar!ID_prod_NF = TBAbrir!Int_codigo
                    TBGravar!ID_carteira = TBFIltro!IDlista
                    TBGravar!Codinterno = TBAbrir!int_Cod_Produto
                    TBGravar!quantidade = TBFIltro!Recebido
                    TBGravar.Update
                    TBGravar.Close
                End If
                TBPedido.Close
                TBFIltro.MoveNext
            Loop
        End If
        TBFIltro.Close
        
        TBAbrir.Close
Proximo:
        Set TBPedido = CreateObject("adodb.recordset")
        TBPedido.Open "Select * from tbl_proposta_nota where id_nota = " & ID_nota & " and proposta = '" & TBReceber!Pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBPedido.EOF = True Then
            TBPedido.AddNew
            TBPedido!Proposta = TBReceber!Pedido
            TBPedido!NF = txtnotafiscal
            TBPedido!ID_nota = ID_nota
            TBPedido.Update
        End If

        Set TBPedido = CreateObject("adodb.recordset")
        TBPedido.Open "Select preco_unitario_desconto, Ordem from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBPedido.EOF = False Then
            valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
            OrdemTexto = IIf(IsNull(TBPedido!Ordem), "", TBPedido!Ordem)
        End If
        TBPedido.Close
        Desenho = TBReceber!Desenho
        TBReceber.MoveNext
    Loop
Else
    USMsgBox ("Não há produto recebido para a nota " & txtnotafiscal & "."), vbExclamation, "CAPRIND v5.0"
    TBFornecedor.Close
    Exit Sub
End If
If Prod = True And Servicos = True Then
    TipoNF = "M1SA"
ElseIf Prod = True And Servicos = False Then
        TipoNF = "M1"
    Else
        TipoNF = "SA"
End If
Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set TipoNF = '" & TipoNF & "' where ID = " & ID_nota

Validada:
    If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then Unload frmFaturamento_Prod_Serv
    If txtuf = "EX" Then
        Faturamento_NF_Saida = True
        Formulario = "Faturamento/Nota fiscal/Própria"
    Else
        Faturamento_NF_Saida = False
        Formulario = "Estoque/Nota fiscal"
    End If
    With frmFaturamento_Prod_Serv
        .Novo_Nota = False
        .Faturamento_Vendas_PI = False
        .TxtID.Text = ID_nota
        .txtNFiscal.Text = txtnotafiscal
        .ProcCarregaDadosNota .TxtID.Text
        .ProcCarregaLista
        .ProcCarregaListaServicos
        .ProcGravarTotaisNota
        .ProcCarregaDadosTransporte
        .ProcCarregaDuplicatas
        .Show
        .txt_DtEmissao.Value = Format(txtdataemissao, "dd/mm/yyyy")
        .txtSerie.Locked = False
        .txtSerie.TabStop = True
        
        CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.dt_Saida_Entrada, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, TN.dbl_Valor_Total_Nota, NF.txt_Razao_Nome, NF.Int_status, NF.Imprimir, NF.ID_empresa, NF.Aplicacao, NF.DtValidacaoOF, NF.DtValidacao"
        .Strsql_Faturamento = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .TxtID
        .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .TxtID & " and NF.Int_status = 1"
        .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .TxtID & " and NF.Int_status = 2"
        .Strsql_FaturamentoNFe = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.TipoNF <> 'SA' AND NF.ID = " & .TxtID
        .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.TipoNF = 'SA' AND NF.ID = " & .TxtID
        .ProcCarregaListaNota (1)
        
        If USMsgBox("Deseja prosseguir com o preenchimento dos dados da nota fiscal?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Unload frmFaturamento_Prod_Serv
    End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'Private Sub ProcSalvarCSTLista()
'On Error GoTo tratar_erro
'
'CST_ICMS = False
'CST_IPI = False
'CST_PIS = False
'CST_Cofins = False
'
''ICMS
'If IsNull(TBAbrir!txt_CST) = False And TBAbrir!txt_CST <> "" Then
'    InicioCST = Left(TBAbrir!txt_CST, 1)
'    If Len(TBAbrir!txt_CST) = 4 Then FimCST = Right(TBAbrir!txt_CST, 3) Else FimCST = Right(TBAbrir!txt_CST, 2)
'    CST_ICMS = True
'    CST_Cofins = False
'    CST_IPI = False
'    CST_PIS = False
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
''IPI
'If IsNull(TBAbrir!CST_IPI) = False And TBAbrir!CST_IPI <> "" Then
'    FimCST = TBAbrir!CST_IPI
'    CST_ICMS = False
'    CST_Cofins = False
'    CST_IPI = True
'    CST_PIS = False
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_IPI where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
''PIS
'If IsNull(TBAbrir!CST_PIS) = False And TBAbrir!CST_PIS <> "" Then
'    FimCST = TBAbrir!CST_PIS
'    CST_ICMS = False
'    CST_Cofins = False
'    CST_IPI = False
'    CST_PIS = True
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_PIS where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
''Cofins
'If IsNull(TBAbrir!CST_Cofins) = False And TBAbrir!CST_Cofins <> "" Then
'    FimCST = TBAbrir!CST_Cofins
'    CST_ICMS = False
'    CST_Cofins = True
'    CST_IPI = False
'    CST_PIS = False
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_Cofins where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
'
'Private Sub ProcEnviadadosListaCST()
'On Error GoTo tratar_erro
'
'If TBCST.EOF = True Then TBCST.AddNew
''ICMS
'If CST_ICMS = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Origem_mercadoria = InicioCST
'    TBCST!Tributacao_ICMS = FimCST
'    If FimCST <> "40" And FimCST <> "41" And FimCST <> "50" And FimCST <> "60" And FimCST <> "101" And FimCST <> "102" And FimCST <> "103" And FimCST <> "300" And FimCST <> "400" And FimCST <> "500" Then
'        If FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
'            If FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
'                Set TBMaquinas = CreateObject("adodb.recordset")
'                TBMaquinas.Open "Select * from regioes where uf = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
'                If TBMaquinas.EOF = False Then
'                    Set TBAfericao = CreateObject("adodb.recordset")
'                    TBAfericao.Open "Select * from tbl_ClassificacaoFiscal where txt_Ref = '" & TBAbrir!Txt_CF & "'", Conexao, adOpenKeyset, adLockOptimistic
'                    If TBAfericao.EOF = False Then
'                        Select Case TBMaquinas!regiao
'                            Case "DE":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTDE
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTDE
'                            Case "SS":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTSS
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTSS
'                            Case "NN":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTNN
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTNN
'                            Case "CO":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTCO
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTCO
'                        End Select
'                    End If
'                    TBAfericao.Close
'                End If
'            End If
'        End If
'
'        If FimCST <> "201" And FimCST <> "202" And FimCST <> "203" Then
'            IntICMS = IIf(IsNull(TBAbrir!int_ICMS), 0, TBAbrir!int_ICMS)
'            If IntICMS <> 0 Then
'                TBCST!Valor_BC = Format(TBAbrir!dbl_ValorTotal, "###,##0.00")
'                Valortotal = TBCST!Valor_BC
'                TBCST!Valor_ICMS = Format((Valortotal * IntICMS) / 100, "###,##0.00")
'            Else
'                TBCST!Valor_BC = 0
'                TBCST!Valor_ICMS = 0
'            End If
'        End If
'    End If
'
'    If FimCST = "101" Or FimCST = "201" Or FimCST = "900" Then
'        TBCST!ICMS_SN = 0
'        TBCST!Valor_ICMS_SN = 0
'
'        IntICMS = IIf(IsNull(TBAbrir!ICMS_SN), 0, TBAbrir!ICMS_SN)
'        If IntICMS <> 0 Then
'            Valortotal = Format(TBAbrir!dbl_ValorTotal, "###,##0.00")
'            TBCST!ICMS_SN = IntICMS
'            TBCST!Valor_ICMS_SN = Format((Valortotal * IntICMS) / 100, "###,##0.00")
'        End If
'    End If
'End If
''IPI
'If CST_IPI = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Codigo_situacaoTributaria = FimCST
'    If FimCST = "00" Or FimCST = "49" Or FimCST = "50" Or FimCST = "99" Then TBCST!Valor_BC = TBAbrir!dbl_ValorTotal
'End If
''PIS
'If CST_PIS = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Codigo_situacaoTributaria = FimCST
'    If FimCST = "01" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBCST!Valor_BC = TBAbrir!dbl_ValorTotal
'End If
''Cofins
'If CST_Cofins = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Codigo_situacaoTributaria = FimCST
'    If FimCST = "01" Or FimCST = "02" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBCST!Valor_BC = TBAbrir!dbl_ValorTotal
'End If
'TBCST.Update
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage <> 2 Then
    If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.PageCount - 1)
    Else
        TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = TBLISTA_Estoque_RecebimentoPedido.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)
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
Quant = DS_RetornarNumeros(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = 1
ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage <> -3 Then
    If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = TBLISTA_Estoque_RecebimentoPedido.PageCount
ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

ProcLimpar
listprod.ListItems.Clear
ProcLimparCamposNF
ProcLimparCamposReq False
Lista_movimentacao.ListItems.Clear
ProcCarregaPedido
ProcBloqueiaFrame
ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCamposNF()
On Error GoTo tratar_erro

txtdataemissao = "__/__/____"
txtnotafiscal = ""
txtSerie = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

Permitido = False
Permitido1 = False
With listprod
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido1 = False Then
                If USMsgBox("Deseja realmente alterar o status do(s) produto(s)/serviço(s)?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            If Permitido1 = False Then
                frmEstoque_Recebimento_aut.Show 1
            End If
            If Permitido = False Then Exit Sub

            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from compras_pedido where pedido = '" & .ListItems.Item(InitFor).SubItems(2) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                Txt_ID_pedido = TBCompras_Pedido!IDpedido
                IDempresa = TBCompras_Pedido!ID_empresa
            End If
            TBCompras_Pedido.Close
            
            If .ListItems.Item(InitFor).SubItems(10) <> "RECEBIDO" Then
                status = "Status_item = 'RECEBIDO'"
            Else
                'Verifica se o status do produto esta aberto, parcial ou recebido
                quantnovo = 0
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select Sum(Recebido) as quantnovo from estoque_controle_recebimento where idpedido = " & Txt_ID_pedido & " and idlista = " & .ListItems.Item(InitFor) & " and Programacao = 'False' and id_empresa = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    quantnovo = IIf(IsNull(TBEstoque!quantnovo), 0, TBEstoque!quantnovo)
                End If
                TBEstoque.Close
                
                quantestoque = .ListItems.Item(InitFor).ListSubItems(7)
                If quantnovo = 0 Then
                    If FunVerifStatusAprovadoPC(txtID_empresa) = True Then status = "Status_item = 'APROVADO'" Else status = "Status_item = 'N_RECEBIDO'"
                ElseIf quantnovo < quantestoque Then
                        status = "Status_item = 'PARCIAL'"
                    Else
                        If quantnovo >= quantestoque Then status = "Status_item = 'RECEBIDO'"
                End If
            End If
            
            Conexao.Execute "Update compras_pedido_lista Set " & status & " where idpedido = " & Txt_ID_pedido & " and idlista = " & .ListItems.Item(InitFor)
            'Verifica status do item
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and (status_item = 'RECEBIDO' or status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = True Then
                If FunVerifStatusAprovadoPC(txtID_empresa) = True Then Status_pedido = "APROVADO" Else Status_pedido = "ABERTO"
            Else
                Set TBCompras = CreateObject("adodb.recordset")
                TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and (status_item = 'N_RECEBIDO' or status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                If TBCompras.EOF = True Then
                    Status_pedido = "ENCERRADO"
                Else
                    Status_pedido = "PARCIAL"
                End If
            End If
            TBCompras.Close
                
            'Grava status do produto na ordem de compra se pedido em Aberto = False
            Conexao.Execute "Update compras_pedido Set Status_pedido = '" & Status_pedido & "' where IDpedido = " & Txt_ID_pedido
            '==================================
            Modulo = "Estoque/Recebimento/Pedido de compra"
            Evento = "Alterar o status"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "N° pedido: " & .ListItems.Item(InitFor).SubItems(2) & " - Cód. interno: " & .ListItems.Item(InitFor).SubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Permitido1 = True
        End If
    Next InitFor
End With
If Permitido1 = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) na lista antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparCamposReq False
    ProcCarregaListaFiltro
    cmdReceber.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Estoque/Recebimento/Pedido de compra"
Direitos
ProcLimpaVariaveisPrincipais
Imprimir = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362ER" Then
    If USMsgBox("Deseja realmente atualizar os dados dos recebimentos?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Estoque_controle_recebimento order by IDPedido, IDLista", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBAbrir.EOF = False
                If IsNull(TBEstoque!local_armaz) = True Or TBEstoque!local_armaz = "" Then TBEstoque!local_armaz = "N/A"
                
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from Compras_programa_item where ID = " & TBAbrir!IDpedido & " and ID_item = " & TBAbrir!IDlista, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then TBAbrir!Programacao = True Else TBAbrir!Programacao = False
                TBAbrir.Update
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBAbrir.Close
        
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "select * from estoque_movimentacao where operacao = 'ENTRADA_NOTA_FISCAL'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBGravar.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBGravar.EOF = False
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "select * from estoque_controle where idestoque = " & TBGravar!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    'verifica se é pedido de compra
                    Set TBPedido = CreateObject("adodb.recordset")
                    TBPedido.Open "select * from compras_pedido where Pedido = '" & TBGravar!LOTE & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBPedido.EOF = False Then
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "select estoque_controle_recebimento.Id,estoque_controle_recebimento.Idlista from estoque_controle_recebimento inner join compras_pedido on estoque_controle_recebimento.IDPedido = compras_pedido.IDPedido where estoque_controle_recebimento.Desenho = '" & TBEstoque!Desenho & "' and estoque_controle_recebimento.certificado = '" & TBEstoque!Certificado & "' and estoque_controle_recebimento.corrida = '" & TBEstoque!Corrida & "' and estoque_controle_recebimento.Local_armaz = '" & TBEstoque!local_armaz & "' and compras_pedido.Pedido = '" & TBEstoque!LOTE & "' and estoque_controle_recebimento.Nota_fiscal = '" & TBGravar!Documento & "' and estoque_controle_recebimento.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            TBGravar!IDEstoque_recebimento = TBItem!ID
                            TBGravar!idlista_recebimento = TBItem!IDlista
                            TBGravar.Update
                        End If
                        TBItem.Close
                    End If
                    TBPedido.Close
                    'verifica se é programação de compra
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "select * from Compras_programa where Programatexto = '" & TBGravar!LOTE & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "select estoque_controle_recebimento.Id from estoque_controle_recebimento inner join Compras_programa on estoque_controle_recebimento.IDPedido = Compras_programa.ID where estoque_controle_recebimento.Desenho = '" & TBEstoque!Desenho & "' and estoque_controle_recebimento.certificado = '" & TBEstoque!Certificado & "' and estoque_controle_recebimento.corrida = '" & TBEstoque!Corrida & "' and estoque_controle_recebimento.Local_armaz = '" & TBEstoque!local_armaz & "' and Compras_programa.Programatexto = '" & TBEstoque!LOTE & "' and estoque_controle_recebimento.Nota_fiscal = '" & TBGravar!Documento & "' and estoque_controle_recebimento.ID_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            TBGravar!IDEstoque_recebimento = TBItem!ID
                            TBGravar.Update
                        End If
                        TBItem.Close
                    End If
                    TBAbrir.Close
                End If
                TBEstoque.Close
                TBGravar.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBGravar.Close
        
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal where int_TipoNota = 2 and Serie <> 'Null' order by int_NotaFiscal", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBGravar.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBGravar.EOF = False
                If TBGravar!Serie <> "" Then Conexao.Execute "Update Estoque_controle_recebimento Set Serie = '" & TBGravar!Serie & "' where Nota_fiscal = '" & TBGravar!int_NotaFiscal & "' and id_empresa = " & txtID_empresa
                TBGravar.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBGravar.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Recebimento/Pedido de compra"
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

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = True
Sit_Data = 1
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_receb_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = True
Sit_Data = 2
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_movimentacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

Permitido1 = False
If ColumnHeader = "" Then
    With Lista_movimentacao
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifExcluirMovimentacao(.ListItems.Item(InitFor), .ListItems.Item(InitFor).SubItems(1), .ListItems.Item(InitFor).SubItems(9), .ListItems.Item(InitFor).SubItems(10), False) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_movimentacao, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifExcluirMovimentacao(IDoperacao As Long, IDestoque As Long, Inspecionado As String, IDestCR As Long, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifExcluirMovimentacao = True
'Verifica se houve saída
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select idoperacao from estoque_movimentacao where Idestoque = " & IDestoque & " and Saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois já houve movimentação de saída neste RE."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If

'Verifica se tem nota fiscal emitida
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select NF.ID from (Estoque_controle_recebimento ECR INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = ECR.Nota_fiscal and NF.Serie = ECR.Serie and NF.txt_Razao_Nome = '" & txtfornecedor & "' and NF.int_TipoNota = 2) INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_nota where ECR.ID = " & IDestCR & " and NFP.int_cod_produto = '" & txtcodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo está sendo utilizado no módulo Estoque/Nota fiscal."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If
TBAbrir.Close

If Inspecionado = "Sim" Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo já foi inspecionado."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    Exit Function
End If

'Verifica se esta amarrado alguma ordem
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select PNC.Ordem from estoque_movimentacao EM INNER JOIN Producao_NF_Consignada PNC ON EM.IDestoque = PNC.Idestoque where EM.idoperacao = " & IDoperacao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo está sendo vinculado a ordem " & TBAbrir!Ordem & "."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If

'Verifica se foi criado instrumento pelo RE
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select I.Codigo from estoque_movimentacao EM INNER JOIN Instrumentos I ON EM.IDestoque = I.Idestoque where EM.idoperacao = " & IDoperacao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo já foi cadastrado como instrumento."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Lista_movimentacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Permitido1 = False
With Lista_movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifExcluirMovimentacao(.ListItems.Item(InitFor), .ListItems.Item(InitFor).SubItems(1), .ListItems.Item(InitFor).SubItems(9), .ListItems.Item(InitFor).SubItems(10), True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With listprod
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView listprod, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listprod.ListItems.Count = 0 Then Exit Sub
ProcLimpar
ProcLimparCamposReq True
ProcHabilitaFrame

With listprod
    txtProg_pedido.Text = .SelectedItem.ListSubItems(2)
    txtidlista = .SelectedItem
End With
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from compras_pedido_lista where idlista = " & txtidlista, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtcodigo = TBProduto!Desenho
    Cmb_codigo_ref = IIf(IsNull(TBProduto!N_referencia), "", TBProduto!N_referencia)
    txtEspecificacoes = TBProduto!Descricao
    txtUN = TBProduto!Un
    
    If TBProduto!Un <> TBProduto!Unidade_com And IsNull(TBProduto!Qtde_estoque) = False And TBProduto!Qtde_estoque <> 0 Then
        txtrequisitado = Format(TBProduto!Qtde_estoque, "###,##0.0000")
    Else
        txtrequisitado = Format(TBProduto!Quant_Comp, "###,##0.0000")
    End If
    txtrequisitado_PC = IIf(IsNull(TBProduto!Quant_Comp_PC), "", TBProduto!Quant_Comp_PC)
    
    If TBProduto!Status_Item = "N_RECEBIDO" Or TBProduto!Status_Item = "APROVADO" Then
        Statusitem = "NÃO_RECEBIDO"
    Else
        Statusitem = TBProduto!Status_Item
    End If
    txtstatus = Statusitem
    
    If Chk_Dt_rcbto.Value = 0 Then Proccarregalocarm
End If

ProcCarregaPedido

'Verifica status do produto
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from estoque_controle_recebimento where idlista = " & txtidlista & " and Programacao = 'False' and id_empresa = " & txtID_empresa & " order by Id desc", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto.RecordCount > 1 Then
        ProcLimparCamposNF
    Else
        txtdataemissao = IIf(IsNull(TBProduto!Data_emissao), "__/__/____", Format(TBProduto!Data_emissao, "dd/mm/yyyy"))
        txtnotafiscal = IIf(IsNull(TBProduto!Nota_fiscal), "", TBProduto!Nota_fiscal)
        txtSerie = IIf(IsNull(TBProduto!Serie), "", TBProduto!Serie)
    End If
    txtcorrida = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
    txtcertificado = IIf(IsNull(TBProduto!Certificado), "", TBProduto!Certificado)
    txtcorrida = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
    txtOBS = IIf(IsNull(TBProduto!Obs), "", TBProduto!Obs)
    If IsNull(TBProduto!local_armaz) = False And TBProduto!local_armaz <> "" Then
        NomeCampo = "o loca de armazenamento (" & TBProduto!local_armaz & ")"
        cmbLocal_armaz = TBProduto!local_armaz
    End If
End If
TBProduto.Close

1:
    'Carrega qtde. recebida e atualiza o saldo
    qtdeliberada = 0
    qtdeliberar = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(ISNULL(Recebido, 0)) as qtdeliberada, Sum(ISNULL(Recebido_PC, 0)) as qtdeliberar from estoque_controle_recebimento where idlista = " & txtidlista & " and Programacao = 'False' and id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
        qtdeliberar = IIf(IsNull(TBAbrir!qtdeliberar), 0, TBAbrir!qtdeliberar)
    End If
    TBAbrir.Close
    
    txtrecebida.Text = Format(qtdeliberada, "###,##0.0000")
    Qtde = txtrequisitado.Text
    qt = Qtde - qtdeliberada
    txtSaldo = IIf(qt < 0, 0, Format(qt, "###,##0.0000"))
    
    If txtrequisitado_PC <> "" Then
        txtrecebida_PC = qtdeliberar
        Qtde = txtrequisitado_PC
        qt = Qtde - qtdeliberar
        txtSaldo_PC = IIf(qt < 0, 0, qt)
    Else
        txtrecebida_PC = ""
        txtSaldo_PC = ""
    End If
    txtQuantidade = txtSaldo
    
    ProcCarregaListaMovimentacao
    Estoquereal = 0
    qtdeliberada = 0

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " desta movimentação, favor revisar."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAreceber_Click()
On Error GoTo tratar_erro

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optRecebidos_Click()
On Error GoTo tratar_erro

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opttodos_Click()
On Error GoTo tratar_erro

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaFiltro()
On Error GoTo tratar_erro

If txtProg_pedido.Text = "" Then
    USMsgBox ("Informe o número do pedido de compras antes de filtrar."), vbExclamation, "CAPRIND v5.0"
    txtProg_pedido.SetFocus
    Exit Sub
End If
If OptAreceber.Value = True Then
    Statusitem = "and (Status_Item = 'N_RECEBIDO' or Status_Item = 'APROVADO' or Status_Item = 'PARCIAL')"
    Statusitemrel = "and ({Estoque_recebimento_pedido.Status_Item} = 'N_RECEBIDO' or {Estoque_recebimento_pedido.Status_Item} = 'APROVADO' or {Estoque_recebimento_pedido.Status_Item} = 'PARCIAL')"
ElseIf optRecebidos.Value = True Then
        Statusitem = "and (Status_Item = 'PARCIAL' or Status_Item = 'RECEBIDO')"
        Statusitemrel = "and ({Estoque_recebimento_pedido.Status_Item} = 'PARCIAL' or {Estoque_recebimento_pedido.Status_Item} = 'RECEBIDO')"
    Else
        Statusitem = "and (Status_Item = 'N_RECEBIDO' or Status_Item = 'APROVADO' or Status_Item = 'PARCIAL' or Status_Item = 'RECEBIDO')"
        Statusitemrel = "and ({Estoque_recebimento_pedido.Status_Item} = 'N_RECEBIDO' or {Estoque_recebimento_pedido.Status_Item} = 'APROVADO' or {Estoque_recebimento_pedido.Status_Item} = 'PARCIAL' or {Estoque_recebimento_pedido.Status_Item} = 'RECEBIDO')"
End If
CamposFiltro = "IDlista, ID_empresa, Pedido, Desenho, Descricao, UN, Unidade_com, preco_unitario, Quant_Comp, Quant_Comp_PC, Prazo, Status_item, Ordem, Qtde_estoque"
StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido where Pedido = '" & txtProg_pedido.Text & "' " & Statusitem & " group by " & CamposFiltro & " order by Pedido, Desenho"
StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, IDlista FROM Estoque_recebimento_pedido where Pedido = '" & txtProg_pedido.Text & "' " & Statusitem & " group by IDlista, Data_recebimento, Nota_fiscal order by IDlista"
FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido.Pedido} = '" & txtProg_pedido.Text & "' " & Statusitemrel
ProcCarregaLista
ProcGravarDataFiltroRel Date, Date, False, 0, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNotaFiscal_LostFocus()
On Error GoTo tratar_erro

If txtnotafiscal <> "" Then txtnotafiscal = FunTamanhoTextoZeroEsq(DS_RetornarNumeros(txtnotafiscal), 9)

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

Private Sub txtProg_pedido_Change()
On Error GoTo tratar_erro

If txtdata <> "" Then listprod.ListItems.Clear
ProcLimpar
ProcLimparCamposReq False
Lista_movimentacao.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

StrSql_Estoque_Recebimento_Localizar = ""
frmEstoque_Recebimento_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcalc_peso_Click()
On Error GoTo tratar_erro

If txtcodigo = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & txtcodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = False
    Compras_Requisicao = False
    Compras_Cotacao = False
    Compras_Pedido = False
    Estoque_recebimento = True
    Vendas_Proposta = False
    Vendas_PI = False
    FrmCalculo_Peso.Show 1
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo tratar_erro
Dim Numero2     As Long 'OK
Dim ID_estoque  As Long 'OK

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido1 = False
With Lista_movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido1 = False Then
                If USMsgBox("Deseja realmente excluir esta(s) movimentação(ões) do produto " & txtcodigo.Text & "?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido1 = True
            quantidade = 0
            'Verifica registro na tabela estoque_movimentacao/estoque_controle
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select idestoque, entrada, Entrada_PC, VlrUnit, IDEstoque_recebimento from estoque_movimentacao where idoperacao = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                ID_estoque = TBAbrir!IDestoque
                qt = TBAbrir!Entrada
                Quant = IIf(IsNull(TBAbrir!Entrada_PC), 0, TBAbrir!Entrada_PC)
                valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
                
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select * from estoque_controle where idestoque = " & TBAbrir!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    'Altera qtde recebida na tabela estoque_controle_recebimento
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from compras_pedido where pedido = '" & txtProg_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then
                        NovoValor = Replace(qt, ",", ".")
                        If IsNull(TBAbrir!IDEstoque_recebimento) = False And TBAbrir!IDEstoque_recebimento <> 0 Then
                            TextoFiltro = "Id = " & TBAbrir!IDEstoque_recebimento
                        Else
                            TextoFiltro = "idpedido = " & TBCompras!IDpedido & " and idlista = " & txtidlista & " and Certificado = '" & txtcertificado & "' and Corrida = '" & txtcorrida & "' and local_armaz = '" & cmbLocal_armaz & "' and Programacao = 'False'"
                            If txtnotafiscal <> "" Then TextoFiltro = TextoFiltro & " and Nota_fiscal = '" & txtnotafiscal & "'"
                        End If
                        Set TBCompras_Lista = CreateObject("adodb.recordset")
                        TBCompras_Lista.Open "Select ID, Recebido, Recebido_PC, IDPedido, IDlista from estoque_controle_recebimento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCompras_Lista.EOF = False Then
                            TBCompras_Lista!Recebido = TBCompras_Lista!Recebido - qt
                            TBCompras_Lista!Recebido_PC = TBCompras_Lista!Recebido_PC - Quant
                            TBCompras_Lista.Update
                            
                            ProcAtualizaQtdeRecebEmp TBCompras_Lista!IDpedido, TBCompras_Lista!IDlista
                            
                            If TBCompras_Lista!Recebido <= 0 Then Conexao.Execute "DELETE from estoque_controle_recebimento where ID = " & TBCompras_Lista!ID
                        End If
                        TBCompras_Lista.Close
                        
                        'Verifica se existe algum produto já recebido para definir o status do item
                        Set TBCompras_Lista = CreateObject("adodb.recordset")
                        TBCompras_Lista.Open "Select ID from estoque_controle_recebimento where idlista = " & txtidlista & " and idpedido = " & TBCompras!IDpedido & " and Programacao = 'False' and id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCompras_Lista.EOF = False Then
                            Status_pedido = "PARCIAL"
                        Else
                            If FunVerifStatusAprovadoPC(txtID_empresa) = True Then Status_pedido = "APROVADO" Else Status_pedido = "N_RECEBIDO"
                        End If
                        TBCompras_Lista.Close
                        
                        'Altera status do item
                        Conexao.Execute "Update compras_pedido_lista Set Status_item = '" & Status_pedido & "' where idlista = " & txtidlista
                        Conexao.Execute "Update Compras_Programacao set Compras_Programacao.Status_prog = '" & Status_pedido & "' from Compras_Programacao INNER JOIN compras_pedido_lista ON Compras_Programacao.ID_prog = compras_pedido_lista.ID_programacao where compras_pedido_lista.idlista = " & txtidlista
                        
                        'Altera status do pedido
                        Set TBPedido = CreateObject("adodb.recordset")
                        TBPedido.Open "Select * from compras_pedido_lista where idpedido = " & TBCompras!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
                        If TBPedido.EOF = False Then
                            Do While TBPedido.EOF = False
                                If TBPedido!Status_Item = "PARCIAL" Or TBPedido!Status_Item = "RECEBIDO" Then
                                    TBCompras!Status_pedido = "PARCIAL"
                                    GoTo 1
                                Else
                                    If FunVerifStatusAprovadoPC(txtID_empresa) = True Then TBCompras!Status_pedido = "APROVADO" Else TBCompras!Status_pedido = "ABERTO"
                                End If
                                TBPedido.MoveNext
                            Loop
                        End If
                        TBPedido.Close
                        
                        'Atualiza qtde recebida e status na programação
                        ProcAtualizaQtdeRecebidaProg "Select Compras_Programacao.* from Compras_Programacao INNER JOIN Compras_pedido_lista ON Compras_Programacao.ID_prog = Compras_pedido_lista.ID_programacao where Compras_pedido_lista.IDlista = " & txtidlista & " and Compras_Programacao.qtderecebida <> 0 order by Compras_Programacao.id_prog desc", True
                        ProcAlteraStatus_prog
                    End If
1:
                    TBCompras.Update
                    TBCompras.Close
                    
                    'Altera a qtde recebida na tabela estoque_controle
                    Permitido = False
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where desenho = '" & txtcodigo & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Permitido = True
                    End If
                    
                    'Verifica se o produto do pedido é remessa
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select CPL.* from Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho where CPL.Idlista = " & txtidlista & " and CPL.Remessa = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then Permitido = False
                    
                    'Verifica se o produto do pedido é mão de obra e se é a ultima fase da ordem e a ordem não controla estoque automaticamente, se for entra no estoque
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select CPL.OS, CPL.Ordem from (Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho) LEFT JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = CPL.ID_CFOP where CPL.Idlista = " & txtidlista & " and CFOP.MaoObra = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockReadOnly
                    If TBProduto.EOF = False Then
                        Permitido = False
                        If IsNull(TBProduto!OS) = False And TBProduto!OS <> "" And IsNull(TBProduto!Ordem) = False And TBProduto!Ordem <> "" Then
                            Set TBOrdemServico = CreateObject("adodb.recordset")
                            TBOrdemServico.Open "Select OS.idproducao from OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.Ordem = " & TBProduto!Ordem & " and P.Entrar_estoque <> 'True' ORDER BY OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockReadOnly
                            If TBOrdemServico.EOF = False Then
                                TBOrdemServico.MoveLast
                                If TBOrdemServico!IDProducao = TBProduto!OS Then Permitido = True
                            End If
                            TBOrdemServico.Close
                        End If
                    End If
                    TBProduto.Close
                    
                    'Verifica se tem centro de custo amarrado ao produto, se tiver não controla estoque
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Compras_pedido_lista_custo.* from Compras_pedido_lista_custo INNER JOIN Compras_pedido ON Compras_pedido_lista_custo.IDPedido = Compras_pedido.IDPedido where Compras_pedido_lista_custo.IDLista = " & txtidlista.Text & " and Compras_pedido.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Permitido = False
                    End If
                    TBProduto.Close
                    
                    If Permitido = True Then
                        TBEstoque!estoque_real = TBEstoque!estoque_real - qt
                        TBEstoque!estoque_real_PC = TBEstoque!estoque_real_PC - Quant
                        TBEstoque!estoque_venda = TBEstoque!estoque_real
                        TBEstoque.Update
                    End If
                End If
                TBEstoque.Close
            End If
            
            Conexao.Execute "DELETE from estoque_movimentacao where idoperacao = " & .ListItems(InitFor)
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select Idoperacao from estoque_movimentacao where idestoque = " & ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = True Then
                Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & ID_estoque
                Conexao.Execute "DELETE from estoque_controle where idestoque = " & ID_estoque
            End If
            TBCompras_Lista.Close
            
            'Centro de custo
            Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & .ListItems(InitFor)
            
            Set TBNivel2 = CreateObject("adodb.recordset")
            TBNivel2.Open "Select sum(Saida) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtcodigo & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
            If TBNivel2.EOF = False Then
                Valor1 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
            End If
            TBNivel2.Close
            Set TBNivel2 = CreateObject("adodb.recordset")
            TBNivel2.Open "Select sum(entrada) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtcodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBNivel2.EOF = False Then
                Valor2 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
            End If
            TBNivel2.Close
            If Valor1 > Valor2 Then Conexao.Execute "UPDATE estoque_movimentacao set Terceiros = 'True' where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtcodigo & "' and destino = 'Terceiros'"
            
            '==================================
            Modulo = "Estoque/Recebimento/Pedido de compra"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & txtcodigo & " - Nº lote: " & txtProg_pedido & " - Nº corrida: " & txtcorrida & " - Nº certificado: " & txtcertificado & " - Local armaz.: " & cmbLocal_armaz
            Documento1 = "Operação: " & .ListItems(InitFor).SubItems(2) & " - Documento: " & .ListItems(InitFor).SubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido1 = False Then
    USMsgBox ("Informe a(s) movimentação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Movimentação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
    ProcLimparCamposNF
    ProcLimparCamposReq False
    ProcCarregaListaMovimentacao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagamento_Click()
On Error GoTo tratar_erro

If txtProg_pedido.Text = "" And txtProg_pedido = "" Then Exit Sub
Aplic = 1
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = True
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdReceber_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
quantnovo = 0
quantestoque = 0
qt = 0
valor = 0
ValorTotal = 0
IDestoque = 0
Acao = "receber no estoque"
If txtProg_pedido = "" Then
    NomeCampo = "o número do pedido/programa"
    ProcVerificaAcao
    txtProg_pedido.SetFocus
    Exit Sub
End If
If txtidlista.Text = "" Or txtcodigo.Text = "" Then
    NomeCampo = "o produto/serviço"
    ProcVerificaAcao
    Exit Sub
End If
If txtstatus = "RECEBIDO" Then
    USMsgBox ("Este produto já foi recebido no estoque."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtuf <> "EX" Then
    If txtnotafiscal = "" Then
        If USMsgBox("O número da nota fiscal não foi informado, deseja prosseguir assim mesmo?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then
            txtnotafiscal.SetFocus
            Exit Sub
        End If
    Else
        If IsDate(txtdataemissao) = False Then
            NomeCampo = "a data de emissão da nota fiscal"
            ProcVerificaAcao
            txtdataemissao.SetFocus
            Exit Sub
        End If
        If txtSerie = "" Then
            NomeCampo = "o número de série da nota fiscal"
            ProcVerificaAcao
            txtSerie.SetFocus
            Exit Sub
        End If
    End If
End If

'Se for serviço, mão de obra ou remessa cria um local de armz padrão
Permitido = True
If cmbLocal_armaz = "SERVIÇOS" Or cmbLocal_armaz = "RETORNO DE MERCADORIA" Or cmbLocal_armaz = "INDUSTRIALIZAÇÃO" Then Permitido = False

TextoLocal = cmbLocal_armaz
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Remessa, Tipo, ID_CFOP from Compras_pedido_lista where IDlista = " & txtidlista, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!Tipo = "S" Then
        Permitido = False
        TextoLocal = "SERVIÇOS"
    ElseIf TBProduto!Remessa = True Then
        Permitido = False
        TextoLocal = "RETORNO DE MERCADORIA"
    ElseIf IsNull(TBProduto!ID_CFOP) = False And TBProduto!ID_CFOP <> "" And TBProduto!ID_CFOP <> "0" Then
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select IDCountCfop from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP & " and MaoObra = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            Permitido = False
            TextoLocal = "INDUSTRIALIZAÇÃO"
        End If
        TBCFOP.Close
    End If
End If
TBProduto.Close

If Permitido = True Then
    If txtcorrida = "" Then txtcorrida = 0
    If txtcertificado = "" Then txtcertificado = 0
    If cmbLocal_armaz = "" Then
        NomeCampo = "o local de armazenamento"
        ProcVerificaAcao
        cmbLocal_armaz.SetFocus
        Exit Sub
    End If
End If

If IsDate(Txt_data_recebimento) = False Then
    NomeCampo = "a data do recebimento"
    ProcVerificaAcao
    Txt_data_recebimento.SetFocus
    Exit Sub
End If
'pega quantidade recebida na caixa de texto
valor = IIf(txtQuantidade = "", 0, txtQuantidade)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQuantidade.SetFocus
    Exit Sub
End If

Valor_Cofins_Prod = IIf(txtQuantidade_PC = "", 0, txtQuantidade_PC)
If txtQuantidade_PC <> "" Then
    If Valor_Cofins_Prod - Int(Valor_Cofins_Prod) > 0 Then
        USMsgBox ("Só é permitido número inteiro na quantidade de peças a receber."), vbExclamation, "CAPRIND v5.0"
        txtQuantidade_PC.SetFocus
        Exit Sub
    End If
Else
    Valor_Cofins_Prod = FunCalculaQtdePC(txtcodigo, txtQuantidade, True, txtUN)
End If
Valor1 = IIf(txtrequisitado_PC = "", 0, txtrequisitado_PC)
If Valor1 > 0 And Valor_Cofins_Prod <= 0 Then
    NomeCampo = "a quantidade de peças"
    ProcVerificaAcao
    txtQuantidade_PC.SetFocus
    Exit Sub
End If

If txtnotafiscal <> "" Then txtnotafiscal = FunTamanhoTextoZeroEsq(DS_RetornarNumeros(txtnotafiscal), 9)
If txtProg_pedido.Text <> "" Then
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from compras_pedido where pedido = '" & txtProg_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        Txt_ID_pedido = TBCompras_Pedido!IDpedido
    End If
    
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from Compras_pedido_lista where IDpedido = " & Txt_ID_pedido & " and IDlista <> " & IDlista & " and Desenho = '" & txtcodigo & "' and Prazo < '" & listprod.SelectedItem.ListSubItems(9) & "' and Status_item <> 'RECEBIDO' and Status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        USMsgBox ("Não é permitido receber este produto/serviço, pois existe(m) outro(s) em aberto com o código interno " & txtcodigo & " e prazo de entrega menor que " & Format(listprod.SelectedItem.ListSubItems(9), "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
        TBCompras_Pedido.Close
        Exit Sub
    End If
    TBCompras_Pedido.Close
End If

'Verifica se o produto já foi recebido
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and idlista = " & txtidlista.Text & " and status_item = '" & "RECEBIDO" & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    USMsgBox ("Este produto/serviço já foi recebido no estoque " & vbCrLf & "Código interno: " & txtcodigo.Text & vbCrLf & " descrição : " & txtEspecificacoes & "."), vbExclamation, "CAPRIND v5.0"
    txtnotafiscal.SetFocus
    TBCompras.Close
    Exit Sub
End If
TBCompras.Close

'Verifica se o código de referencia está vinculado a outro produto
'If Cmb_codigo_ref <> "" Then If FunVerifiCodRefUtilizado(txtCodigo, Cmb_codigo_ref) = True Then Exit Sub

'If txtnotafiscal <> "" Then
'    If usMsgbox("Confirma o recebimento deste produto/serviço " & vbCrLf & "Cód. interno: " & txtcodigo & vbCrLf & "Descrição: " & listprod.SelectedItem.ListSubItems(4) & vbCrLf & "Dt. de recebimento: " & Txt_data_recebimento & vbCrLf & "Qtde. recebida: " & Format(txtQuantidade, "###,##0.0000") & vbCrLf & "Dt. emissão NF: " & txtdataemissao & vbCrLf & "Nota fiscal: " & txtnotafiscal & vbCrLf & "Série: " & txtSerie, vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
'Else
'    If usMsgbox("Confirma o recebimento deste produto/serviço " & vbCrLf & "Cód. interno: " & txtcodigo & vbCrLf & "Descrição: " & listprod.SelectedItem.ListSubItems(4) & vbCrLf & "Dt. de recebimento: " & Txt_data_recebimento & vbCrLf & "Qtde. recebida: " & Format(txtQuantidade, "###,##0.0000"), vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
'End If

'Grava movimentação na tabela estoque_controle
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!Desenho = txtcodigo.Text
TBEstoque!Ref = IIf(Cmb_codigo_ref = "", Null, Cmb_codigo_ref)
TBEstoque!Descricao = IIf(txtEspecificacoes.Text = "", Null, txtEspecificacoes.Text)
TBEstoque!LOTE = txtProg_pedido
TBEstoque!data = Txt_data_recebimento
TBEstoque!Responsavel = pubUsuario
TBEstoque!Certificado = IIf(txtcertificado = "", 0, txtcertificado)
TBEstoque!Corrida = IIf(txtcorrida = "", 0, txtcorrida)
TBEstoque!local_armaz = cmbLocal_armaz
TBEstoque!Fornecedor = txtfornecedor
TBEstoque!ID_empresa = txtID_empresa
TBEstoque!Un = txtUN.Text
TBEstoque!Numero_serie = Txt_numero_serie

Permitido = False
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtcodigo & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then Permitido = True

'Verifica se o produto do pedido é remessa
'Set TBProduto = CreateObject("adodb.recordset")
'TBProduto.Open "Select CPL.IDlista from Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho where CPL.Idlista = " & TXTIDLista & " and CPL.Remessa = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockOptimistic
'If TBProduto.EOF = False Then Permitido = False
'
''Verifica se o produto do pedido é mão de obra e se é a ultima fase da ordem e a ordem não controla estoque automaticamente, se for entra no estoque
'Set TBProduto = CreateObject("adodb.recordset")
'TBProduto.Open "Select CPL.OS, CPL.Ordem from (Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho) LEFT JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = CPL.ID_CFOP where CPL.Idlista = " & TXTIDLista & " and CFOP.MaoObra = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockReadOnly
'If TBProduto.EOF = False Then
'    Permitido = False
'    If IsNull(TBProduto!OS) = False And TBProduto!OS <> "" And IsNull(TBProduto!Ordem) = False And TBProduto!Ordem <> "" Then
'        Set TBOrdemServico = CreateObject("adodb.recordset")
'        TBOrdemServico.Open "Select OS.idproducao from OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.Ordem = " & TBProduto!Ordem & " and P.Entrar_estoque <> 'True' ORDER BY OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockReadOnly
'        If TBOrdemServico.EOF = False Then
'            TBOrdemServico.MoveLast
'            If TBOrdemServico!IDProducao = TBProduto!OS Then Permitido = True
'        End If
'        TBOrdemServico.Close
'    End If
'End If
'TBProduto.Close

'If Permitido = True Then
    TBEstoque!estoque_venda = txtQuantidade.Text
    TBEstoque!estoque_real = txtQuantidade.Text
    TBEstoque!estoque_real_PC = Valor_Cofins_Prod
'End If

Qtd = IIf(txtQuantidade.Text = "", 0, txtQuantidade.Text)
TBEstoque!Qtde = Qtd
TBEstoque!status = "ENTRADA_NOTA_FISCAL"

IDFase = 0
IDPlano = 0
'Grava familia do produto na tabela estoque_controle e Atualiza valor do material no estoque
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select classe, Codproduto, ID_PC from projproduto where desenho = '" & txtcodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBEstoque!Classe = TBAbrir!Classe
    IDFase = TBAbrir!Codproduto
    IDPlano = IIf(IsNull(TBAbrir!ID_PC), 0, TBAbrir!ID_PC)
    Familiatext = TBAbrir!Classe
    
    'Grava código de referência no produto
    If Cmb_codigo_ref <> "" Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from item_aplicacoes where Codproduto = " & IDFase & " and n_referencia = '" & Cmb_codigo_ref & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then TBProduto.AddNew
        TBProduto!Codproduto = IDFase
        TBProduto!N_referencia = Cmb_codigo_ref
        TBProduto!Descricao = IIf(txtEspecificacoes.Text = "", Null, txtEspecificacoes.Text)
        TBProduto!ID_cliente_forn = Txt_ID_forn
        TBProduto!Tipo = "F"
        TBProduto.Update
        TBProduto.Close
    End If
End If
TBAbrir.Close

TBEstoque.Update

'Cria o empenho no RE para o pedido interno se o produto controlar estoque
If Permitido = True Then
    qt = txtQuantidade
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IDcarteira, Qtde_empenho - ISNULL(Qtde_recebida, 0) AS Qtde_empenhada from Compras_pedido_lista_empenhos where IDlista = " & txtidlista & " and Qtde_empenho - ISNULL(Qtde_recebida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False And qt > 0
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Estoque_Controle_Empenho_Vendas", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!data = Txt_data_recebimento
            TBGravar!Responsavel = pubUsuario
            If qt >= TBAbrir!Qtde_empenhada Then
                TBGravar!Qtde_empenhada = TBAbrir!Qtde_empenhada
                qt = qt - TBAbrir!Qtde_empenhada
            Else
                TBGravar!Qtde_empenhada = qt
                qt = 0
            End If
            TBGravar.Update
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End If
frmEstoque_Recebimento.Tag = TBEstoque!IDestoque
IDestoque = TBEstoque!IDestoque
TBEstoque.Close

ValorTotal = 0
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select CPL.preco_unitario_desconto, CPL.preco_total, CPL.vlrICMS, CPL.Quant_Comp, CPL.UN, CPL.Unidade_com, ISNULL(CPL.Qtde_estoque, 0) AS Qtde_estoque from compras_pedido_lista CPL inner join compras_pedido CP on CPL.idpedido = CP.idpedido where CP.pedido = '" & txtProg_pedido & "' and CPL.idlista = " & txtidlista, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
    qt = 1
    Valor1 = Format(IIf(IsNull(TBPedido!vlrICMS), 0, TBPedido!vlrICMS) / IIf(IsNull(TBPedido!Quant_Comp), 0, TBPedido!Quant_Comp), "0.00")
    If TBPedido!Un <> TBPedido!Unidade_com And TBPedido!Qtde_estoque > 0 Then qt = TBPedido!Quant_Comp / TBPedido!Qtde_estoque
    ValorTotal = Format(qt * Format(IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)) - (qt * Valor1), "0.0000000000")
End If

quantestoque = txtQuantidade
NovoValor = Replace(ValorTotal, ",", ".")
Conexao.Execute "Update estoque_controle Set valor_unitario = " & NovoValor & " where IDestoque = " & IDestoque
If Permitido = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Valor_total from estoque_controle where IDestoque = " & IDestoque, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir!Valor_total = Format((quantestoque * ValorTotal) + IIf(IsNull(TBAbrir!Valor_total), 0, TBAbrir!Valor_total), "###.##0.00")
        TBAbrir.Update
    End If
    TBAbrir.Close
End If

quantestoque = 0
'Grava movimentação na tabela estoque_controle_recebimento
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle_recebimento", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!Data_recebimento = Txt_data_recebimento
TBEstoque!IDpedido = Txt_ID_pedido
TBEstoque!IDlista = txtidlista.Text
TBEstoque!Desenho = txtcodigo.Text
TBEstoque!Certificado = IIf(txtcertificado = "", 0, txtcertificado)
TBEstoque!Corrida = IIf(txtcorrida = "", 0, txtcorrida)
TBEstoque!local_armaz = cmbLocal_armaz
TBEstoque!Nota_fiscal = txtnotafiscal.Text
TBEstoque!Serie = txtSerie
If txtnotafiscal <> "" And txtdataemissao <> "__/__/____" Then TBEstoque!Data_emissao = txtdataemissao Else TBEstoque!Data_emissao = Null
TBEstoque!Responsavel = pubUsuario
TBEstoque!Recebido = Format(txtQuantidade.Text, "###.##0.000")
TBEstoque!Recebido_PC = Format(Valor_Cofins_Prod, "###.##0.000")
 
If Valor_Cofins_Prod = 0 Then
    quantestoque = Format(txtrequisitado.Text, "###.##0.000")
    quantnovo = Format(txtQuantidade.Text, "###.##0.000")
Else
    quantestoque = Format(IIf(txtrequisitado_PC = "", txtrequisitado, txtrequisitado_PC), "###.##0.000")
    quantnovo = Format(Valor_Cofins_Prod, "###.##0.000")
End If
If quantnovo < quantestoque Then TBEstoque!Parcial = True Else TBEstoque!Parcial = False
TBEstoque!Programacao = False
TBEstoque!ID_empresa = txtID_empresa
TBEstoque!Obs = txtOBS
TBEstoque.Update
IDEstoque_recebimento = TBEstoque!ID
TBEstoque.Close

ProcAtualizaQtdeRecebEmp Txt_ID_pedido, txtidlista

quantestoque = 0
quantnovo = 0
'Grava movimentação na tabela estoque_movimentacao
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!Destino = "Interno"
TBEstoque!Terceiros = False
TBEstoque!Operacao = "ENTRADA_NOTA_FISCAL"
TBEstoque!IDestoque = IDestoque
TBEstoque!Documento = txtnotafiscal.Text
TBEstoque!DtEmissao = Txt_data_recebimento
TBEstoque!LOTE = txtProg_pedido
TBEstoque!Responsavel = pubUsuario
TBEstoque!data = Txt_data_recebimento
TBEstoque!Descricao = txtEspecificacoes.Text
TBEstoque!Desenho = txtcodigo.Text
TBEstoque!estoque_venda = Format(txtQuantidade.Text, "###.##0.000")
TBEstoque!Entrada = Format(txtQuantidade.Text, "###.##0.000")
TBEstoque!Entrada_PC = Format(Valor_Cofins_Prod, "###.##0.000")
TBEstoque!Familia = Familiatext
TBEstoque!Obs = txtOBS
quantestoque = txtQuantidade

'Atualiza valor do material no estoque
TBEstoque!VlrUnit = Format(ValorTotal, "###.##0.00000")
TBEstoque!VlrTotal = Format(quantestoque * ValorTotal, "###.##0.00")

TBEstoque!IDEstoque_recebimento = IDEstoque_recebimento
TBEstoque!idlista_recebimento = txtidlista.Text
TBEstoque!Destino = "Interno"
TBEstoque!Terceiros = False

Set TBNivel1 = CreateObject("adodb.recordset")
TBNivel1.Open "Select * from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtcodigo & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel1.EOF = False Then
    Set TBNivel2 = CreateObject("adodb.recordset")
    TBNivel2.Open "Select sum(Saida) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtcodigo & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel2.EOF = False Then
        Valor1 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
    End If
    TBNivel2.Close
    Set TBNivel2 = CreateObject("adodb.recordset")
    TBNivel2.Open "Select sum(entrada) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtcodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel2.EOF = False Then
        Valor2 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
    End If
    TBNivel2.Close
    Valor2 = Valor2 + Qtd
    If Valor1 <= Valor2 Then
        Conexao.Execute "UPDATE estoque_movimentacao set Terceiros = 'False' where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtcodigo & "' and destino = 'Terceiros'"
    End If
    TBEstoque!Pedidocompra = txtProg_pedido
    TBEstoque!IDpedido = Txt_ID_pedido
Else
    TBEstoque!Pedidocompra = Null
    TBEstoque!IDpedido = Null
End If
TBNivel1.Close
TBEstoque.Update

'Atualiza o RE no instrumento
Conexao.Execute "Update I set I.IDestoque = " & TBEstoque!IDestoque & " from Instrumentos I INNER JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where EC.Desenho = '" & txtcodigo & "' and EC.Numero_serie = '" & Txt_numero_serie & "'"

'Centro de custo
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CPLC.Valor, CPLC.ID_CC, CP.Data from Compras_pedido_lista_custo CPLC INNER JOIN Compras_pedido CP ON CPLC.IDPedido = CP.IDPedido where CPLC.IDLista = " & txtidlista.Text & " and CP.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Valor3 = TBAbrir!valor
        
        qt = txtrequisitado
        Qtde = TBEstoque!Entrada
        
        'Calcula quantidade se for com unidade diferente (inverti a un de estoque com a un comercial porque preciso voltar a conversão)
        Set TBPedido = CreateObject("adodb.recordset")
        TBPedido.Open "Select CPL.UN, CPL.Unidade_com, CPL.Qtde_estoque, CPL.Quant_comp from compras_pedido_lista CPL INNER JOIN compras_pedido CP on CPL.idpedido = CP.idpedido where CP.pedido = '" & txtProg_pedido & "' and CPL.idlista = " & txtidlista & " and CPL.Qtde_Estoque IS NOT NULL and CPL.Qtde_estoque <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBPedido.EOF = False Then
            If TBPedido!Un <> TBPedido!Unidade_com Then
                qt = TBPedido!Quant_Comp
                If FunVerifUNConversao(TBPedido!Un, TBPedido!Unidade_com) = True Then
                    Qtde = FunConverteUN(TBPedido!Unidade_com, TBPedido!Un, TBEstoque!Entrada, txtcodigo)
                Else
                    Qtde = TBEstoque!Entrada * FunVerificaTabelaConversaoUnidade(TBPedido!Un, TBPedido!Unidade_com)
                End If
            End If
        End If
        valor = Format((Valor3 / qt) * Qtde, "###,##0.00")
        
        'Verifica se tem CC amarrado no produto, se for diferente do informado no pedido cria débito e crédito no CC do produto
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select ID_CC from projproduto where codproduto = " & IDFase & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            If TBProduto!ID_CC <> "" Then
                If TBAbrir!ID_CC <> TBProduto!ID_CC Then
                    ProcSalvarCCRealizado TBAbrir!data, txtID_empresa, "Débito", TBProduto!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, txtidlista, valor, True, False
                    
                    'Grava movimentação no centro consolidado
                    ProcSalvarRealCCConsolidado TBProduto!ID_CC, "Débito", False, False, False
                    
                    ProcSalvarCCRealizado TBAbrir!data, txtID_empresa, "Crédito", TBProduto!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, txtidlista, valor, True, False
                    
                    'Grava movimentação no centro consolidado
                    ProcSalvarRealCCConsolidado TBProduto!ID_CC, "Crédito", True, True, False
                End If
            End If
        End If
        TBProduto.Close
        
        ProcSalvarCCRealizado TBAbrir!data, txtID_empresa, "Débito", TBAbrir!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, txtidlista, valor, False, False
        
        'Grava movimentação no centro consolidado
        ProcSalvarRealCCConsolidado TBAbrir!ID_CC, "Débito", False, False, False
        
        TBAbrir.MoveNext
    Loop
Else
    'Verifica se tem CC amarrado no produto e cria um débito no CC do produto
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select ID_CC from projproduto where codproduto = " & IDFase & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        qt = txtrequisitado
        Qtde = TBEstoque!Entrada
        valor = Format((Valor3 / qt) * Qtde, "###,##0.00")
        
        ProcSalvarCCRealizado Txt_data_recebimento, txtID_empresa, "Débito", TBProduto!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, txtidlista, valor, False, False
        
        'Grava movimentação no centro consolidado
        ProcSalvarRealCCConsolidado TBProduto!ID_CC, "Débito", False, False, False
    End If
    TBProduto.Close
End If
TBAbrir.Close
TBEstoque.Close

'Atualiza o status do pedido, status e qtde recebida da programação
ProcAtualizaQtdeRecebidaProg "Select CP.qtderecebida, CP.Quantidade from Compras_Programacao CP INNER JOIN Compras_pedido_lista CPL ON CP.ID_prog = CPL.ID_programacao where CPL.IDlista = " & txtidlista & " order by CP.data_inicio", False
ProcAlteraStatus_pedido
ProcAlteraStatus_prog

USMsgBox ("Produto recebido ao estoque com sucesso."), vbInformation, "CAPRIND v5.0"
cmdReceber.Enabled = False
ProcBloqueiaFrame
ProcCarregaListaMovimentacao
'==================================
Modulo = "Estoque/Recebimento/Pedido de compra"
Evento = "Receber"
ID_documento = txtidlista
Documento = "Cód. interno: " & txtcodigo & " - Nº lote: " & txtProg_pedido & " - Nº corrida: " & IIf(txtcorrida = "", 0, txtcorrida) & " - Nº certificado: " & IIf(txtcertificado = "", 0, txtcertificado) & " - Local armaz.: " & cmbLocal_armaz
Documento1 = "Operação: " & Lista_movimentacao.SelectedItem.SubItems(2) & " - Documento: " & Lista_movimentacao.SelectedItem.SubItems(6)
ProcGravaEvento
'==================================
If txtnotafiscal <> "" Then ProcAtualizaVlrEntradaEstoque True
ProcCarregaListaFiltro
ProcLimparCamposReq True
Lista_movimentacao.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaQtdeRecebEmp(IDpedido As Long, IDlista As Long)
On Error GoTo tratar_erro

valor = 0
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select ROUND(Sum(ISNULL(Recebido, 0)), 4) as Valor from estoque_controle_recebimento where idpedido = " & IDpedido & " and idlista = " & IDlista & " and Programacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    valor = TBCFOP!valor
End If
TBCFOP.Close
NovoValor = Replace(valor, ",", ".")
Conexao.Execute "Update Compras_pedido_lista_empenhos Set Qtde_recebida = " & NovoValor & " where IDlista = " & IDlista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaQtdeRecebidaProg(TextoFiltro As String, Excluir As Boolean)
On Error GoTo tratar_erro

If Excluir = False Then
    QuantEmpenho = qt
    quantidade = 0
    ValorTotal = 0
    Set TBProgramas = CreateObject("adodb.recordset")
    TBProgramas.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBProgramas.EOF = False Then
        Do While TBProgramas.EOF = False
            If TBProgramas!Qtderecebida <> TBProgramas!quantidade Then
                If QuantEmpenho > TBProgramas!quantidade - TBProgramas!Qtderecebida Then  'Se qtde recebida for maior q a qtde programada
                    quantidade = TBProgramas!quantidade - TBProgramas!Qtderecebida
                    TBProgramas!Qtderecebida = TBProgramas!quantidade
                    QuantEmpenho = QuantEmpenho - quantidade
                Else
                    If TBProgramas!Qtderecebida = 0 Then 'Se qtde recebida for = 0
                        TBProgramas!Qtderecebida = QuantEmpenho
                        QuantEmpenho = QuantEmpenho - QuantEmpenho
                    Else
                        'Se qtde recebida for menor q a qtde programada
                        TBProgramas!Qtderecebida = QuantEmpenho + TBProgramas!Qtderecebida
                        QuantEmpenho = QuantEmpenho - QuantEmpenho
                    End If
                End If
            End If
            TBProgramas.Update
            If QuantEmpenho <= 0 Then GoTo Sair
            TBProgramas.MoveNext
        Loop
    End If
Sair:
        TBProgramas.Close
    
Else
    QuantEmpenho = qt
    quantidade = 0
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        Do While TBCompras_Lista.EOF = False
            If QuantEmpenho <= 0 Then GoTo Sair1
            quantidade = TBCompras_Lista!Qtderecebida
            If QuantEmpenho >= TBCompras_Lista!Qtderecebida Then
                TBCompras_Lista!Qtderecebida = TBCompras_Lista!Qtderecebida - TBCompras_Lista!Qtderecebida
            Else
                TBCompras_Lista!Qtderecebida = TBCompras_Lista!Qtderecebida - QuantEmpenho
            End If
            If TBCompras_Lista!Qtderecebida = 0 Then
                 If TBCompras_Lista!Firme = True Then
                    TBCompras_Lista!Status_prog = "ABERTO"
                    TBCompras_Lista!Ordenar = 2
                Else
                    TBCompras_Lista!Status_prog = "PREVISÃO FUTURA"
                    TBCompras_Lista!Ordenar = 3
                End If
            End If
            If TBCompras_Lista!Qtderecebida <> 0 And TBCompras_Lista!Qtderecebida < TBCompras_Lista!quantidade Then
                TBCompras_Lista!Status_prog = "PARCIAL"
                TBCompras_Lista!Ordenar = 1
            End If
            If TBCompras_Lista!Qtderecebida >= TBCompras_Lista!quantidade Then
                TBCompras_Lista!Status_prog = "RECEBIDO"
                TBCompras_Lista!Ordenar = 4
            End If
            TBCompras_Lista.Update
            QuantEmpenho = QuantEmpenho - quantidade
Sair1:
            TBCompras_Lista.MoveNext
        Loop
    End If
    TBCompras_Lista.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: cmdReceber_Click
    Case vbKeyF4: cmdCancelar_Click
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyF7: ProcStatus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 8, True
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Formulario = "Estoque/Recebimento/Pedido de compra"
Direitos

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaPedido()
On Error GoTo tratar_erro

If txtProg_pedido.Text = "" Then Exit Sub
ValorTotal = 0
Qtd = 0
Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select ID_empresa, Data, IDpedido, idfornecedor, Fornecedor, Estado from compras_pedido where pedido = '" & txtProg_pedido & "' and (Status_pedido = 'ABERTO' or Status_pedido = 'APROVADO' or Status_pedido = 'PARCIAL' or Status_pedido = 'ENCERRADO')", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    Txt_ID_pedido = TBCompras_Pedido!IDpedido
    Txt_ID_forn = TBCompras_Pedido!IDfornecedor
    txtfornecedor.Text = IIf(IsNull(TBCompras_Pedido!Fornecedor) = False, TBCompras_Pedido!Fornecedor, "")
    txtuf = IIf(IsNull(TBCompras_Pedido!Estado), "", TBCompras_Pedido!Estado)
    
    Set TBExecucao = CreateObject("adodb.recordset")
    TBExecucao.Open "Select CODIGO, Empresa from Empresa where codigo = " & TBCompras_Pedido!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBExecucao.EOF = False Then
        txtID_empresa.Text = IIf(IsNull(TBExecucao!CODIGO), "", TBExecucao!CODIGO)
        txtEmpresa.Text = IIf(IsNull(TBExecucao!Empresa), "", TBExecucao!Empresa)
    End If
    TBExecucao.Close
    
    txtdata.Text = Format(TBCompras_Pedido!data, "dd/mm/yy")
End If
TBCompras_Pedido.Close
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from compras_comercial where idpedido = " & Txt_ID_pedido, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtcondpagamento.Text = IIf(IsNull(TBFornecedor!condicoes) = False, TBFornecedor!condicoes, "")
End If
TBFornecedor.Close
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Sum(preco_total) as Valortotal from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ValorTotal = IIf(IsNull(TBLISTA!ValorTotal), 0, TBLISTA!ValorTotal)
End If
TBLISTA.Close
txtvalortotal.Text = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmEstoque_Recebimento_Imprimir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro
  
txtidlista = ""
txtcondpagamento = ""
txtdata = ""
txtvalortotal = ""
Txt_ID_forn = ""
txtfornecedor = ""
txtuf = ""
Txt_ID_pedido = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparCamposReq(Receber As Boolean)
On Error GoTo tratar_erro

txtcodigo.Text = ""
Cmb_codigo_ref.Clear
txtUN.Text = ""
txtstatus.Text = ""
txtEspecificacoes = ""
txtcertificado = ""
txtcorrida = ""
If Receber = True Then
    If Chk_LA.Value = 0 Then cmbLocal_armaz.ListIndex = -1
Else
    cmbLocal_armaz.ListIndex = -1
End If
txtQuantidade = ""
txtQuantidade_PC = ""
Txt_numero_serie = ""
If Receber = True Then
    If Chk_Dt_rcbto.Value = 0 Then Txt_data_recebimento = "__/__/____"
Else
    Txt_data_recebimento = "__/__/____"
End If
txtrequisitado = "0,0000"
txtrequisitado_PC = "0,0000"
txtrecebida = "0,0000"
txtrecebida_PC = "0,0000"
txtSaldo = "0,0000"
txtOBS = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_movimentacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ECR.*, EM.Data, EC.Numero_serie from (Estoque_controle_recebimento ECR INNER JOIN Estoque_movimentacao EM ON ECR.Id = EM.IDEstoque_recebimento) INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.Idoperacao = " & Lista_movimentacao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Nota_fiscal) = False And TBAbrir!Nota_fiscal <> "" Then
        txtnotafiscal.Text = TBAbrir!Nota_fiscal
        txtSerie = IIf(IsNull(TBAbrir!Serie), "", TBAbrir!Serie)
        txtdataemissao = Format(TBAbrir!Data_emissao, "dd/mm/yyyy")
    End If
    txtcertificado.Text = IIf(IsNull(TBAbrir!Certificado), "", TBAbrir!Certificado)
    Txt_data_recebimento = Format(TBAbrir!data, "dd/mm/yyyy")
    txtQuantidade.Text = Format(Lista_movimentacao.SelectedItem.ListSubItems(4), "###,##0.0000")
    txtQuantidade_PC.Text = Lista_movimentacao.SelectedItem.ListSubItems(5)
    txtcorrida.Text = IIf(IsNull(TBAbrir!Corrida), "", TBAbrir!Corrida)
    Txt_numero_serie = IIf(IsNull(TBAbrir!Numero_serie), "", TBAbrir!Numero_serie)
    
    If TBAbrir!local_armaz <> "" Then cmbLocal_armaz.Text = TBAbrir!local_armaz
1:
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o local de armazenamento (" & TBAbrir!local_armaz & ") desta movimentação, favor revisar."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNotaFiscal_Change()
On Error GoTo tratar_erro
    
If txtnotafiscal.Text <> "" Then
    VerifNumero = txtnotafiscal.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtnotafiscal.Text = ""
        txtnotafiscal.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

If txtQuantidade.Text <> "" Then
    VerifNumero = txtQuantidade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade.Text = ""
        txtQuantidade.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantidade_LostFocus()
On Error GoTo tratar_erro

txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAlteraStatus_pedido()
On Error GoTo tratar_erro
quantnovo = 0
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from compras_pedido where pedido = '" & txtProg_pedido.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    Txt_ID_pedido = TBCompras!IDpedido
    
    quantnovo = 0
    Valor_Cofins_Prod = 0
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select Sum(Recebido) as quantnovo, Sum(ISNULL(Recebido_PC, 0)) as Valor_Cofins_Prod from estoque_controle_recebimento where idpedido = " & Txt_ID_pedido & " and idlista = " & txtidlista.Text & " and Programacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        quantnovo = IIf(IsNull(TBEstoque!quantnovo), 0, TBEstoque!quantnovo)
        Valor_Cofins_Prod = IIf(IsNull(TBEstoque!Valor_Cofins_Prod), 0, TBEstoque!Valor_Cofins_Prod)
    End If
    TBEstoque.Close
    
    If Valor_Cofins_Prod > 0 Then
        quantnovo = Valor_Cofins_Prod
        quantestoque = IIf(txtrequisitado_PC = "", txtrequisitado, txtrequisitado_PC)
    Else
        quantestoque = txtrequisitado.Text
    End If
    
'=======================================================
'Acerta Status dos itns recebidos de forma parcial
'=======================================================
    If quantnovo < quantestoque Then
        If USMsgBox("Este produto/serviço será recebido parcialmente, deseja encerrar o mesmo no pedido de compra?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then
            Status_Item = "RECEBIDO"
        Else
            Status_Item = "PARCIAL"
        End If
    End If
'======================================================
'Acerta Status dos itens do pedido
'======================================================
    If quantnovo >= quantestoque Then Status_Item = "RECEBIDO"
    Conexao.Execute "Update compras_pedido_lista Set Status_item = '" & Status_Item & "' where idpedido = " & Txt_ID_pedido & " and idlista = " & txtidlista.Text
    Conexao.Execute "Update Compras_Programacao set Compras_Programacao.Status_prog = '" & Status_Item & "' from Compras_Programacao INNER JOIN compras_pedido_lista ON Compras_Programacao.ID_prog = compras_pedido_lista.ID_programacao where compras_pedido_lista.idpedido = " & Txt_ID_pedido & " and compras_pedido_lista.idlista = " & txtidlista.Text
'=====================================================
'Verifica Status dos itens no pedido
'=====================================================
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and status_item <> 'RECEBIDO' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = True Then
        Status_pedido = "ENCERRADO"
    Else
        Status_pedido = "PARCIAL"
    End If
    TBCompras.Close
'======================================================
'Acerta Status do Pedido de compra
'======================================================
    Conexao.Execute "Update compras_pedido Set Status_pedido = '" & Status_pedido & "' where idpedido = '" & Txt_ID_pedido.Text & "'"
End If
'======================================================
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAlteraStatus_prog()
On Error GoTo tratar_erro

'Produto
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select compras_programa_item.* from (Compras_Programacao INNER JOIN Compras_pedido_lista ON Compras_Programacao.ID_prog = Compras_pedido_lista.ID_programacao) INNER JOIN compras_programa_item ON compras_programa_item.ID_item = Compras_Programacao.Id_item where Compras_pedido_lista.IDlista = " & txtidlista, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!Status_Item = "PREVISÃO FUTURA"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from compras_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!Status_Item = "N_RECEBIDO"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from compras_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!Status_Item = "RECEBIDO"
            Else
                TBItem!Status_Item = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    Txt_ID_pedido = TBItem!ID
    TBItem.Update
End If

'Programa
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from compras_programa where id = " & Txt_ID_pedido, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_programa_item where id = " & Txt_ID_pedido & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!status = "PREVISÃO FUTURA"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from compras_programa_item where id = " & Txt_ID_pedido & " and Status_Item <> 'N_RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!status = "ABERTO"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from compras_programa_item where id = " & Txt_ID_pedido & " and Status_Item <> 'RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!status = "RECEBIDO"
            Else
                TBItem!status = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    TBItem.Update
End If
TBItem.Close
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

If StrSql_Estoque_Recebimento_Localizar = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
listprod.ListItems.Clear
Set TBLISTA_Estoque_RecebimentoPedido = CreateObject("adodb.recordset")
TBLISTA_Estoque_RecebimentoPedido.Open StrSql_Estoque_Recebimento_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Estoque_RecebimentoPedido.EOF = False Then ProcExibePagina (1)
ProcCarregaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

listprod.ListItems.Clear
TBLISTA_Estoque_RecebimentoPedido.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_RecebimentoPedido.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_RecebimentoPedido.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_RecebimentoPedido.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_RecebimentoPedido.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Estoque_RecebimentoPedido.EOF = False And (ContadorReg <= TamanhoPagina)
    With listprod.ListItems
        .Add , , TBLISTA_Estoque_RecebimentoPedido!IDlista
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Empresa from Empresa where codigo = " & TBLISTA_Estoque_RecebimentoPedido!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!Empresa), "", TBFI!Empresa)
        TBFI.Close
        .Item(.Count).SubItems(2) = TBLISTA_Estoque_RecebimentoPedido!Pedido
        .Item(.Count).SubItems(3) = TBLISTA_Estoque_RecebimentoPedido!Desenho
        .Item(.Count).SubItems(4) = TBLISTA_Estoque_RecebimentoPedido!Descricao
        .Item(.Count).SubItems(5) = TBLISTA_Estoque_RecebimentoPedido!Un
        
        If TBLISTA_Estoque_RecebimentoPedido!Un <> TBLISTA_Estoque_RecebimentoPedido!Unidade_com And IsNull(TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque) = False And TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque <> 0 Then
            valor = TBLISTA_Estoque_RecebimentoPedido!Quant_Comp / TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque
             
            .Item(.Count).SubItems(6) = Format(valor * TBLISTA_Estoque_RecebimentoPedido!preco_unitario, "###,##0.0000")
            .Item(.Count).SubItems(7) = Format(TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque, "###,##0.0000")
        Else
            .Item(.Count).SubItems(6) = Format(TBLISTA_Estoque_RecebimentoPedido!preco_unitario, "###,##0.0000")
            .Item(.Count).SubItems(7) = Format(TBLISTA_Estoque_RecebimentoPedido!Quant_Comp, "###,##0.0000")
        End If
        
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Estoque_RecebimentoPedido!Quant_Comp_PC), "", TBLISTA_Estoque_RecebimentoPedido!Quant_Comp_PC)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Estoque_RecebimentoPedido!Prazo), "", Format(TBLISTA_Estoque_RecebimentoPedido!Prazo, "dd/mm/yy"))
        
        If TBLISTA_Estoque_RecebimentoPedido!Status_Item = "N_RECEBIDO" Or TBLISTA_Estoque_RecebimentoPedido!Status_Item = "APROVADO" Then
            Statusitem = "NÃO_RECEBIDO"
        Else
            Statusitem = TBLISTA_Estoque_RecebimentoPedido!Status_Item
        End If
        .Item(.Count).SubItems(10) = Statusitem
        
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Estoque_RecebimentoPedido!Ordem), "", TBLISTA_Estoque_RecebimentoPedido!Ordem)
    End With
    TBLISTA_Estoque_RecebimentoPedido.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_RecebimentoPedido.RecordCount
If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_RecebimentoPedido.PageCount
ElseIf TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_RecebimentoPedido.PageCount & " de: " & TBLISTA_Estoque_RecebimentoPedido.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_RecebimentoPedido.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_RecebimentoPedido.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

TotContas = 0
IDlista = 0
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open StrSql_Estoque_Recebimento_LocalizarTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Do While TBContas.EOF = False
        Permitido = False
        If IDlista <> TBContas!IDlista Then Permitido = True
        If Permitido = True Then TotContas = TotContas + IIf(IsNull(TBContas!TotContas), 0, TBContas!TotContas)
        IDlista = TBContas!IDlista
        TBContas.MoveNext
    Loop
End If
TBContas.Close
txtQtde_total.Text = Format(TotContas, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proccarregalocarm()
On Error GoTo tratar_erro

cmbLocal_armaz.Clear
TextoLocal = ""
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select ELC.Descricao from Estoque_Localarmazenamento_criar ELC INNER JOIN Estoque_Localarmazenamento EL ON ELC.ID = EL.idemb_locarm where EL.codinterno = '" & txtcodigo & "' and ELC.Descricao is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    Do While TBAliquota.EOF = False
        If IsNull(TBAliquota!Descricao) = False Then cmbLocal_armaz.AddItem TBAliquota!Descricao
        TBAliquota.MoveNext
    Loop
Else
    ProcCarregaComboLA cmbLocal_armaz, False, False
    
    'Se for serviço, mão de obra ou remessa cria um local de armz padrão
    If TBProduto!Tipo = "S" Then
        TextoLocal = "SERVIÇOS"
    ElseIf TBProduto!Remessa = True Then
        TextoLocal = "RETORNO DE MERCADORIA"
    ElseIf IsNull(TBProduto!ID_CFOP) = False And TBProduto!ID_CFOP <> "" And TBProduto!ID_CFOP <> "0" Then
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select IDCountCfop from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP & " and MaoObra = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            TextoLocal = "INDUSTRIALIZAÇÃO"
        End If
        TBCFOP.Close
    End If
End If
TBAliquota.Close
    
With cmbLocal_armaz
    If TextoLocal <> "" Then
        .Text = TextoLocal
        .Locked = True
        .TabStop = False
    Else
        .Locked = False
        .TabStop = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_PC_Change()
On Error GoTo tratar_erro

If txtQuantidade_PC <> "" Then
    VerifNumero = txtQuantidade_PC
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade_PC = ""
        txtQuantidade_PC.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtstatus_Change()
On Error GoTo tratar_erro

If txtstatus = "RECEBIDO" Then cmdReceber.Enabled = False Else cmdReceber.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaFrame()
On Error GoTo tratar_erro

Frame2.Enabled = False
Frame6.Enabled = False
Frame11.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaFrame()
On Error GoTo tratar_erro

Frame2.Enabled = True
Frame6.Enabled = True
Frame11.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarRealCCConsolidado(ID_CC As Long, Operacao As String, Credito As Boolean, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & ID_CC, Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
    Do While TBAfericao.EOF = False
        ProcSalvarCCRealizado Txt_data_recebimento, txtID_empresa, Operacao, TBAfericao!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, txtidlista, valor, CC_produto, Bloqueado
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                ProcSalvarCCRealizado Txt_data_recebimento, txtID_empresa, Operacao, TBCiclo!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, txtidlista, valor, CC_produto, Bloqueado
                TBCiclo.MoveNext
            Loop
        End If
        TBCiclo.Close
        
        TBAfericao.MoveNext
    Loop
End If
TBAfericao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarCCRealizado(Data1 As Date, ID_empresa As Integer, Operacao As String, ID_CC As Long, Cod_produto As Long, ID_plano_contas As Long, ID_estoque As Long, ID_lista As Long, valor As Double, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

NovoValor = Replace(valor, ",", ".")
ProcINSERTINTO "CC_realizado", "Data, Responsavel, ID_empresa, Operacao, ID_CC, Cod_produto, ID_PC, ID_estoque, ID_lista, Valor, Bloqueado", "'" & data & "', '" & pubUsuario & "', " & ID_empresa & ", '" & Operacao & "', " & ID_CC & ", " & Cod_produto & ", " & ID_plano_contas & ", " & IIf(ID_estoque = 0, "NULL", ID_estoque) & ", " & ID_lista & ", " & NovoValor & ", " & IIf(Bloqueado = True, 1, 0) & ""

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select ID from CC_realizado where ID_estoque = " & ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If CC_produto = True And Operacao = "Crédito" Then Conexao.Execute "Update CC_realizado Set ID_ref_debito = " & TBGravar!ID - 1 & " where ID = " & TBGravar!ID
End If
TBGravar.Close

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
    Case 3: ProcStatus
    Case 4: procAtualiza
    Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

