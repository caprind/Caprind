VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmestoque_entrada 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Movimentação - Entrada"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   420
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
   ForeColor       =   &H00000000&
   Icon            =   "frmestoque_entrada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quantidades no estoque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1065
      Left            =   90
      TabIndex        =   57
      Top             =   8850
      Width           =   15165
      Begin DrawSuite2022.USButton btnEntrada 
         Height          =   795
         Left            =   13050
         TabIndex        =   66
         Top             =   150
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1402
         DibPicture      =   "frmestoque_entrada.frx":0442
         Caption         =   "Executar a entrada"
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
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.TextBox txtestoqueatualizado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "0,000"
         ToolTipText     =   "Estoque atualizado."
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txtentrada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   11610
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "0,000"
         ToolTipText     =   "Quantidade de entrada."
         Top             =   510
         Width           =   1365
      End
      Begin VB.TextBox txtestoquereal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "0,000"
         ToolTipText     =   "Quantidade em estoque."
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txtVlrTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. estoque"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   930
         TabIndex        =   65
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. entrada*"
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
         Left            =   11685
         TabIndex        =   64
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. atualiza."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2385
         TabIndex        =   63
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4425
         TabIndex        =   62
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escolha o tipo da entrada no estoque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   60
      TabIndex        =   53
      Top             =   990
      Width           =   15195
      Begin VB.OptionButton opt_ordem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrada em estoque por ordem de produção"
         DisabledPicture =   "frmestoque_entrada.frx":20EF
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   56
         Top             =   420
         Value           =   -1  'True
         Width           =   4245
      End
      Begin VB.OptionButton Opt_devolucao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrada em estoque por devolução nota fiscal"
         DisabledPicture =   "frmestoque_entrada.frx":24C031
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5220
         TabIndex        =   55
         Top             =   420
         Width           =   4875
      End
      Begin VB.OptionButton Opt_outras 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outras tipos de entradas no estoque"
         DisabledPicture =   "frmestoque_entrada.frx":495F73
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   11070
         TabIndex        =   54
         Top             =   420
         Width           =   3855
      End
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
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.ComboBox cmbReferencia 
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
      Left            =   10575
      Sorted          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Código de referência."
      Top             =   2130
      Width           =   3120
   End
   Begin VB.ComboBox txtLocal_armaz 
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
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "Local de armazenamento."
      Top             =   3330
      Width           =   5865
   End
   Begin VB.Frame Frame1 
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
      Height          =   7110
      Left            =   55
      TabIndex        =   26
      Top             =   1740
      Width           =   15195
      Begin VB.CommandButton cmdNF 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1380
         Picture         =   "frmestoque_entrada.frx":6DFEB5
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Pesquisar por número da nota fiscal."
         Top             =   2190
         Width           =   315
      End
      Begin VB.TextBox Txt_numero_serie 
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
         Left            =   12990
         TabIndex        =   20
         ToolTipText     =   "Número de série."
         Top             =   1590
         Width           =   1995
      End
      Begin VB.CommandButton Cmd_localizar_produtos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9855
         Picture         =   "frmestoque_entrada.frx":6E02D0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Localizar produtos."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_qtde_refugada 
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
         Left            =   10860
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade refugada."
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   13620
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Saldo."
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox txtQtde_produzida 
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade produzida."
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox txtQtde_entrada 
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de entrada."
         Top             =   990
         Width           =   1365
      End
      Begin VB.CommandButton Cmd_visualizar_arquivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   10170
         Picture         =   "frmestoque_entrada.frx":6E03D2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Visualizar arquivo."
         Top             =   390
         Width           =   315
      End
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
         ItemData        =   "frmestoque_entrada.frx":6E0994
         Left            =   180
         List            =   "frmestoque_entrada.frx":6E0996
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   5040
      End
      Begin VB.TextBox txtVlrUnit 
         Alignment       =   2  'Center
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
         Left            =   8340
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "0,0000000000"
         ToolTipText     =   "Valor unitário."
         Top             =   2190
         Width           =   1275
      End
      Begin VB.TextBox txtdescricao 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   990
         Width           =   7275
      End
      Begin VB.TextBox txtunidade 
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
         Left            =   7470
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   990
         Width           =   620
      End
      Begin VB.CommandButton cmdlote 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6810
         Picture         =   "frmestoque_entrada.frx":6E0998
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Pesquisar por número do lote."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtQtde 
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
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox txtObservacoes 
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
         Height          =   4125
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         ToolTipText     =   "Observações."
         Top             =   2790
         Width           =   14805
      End
      Begin VB.TextBox txtNota_Fiscal 
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
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Número da nota fiscal."
         Top             =   2190
         Width           =   1185
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Cliente."
         Top             =   2190
         Width           =   6615
      End
      Begin VB.TextBox txtCorrida 
         Alignment       =   2  'Center
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
         Left            =   8340
         MaxLength       =   50
         TabIndex        =   18
         ToolTipText     =   "Número da corrida."
         Top             =   1590
         Width           =   2250
      End
      Begin VB.TextBox txtPeso 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Peso unitário."
         Top             =   1590
         Width           =   1170
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   2640
      End
      Begin VB.TextBox txtresponsavel 
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
         Left            =   9630
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   2190
         Width           =   5355
      End
      Begin VB.TextBox txtlote 
         Alignment       =   2  'Center
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
         Left            =   5250
         MaxLength       =   12
         TabIndex        =   1
         ToolTipText     =   "Número do lote."
         Top             =   390
         Width           =   1545
      End
      Begin VB.TextBox txtUN 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Unidade por kilograma."
         Top             =   1590
         Width           =   750
      End
      Begin VB.TextBox txtcertificado 
         Alignment       =   2  'Center
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
         Left            =   10605
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Número do certificado."
         Top             =   1590
         Width           =   2370
      End
      Begin MSComCtl2.DTPicker txtdata 
         Height          =   315
         Left            =   13650
         TabIndex        =   7
         ToolTipText     =   "Data da movimentação."
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   197066753
         CurrentDate     =   39057
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de série"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   13402
         TabIndex        =   51
         Top             =   1380
         Width           =   1170
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. refugada"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10980
         TabIndex        =   50
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   14070
         TabIndex        =   49
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. entrada"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12405
         TabIndex        =   48
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. produzida"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9570
         TabIndex        =   47
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa*"
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
         Left            =   2280
         TabIndex        =   45
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor unit."
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
         Left            =   8550
         TabIndex        =   44
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7650
         TabIndex        =   43
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un/Kg"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1815
         TabIndex        =   42
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1410
         TabIndex        =   41
         Top             =   1710
         Width           =   195
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8565
         TabIndex        =   40
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Local de armazenamento*"
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
         Left            =   4260
         TabIndex        =   39
         Top             =   1380
         Width           =   2250
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7110
         TabIndex        =   38
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota fiscal"
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
         Index           =   3
         Left            =   390
         TabIndex        =   37
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   4770
         TabIndex        =   36
         Top             =   1980
         Width           =   585
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Corrida"
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
         Left            =   9158
         TabIndex        =   35
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Peso unit."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   34
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11475
         TabIndex        =   33
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição técnica"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3195
         TabIndex        =   32
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7950
         TabIndex        =   31
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° do lote*"
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
         Height          =   210
         Index           =   0
         Left            =   5595
         TabIndex        =   30
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   14175
         TabIndex        =   29
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   11805
         TabIndex        =   28
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado"
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
         Left            =   11333
         TabIndex        =   27
         Top             =   1380
         Width           =   915
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   46
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar (F3)"
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
      ButtonWidth1    =   38
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   42
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "4"
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
      ButtonLeft3     =   46
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "5"
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
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "6"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3930
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmestoque_entrada.frx":6E0DB3
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmestoque_entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnEntrada_Click()
On Error GoTo tratar_erro

ProcSalvar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_produtos_Click()
On Error GoTo tratar_erro

CadMaquinas = False
Estoque_entrada = True
frmEstoque_fisico_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txtCodigo = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtCodigo & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdlote_Click()
On Error GoTo tratar_erro

If Cmb_empresa = "" Then
    USMsgBox ("Informe a empresa antes de pesquisar."), vbExclamation, "CAPRIND v5.0"
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtLote = "" Then
    If Opt_devolucao.Value = True Then MsgTexto = "RE" Else MsgTexto = "lote"
    USMsgBox ("Informe o número do " & MsgTexto & " antes de pesquisar."), vbExclamation, "CAPRIND v5.0"
    txtLote.SetFocus
    Exit Sub
End If
ProcLimpaCampos

If opt_ordem.Value = True Then
    If FunPuxaDadosOrdem = False Then Exit Sub
ElseIf Opt_devolucao.Value = True Then
        If FunPuxaDadosEstoque = False Then Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Function FunPuxaDadosOrdem() As Boolean
On Error GoTo tratar_erro

FunPuxaDadosOrdem = True
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select P.*, PP.Unidade, PP.Peso_metro, PP.Un_Kg from producao P INNER JOIN projproduto PP ON PP.Desenho = P.Desenho where P.Ordem = " & txtLote.Text & " and P.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and P.DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    If TBproducao!status = "Cancelada" Then
        USMsgBox ("Esta ordem está cancelada."), vbExclamation, "CAPRIND v5.0"
        txtLote.SetFocus
        TBproducao.Close
        FunPuxaDadosOrdem = False
        Exit Function
    End If
    
    With cmbReferencia
        .Clear
        If IsNull(TBproducao!N_referencia) = False And TBproducao!N_referencia <> "" Then
            .AddItem TBproducao!N_referencia
            .Text = TBproducao!N_referencia
        End If
        .Locked = False
        .TabStop = True
    End With
    
    txtCodigo.Text = IIf(IsNull(TBproducao!Desenho), "", TBproducao!Desenho)
    txtdescricao = IIf(IsNull(TBproducao!Produto), "", TBproducao!Produto)
    txtunidade = IIf(IsNull(TBproducao!Unidade), "", TBproducao!Unidade)
    
    txtQtde.Text = Format(IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant), "###,##0.0000")
    txtQtde_produzida = Format(IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd), "###,##0.0000")
    Txt_qtde_refugada = Format(IIf(IsNull(TBproducao!QuantNC), 0, TBproducao!QuantNC), "###,##0.0000")
    
    quantidade = 0
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Qtde_entrada from Qtde_entrada_estoque_produto_produzido_lote where lote = '" & txtLote & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        quantidade = IIf(IsNull(TBFIltro!Qtde_entrada), 0, TBFIltro!Qtde_entrada)
    End If
    TBFIltro.Close
    txtQtde_entrada = Format(IIf(quantidade < 0, 0, quantidade), "###,##0.0000")
    
    If IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd) > 0 Or TBproducao!Processo_controlado = True Then
        txtSaldo = Format((IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd) - IIf(IsNull(TBproducao!QuantNC), 0, TBproducao!QuantNC)) - quantidade, "###,##0.0000")
    Else
        txtSaldo = Format(IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant) - quantidade, "###,##0.0000")
    End If
    
    txtpeso = IIf(IsNull(TBproducao!peso_metro), "", TBproducao!peso_metro)
    txtUN = IIf(IsNull(TBproducao!Un_Kg), "", TBproducao!Un_Kg)
    txtCliente.Text = IIf(IsNull(TBproducao!Cliente), "", TBproducao!Cliente)
                                      'ORDEM             QTDE. PREVISTA                                      QTDE. OK                                                    QT. PROD.(OK+NC)                                                                                                     CUSTO LOTE                                              CUSTO PEÇA                                      CUSTO TERCEIROS                                             CUSTO MATERIAL                                                CUSTO OUTRAS                                              ORDEM CONSIGNADA
    txtvlrUnit = FunCalculaValorUnitOrdem(TBproducao!Ordem, IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant), IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd), IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd) + IIf(IsNull(TBproducao!QuantNC), 0, TBproducao!QuantNC), IIf(IsNull(TBproducao!CTTReal), 0, TBproducao!CTTReal), IIf(IsNull(TBproducao!CPR), 0, TBproducao!CPR), IIf(IsNull(TBproducao!CTServico), 0, TBproducao!CTServico), IIf(IsNull(TBproducao!CTMaterial), 0, TBproducao!CTMaterial), IIf(IsNull(TBproducao!CTOutras), 0, TBproducao!CTOutras), TBproducao!Consignacao)
    
Else
    USMsgBox ("Não foi encontrado nenhuma ordem validada com este número."), vbExclamation, "CAPRIND v5.0"
    txtLote.SetFocus
    TBproducao.Close
    FunPuxaDadosOrdem = False
    Exit Function
End If
TBproducao.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Function FunPuxaDadosEstoque() As Boolean
On Error GoTo tratar_erro

FunPuxaDadosEstoque = True
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select EC.*, P.Un_Kg from Estoque_Controle EC INNER JOIN projproduto P ON EC.Desenho = P.Desenho where EC.IDestoque = " & txtLote.Text & " and EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    If IsNull(TBproducao!Ref) = False And TBproducao!Ref <> "" Then
        With cmbReferencia
            .Clear
            .AddItem TBproducao!Ref
            .Text = TBproducao!Ref
            .Locked = True
            .TabStop = False
        End With
    End If
    txtCodigo.Text = IIf(IsNull(TBproducao!Desenho), "", TBproducao!Desenho)
    txtdescricao = IIf(IsNull(TBproducao!Descricao), "", TBproducao!Descricao)
    txtunidade = IIf(IsNull(TBproducao!Un), "", TBproducao!Un)
    txtQtde.Text = Format(IIf(IsNull(TBproducao!estoque_real), 0, TBproducao!estoque_real), "###,##0.0000")
    txtpeso = IIf(IsNull(TBproducao!peso_unit), "", TBproducao!peso_unit)
    txtUN = IIf(IsNull(TBproducao!Un_Kg), "", TBproducao!Un_Kg)
    If IsNull(TBproducao!local_armaz) = False Then txtLocal_armaz = TBproducao!local_armaz
    txtcorrida = IIf(IsNull(TBproducao!Corrida), "", TBproducao!Corrida)
    txtCertificado = IIf(IsNull(TBproducao!Certificado), "", TBproducao!Certificado)
    Txt_numero_serie = IIf(IsNull(TBproducao!Numero_serie), "", TBproducao!Numero_serie)
    txtCliente.Text = IIf(IsNull(TBproducao!Cliente), "", TBproducao!Cliente)
    txtvlrUnit = Format(IIf(IsNull(TBproducao!valor_unitario), 0, TBproducao!valor_unitario), "###,##0.0000000000")
Else
    USMsgBox ("Não foi encontrado nenhuma RE com este número."), vbExclamation, "CAPRIND v5.0"
    txtLote.SetFocus
    TBproducao.Close
    FunPuxaDadosEstoque = False
    Exit Function
End If
TBproducao.Close

Exit Function
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o local de armazenamento deste RE, favor revisar."), vbExclamation, "CAPRIND v5.0"
        FunPuxaDadosEstoque = False
        Exit Function
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcSalvar()
On Error GoTo tratar_erro

Acao = "salvar"

'If Opt_devolucao.Value = True And txtCliente.Text = "" Then
'    USMsgBox "Favor informar o cliente que está fazendo a devolução cadastrando a nota fiscal no sistema", vbInformation, "CAPRIND v5.0"
'    Exit Sub
'End If

If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtLote.Text = "" Then
    If Opt_devolucao.Value = True Then NomeCampo = "o número do RE" Else NomeCampo = "o número do lote"
    ProcVerificaAcao
    txtLote.SetFocus
    Exit Sub
End If
If txtCodigo.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodigo.SetFocus
    Exit Sub
End If

Dataini = txtData
If Dataini > Date Then
    USMsgBox ("A data da entrada não pode ser maior que a data de hoje."), vbExclamation, "CAPRIND v5.0"
    txtData = Date
    Exit Sub
End If

If txtLocal_armaz = "" Then
    NomeCampo = "o local de armazenamento"
    ProcVerificaAcao
    txtLocal_armaz.SetFocus
    Exit Sub
End If
If Opt_devolucao.Value = True Then
    If txtNota_Fiscal = "" Then
        NomeCampo = "a nota fiscal"
        ProcVerificaAcao
        txtNota_Fiscal.SetFocus
        Exit Sub
    End If
    txtNota_Fiscal = FunTamanhoTextoZeroEsq(txtNota_Fiscal, 9)
End If
If txtcorrida.Text = "" Then txtcorrida = 0
If txtCertificado.Text = "" Then txtCertificado = 0
valor = IIf(txtvlrUnit = "", 0, txtvlrUnit)
If valor < 0 Or txtvlrUnit = "" Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    txtvlrUnit.SetFocus
    Exit Sub
End If
If txtResponsavel = "" Then
    NomeCampo = "o responsável"
    ProcVerificaAcao
    txtResponsavel.SetFocus
    Exit Sub
End If
Qtde = IIf(txtEntrada = "", 0, txtEntrada)
If Qtde <= 0 Then
    NomeCampo = "a quantidade de entrada"
    ProcVerificaAcao
    txtEntrada.SetFocus
    Exit Sub
End If

If opt_ordem.Value = True Then
    valor = IIf(txtQtde_produzida = "", 0, txtQtde_produzida)
    Valor1 = IIf(txtSaldo = "", 0, txtSaldo)
    If valor > 0 And Qtde > Valor1 Then
        USMsgBox ("A quantidade de entrada não pode ser maior que o saldo, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtEntrada.SetFocus
        Exit Sub
    End If
    If Valor1 <= 0 Then
        USMsgBox ("Não é permitido efetuar a entrada no estoque, pois não existe mais saldo nesta ordem."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    'Verifica se o código de referencia está vinculado a outro produto
    'If cmbreferencia <> "" Then If FunVerifiCodRefUtilizado(txtCodigo, cmbreferencia) = True Then Exit Sub
End If

ValorTotal = 0
quantestoque = 0
If Opt_devolucao.Value = True Then TextoFiltro = "IDestoque = " & txtLote Else TextoFiltro = "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and desenho = '" & txtCodigo.Text & "' and lote = '" & txtLote.Text & "' and local_armaz = '" & txtLocal_armaz & "' and corrida = '" & txtcorrida.Text & "' and certificado = '" & txtCertificado.Text & "'"
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = True Then
    TBEstoque.AddNew
    TBEstoque!LOTE = txtLote.Text
End If


Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBProduto.AddNew

TBProduto!Destino = "Interno"
TBProduto!Terceiros = False

TBProduto!Bloqueado = False
TBProduto!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBProduto!Saida = 0

TBEstoque!Bloqueado = False
TBEstoque!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)


TBProduto!LOTE = TBEstoque!LOTE
    
If Opt_devolucao.Value = True Then TBProduto!Documento = txtNota_Fiscal.Text Else TBProduto!Documento = txtLote.Text
TBEstoque!Desenho = txtCodigo.Text
TBProduto!Desenho = txtCodigo.Text

'Atualiza valor do produto no estoque
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select Codproduto, Estoque, classe from projproduto where desenho = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    If TBItem!Estoque = True Then ControlaEstoque = True Else ControlaEstoque = False
    TBEstoque!Classe = TBItem!Classe
    TBProduto!Familia = TBItem!Classe
    
    'Grava código de referência no produto
    If cmbReferencia <> "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where Codproduto = " & TBItem!Codproduto & " and n_referencia = '" & cmbReferencia & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then TBAbrir.AddNew
        TBAbrir!Codproduto = TBItem!Codproduto
        TBAbrir!N_referencia = cmbReferencia
        TBAbrir!Descricao = IIf(txtdescricao = "", Null, txtdescricao)
        TBAbrir.Update
        TBAbrir.Close
    End If
    
    TBEstoque!Ref = IIf(cmbReferencia = "", Null, cmbReferencia)
End If
TBItem.Close

ValorTotal = txtvlrUnit

'Estoque_movimentação
quantestoque = txtEntrada
TBProduto!VlrUnit = Format(ValorTotal, "###,##0.0000000000")
TBProduto!vlrTotal = Format(ValorTotal * quantestoque, "###,##0.00")

TBEstoque!Descricao = txtdescricao.Text
TBProduto!Descricao = txtdescricao.Text
TBProduto!Data = txtData.Value
TBEstoque!Data = txtData.Value
TBEstoque!Responsavel = txtResponsavel.Text
TBProduto!Responsavel = pubUsuario
TBEstoque!Certificado = txtCertificado.Text
TBEstoque!Numero_serie = Txt_numero_serie
TBEstoque!Corrida = txtcorrida.Text
If txtNota_Fiscal <> "" Then TBEstoque!NF = txtNota_Fiscal
If txtpeso <> "" Then TBEstoque!peso_unit = txtpeso
If txtunidade <> "" Then TBEstoque!Un = txtunidade
If txtLocal_armaz <> "" Then TBEstoque!local_armaz = txtLocal_armaz
TBProduto!Entrada = txtEntrada.Text
TBProduto!Entrada_PC = txtEntrada.Text
TBEstoque!Qtde = txtEntrada.Text
Qtde = txtQtde.Text
Entrada = txtEntrada.Text
If opt_ordem.Value = True Then
    If Entrada >= Qtde Then StatusTexto = "ENTRADA_ORDEM"
    If Entrada < Qtde Then StatusTexto = "ENTRADA_ORDEM_PARCIAL"
    TBEstoque!status = StatusTexto
    TBProduto!Operacao = StatusTexto
ElseIf Opt_devolucao.Value = True Then
        TBEstoque!status = "ENTRADA_DEVOLUÇÃO"
        TBProduto!Operacao = "ENTRADA_DEVOLUÇÃO"
    Else
        TBEstoque!status = "ENTRADA_OUTRAS"
        TBProduto!Operacao = "ENTRADA_OUTRAS"
End If
If txtCliente <> "" Then TBEstoque!Cliente = txtCliente.Text

If ControlaEstoque = True Then QtdeEstoque = txtestoqueatualizado Else QtdeEstoque = 0
TBProduto!estoque_venda = QtdeEstoque
TBProduto!Obs = IIf(txtObservacoes = "", Null, txtObservacoes)
TBProduto.Update

Conexao.Execute "update Estoque_movimentacao Set ID_Tipo = projproduto.ID_Tipo from Estoque_movimentacao inner join projproduto on Estoque_movimentacao.Desenho= projproduto.desenho where Estoque_movimentacao.Idoperacao=" & TBProduto!IDoperacao

Valor1 = txtvlrTotal
If Valor1 > 0 And Opt_devolucao.Value = True Then
    'Verifica se tem CC adicionado no produto
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select codproduto, ID_CC, ID_PC from projproduto where Desenho = '" & txtCodigo & "' and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If TBAbrir!ID_CC <> "" Then
            ProcSalvarCCRealizado txtData, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "Débito", TBAbrir!ID_CC, TBAbrir!ID_CC, IIf(IsNull(TBAbrir!ID_PC), 0, TBAbrir!ID_PC), TBProduto!IDoperacao, 0, Valor1, False, False
        End If
    End If
    TBAbrir.Close
End If

TBEstoque!estoque_real = QtdeEstoque
TBEstoque!estoque_real_PC = QtdeEstoque
TBEstoque!estoque_venda = QtdeEstoque
TBEstoque!valor_unitario = Format(ValorTotal, "###,##0.0000000000")
TBEstoque!Valor_total = Format(QtdeEstoque * ValorTotal, "###,##0.00")

'Verifica se a ordem é consignada
If IsNumeric(txtLote) = True Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from producao where Ordem = " & txtLote.Text & " and Status <> 'Cancelada' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        TBEstoque!ID_Cliente = TBproducao!IDCliente
        TBEstoque!Cliente = TBproducao!Cliente
        If TBproducao!Consignacao = True Then TBEstoque!Consignacao = True Else TBEstoque!Consignacao = False
    End If
    TBproducao.Close
End If
TBEstoque.Update
'===================================
' Atualiza tipo item na movimentação
'===================================
Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDEstoque & " where IDoperacao = " & TBProduto!IDoperacao
'===================================
IDEstoque = TBEstoque!IDEstoque

If opt_ordem.Value = True Then ProcEmpenhaProdutoeAtualQtdeEntEmpOrdem txtLote, txtCodigo, Entrada, TBEstoque!IDEstoque

'==================================
Modulo = "Estoque/Movimentação/Entrada"
Evento = "Entrar"
ID_documento = IDEstoque
Documento = "Cód. interno: " & txtCodigo & " - Lote: " & txtLote & " - Corrida: " & txtcorrida & " - Certificado: " & txtCertificado & " - Local armaz.: " & txtLocal_armaz & " - Qtde.: " & Format(txtEntrada, "###,##0.0000")
Documento1 = ""
ProcGravaEvento
'==================================

If opt_ordem.Value = True Then
    ProcEmpenharREAutomOrdem TBEstoque!IDEstoque, txtEntrada, txtLote, txtData, txtResponsavel, txtCodigo, False
    
    'Marca ordem como concluida
    Qtd = txtQtde_entrada
    qt = txtEntrada
    Qtde = txtQtde
    qt = Qtd + qt
    Qtde = Format(Qtde - qt, "###,##0.0000")
    If Qtde <= 0 Then
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select NOF, Ordem, dataentrega, Concluida, pronta, Status, Ap_backup from producao where Ordem = " & txtLote.Text & " and Concluida = 'False'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            If USMsgBox("Deseja encerrar a ordem " & txtLote.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                If TBCiclo!AP_backup = True Then NomeTabelaAp = "ProducaoFases_Backup" Else NomeTabelaAp = "ProducaoFases"
                
                TBCiclo!DataEntrega = Date
                TBCiclo!Concluida = True
                TBCiclo!pronta = "SIM"
                If TBCiclo!status <> "Entregue" Then TBCiclo!status = "Concluída"
                TBCiclo.Update
                
                Set TBproducao = CreateObject("adodb.recordset")
                TBproducao.Open "Select IDProducao, maquina, Pronto, DataConclusao, Status from ordemservico where Ordem = " & TBCiclo!Ordem & " and pronto = 'NÃO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBproducao.EOF = False Then
                    Do While TBproducao.EOF = False
                        TBproducao!Pronto = "SIM"
                        TBproducao!DataConclusao = Date
                        TBproducao!status = "Concluída"
                        TBproducao.Update
                        'Filtra todos os eventos desta OS na tabela producaofases para marcar como fase pronta
                        Conexao.Execute "Update " & NomeTabelaAp & " Set pronto = 'SIM' where idfase = " & TBproducao!IDProducao
                        'Libera máquina
                        Conexao.Execute "Update cadmaquinas Set cadmaquinas.Liberada = 'True' from cadmaquinas INNER JOIN cadmaquinas_Monitor ON cadmaquinas.Maquina = cadmaquinas_Monitor.Maquina where cadmaquinas_Monitor.maquina = '" & TBproducao!maquina & "' and cadmaquinas_Monitor.OS = " & TBproducao!IDProducao
                        TBproducao.MoveNext
                    Loop
                End If
                TBproducao.Close
                '==================================
                Modulo = "Estoque/Movimentação/Entrada"
                Evento = "Alterar ordem p/ concluída"
                ID_documento = TBCiclo!NOF
                Documento = "Ordem: " & TBCiclo!Ordem
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
        End If
        TBCiclo.Close
    End If
End If
USMsgBox ("Produto acrescentado ao estoque com sucesso."), vbInformation, "CAPRIND v5.0"
cmdlote_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarCCRealizado(Data1 As Date, ID_empresa As Integer, Operacao As String, ID_CC As Long, Cod_produto As Long, ID_plano_contas As Long, ID_estoque As Long, ID_lista As Long, valor As Double, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

ProcEnviaDadosCCRealizado Data, ID_empresa, Operacao, ID_CC, Cod_produto, ID_plano_contas, ID_estoque, ID_lista, valor, CC_produto, Bloqueado

'Grava movimentação no centro consolidado
Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & ID_CC, Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
    Do While TBAfericao.EOF = False
        ProcEnviaDadosCCRealizado Data, ID_empresa, Operacao, TBAfericao!ID_CC, Cod_produto, ID_plano_contas, ID_estoque, ID_lista, valor, CC_produto, Bloqueado
       
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                ProcEnviaDadosCCRealizado Data, ID_empresa, Operacao, TBCiclo!ID_CC, Cod_produto, ID_plano_contas, ID_estoque, ID_lista, valor, CC_produto, Bloqueado
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

Private Sub ProcEnviaDadosCCRealizado(Data1 As Date, ID_empresa As Integer, Operacao As String, ID_CC As Long, Cod_produto As Long, ID_plano_contas As Long, ID_estoque As Long, ID_lista As Long, valor As Double, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

NovoValor = Replace(valor, ",", ".")
ProcINSERTINTO "CC_realizado", "Data, Responsavel, ID_empresa, Operacao, ID_CC, Cod_produto, ID_PC, ID_estoque, ID_lista, Valor, Bloqueado", "'" & Data & "', '" & pubUsuario & "', " & ID_empresa & ", '" & Operacao & "', " & ID_CC & ", " & Cod_produto & ", " & ID_plano_contas & ", " & IIf(ID_estoque = 0, "NULL", ID_estoque) & ", " & ID_lista & ", " & NovoValor & ", " & IIf(Bloqueado = True, 1, 0) & ""

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

Private Sub cmdNF_Click()
On Error GoTo tratar_erro

If txtNota_Fiscal.Text <> "" Then
     VerifNumero = txtNota_Fiscal.Text
     ProcVerificaNumero
     If VerifNumero = False Then
         txtNota_Fiscal.Text = ""
         txtNota_Fiscal.SetFocus
         Exit Sub
     Else
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "Select DN.int_NotaFiscal,  DN.txt_Razao_Nome, DTN.int_Cod_Produto, DTN.int_Qtd, DTN.dbl_ValorUnitario, DTN.dbl_ValorTotal from tbl_Dados_Nota_Fiscal DN inner join tbl_Detalhes_Nota DTN on DTN.ID_Nota = DN.ID where DN.int_NotaFiscal = '" & txtNota_Fiscal.Text & "' and DTN.int_Cod_produto = '" & txtCodigo & "'"
    'Debug.print StrSql
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
    txtCliente.Text = TBAbrir!txt_Razao_Nome
    txtEntrada = TBAbrir!int_Qtd
    txtvlrUnit = TBAbrir!dbl_ValorUnitario
    'TxtVlrTotal = TBAbrir!dbl_ValorTotal
    Else
    USMsgBox "Nota fiscal não encontrada, favor executar a entrada da nota no sistema antes de fazer a devolução", vbCritical, "CAPRIND v5.0"
    txtCliente.Text = ""
    End If
    TBAbrir.Close
     
     End If
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
    Case vbKeyF3: ProcSalvar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
txtData.Value = Date
txtResponsavel.Text = pubUsuario
ProcCarregaComboEmpresa Cmb_empresa, False

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proccarregalocarm()
On Error GoTo tratar_erro

With txtLocal_armaz
    .Clear
    If txtCodigo.Text <> "" Then
        LATexto = ""
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select ELAC.Descricao, ELA.Padrao from Estoque_Localarmazenamento ELA INNER JOIN Estoque_Localarmazenamento_criar ELAC ON ELAC.ID = ELA.idemb_locarm where ELA.codinterno = '" & txtCodigo & "' and ELAC.DtValidacao IS NOT NULL and ELAC.Status = 'Liberado'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                If IsNull(TBFI!Descricao) = False Then
                    .AddItem TBFI!Descricao
                    If TBFI!Padrao = True Then LATexto = TBFI!Descricao
                End If
                TBFI.MoveNext
            Loop
            
            Set TBTempo = CreateObject("adodb.recordset")
            TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Carregar_LAentrada = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBTempo.EOF = False Then
                Set TBAliquota = CreateObject("adodb.recordset")
                TBAliquota.Open "Select ELAC.Descricao from (Estoque_Localarmazenamento ELA RIGHT JOIN Estoque_Localarmazenamento_criar ELAC ON ELAC.ID = ELA.idemb_locarm) LEFT JOIN Estoque_Controle EC on EC.local_armaz = ELAC.Descricao where ELA.ID IS NULL group by ELAC.Descricao HAVING SUM(ISNULL(EC.Estoque_Real, 0)) = 0 ", Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.EOF = False Then
                    Do While TBAliquota.EOF = False
                        If IsNull(TBAliquota!Descricao) = False Then .AddItem TBAliquota!Descricao
                        TBAliquota.MoveNext
                    Loop
                End If
                TBAliquota.Close
            End If
            TBTempo.Close
            
            If LATexto <> "" Then
                .Text = LATexto
            Else
                TBFI.MoveFirst
                .Text = TBFI!Descricao
            End If
            
        Else
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Descricao from Estoque_Localarmazenamento_criar where DtValidacao IS NOT NULL and Status = 'Liberado' order by descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Descricao) = False Then .AddItem TBFI!Descricao
                    TBFI.MoveNext
                Loop
            End If
        End If
        TBFI.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_devolucao_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
If Opt_devolucao.Value = True Then
    Label7(0).Caption = "N° do RE*"
    Label7(3).Caption = "Nota fiscal*"
    Label23.Caption = "Valor unit."
    Cmd_localizar_produtos.Enabled = False
    With txtLote
        .Text = ""
        .ToolTipText = "Número do RE."
    End With
    With txtLocal_armaz
      '  .Locked = True
        .TabStop = False
    End With
    With txtcorrida
        .Locked = True
        .TabStop = False
    End With
    With txtCertificado
        .Locked = True
        .TabStop = False
    End With
    With Txt_numero_serie
        .Locked = True
        .TabStop = False
    End With
    With txtNota_Fiscal
        .Locked = False
        .TabStop = True
    End With
    With txtvlrUnit
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt_ordem_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
If opt_ordem = True Then
    Label7(0).Caption = "N° do lote*"
    Label7(3).Caption = "Nota fiscal"
    Label23.Caption = "Valor unit."
    txtLote.ToolTipText = "Número do lote."
    Cmd_localizar_produtos.Enabled = False
    With txtLocal_armaz
        .Locked = False
        .TabStop = True
    End With
    With txtcorrida
        .Locked = False
        .TabStop = True
    End With
    With txtCertificado
        .Locked = False
        .TabStop = True
    End With
    With Txt_numero_serie
        .Locked = False
        .TabStop = True
    End With
    With txtNota_Fiscal
        .Locked = True
        .TabStop = False
    End With
    With txtvlrUnit
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_outras_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
If Opt_outras = True Then
    Label7(0).Caption = "N° do lote*"
    Label7(3).Caption = "Nota fiscal"
    Label23.Caption = "Valor unit.*"
    txtLote.ToolTipText = "Número do lote."
    Cmd_localizar_produtos.Enabled = True
    With txtLocal_armaz
        .Locked = False
        .TabStop = True
    End With
    With txtcorrida
        .Locked = False
        .TabStop = True
    End With
    With txtCertificado
        .Locked = False
        .TabStop = True
    End With
    With Txt_numero_serie
        .Locked = False
        .TabStop = True
    End With
    With txtNota_Fiscal
        .Locked = True
        .TabStop = False
    End With
    With txtvlrUnit
        .Locked = False
        .TabStop = True
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCertificado_Change()
On Error GoTo tratar_erro

ProcCalculaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

Proccarregalocarm

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcorrida_Change()
On Error GoTo tratar_erro

ProcCalculaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTotais()
On Error GoTo tratar_erro

If txtLote <> "" And txtLocal_armaz <> "" Then
    Pesoestoque = 0
    Qtd = 0
    If Opt_devolucao.Value = True Then TextoFiltro = "IDestoque = " & txtLote Else TextoFiltro = "desenho = '" & txtCodigo.Text & "' AND lote = '" & txtLote.Text & "' AND certificado = '" & IIf(txtCertificado = "", 0, txtCertificado) & "' and corrida = '" & IIf(txtcorrida = "", 0, txtcorrida) & "' and local_armaz = '" & txtLocal_armaz & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    Set TBMaterial = CreateObject("adodb.recordset")
    TBMaterial.Open "Select Estoque_real from estoque_controle where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBMaterial.EOF = False Then
        Pesoestoque = IIf(IsNull(TBMaterial!estoque_real), 0, TBMaterial!estoque_real)
    End If
    TBMaterial.Close
    txtestoquereal = Format(Pesoestoque, "###,##0.0000")
    
    'Estoque atualizado
    Pesolote = IIf(txtEntrada.Text = "", 0, txtEntrada.Text)
    txtestoqueatualizado = Format(Pesoestoque + Pesolote, "###,##0.0000")
    
    'Valor total
    qt = txtestoqueatualizado
    Qtde = IIf(txtvlrUnit = "", 0, txtvlrUnit)
    txtvlrTotal = Format(Qtde * qt, "###,##0.00")
        
    With txtEntrada
        .Locked = False
        .TabStop = True
    End With
Else
    txtestoquereal.Text = "0,0000"
    txtestoqueatualizado = "0,0000"
    With txtEntrada
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtentrada_Change()
On Error GoTo tratar_erro

If txtEntrada.Text <> "" Then
    VerifNumero = txtEntrada.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtEntrada.Text = ""
        txtEntrada.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtData.Value = Format(Date, "dd/mm/yyyy")
txtResponsavel = pubUsuario
txtCodigo = ""
txtQtde = "0,0000"
txtQtde_produzida = "0,0000"
Txt_qtde_refugada = "0,0000"
txtQtde_entrada = "0,0000"
txtSaldo = "0,0000"
cmbReferencia.Clear
txtpeso = ""
txtUN = ""
txtLocal_armaz.ListIndex = -1
txtdescricao = ""
txtunidade.Text = ""
txtCertificado = ""
Txt_numero_serie = ""
txtcorrida = ""
txtCliente = ""
txtNota_Fiscal = ""
txtObservacoes = ""
txtEntrada = "0,0000"
txtestoquereal = "0,0000"
txtestoqueatualizado = "0,0000"
txtvlrUnit = "0,0000"
txtvlrTotal = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtentrada_GotFocus()
On Error GoTo tratar_erro

If txtEntrada.Text = "0,0000" Then txtEntrada.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtentrada_LostFocus()
On Error GoTo tratar_erro

txtEntrada = IIf(txtEntrada.Text = "", "0,0000", Format(txtEntrada.Text, "###,##0.0000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLocal_armaz_Click()
On Error GoTo tratar_erro

ProcCalculaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLote_Change()
On Error GoTo tratar_erro

If txtLote <> "" Then
    VerifNumero = txtLote
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLote = ""
        txtLote.SetFocus
        Exit Sub
    End If
End If
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNota_Fiscal_LostFocus()
On Error GoTo tratar_erro

If Opt_devolucao.Value = True And txtNota_Fiscal <> "" Then txtNota_Fiscal = FunTamanhoTextoZeroEsq(txtNota_Fiscal, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlrunit_Change()
On Error GoTo tratar_erro

txtvlrTotal.Text = "0,00"
If txtvlrUnit.Text <> "" Then
    VerifNumero = txtvlrUnit.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvlrUnit.Text = ""
        txtvlrUnit.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlrunit_GotFocus()
On Error GoTo tratar_erro

If txtvlrUnit.Text = "0,0000000000" Then txtvlrUnit.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlrunit_LostFocus()
On Error GoTo tratar_erro

txtvlrUnit = IIf(txtvlrUnit.Text = "", "0,0000", Format(txtvlrUnit.Text, "###,##0.0000000000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

