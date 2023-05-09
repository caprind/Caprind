VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQualidadePPAP_PlanoControle 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - PPAP - Plano de controle"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11880
   Begin MSComctlLib.ProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   85
      Top             =   6960
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton imgSair 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10995
      MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmQualidadePPAP_PlanoControle.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Sair (Esc)"
      Top             =   510
      Width           =   630
   End
   Begin VB.CommandButton cmdAjuda 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10350
      MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":0925
      MousePointer    =   99  'Custom
      Picture         =   "frmQualidadePPAP_PlanoControle.frx":0A77
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Ajuda (F1)"
      Top             =   510
      Width           =   630
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados principais"
      TabPicture(0)   =   "frmQualidadePPAP_PlanoControle.frx":0F19
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtIDProduto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Dimensões"
      TabPicture(1)   =   "frmQualidadePPAP_PlanoControle.frx":0F35
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIDdimensoes"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Lista2"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtIDdimensoes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73560
         TabIndex        =   98
         Text            =   "0"
         Top             =   5880
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   74
         Text            =   "0"
         Top             =   6090
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtIDProduto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1500
         TabIndex        =   73
         Text            =   "0"
         Top             =   6090
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   -74925
         TabIndex        =   71
         Top             =   330
         Width           =   11745
         Begin VB.CommandButton cmdImprimir_resultado_dimensional1 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   1440
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":0F51
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":10A3
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Visualizar impressão resultado dimensional (F6)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdProximo2 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2700
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":1867
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":19B9
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Próximo registro."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdanterior2 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2070
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":1D02
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":1E54
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Registro anterior."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdImprimirDimensoes 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   810
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":219D
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":22EF
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Visualizar impressão (F5)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdSalvarDimensoes 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   180
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":2ADE
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":2C30
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Salvar (F3)"
            Top             =   180
            Width           =   630
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   75
         TabIndex        =   60
         Top             =   330
         Width           =   11745
         Begin VB.CommandButton cmdImprimir_resultado_dimensional 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   3330
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":3409
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":355B
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Visualizar impressão resultado dimensional (F6)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdCopiar 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   5220
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":3D1F
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":3E71
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Copiar (F7)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdRevisao 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   5850
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":4373
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":44C5
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Revisar (F8)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdSalvar 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   1440
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":4968
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":4ABA
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Salvar (F3)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdExcluir 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2070
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":5293
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":53E5
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Excluir (F4)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdNovo 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   180
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":5C34
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":5D86
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Novo (Insert)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdLocalizar 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   810
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":62AC
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":63FE
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Localizar (F2)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdImprimir 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2700
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":6BBF
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":6D11
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Visualizar impressão (F5)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdanterior 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   3960
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":7500
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":7652
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Registro anterior."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdProximo 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   4590
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":799B
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":7AED
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Próximo registro."
            Top             =   180
            Width           =   630
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   1785
         Left            =   75
         TabIndex        =   25
         Top             =   5160
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   3149
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
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "P. controle"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   1058
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
            Text            =   "Descrição"
            Object.Width           =   8035
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   2895
         Left            =   -74925
         TabIndex        =   72
         Top             =   1200
         Width           =   11745
         Begin VB.CheckBox chkOK 
            BackColor       =   &H00E0E0E0&
            Caption         =   "OK"
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
            Left            =   11010
            TabIndex        =   52
            Top             =   2490
            Width           =   555
         End
         Begin VB.TextBox txtResultados 
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
            MaxLength       =   255
            TabIndex        =   51
            ToolTipText     =   "Resultados (dados) das madições do fornecedor."
            Top             =   2430
            Width           =   7215
         End
         Begin VB.TextBox txtQtdeEnsaiado 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   30
            TabIndex        =   50
            ToolTipText     =   "Quantidade ensaiada."
            Top             =   2430
            Width           =   1905
         End
         Begin VB.TextBox txtPlanoReacao 
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
            Left            =   10080
            MaxLength       =   50
            TabIndex        =   48
            ToolTipText     =   "Plano de reação."
            Top             =   1785
            Width           =   1485
         End
         Begin VB.TextBox txtMetodoControle 
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
            Left            =   8580
            MaxLength       =   50
            TabIndex        =   47
            ToolTipText     =   "Método de controle."
            Top             =   1785
            Width           =   1485
         End
         Begin VB.TextBox txtAmostraFreq 
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
            Left            =   6090
            MaxLength       =   50
            TabIndex        =   46
            ToolTipText     =   "Amostra frequencia."
            Top             =   1785
            Width           =   2475
         End
         Begin VB.TextBox txtProcesso 
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
            Left            =   7260
            MaxLength       =   50
            TabIndex        =   41
            ToolTipText     =   "Característica processo."
            Top             =   390
            Width           =   4275
         End
         Begin VB.TextBox txtProduto 
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
            Left            =   4950
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Característica produto."
            Top             =   390
            Width           =   2295
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Característica número."
            Top             =   390
            Width           =   1335
         End
         Begin VB.TextBox txtTecnicaAvaliacao 
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
            TabIndex        =   45
            ToolTipText     =   "Técnica de avaliação/medição"
            Top             =   1785
            Width           =   2385
         End
         Begin VB.TextBox txtEspc 
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
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Especificação/tolerancia do produto/processo"
            Top             =   1785
            Width           =   2265
         End
         Begin VB.TextBox txtCarac 
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
            MaxLength       =   50
            TabIndex        =   43
            ToolTipText     =   "Característica esp."
            Top             =   1785
            Width           =   1215
         End
         Begin VB.TextBox txtInstrucao 
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
            Height          =   465
            Left            =   180
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":7E36
            MousePointer    =   99  'Custom
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            ToolTipText     =   "Nome do processo/descrição da operação."
            Top             =   1008
            Width           =   11355
         End
         Begin VB.TextBox txtGrupo 
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
            Left            =   930
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Grupo/operação."
            Top             =   390
            Width           =   1125
         End
         Begin VB.TextBox txtPosto 
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
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Posto de trabalho."
            Top             =   390
            Width           =   1515
         End
         Begin VB.TextBox txtFase 
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
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Fase."
            Top             =   390
            Width           =   735
         End
         Begin MSMask.MaskEdBox txtDataEnsaio 
            Height          =   315
            Left            =   180
            TabIndex        =   49
            ToolTipText     =   "Prazo de entrega."
            Top             =   2430
            Width           =   1215
            _ExtentX        =   2143
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
         Begin VB.Image ImgCalendario1 
            Height          =   360
            Left            =   1390
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":8140
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":8292
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   2400
            Width           =   330
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data ensaio"
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
            Left            =   360
            TabIndex        =   101
            Top             =   2220
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. ensaiada"
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
            Left            =   2167
            TabIndex        =   100
            Top             =   2220
            Width           =   1110
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resultados (dados) das medições do fornecedor"
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
            Left            =   5565
            TabIndex        =   99
            Top             =   2220
            Width           =   3465
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plano de reação"
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
            Left            =   10245
            TabIndex        =   97
            Top             =   1590
            Width           =   1155
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Método de controle"
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
            Left            =   8625
            TabIndex        =   96
            Top             =   1590
            Width           =   1395
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amostra freq."
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
            Left            =   6825
            TabIndex        =   95
            Top             =   1590
            Width           =   1005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Característica processo"
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
            Left            =   8557
            TabIndex        =   94
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Característica produto"
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
            Left            =   5295
            TabIndex        =   93
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carac. número"
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
            Left            =   3735
            TabIndex        =   92
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Técnica de avaliação/medição"
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
            Left            =   3810
            TabIndex        =   91
            Top             =   1590
            Width           =   2130
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Espec./toler. do prod./proc."
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
            Left            =   1530
            TabIndex        =   90
            Top             =   1590
            Width           =   2025
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carac. esp."
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
            Left            =   360
            TabIndex        =   89
            Top             =   1590
            Width           =   840
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do processo/descrição da operação"
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
            Left            =   4357
            TabIndex        =   88
            Top             =   810
            Width           =   3000
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo/op."
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
            Left            =   1125
            TabIndex        =   87
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posto de trabalho"
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
            Left            =   2190
            TabIndex        =   86
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fase"
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
            Left            =   375
            TabIndex        =   75
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   3945
         Left            =   75
         TabIndex        =   61
         Top             =   1200
         Width           =   11745
         Begin VB.TextBox Txt_doc_engenharia 
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
            Left            =   8280
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":8715
            MousePointer    =   99  'Custom
            TabIndex        =   23
            ToolTipText     =   "Documentos de alteração da engenharia."
            Top             =   3480
            Width           =   3255
         End
         Begin VB.TextBox Txt_local_inspecao 
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
            Left            =   5010
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":8A1F
            MousePointer    =   99  'Custom
            TabIndex        =   22
            ToolTipText     =   "Local de inspeção."
            Top             =   3480
            Width           =   3255
         End
         Begin VB.ComboBox Cmb_tipo 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmQualidadePPAP_PlanoControle.frx":8D29
            Left            =   6150
            List            =   "frmQualidadePPAP_PlanoControle.frx":8D36
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":8D5F
            MousePointer    =   99  'Custom
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Tipo."
            Top             =   390
            Width           =   1580
         End
         Begin VB.CommandButton cmdFiltrar_codigo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1920
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":9069
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":91BB
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Filtrar por código interno."
            Top             =   1008
            Width           =   315
         End
         Begin VB.TextBox txtOutraAprovacao2 
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
            Left            =   2610
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":95D6
            MousePointer    =   99  'Custom
            TabIndex        =   21
            ToolTipText     =   "Outra aprovação/data."
            Top             =   3480
            Width           =   2385
         End
         Begin VB.TextBox txtOutraAprovacao 
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
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":98E0
            MousePointer    =   99  'Custom
            TabIndex        =   20
            ToolTipText     =   "Outra aprovação/data."
            Top             =   3480
            Width           =   2415
         End
         Begin VB.TextBox txtOrganizacaoAprovacao 
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
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":9BEA
            MousePointer    =   99  'Custom
            TabIndex        =   18
            ToolTipText     =   "Organização/aprovação da planta/data."
            Top             =   2862
            Width           =   5715
         End
         Begin VB.TextBox txtCodOrganizacao 
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
            Left            =   5910
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":9EF4
            MousePointer    =   99  'Custom
            TabIndex        =   17
            ToolTipText     =   "Código da organização."
            Top             =   2244
            Width           =   5625
         End
         Begin VB.TextBox txtAprovacaoQualidade 
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
            Left            =   5910
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":A1FE
            MousePointer    =   99  'Custom
            TabIndex        =   19
            ToolTipText     =   "Aprovação da qualidade do cliente/data."
            Top             =   2862
            Width           =   5625
         End
         Begin VB.TextBox txtOrganizacaoPlanta 
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
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":A508
            MousePointer    =   99  'Custom
            TabIndex        =   16
            ToolTipText     =   "Organização/planta."
            Top             =   2244
            Width           =   5715
         End
         Begin VB.TextBox txtAprovacao_engenharia 
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
            Left            =   5910
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":A812
            MousePointer    =   99  'Custom
            TabIndex        =   15
            ToolTipText     =   "Aprovação da engenharia do cliente/data."
            Top             =   1620
            Width           =   5625
         End
         Begin VB.TextBox txtContato 
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
            Left            =   7740
            MaxLength       =   50
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":AB1C
            MousePointer    =   99  'Custom
            TabIndex        =   7
            ToolTipText     =   "Contato pricipal/telefone."
            Top             =   390
            Width           =   3795
         End
         Begin VB.TextBox txtDataRev 
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
            Left            =   2175
            Locked          =   -1  'True
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":AE26
            MousePointer    =   99  'Custom
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data da revisão."
            Top             =   390
            Width           =   960
         End
         Begin VB.TextBox txtDescricaoProduto 
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
            Left            =   5715
            Locked          =   -1  'True
            MaxLength       =   255
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":B130
            MousePointer    =   99  'Custom
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   1008
            Width           =   5820
         End
         Begin VB.TextBox txtRevProduto 
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
            Left            =   2640
            Locked          =   -1  'True
            MaxLength       =   10
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":B43A
            MousePointer    =   99  'Custom
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   1008
            Width           =   530
         End
         Begin VB.TextBox txtCodInterno 
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
            MaxLength       =   50
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":B744
            MousePointer    =   99  'Custom
            TabIndex        =   8
            ToolTipText     =   "Código interno."
            Top             =   1008
            Width           =   1725
         End
         Begin VB.CommandButton cmdLocalizarProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2250
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":BA4E
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_PlanoControle.frx":BBA0
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Localizar código interno."
            Top             =   1008
            Width           =   315
         End
         Begin VB.ComboBox cmbReferencia_prod 
            Appearance      =   0  'Flat
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
            Left            =   3180
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":BCA2
            MousePointer    =   99  'Custom
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Código de referência."
            Top             =   1008
            Width           =   2520
         End
         Begin VB.TextBox txtDataEmissao 
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
            Left            =   3150
            Locked          =   -1  'True
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":BFAC
            MousePointer    =   99  'Custom
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   855
         End
         Begin VB.TextBox txtResp 
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
            Left            =   4020
            Locked          =   -1  'True
            MaxLength       =   50
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":C2B6
            MousePointer    =   99  'Custom
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2115
         End
         Begin VB.TextBox txtPlano 
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
            Left            =   180
            MaxLength       =   15
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":C5C0
            MousePointer    =   99  'Custom
            TabIndex        =   1
            ToolTipText     =   "Número do plano de controle."
            Top             =   390
            Width           =   1395
         End
         Begin VB.TextBox txtRev 
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
            Left            =   1590
            Locked          =   -1  'True
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":C8CA
            MousePointer    =   99  'Custom
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   390
            Width           =   570
         End
         Begin VB.TextBox txtEquipe 
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
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_PlanoControle.frx":CBD4
            MousePointer    =   99  'Custom
            TabIndex        =   14
            ToolTipText     =   "Equipe principal."
            Top             =   1626
            Width           =   5715
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local de inspeção"
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
            Left            =   6150
            TabIndex        =   104
            Top             =   3270
            Width           =   1260
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Docs. de alt. engenharia"
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
            Left            =   9022
            TabIndex        =   103
            Top             =   3270
            Width           =   1770
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
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
            Left            =   6783
            TabIndex        =   102
            Top             =   180
            Width           =   315
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Outra aprovação/data"
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
            Left            =   2992
            TabIndex        =   84
            Top             =   3270
            Width           =   1620
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Outra aprovação/data"
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
            Left            =   577
            TabIndex        =   83
            Top             =   3270
            Width           =   1620
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Organização/aprovação da planta/data"
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
            Left            =   1620
            TabIndex        =   82
            Top             =   2670
            Width           =   2835
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. da organização"
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
            Left            =   7980
            TabIndex        =   81
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aprovação da qualidade do cliente/data"
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
            Left            =   7290
            TabIndex        =   80
            Top             =   2670
            Width           =   2865
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Organização/planta"
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
            Left            =   2332
            TabIndex        =   79
            Top             =   2040
            Width           =   1410
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aprovação da engenharia do cliente/data"
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
            Left            =   7230
            TabIndex        =   78
            Top             =   1410
            Width           =   2985
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contato principal/telefone"
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
            Left            =   8685
            TabIndex        =   77
            Top             =   180
            Width           =   1905
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. revisão"
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
            Left            =   2258
            TabIndex        =   76
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
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
            Left            =   8280
            TabIndex        =   70
            Top             =   810
            Width           =   690
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   2715
            TabIndex        =   69
            Top             =   810
            Width           =   375
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código de referência"
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
            Left            =   3690
            TabIndex        =   68
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno"
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
            Left            =   517
            TabIndex        =   67
            Top             =   810
            Width           =   1050
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Left            =   3390
            TabIndex        =   66
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
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
            Left            =   4605
            TabIndex        =   65
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plano de cont."
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
            Left            =   285
            TabIndex        =   64
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   1695
            TabIndex        =   63
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Equipe principal"
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
            Left            =   2482
            TabIndex        =   62
            Top             =   1410
            Width           =   1110
         End
      End
      Begin MSComctlLib.ListView Lista2 
         Height          =   2835
         Left            =   -74925
         TabIndex        =   54
         Top             =   4110
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   5001
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
         MousePointer    =   99
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
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "IDplano"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Fase"
            Object.Width           =   1060
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Grupo/op."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Posto de trabalho"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Carac. núm."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Carac. produto"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Carac. processo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Carac. esp."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Espec./toler. do prod./proc."
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Téc. de avaliação/medição"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "Amostra freq."
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Tag             =   "T"
            Text            =   "Método de controle"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Plano de reação"
            Object.Width           =   2646
         EndProperty
      End
   End
End
Attribute VB_Name = "frmQualidadePPAP_PlanoControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_PlanoControle As Boolean 'OK
Public SQL_PlanoControle  As String 'OK

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtidproduto = 0
txtPlano = ""
txtRev = 0
txtDataRev = ""
txtDataemissao = Format(Date, "dd/mm/yy")
txtResp = pubUsuario
Cmb_tipo.ListIndex = -1
txtContato = ""
txtCodinterno = ""
txtRevProduto = ""
cmbReferencia_prod.Clear
txtDescricaoProduto = ""
txtEquipe = ""
txtAprovacao_engenharia = ""
txtOrganizacaoPlanta = ""
txtCodOrganizacao = ""
txtOrganizacaoAprovacao = ""
txtAprovacaoQualidade = ""
txtOutraAprovacao = ""
txtOutraAprovacao2 = ""
Txt_local_inspecao = ""
Txt_doc_engenharia = ""
CodigoLista = 0
Caption = "Qualidade - PPAP - Plano de controle"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!Plano = txtPlano
TBGravar!Rev = txtRev.Text
If IsNull(TBGravar!DtEmissao) = True Or TBGravar!DtEmissao = "" Then TBGravar!DtEmissao = Date Else TBGravar!DtEmissao = txtDataemissao
If IsNull(TBGravar!Responsavel) = True Or TBGravar!Responsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResp
TBGravar!Tipo = Cmb_tipo
TBGravar!IDProduto = txtidproduto.Text
TBGravar!contato = txtContato
If Novo_PlanoControle = False Then
    If txtCodinterno.Text <> TBGravar!Codinterno Then
        Conexao.Execute "DELETE from QualidadePPAP_PlanoControle_Dimensoes where idplanoControle = " & txtId
        ProcSalvarPlanoInspecao
    End If
End If
TBGravar!Codinterno = txtCodinterno.Text
TBGravar!N_referencia = cmbReferencia_prod
TBGravar!Aprovacao_engenharia = txtAprovacao_engenharia.Text
TBGravar!Equipe = txtEquipe.Text
TBGravar!Organizacao_Planta = txtOrganizacaoPlanta.Text
TBGravar!CodOrganizacao = txtCodOrganizacao.Text
TBGravar!Organizacao_Aprovacao = txtOrganizacaoAprovacao.Text
TBGravar!Aprovacao_Qualidade = txtAprovacaoQualidade.Text
TBGravar!Outra_Aprovacao = txtOutraAprovacao.Text
TBGravar!Outra_Aprovacao2 = txtOutraAprovacao2.Text
TBGravar!Local_inspecao = Txt_local_inspecao
TBGravar!Doc_engenharia = Txt_doc_engenharia

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDadosDimensoes()
On Error GoTo tratar_erro

TBGravar!Processo = txtProcesso
TBGravar!Descricao = txtInstrucao
TBGravar!Carac = txtCarac.Text
TBGravar!TecnicaAvaliacao = txtTecnicaAvaliacao.Text
TBGravar!AmostraFreq = txtAmostraFreq.Text
TBGravar!MetodoControle = txtMetodoControle.Text
TBGravar!PlanoReacao = txtPlanoReacao.Text
TBGravar!DataEnsaio = IIf(txtDataEnsaio = "__/__/____", Null, txtDataEnsaio)
TBGravar!QtdeEnsaiado = IIf(txtQtdeEnsaiado = "", Null, txtQtdeEnsaiado)
TBGravar!Resultados = txtResultados
If chkOK.Value = 1 Then TBGravar!ok = True Else TBGravar!ok = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposDimensoes()
On Error GoTo tratar_erro

txtIDdimensoes = 0
txtFase.Text = ""
txtGrupo = ""
txtPosto.Text = ""
txtNumero = ""
txtProduto.Text = ""
txtProcesso = ""
txtInstrucao.Text = ""
txtCarac.Text = ""
txtEspc.Text = ""
txtTecnicaAvaliacao.Text = ""
txtAmostraFreq.Text = ""
txtMetodoControle.Text = ""
txtPlanoReacao.Text = ""
txtDataEnsaio = "__/__/____"
txtQtdeEnsaiado = ""
txtResultados = ""
chkOK.Value = 0
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadadosDimensoes()
On Error GoTo tratar_erro

txtIDdimensoes = TBAbrir!ID
txtProcesso.Text = IIf(IsNull(TBAbrir!Processo), "", TBAbrir!Processo)
txtCarac.Text = IIf(IsNull(TBAbrir!Carac), "", TBAbrir!Carac)
txtTecnicaAvaliacao = IIf(IsNull(TBAbrir!TecnicaAvaliacao), "", TBAbrir!TecnicaAvaliacao)
txtAmostraFreq = IIf(IsNull(TBAbrir!AmostraFreq), "", TBAbrir!AmostraFreq)
txtMetodoControle = IIf(IsNull(TBAbrir!MetodoControle), "", TBAbrir!MetodoControle)
txtPlanoReacao = IIf(IsNull(TBAbrir!PlanoReacao), "", TBAbrir!PlanoReacao)
txtDataEnsaio = IIf(IsNull(TBAbrir!DataEnsaio), "__/__/____", Format(TBAbrir!DataEnsaio, "dd/mm/yyyy"))
txtQtdeEnsaiado = IIf(IsNull(TBAbrir!QtdeEnsaiado), "", Format(TBAbrir!QtdeEnsaiado, "###,##0.0000"))
txtResultados = IIf(IsNull(TBAbrir!Resultados), "", TBAbrir!Resultados)
If TBAbrir!ok = True Then chkOK.Value = 1 Else chkOK.Value = 0
txtInstrucao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)

Set TBplanomedicao = CreateObject("adodb.recordset")
TBplanomedicao.Open "Select * from plano where idplano = " & TBAbrir!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
If TBplanomedicao.EOF = False Then
    txtFase = IIf(IsNull(TBplanomedicao!Fase), "", TBplanomedicao!Fase)
    txtGrupo = IIf(IsNull(TBplanomedicao!Grupo_op), "", TBplanomedicao!Grupo_op)
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "select Fases.* from fases INNER JOIN processos on fases.IDProcesso = processos.IDProcesso where processos.CodProduto = " & txtidproduto & " and fases.Fase = " & TBplanomedicao!Fase, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        txtPosto = IIf(IsNull(TBFases!maquina), "", TBFases!maquina)
    End If
    TBFases.Close
End If
Set TBplanomedicao = CreateObject("adodb.recordset")
TBplanomedicao.Open "Select * from Planodimensao where IDdimensao = " & TBAbrir!idDimensao, Conexao, adOpenKeyset, adLockOptimistic
If TBplanomedicao.EOF = False Then
    txtNumero.Text = IIf(IsNull(TBplanomedicao!Numero), "", TBplanomedicao!Numero)
    txtProduto = IIf(IsNull(TBplanomedicao!Tipo), "", TBplanomedicao!Tipo)
    Texto = IIf(IsNull(TBplanomedicao!dimdesejada), "", Format(TBplanomedicao!dimdesejada, "###,##0.0000"))
    Texto1 = IIf(IsNull(TBplanomedicao!TolSup), "", Format(TBplanomedicao!TolSup, "###,##0.0000"))
    Texto2 = IIf(IsNull(TBplanomedicao!TolInf), "", Format(TBplanomedicao!TolInf, "###,##0.0000"))
    txtEspc.Text = Texto & "  " & Texto1 & " / " & Texto2
End If
TBplanomedicao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdAnterior_Click()
On Error GoTo tratar_erro

ProcAnterior

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdAnterior2_Click()
On Error GoTo tratar_erro

ProcAnterior

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCopiar_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtId.Text = 0 Then
    USMsgBox ("Informe o plano de controle antes de copiar."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
frmQualidadePPAP_PlanoControle_Copiar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_codigo_Click()
On Error GoTo tratar_erro

If txtCodinterno <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtCodinterno & "' and (tipo = 'P' or tipo = 'I' or tipo = 'PI')", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtidproduto = 0
        txtDescricaoProduto = ""
        txtRevProduto = ""
        cmbReferencia_prod.Clear
        txtidproduto = TBProduto!Codproduto
        txtCodinterno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
        txtRevProduto = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
        txtDescricaoProduto = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select N_Referencia from item_aplicacoes where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            cmbReferencia_prod.AddItem ""
            Do While TBItem.EOF = False
                cmbReferencia_prod.AddItem TBItem!N_referencia
                TBItem.MoveNext
            Loop
            TBItem.MoveFirst
            cmbReferencia_prod = TBItem!N_referencia
        End If
        TBItem.Close
    Else
        USMsgBox ("Não foi encontrado nenhum produto com este código interno."), vbExclamation, "CAPRIND v5.0"
        txtidproduto = 0
        txtDescricaoProduto = ""
        txtRevProduto = ""
        cmbReferencia_prod.Clear
    End If
    TBProduto.Close
Else
    txtidproduto = 0
    txtDescricaoProduto = ""
    txtRevProduto = ""
    cmbReferencia_prod.Clear
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimir_resultado_dimensional_Click()
On Error GoTo tratar_erro

ProcImprimir1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimir_resultado_dimensional1_Click()
On Error GoTo tratar_erro

ProcImprimir1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimirDimensoes_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarProduto_Click()
On Error GoTo tratar_erro

frmQualidadePPAP_PlanoControle_LocalizarProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProximo_Click()
On Error GoTo tratar_erro

ProcProximo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProximo2_Click()
On Error GoTo tratar_erro

ProcProximo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdrevisao_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtId.Text = 0 Then
    USMsgBox ("Informe o plano de controle antes de criar a revisão."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente criar uma revisão do plano " & txtPlano.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If txtDataRev <> "" Then
        USMsgBox ("Não é permitido revisar o plano de controle revisado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Contador = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from QualidadePPAP_PlanoControle where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcRevisar
        USMsgBox ("Plano de controle revisado com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
    TBAbrir.Close
    Lista.ListItems.Clear
    ProcCarregaLista
    '==================================
    Modulo = "Qualidade/Plano de controle"
    Evento = "Revisar"
    ID_documento = txtId
    Documento = "Plano de controle: " & txtPlano & " - Cód. interno: " & txtCodinterno
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyInsert: cmdNovo_Click
            Case vbKeyF2: cmdLocalizar_Click
            Case vbKeyF3: CmdSalvar_Click
            Case vbKeyF4: cmdExcluir_Click
            Case vbKeyF5: cmdImprimir_Click
            Case vbKeyF6: cmdImprimir_resultado_dimensional_Click
            Case vbKeyF7: cmdCopiar_Click
            Case vbKeyF8: cmdrevisao_Click
            Case vbKeyEscape: imgSair_Click
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: cmdSalvarDimensoes_Click
            Case vbKeyF5: cmdImprimirDimensoes_Click
            Case vbKeyF6: cmdImprimir_resultado_dimensional1_Click
            Case vbKeyEscape: imgSair_Click
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Formulario = "Qualidade/PPAP/Plano de controle"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_Click()
On Error GoTo tratar_erro

frmQualidadePPAP_PlanoControle_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNovo_Click()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos
Frame2.Enabled = True
Novo_PlanoControle = True
frmQualidadePPAP_PlanoControle_LocalizarProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If txtId.Text = 0 Then
    USMsgBox ("Informe o plano de controle antes de excluir."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir o nº " & txtPlano.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If txtDataRev <> "" Then
        USMsgBox ("Não é permitida a exclusão do plano de controle revisado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Conexao.Execute "Update QualidadePPAP_PlanoControle Set DataRevisao = NULL, IDRevisao = 0 where idRevisao = " & IIf(txtId = "", 0, txtId)
    Conexao.Execute "DELETE from QualidadePPAP_PlanoControle WHERE ID = " & txtId
    Conexao.Execute "DELETE from QualidadePPAP_PlanoControle_Dimensoes WHERE IDPlanoControle = " & txtId
    USMsgBox ("Plano de controle excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/Plano de controle"
    Evento = "Excluir"
    ID_documento = txtId
    Documento = "Plano de controle: " & txtPlano & " - Cód. interno: " & txtCodinterno
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos
    Lista.ListItems.Clear
    ProcCarregaLista
    Frame2.Enabled = False
    Novo_PlanoControle = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/PPAP/Plano de controle"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ImgCalendario1_Click()
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
Qualidade_PPAP_Plano = True
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgSair_Click()
On Error GoTo tratar_erro

If Novo_PlanoControle = True Then
    If USMsgBox("O plano de controle ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        CmdSalvar_Click
        If Novo_PlanoControle = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_PlanoControle = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdSalvar_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    cmdNovo.SetFocus
    Exit Sub
End If
Acao = "salvar"
If txtPlano = "" Then
    NomeCampo = "o número do plano"
    ProcVerificaAcao
    txtPlano.SetFocus
    Exit Sub
End If
If Cmb_tipo = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    Cmb_tipo.SetFocus
    Exit Sub
End If
If txtidproduto = 0 Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    txtCodinterno.SetFocus
    Exit Sub
End If
If txtDataRev <> "" Then
    USMsgBox ("Não é permitido alterar o plano de controle revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido2 = False
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * from QualidadePPAP_PlanoControle where ID = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    Permitido2 = True
End If
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
If Permitido2 = True Then ProcSalvarPlanoInspecao
TBGravar.Close
Lista.ListItems.Clear
ProcCarregaLista
procCarregalistaDimensoes
If Novo_PlanoControle = True Then
    USMsgBox ("Novo plano de controle cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Plano de controle"
ID_documento = txtId
Documento = "Plano de controle: " & txtPlano & " - Cód. interno: " & txtCodinterno
Documento1 = ""
ProcGravaEvento
'==================================
Novo_PlanoControle = False
Caption = "Qualidade - PPAP - Plano de controle (" & txtPlano & " - Rev.: " & txtRev & ")"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSalvarDimensoes_Click()
On Error GoTo tratar_erro

If txtDataRev <> "" Then
    USMsgBox ("Não é permitido alteração no plano de controle revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDdimensoes = 0 Then
    USMsgBox ("Informe a dimensão antes de salvar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtDataRev <> "" Then
    USMsgBox ("Não é permitido alterar dimensão do plano de controle revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If IsNumeric(txtQtdeEnsaiado) = False And txtQtdeEnsaiado <> "" Then
    NomeCampo = "a quantidade ensaiada"
    ProcVerificaAcao
    txtQtdeEnsaiado.SetFocus
    Exit Sub
End If
If IsDate(txtDataemissao) = False And txtDataemissao <> "__/__/____" Then
    NomeCampo = "a data do ensaio"
    ProcVerificaAcao
    txtDataemissao.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * from QualidadePPAP_PlanoControle_Dimensoes where ID = " & txtIDdimensoes.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar dimensão"
End If
ProcEnviaDadosDimensoes
TBGravar.Update
TBGravar.Close
'==================================
Modulo = "Qualidade/Plano de controle"
ID_documento = txtIDdimensoes
Documento = "Plano de controle: " & txtPlano & " - Cód. interno: " & txtCodinterno
Documento1 = "Fase: " & txtFase
ProcGravaEvento
'==================================
procCarregalistaDimensoes
If CodigoLista1 <> 0 And Lista2.ListItems.Count <> 0 Then
    Lista2.SelectedItem = Lista2.ListItems(CodigoLista1)
    Lista2.SetFocus
End If

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

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from QualidadePPAP_PlanoControle where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista2, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista2.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from QualidadePPAP_PlanoControle_Dimensoes where id = " & Lista2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposDimensoes
    ProcPuxadadosDimensoes
    CodigoLista1 = Lista2.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId.Text = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If

Select Case SSTab1.Tab
    Case 0:
        cmdNovo.SetFocus
    Case 1:
        If Novo_PlanoControle = True Then
            USMsgBox ("Salve o plano de controle antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            CmdSalvar.SetFocus
            Exit Sub
        End If
        cmdSalvarDimensoes.SetFocus
        ProcLimpaCamposDimensoes
        procCarregalistaDimensoes
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtId.Text = IIf(IsNull(TBAbrir!ID), "", TBAbrir!ID)
txtPlano.Text = IIf(IsNull(TBAbrir!Plano), "", TBAbrir!Plano)
txtRev.Text = IIf(IsNull(TBAbrir!Rev), "", TBAbrir!Rev)
txtDataRev.Text = IIf(IsNull(TBAbrir!DataRevisao), "", Format(TBAbrir!DataRevisao, "dd/mm/yy"))
txtDataemissao.Text = IIf(IsNull(TBAbrir!DtEmissao), "", Format(TBAbrir!DtEmissao, "dd/mm/yy"))
txtResp.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
If IsNull(TBAbrir!Tipo) = False And TBAbrir!Tipo <> "" Then Cmb_tipo = TBAbrir!Tipo
txtContato = IIf(IsNull(TBAbrir!contato), "", TBAbrir!contato)
txtCodinterno.Text = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
txtidproduto.Text = IIf(IsNull(TBAbrir!IDProduto), "0", TBAbrir!IDProduto)
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtRevProduto = IIf(IsNull(TBItem!RevDesenho), "", TBItem!RevDesenho)
    txtDescricaoProduto = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
    Set TBCiclo = CreateObject("adodb.recordset")
    TBCiclo.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto & " order by n_referencia", Conexao, adOpenKeyset, adLockOptimistic
    If TBCiclo.EOF = False Then
        Do While TBCiclo.EOF = False
            cmbReferencia_prod.AddItem IIf(IsNull(TBCiclo!N_referencia), "", TBCiclo!N_referencia)
            TBCiclo.MoveNext
        Loop
    End If
    TBCiclo.Close
    If IsNull(TBAbrir!N_referencia) = False And TBAbrir!N_referencia <> "" Then cmbReferencia_prod = TBAbrir!N_referencia
End If
TBItem.Close
txtEquipe = IIf(IsNull(TBAbrir!Equipe), "", TBAbrir!Equipe)
txtAprovacao_engenharia = IIf(IsNull(TBAbrir!Aprovacao_engenharia), "", TBAbrir!Aprovacao_engenharia)
txtOrganizacaoPlanta = IIf(IsNull(TBAbrir!Organizacao_Planta), "", TBAbrir!Organizacao_Planta)
txtCodOrganizacao = IIf(IsNull(TBAbrir!CodOrganizacao), "", TBAbrir!CodOrganizacao)
txtOrganizacaoAprovacao = IIf(IsNull(TBAbrir!Organizacao_Aprovacao), "", TBAbrir!Organizacao_Aprovacao)
txtAprovacaoQualidade = IIf(IsNull(TBAbrir!Aprovacao_Qualidade), "", TBAbrir!Aprovacao_Qualidade)
txtOutraAprovacao = IIf(IsNull(TBAbrir!Outra_Aprovacao), "", TBAbrir!Outra_Aprovacao)
txtOutraAprovacao2 = IIf(IsNull(TBAbrir!Outra_Aprovacao2), "", TBAbrir!Outra_Aprovacao2)
Txt_local_inspecao = IIf(IsNull(TBAbrir!Local_inspecao), "", TBAbrir!Local_inspecao)
Txt_doc_engenharia = IIf(IsNull(TBAbrir!Doc_engenharia), "", TBAbrir!Doc_engenharia)
Caption = "Qualidade - PPAP - Plano de controle (Plano de controle : " & TBAbrir!Plano & " - Rev. : " & TBAbrir!Rev & ")"
Frame2.Enabled = True
Novo_PlanoControle = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

If SQL_PlanoControle = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open SQL_PlanoControle, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Plano), "", TBLISTA!Plano)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Rev), "", TBLISTA!Rev)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "select * from projproduto where codproduto = " & TBLISTA!IDProduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .Item(.Count).SubItems(4) = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
            End If
            TBItem.Close
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!DtEmissao), "", Format(TBLISTA!DtEmissao, "dd/mm/yy"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from QualidadePPAP_PlanoControle order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("ID = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId.Text = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from QualidadePPAP_PlanoControle where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCamposDimensoes
        ProcPuxaDados
        procCarregalistaDimensoes
    Else
        USMsgBox ("Fim dos cadastros de plano de controle."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_PlanoControle = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from QualidadePPAP_PlanoControle order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("ID = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId.Text = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from QualidadePPAP_PlanoControle where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCamposDimensoes
        ProcPuxaDados
        procCarregalistaDimensoes
    Else
        USMsgBox ("Fim dos cadastros de plano de controle."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_PlanoControle = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRevisar()
On Error GoTo tratar_erro

Contador = 0
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from QualidadePPAP_PlanoControle where Plano = '" & txtPlano & "' order by Rev", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.MoveLast
Contador = TBAbrir!Rev
Contador = Contador + 1

TBGravar.AddNew
TBGravar!Plano = TBAbrir!Plano
TBGravar!Rev = Contador
TBGravar!DtEmissao = Date
TBGravar!Responsavel = pubUsuario
TBGravar!Tipo = TBAbrir!Tipo
TBGravar!contato = TBAbrir!contato
TBGravar!IDProduto = TBAbrir!IDProduto
TBGravar!Codinterno = TBAbrir!Codinterno
TBGravar!N_referencia = TBAbrir!N_referencia
TBGravar!Aprovacao_engenharia = TBAbrir!Aprovacao_engenharia
TBGravar!Equipe = TBAbrir!Equipe
TBGravar!Organizacao_Planta = TBAbrir!Organizacao_Planta
TBGravar!CodOrganizacao = TBAbrir!CodOrganizacao
TBGravar!Organizacao_Aprovacao = TBAbrir!Organizacao_Aprovacao
TBGravar!Aprovacao_Qualidade = TBAbrir!Aprovacao_Qualidade
TBGravar!Outra_Aprovacao = TBAbrir!Outra_Aprovacao
TBGravar!Outra_Aprovacao2 = TBAbrir!Outra_Aprovacao2

TBAbrir!DataRevisao = Date
TBAbrir!IDRevisao = TBGravar!ID
TBAbrir.Update

TBGravar.Update
txtId = TBGravar!ID
Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "select * from QualidadePPAP_PlanoControle_Dimensoes where IDPlanoControle = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then
    Do While TBCiclo.EOF = False
        Set TBExecucao = CreateObject("adodb.recordset")
        TBExecucao.Open "select * from QualidadePPAP_PlanoControle_Dimensoes", Conexao, adOpenKeyset, adLockOptimistic
        TBExecucao.AddNew
        TBExecucao!IdPlanoControle = TBGravar!ID
        TBExecucao!IDPlano = TBCiclo!IDPlano
        TBExecucao!idDimensao = TBCiclo!idDimensao
        TBExecucao!Data = Date
        TBExecucao!Responsavel = pubUsuario
        TBExecucao!Processo = TBCiclo!Processo
        TBExecucao!Carac = TBCiclo!Carac
        TBExecucao!TecnicaAvaliacao = TBCiclo!TecnicaAvaliacao
        TBExecucao!AmostraFreq = TBCiclo!AmostraFreq
        TBExecucao!MetodoControle = TBCiclo!MetodoControle
        TBExecucao!PlanoReacao = TBCiclo!PlanoReacao
        TBExecucao!DataEnsaio = TBCiclo!DataEnsaio
        TBExecucao!QtdeEnsaiado = TBCiclo!QtdeEnsaiado
        TBExecucao!Resultados = TBCiclo!Resultados
        If TBCiclo!ok = True Then TBExecucao!ok = True Else TBExecucao!ok = False
        TBExecucao.Update
        TBExecucao.Close
        TBCiclo.MoveNext
    Loop
End If
TBCiclo.Close
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = 0 Then
    USMsgBox ("Informe o plano de controle antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
NomeRel = "CQ_PPAP_Plano de controle.rpt"
ProcImprimirRel "{QualidadePPAP_PlanoControle.ID} = " & txtId, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir1()
On Error GoTo tratar_erro

If txtId = 0 Then
    USMsgBox ("Informe o plano de controle antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
NomeRel = "CQ_PPAP_Plano de controle_dimensional.rpt"
ProcImprimirRel "{QualidadePPAP_PlanoControle.ID} = " & txtId, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procCarregalistaDimensoes()
On Error GoTo tratar_erro

Lista2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select QualidadePPAP_PlanoControle_Dimensoes.*, Planodimensao.IDPlano, Planodimensao.Numero, Planodimensao.Tipo, Planodimensao.dimdesejada, Planodimensao.TolSup, Planodimensao.TolInf from QualidadePPAP_PlanoControle_Dimensoes INNER JOIN Planodimensao on QualidadePPAP_PlanoControle_Dimensoes.IDdimensao = Planodimensao.idDimensao where QualidadePPAP_PlanoControle_Dimensoes.idplanoControle = " & txtId & " order by Planodimensao.Numero", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista2.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDPlano), "", TBLISTA!IDPlano)
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from plano where idplano = " & TBLISTA!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Fase), "", TBAbrir!Fase)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Grupo_op), "", TBAbrir!Grupo_op)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Numero), "", TBLISTA!Numero)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
                    
                Texto = IIf(IsNull(TBLISTA!dimdesejada), "", Format(TBLISTA!dimdesejada, "###,##0.0000"))
                Texto1 = IIf(IsNull(TBLISTA!TolSup), "", Format(TBLISTA!TolSup, "###,##0.0000"))
                Texto2 = IIf(IsNull(TBLISTA!TolInf), "", Format(TBLISTA!TolInf, "###,##0.0000"))
                .Item(.Count).SubItems(9) = Texto & "  " & Texto1 & " / " & Texto2
                    
                Set TBFases = CreateObject("adodb.recordset")
                TBFases.Open "select Fases.* from fases INNER JOIN processos on fases.IDProcesso = processos.IDProcesso where processos.CodProduto = " & txtidproduto & " and fases.Fase = " & IIf(IsNull(TBAbrir!Fase), 0, TBAbrir!Fase), Conexao, adOpenKeyset, adLockOptimistic
                If TBFases.EOF = False Then
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBFases!maquina), "", TBFases!maquina)
                End If
                TBFases.Close
            End If
            TBAbrir.Close
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Processo), "", TBLISTA!Processo)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Carac), "", TBLISTA!Carac)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!TecnicaAvaliacao), "", TBLISTA!TecnicaAvaliacao)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!AmostraFreq), "", TBLISTA!AmostraFreq)
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!MetodoControle), "", TBLISTA!MetodoControle)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!PlanoReacao), "", TBLISTA!PlanoReacao)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarPlanoInspecao()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select plano.idplano, plano.Fase, Planodimensao.IDdimensao, Planodimensao.Numero, Planodimensao.Freq from plano inner join Planodimensao on plano.IdPlano = Planodimensao.IdPlano where plano.Desenho = '" & txtCodinterno & "' order by Planodimensao.Numero", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If IsNull(TBAbrir!Numero) = False And TBAbrir!Numero <> "" Then
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select * from QualidadePPAP_PlanoControle_Dimensoes where IDPlano = " & txtId & " AND Iddimensao = " & TBAbrir!idDimensao, Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = True Then TBFIltro.AddNew
            ProcEnviaDadosPlanoInspecao
            Familiatext = ""
            Set TBInstrumentos = CreateObject("adodb.recordset")
            TBInstrumentos.Open "Select * from Planodimensao_instrumentos where ID_dimensao = " & TBAbrir!idDimensao & " order by Familia", Conexao, adOpenKeyset, adLockOptimistic
            If TBInstrumentos.EOF = False Then
                Do While TBInstrumentos.EOF = False
                    If Familiatext <> "" Then Familiatext = Familiatext & " / " & Trim(TBInstrumentos!Familia) Else Familiatext = Trim(TBInstrumentos!Familia)
                    TBInstrumentos.MoveNext
                Loop
                TBFIltro!TecnicaAvaliacao = Familiatext
            End If
            TBInstrumentos.Close
            TBFIltro.Update
            TBFIltro.Close
        End If
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosPlanoInspecao()
On Error GoTo tratar_erro

TBFIltro!IdPlanoControle = txtId
TBFIltro!IDPlano = TBAbrir!IDPlano
TBFIltro!idDimensao = TBAbrir!idDimensao
TBFIltro!AmostraFreq = TBAbrir!Freq


Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Fases.Descricao from Fases inner join Processos on Processos.IDProcesso = Fases.IDProcesso where Processos.CodProduto = " & txtidproduto & " and Fases.Fase = " & IIf(IsNull(TBAbrir!Fase), 0, TBAbrir!Fase), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    TBFIltro!Descricao = TBFI!Descricao
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodInterno_Change()
On Error GoTo tratar_erro

txtidproduto = 0
txtRevProduto = ""
cmbReferencia_prod.Clear
txtDescricaoProduto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDataEnsaio_LostFocus()
On Error GoTo tratar_erro

If txtDataEnsaio <> "__/__/____" Then
    VerifData = txtDataEnsaio
    ProcVerificaData
    If VerifData = False Then
        txtDataEnsaio = "__/__/____"
        txtDataEnsaio.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtdeEnsaiado_LostFocus()
On Error GoTo tratar_erro

If txtQtdeEnsaiado.Text <> "" Then
    VerifNumero = txtQtdeEnsaiado.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtdeEnsaiado.Text = ""
        txtQtdeEnsaiado.SetFocus
        Exit Sub
    End If
    txtQtdeEnsaiado = Format(txtQtdeEnsaiado, "###,##0.0000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
