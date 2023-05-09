VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQualidadePPAP_FMEA 
   Caption         =   "Qualidade - PPAP - FMEA"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   495
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
      Left            =   75
      TabIndex        =   93
      Top             =   6960
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   75
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
      TabPicture(0)   =   "frmQualidadePPAP_FMEA.frx":0000
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
      TabPicture(1)   =   "frmQualidadePPAP_FMEA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListaFMEA"
      Tab(1).Control(1)=   "SSTab2"
      Tab(1).ControlCount=   2
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
         TabIndex        =   86
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
         TabIndex        =   85
         Text            =   "0"
         Top             =   6090
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   75
         TabIndex        =   76
         Top             =   330
         Width           =   11745
         Begin VB.CommandButton cmdAjuda 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   10290
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":0038
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":018A
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Ajuda (F1)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton imgSair 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   10935
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":062C
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":077E
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Sair (Esc)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdCopiar 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   4590
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":0F51
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":10A3
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Copiar (F7)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdRevisao 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   5220
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":15A5
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":16F7
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Revisar (F8)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdSalvar 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   1440
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":1B9A
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":1CEC
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
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":24C5
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":2617
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
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":2E66
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":2FB8
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
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":34DE
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":3630
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
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":3DF1
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":3F43
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Visualizar impressão (F5)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdanterior 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   3330
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":4732
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":4884
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Registro anterior."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdProximo 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   3960
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":4BCD
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":4D1F
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Próximo registro."
            Top             =   180
            Width           =   630
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   1845
         Left            =   75
         TabIndex        =   25
         Top             =   5100
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   3254
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
            Text            =   "Nº FMEA"
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
         Height          =   3885
         Left            =   75
         TabIndex        =   77
         Top             =   1200
         Width           =   11745
         Begin VB.CommandButton cmdFiltrar_codigo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2460
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":5068
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":51BA
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Filtrar por código interno."
            Top             =   2235
            Width           =   315
         End
         Begin VB.TextBox txtIDforn 
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
            TabIndex        =   10
            ToolTipText     =   "Id do fornecedor."
            Top             =   1620
            Width           =   525
         End
         Begin VB.TextBox txtIDCliente 
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
            TabIndex        =   7
            ToolTipText     =   "Id do cliente."
            Top             =   1025
            Width           =   525
         End
         Begin VB.TextBox txtResponsabilidade 
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
            Left            =   7980
            MaxLength       =   50
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":55D5
            MousePointer    =   99  'Custom
            TabIndex        =   6
            ToolTipText     =   "Responsabilidade pelo processo."
            Top             =   390
            Width           =   3585
         End
         Begin VB.CommandButton cmdFornecedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   11250
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":58DF
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":5A31
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Localizar fornecedor."
            Top             =   1620
            Width           =   315
         End
         Begin VB.CommandButton cmdCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   11250
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":5B33
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":5C85
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Localizar cliente."
            Top             =   1025
            Width           =   315
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
            Left            =   1650
            Locked          =   -1  'True
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":5D87
            MousePointer    =   99  'Custom
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   390
            Width           =   570
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
            Left            =   2235
            Locked          =   -1  'True
            MaxLength       =   100
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":6091
            MousePointer    =   99  'Custom
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data da revisão."
            Top             =   390
            Width           =   1020
         End
         Begin VB.TextBox txtAprovado 
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
            Left            =   8130
            MaxLength       =   50
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":639B
            MousePointer    =   99  'Custom
            TabIndex        =   21
            ToolTipText     =   "Aprovado por."
            Top             =   2850
            Width           =   3435
         End
         Begin VB.TextBox txtObs 
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
            MaxLength       =   255
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":66A5
            MousePointer    =   99  'Custom
            TabIndex        =   22
            ToolTipText     =   "Observações."
            Top             =   3450
            Width           =   9945
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
            MaxLength       =   255
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":69AF
            MousePointer    =   99  'Custom
            TabIndex        =   20
            ToolTipText     =   "Equipe."
            Top             =   2850
            Width           =   7935
         End
         Begin VB.TextBox txtFornecedor 
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
            Left            =   720
            MaxLength       =   255
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":6CB9
            MousePointer    =   99  'Custom
            TabIndex        =   11
            ToolTipText     =   "Fornecedor."
            Top             =   1620
            Width           =   10515
         End
         Begin VB.TextBox txtCliente 
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
            Left            =   720
            MaxLength       =   255
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":6FC3
            MousePointer    =   99  'Custom
            TabIndex        =   8
            ToolTipText     =   "Cliente."
            Top             =   1025
            Width           =   10515
         End
         Begin VB.TextBox txtDescricao 
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
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":72CD
            MousePointer    =   99  'Custom
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   2235
            Width           =   4095
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
            Left            =   3180
            Locked          =   -1  'True
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":75D7
            MousePointer    =   99  'Custom
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   2235
            Width           =   530
         End
         Begin VB.TextBox txtCodInterno 
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
            MaxLength       =   50
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":78E1
            MousePointer    =   99  'Custom
            TabIndex        =   13
            ToolTipText     =   "Código interno."
            Top             =   2235
            Width           =   2265
         End
         Begin VB.CommandButton cmdLocalizarProduto_cliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2790
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":7BEB
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":7D3D
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Localizar código interno."
            Top             =   2235
            Width           =   315
         End
         Begin VB.ComboBox cmbReferencia 
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
            Left            =   3750
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":7E3F
            MousePointer    =   99  'Custom
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Código de referência."
            Top             =   2235
            Width           =   2460
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
            Left            =   3270
            Locked          =   -1  'True
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":8149
            MousePointer    =   99  'Custom
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1155
         End
         Begin VB.TextBox txtResp 
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
            Left            =   4440
            Locked          =   -1  'True
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":8453
            MousePointer    =   99  'Custom
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3525
         End
         Begin VB.TextBox txtFMEA 
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
            MaxLength       =   50
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":875D
            MousePointer    =   99  'Custom
            TabIndex        =   1
            ToolTipText     =   "Número FMEA."
            Top             =   390
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txtDatachave 
            Height          =   315
            Left            =   10140
            TabIndex        =   23
            ToolTipText     =   "Data chave."
            Top             =   3450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker txtDataCod 
            Height          =   315
            Left            =   10365
            TabIndex        =   19
            ToolTipText     =   "Data do código interno."
            Top             =   2235
            Width           =   1200
            _ExtentX        =   2117
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
            Format          =   199557123
            CurrentDate     =   39057
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsabilidade pelo processo"
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
            TabIndex        =   97
            Top             =   180
            Width           =   2295
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
            Index           =   3
            Left            =   1748
            TabIndex        =   96
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Revisado em"
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
            Left            =   2288
            TabIndex        =   95
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data do cód."
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
            Left            =   10500
            TabIndex        =   94
            Top             =   2040
            Width           =   930
         End
         Begin VB.Image imgCalendario 
            Height          =   360
            Left            =   11235
            MouseIcon       =   "frmQualidadePPAP_FMEA.frx":8A67
            MousePointer    =   99  'Custom
            Picture         =   "frmQualidadePPAP_FMEA.frx":8BB9
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   3420
            Width           =   330
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data chave"
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
            Left            =   10275
            TabIndex        =   92
            Top             =   3240
            Width           =   825
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
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
            Left            =   4680
            TabIndex        =   91
            Top             =   3240
            Width           =   945
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Equipe"
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
            Left            =   3907
            TabIndex        =   90
            Top             =   2640
            Width           =   480
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aprovado por"
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
            Left            =   9345
            TabIndex        =   89
            Top             =   2640
            Width           =   990
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
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
            TabIndex        =   88
            Top             =   1410
            Width           =   825
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
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
            Left            =   5730
            TabIndex        =   87
            Top             =   810
            Width           =   495
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
            Left            =   7935
            TabIndex        =   84
            Top             =   2040
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
            Left            =   3258
            TabIndex        =   83
            Top             =   2040
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
            Left            =   4230
            TabIndex        =   82
            Top             =   2040
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
            Left            =   787
            TabIndex        =   81
            Top             =   2040
            Width           =   1050
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data emissão"
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
            TabIndex        =   80
            Top             =   180
            Width           =   975
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
            Left            =   5730
            TabIndex        =   79
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FMEA"
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
            Left            =   682
            TabIndex        =   78
            Top             =   180
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView ListaFMEA 
         Height          =   1340
         Left            =   -74910
         TabIndex        =   102
         Top             =   5610
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   2355
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
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   20
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Fase"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Grupo/op."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Requisito"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Modo de falha potencial"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Efeito potencial de falha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Severidade"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Class"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Causa potencial de falha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Controles prevenção"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Ocorrência"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Controles detecção"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Deteccção"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "NPR"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Ações recomendadas"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Tag             =   "T"
            Text            =   "Resp.. e data de concl. preten."
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Tag             =   "T"
            Text            =   "Ações implem. data de concl."
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Object.Tag             =   "N"
            Text            =   "S"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Object.Tag             =   "N"
            Text            =   "O"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Object.Tag             =   "N"
            Text            =   "D"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Object.Tag             =   "N"
            Text            =   "NPR"
            Object.Width           =   882
         EndProperty
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   7335
         Left            =   -74985
         TabIndex        =   35
         Top             =   300
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   12938
         _Version        =   393216
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
         TabCaption(0)   =   "Etapa do processo/função"
         TabPicture(0)   =   "frmQualidadePPAP_FMEA.frx":903C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Lista2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtIDFase"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Modo de falha potencial"
         TabPicture(1)   =   "frmQualidadePPAP_FMEA.frx":9058
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkAcoes"
         Tab(1).Control(1)=   "txtID_ModoFalha"
         Tab(1).Control(2)=   "Frame4"
         Tab(1).Control(3)=   "Lista3"
         Tab(1).Control(4)=   "Frame8"
         Tab(1).Control(5)=   "Frame6"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Efeito potencial de falha"
         TabPicture(2)   =   "frmQualidadePPAP_FMEA.frx":9074
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtID_Efeitofalha"
         Tab(2).Control(1)=   "Frame7"
         Tab(2).Control(2)=   "Lista4"
         Tab(2).Control(3)=   "Frame10"
         Tab(2).ControlCount=   4
         Begin VB.CheckBox chkAcoes 
            BackColor       =   &H00E0E0E0&
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
            Height          =   225
            Left            =   -72900
            TabIndex        =   53
            Top             =   2700
            Width           =   195
         End
         Begin VB.TextBox txtIDFase 
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
            Left            =   1110
            TabIndex        =   120
            Text            =   "0"
            Top             =   2490
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtID_Efeitofalha 
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
            Left            =   -72240
            TabIndex        =   106
            Text            =   "0"
            Top             =   3150
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtID_ModoFalha 
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
            Left            =   -73170
            TabIndex        =   105
            Text            =   "0"
            Top             =   3810
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Height          =   885
            Left            =   -74925
            TabIndex        =   104
            Top             =   330
            Width           =   11745
            Begin VB.CommandButton cmdAjuda_Efeitofalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   10290
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":9090
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":91E2
               Style           =   1  'Graphical
               TabIndex        =   73
               ToolTipText     =   "Ajuda (F1)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdSair_Efeitofalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   10935
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":9684
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":97D6
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Sair (Esc)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdSalvar_Efeitofalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   810
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":9FA9
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":A0FB
               Style           =   1  'Graphical
               TabIndex        =   69
               ToolTipText     =   "Salvar (F3)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdExcluir_Efeitofalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   1440
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":A8D4
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":AA26
               Style           =   1  'Graphical
               TabIndex        =   71
               ToolTipText     =   "Excluir (F4)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdNovo_Efeitofalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   180
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":B275
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":B3C7
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "Novo (Insert)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdImprimir_Efeitofalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   2070
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":B8ED
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":BA3F
               Style           =   1  'Graphical
               TabIndex        =   72
               ToolTipText     =   "Visualizar impressão (F5)"
               Top             =   180
               Width           =   630
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   885
            Left            =   -74925
            TabIndex        =   103
            Top             =   330
            Width           =   11745
            Begin VB.CommandButton cmdImprimir_ModoFalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   2070
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":C22E
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":C380
               Style           =   1  'Graphical
               TabIndex        =   62
               ToolTipText     =   "Visualizar impressão (F5)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdNovo_ModoFalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   180
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":CB6F
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":CCC1
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Novo (Insert)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdExcluir_ModoFalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   1440
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":D1E7
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":D339
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Excluir (F4)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdSalvar_ModoFalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   810
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":DB88
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":DCDA
               Style           =   1  'Graphical
               TabIndex        =   59
               ToolTipText     =   "Salvar (F3)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdSair_ModoFalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   10935
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":E4B3
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":E605
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Sair (Esc)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdAjuda_ModoFalha 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   10290
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":EDD8
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":EF2A
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Ajuda (F1)"
               Top             =   180
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   885
            Left            =   75
            TabIndex        =   98
            Top             =   330
            Width           =   11745
            Begin VB.CommandButton cmdAjuda_Fases 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   10290
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":F3CC
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":F51E
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Ajuda (F1)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdSair_Fases 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   10935
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":F9C0
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":FB12
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Sair (Esc)"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton cmdImprimir_Fases 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   180
               MouseIcon       =   "frmQualidadePPAP_FMEA.frx":102E5
               MousePointer    =   99  'Custom
               Picture         =   "frmQualidadePPAP_FMEA.frx":10437
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Visualizar impressão (F5)"
               Top             =   180
               Width           =   630
            End
         End
         Begin MSComctlLib.ListView Lista2 
            Height          =   3225
            Left            =   75
            TabIndex        =   40
            Top             =   2065
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   5689
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "Fase"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Grupo/op."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Requisito"
               Object.Width           =   16501
            EndProperty
         End
         Begin MSComctlLib.ListView Lista4 
            Height          =   3225
            Left            =   -74925
            TabIndex        =   70
            Top             =   2070
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   5689
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Efeito potencial de falha"
               Object.Width           =   8691
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Class"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Causa potencial de falha"
               Object.Width           =   8691
            EndProperty
         End
         Begin MSComctlLib.ListView Lista3 
            Height          =   1665
            Left            =   -74925
            TabIndex        =   60
            Top             =   3625
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   2937
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Modo de falha poten."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Object.Tag             =   "N"
               Text            =   "Sever."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Controle prevenção"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Ocorrência"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Controle detecção"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Detecção"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "NPR"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Ações recomendadas"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Resp. e data de concl. preten."
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Tag             =   "T"
               Text            =   "Ações impl. data de concl."
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "S"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Object.Tag             =   "N"
               Text            =   "O"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   13
               Object.Tag             =   "N"
               Text            =   "D"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   14
               Object.Tag             =   "N"
               Text            =   "NPR"
               Object.Width           =   882
            EndProperty
         End
         Begin VB.Frame Frame8 
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
            Height          =   1515
            Left            =   -74925
            TabIndex        =   108
            Top             =   1200
            Width           =   11745
            Begin VB.TextBox txtNPR 
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
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   50
               TabStop         =   0   'False
               ToolTipText     =   "NPR."
               Top             =   1050
               Width           =   945
            End
            Begin VB.TextBox txtSever 
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
               TabIndex        =   45
               ToolTipText     =   "Severidade."
               Top             =   390
               Width           =   945
            End
            Begin VB.TextBox txtModo_falha 
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
               MaxLength       =   100
               TabIndex        =   44
               ToolTipText     =   "Modo de falha potencial."
               Top             =   390
               Width           =   2955
            End
            Begin VB.TextBox txtControle_prevencao 
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
               Left            =   4110
               MaxLength       =   50
               TabIndex        =   46
               ToolTipText     =   "Controle prevenção."
               Top             =   390
               Width           =   3225
            End
            Begin VB.TextBox txtOcorrencia 
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
               Left            =   7350
               TabIndex        =   47
               ToolTipText     =   "Ocorrência."
               Top             =   390
               Width           =   945
            End
            Begin VB.TextBox txtControle_deteccao 
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
               Left            =   8310
               TabIndex        =   48
               ToolTipText     =   "Controle detecção."
               Top             =   390
               Width           =   3255
            End
            Begin VB.TextBox txtDeteccao 
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
               TabIndex        =   49
               ToolTipText     =   "Detecção."
               Top             =   1050
               Width           =   945
            End
            Begin VB.TextBox txtAcoes_recomendadas 
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
               Left            =   2100
               MaxLength       =   100
               TabIndex        =   51
               ToolTipText     =   "Ações recomendadas."
               Top             =   1050
               Width           =   4785
            End
            Begin VB.TextBox txtResp_conclusao 
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
               Left            =   6900
               MaxLength       =   100
               TabIndex        =   52
               ToolTipText     =   "Responsabilidade de data de conclusão pretendida."
               Top             =   1050
               Width           =   4665
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NPR"
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
               Left            =   1462
               TabIndex        =   128
               Top             =   840
               Width           =   300
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Severidade"
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
               Left            =   3217
               TabIndex        =   121
               Top             =   180
               Width           =   810
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modo de falha potencial"
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
               Left            =   802
               TabIndex        =   115
               Top             =   180
               Width           =   1710
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Controle prevenção"
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
               Left            =   5010
               TabIndex        =   114
               Top             =   180
               Width           =   1425
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ocorrência"
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
               Left            =   7432
               TabIndex        =   113
               Top             =   180
               Width           =   780
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Controle detecção"
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
               Left            =   9277
               TabIndex        =   112
               Top             =   180
               Width           =   1320
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Detecção"
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
               Left            =   315
               TabIndex        =   111
               Top             =   840
               Width           =   675
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ações recomendadas"
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
               Left            =   3727
               TabIndex        =   110
               Top             =   840
               Width           =   1530
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Respon. e data de conclusão preten."
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
               Left            =   7897
               TabIndex        =   109
               Top             =   840
               Width           =   2670
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Resultados das ações "
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
            Height          =   915
            Left            =   -74925
            TabIndex        =   122
            Top             =   2700
            Width           =   11745
            Begin VB.TextBox txtAcoes_implementacoes 
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
               MaxLength       =   100
               TabIndex        =   54
               ToolTipText     =   "Ações implementadas data de conclusão."
               Top             =   450
               Width           =   5840
            End
            Begin VB.TextBox txtDeteccao_acoes 
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
               Left            =   8790
               TabIndex        =   57
               ToolTipText     =   "Detecção."
               Top             =   450
               Width           =   1335
            End
            Begin VB.TextBox txtOcorrencia_acoes 
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
               Left            =   7410
               TabIndex        =   56
               ToolTipText     =   "Ocorrência."
               Top             =   450
               Width           =   1365
            End
            Begin VB.TextBox txtSeveridade_acoes 
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
               Left            =   6030
               TabIndex        =   55
               ToolTipText     =   "Severidade."
               Top             =   450
               Width           =   1365
            End
            Begin VB.TextBox txtNPR_acoes 
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
               Left            =   10140
               Locked          =   -1  'True
               TabIndex        =   58
               TabStop         =   0   'False
               ToolTipText     =   "NPR."
               Top             =   450
               Width           =   1425
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NPR"
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
               Left            =   10702
               TabIndex        =   127
               Top             =   240
               Width           =   300
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ações implementadas data de conclusão"
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
               Left            =   1653
               TabIndex        =   126
               Top             =   240
               Width           =   2895
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Detecção"
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
               Left            =   9120
               TabIndex        =   125
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ocorrência"
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
               Left            =   7702
               TabIndex        =   124
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Severidade"
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
               Left            =   6307
               TabIndex        =   123
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame Frame10 
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
            Height          =   855
            Left            =   -74925
            TabIndex        =   116
            Top             =   1200
            Width           =   11745
            Begin VB.TextBox txtEfeito_potencial 
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
               TabIndex        =   66
               ToolTipText     =   "Efeito potencial de falha."
               Top             =   390
               Width           =   5685
            End
            Begin VB.TextBox txtClass 
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
               Left            =   5880
               TabIndex        =   67
               ToolTipText     =   "Class."
               Top             =   390
               Width           =   1365
            End
            Begin VB.TextBox txtCausa_potencial 
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
               MaxLength       =   100
               TabIndex        =   68
               ToolTipText     =   "Causa potencial de falha."
               Top             =   390
               Width           =   4305
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Efeito potencial de falha"
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
               Left            =   2152
               TabIndex        =   119
               Top             =   180
               Width           =   1740
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Class"
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
               Left            =   6375
               TabIndex        =   118
               Top             =   180
               Width           =   375
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Causa potencial de falha"
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
               Left            =   8527
               TabIndex        =   117
               Top             =   180
               Width           =   1770
            End
         End
         Begin VB.Frame Frame3 
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
            Height          =   855
            Left            =   75
            TabIndex        =   99
            Top             =   1200
            Width           =   11745
            Begin VB.TextBox txtRequisito 
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
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               ToolTipText     =   "Requisito."
               Top             =   390
               Width           =   7485
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
               Left            =   2130
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               ToolTipText     =   "Grupo/operação."
               Top             =   390
               Width           =   1935
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
               TabIndex        =   37
               TabStop         =   0   'False
               ToolTipText     =   "Fase."
               Top             =   390
               Width           =   1935
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Requisito"
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
               Left            =   7492
               TabIndex        =   107
               Top             =   180
               Width           =   660
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
               Left            =   2730
               TabIndex        =   101
               Top             =   180
               Width           =   735
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
               Left            =   975
               TabIndex        =   100
               Top             =   180
               Width           =   345
            End
         End
      End
   End
End
Attribute VB_Name = "frmQualidadePPAP_FMEA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_FMEA           As Boolean 'OK
Public Novo_FMEA1          As Boolean 'OK
Public Novo_FMEA2          As Boolean 'OK
Public SQL_FMEA            As String 'OK

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtidproduto = 0
txtFMEA.Text = ""
txtRev = 0
txtDataRev = ""
txtDataemissao = Format(Date, "dd/mm/yy")
txtResp = pubUsuario
txtResponsabilidade = ""
txtIDcliente = ""
txtCliente = ""
txtDataCod = Date
txtCodinterno.Text = ""
txtRevProduto.Text = ""
cmbReferencia.Clear
txtdescricao.Text = ""
txtIDforn = ""
txtFornecedor = ""
txtEquipe = ""
txtAprovado = ""
txtObs = ""
txtDatachave = "__/__/____"
CodigoLista = 0
Caption = "Qualidade - PPAP - FMEA"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!FMEA = txtFMEA
TBGravar!Rev = txtRev.Text
If IsNull(TBGravar!Data) = True Or TBGravar!Data = "" Then TBGravar!Data = Date Else TBGravar!Data = txtDataemissao
If IsNull(TBGravar!Responsavel) = True Or TBGravar!Responsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResp
TBGravar!Responsabilidade = txtResponsabilidade
TBGravar!IDCliente = IIf(txtIDcliente = "", 0, txtIDcliente)
TBGravar!IDforn = IIf(txtIDforn = "", 0, txtIDforn)
If Novo_FMEA = False Then
    If txtidproduto.Text <> TBGravar!IDProduto Then
        Conexao.Execute "DELETE from qualidadePPAP_FMEA_FASES where idFMEA = " & txtId
        ProcSalvarPlanoInspecao
    End If
End If
TBGravar!IDProduto = txtidproduto.Text
TBGravar!N_referencia = IIf(cmbReferencia = "", Null, cmbReferencia)
TBGravar!datacod = txtDataCod
TBGravar!Equipe = IIf(txtEquipe.Text = "", Null, txtEquipe.Text)
TBGravar!Aprovado = txtAprovado
TBGravar!Obs = IIf(txtObs = "", Null, txtObs)
TBGravar!DataChave = IIf(txtDatachave = "__/__/____", Null, txtDatachave)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposFases()
On Error GoTo tratar_erro

txtIDFase = 0
txtFase.Text = ""
txtGrupo = ""
txtRequisito.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposModoFalha()
On Error GoTo tratar_erro

txtID_ModoFalha = 0
txtModo_falha.Text = ""
txtSever = ""
txtControle_prevencao = ""
txtOcorrencia = ""
txtControle_deteccao = ""
txtDeteccao = ""
txtNPR = ""
txtAcoes_recomendadas = ""
txtResp_conclusao = ""

chkAcoes.Value = 0
txtAcoes_implementacoes = ""
txtOcorrencia_acoes = ""
txtSeveridade_acoes = ""
txtDeteccao_acoes = ""
txtNPR_acoes = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposEfeitoFalha()
On Error GoTo tratar_erro

txtID_Efeitofalha = 0
txtEfeito_potencial.Text = ""
txtClass.Text = ""
txtCausa_potencial = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadadosFases()
On Error GoTo tratar_erro

txtIDFase = TBAbrir!ID
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select * from plano where idplano = " & TBAbrir!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = False Then
    txtFase = IIf(IsNull(TBplano!Fase), "", TBplano!Fase)
    txtGrupo = IIf(IsNull(TBplano!Grupo_op), "", TBplano!Grupo_op)
End If
TBplano.Close

Set TBplanomedicao = CreateObject("adodb.recordset")
TBplanomedicao.Open "select * from Planodimensao where IDdimensao = " & TBAbrir!idDimensao, Conexao, adOpenKeyset, adLockOptimistic
If TBplanomedicao.EOF = False Then
    Texto = IIf(IsNull(TBplanomedicao!Tipo), "", TBplanomedicao!Tipo)
    Texto1 = IIf(IsNull(TBplanomedicao!dimdesejada), "", Format(TBplanomedicao!dimdesejada, "###,##0.0000"))
    txtRequisito.Text = Texto & " - " & Texto1
End If
TBplanomedicao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadadosModoFalha()
On Error GoTo tratar_erro

txtID_ModoFalha = TBAbrir!ID
txtModo_falha = IIf(IsNull(TBAbrir!ModoFalha), "", TBAbrir!ModoFalha)
txtSever = IIf(IsNull(TBAbrir!Sever), "", TBAbrir!Sever)
txtControle_prevencao = IIf(IsNull(TBAbrir!ControlePrevencao), "", TBAbrir!ControlePrevencao)
txtOcorrencia = IIf(IsNull(TBAbrir!Ocorrencia), "", TBAbrir!Ocorrencia)
txtControle_deteccao = IIf(IsNull(TBAbrir!ControleDeteccao), "", TBAbrir!ControleDeteccao)
txtDeteccao = IIf(IsNull(TBAbrir!Deteccao), "", TBAbrir!Deteccao)
txtNPR = IIf(IsNull(TBAbrir!NPR), "", TBAbrir!NPR)
txtAcoes_recomendadas = IIf(IsNull(TBAbrir!AcoesRecomendadas), "", TBAbrir!AcoesRecomendadas)
txtResp_conclusao = IIf(IsNull(TBAbrir!RespConclusao), "", TBAbrir!RespConclusao)

If TBAbrir!chkAcoes = True Then chkAcoes.Value = 1 Else chkAcoes.Value = 0
txtAcoes_implementacoes = IIf(IsNull(TBAbrir!AcoesImplementacoes), "", TBAbrir!AcoesImplementacoes)
txtSeveridade_acoes = IIf(IsNull(TBAbrir!Sever_acoes), "", TBAbrir!Sever_acoes)
txtOcorrencia_acoes = IIf(IsNull(TBAbrir!Ocorrencia_acoes), "", TBAbrir!Ocorrencia_acoes)
txtDeteccao_acoes = IIf(IsNull(TBAbrir!Deteccao_acoes), "", TBAbrir!Deteccao_acoes)
txtNPR_acoes = IIf(IsNull(TBAbrir!NPR_acoes), "", TBAbrir!NPR_acoes)
Frame8.Enabled = True
chkAcoes.Enabled = True
Novo_FMEA1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadadosEfeitoFalha()
On Error GoTo tratar_erro

txtID_Efeitofalha = TBAbrir!ID
txtEfeito_potencial = IIf(IsNull(TBAbrir!EfeitoPotencial), "", TBAbrir!EfeitoPotencial)
txtClass = IIf(IsNull(TBAbrir!Class), "", TBAbrir!Class)
txtCausa_potencial = IIf(IsNull(TBAbrir!CausaPotencial), "", TBAbrir!CausaPotencial)
Frame10.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkAcoes_Click()
On Error GoTo tratar_erro

If chkAcoes.Value = 0 Then
    txtAcoes_implementacoes = ""
    txtSeveridade_acoes = ""
    txtOcorrencia_acoes = ""
    txtDeteccao_acoes = ""
    txtNPR_acoes = ""
    Frame6.Enabled = False
Else
    Frame6.Enabled = True
End If

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

Private Sub cmdcliente_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False
frmVendas_LocalizarCliente.Show 1

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
    USMsgBox ("Informe o FMEA antes de copiar."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
frmQualidadePPAP_FMEA_Copiar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdExcluir_Efeitofalha_Click()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If txtID_Efeitofalha.Text = 0 Then
    USMsgBox ("Informe o efeito de falha antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir este efeito de falha?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "DELETE from QualidadePPAP_FMEA_EfeitoFalha WHERE ID = " & txtID_Efeitofalha
    USMsgBox ("Efeito de falha excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/FMEA"
    Evento = "Excluir efeito de falha"
    ID_documento = txtID_Efeitofalha
    Documento = "FMEA: " & txtFMEA & " - Cód. interno: " & txtCodinterno
    Documento1 = "Efeito potencial de falha: " & txtEfeito_potencial
    ProcGravaEvento
    '==================================
    ProcLimpaCamposEfeitoFalha
    Lista4.ListItems.Clear
    ProcCarregaListaEfeitofalha
    procCarregalistaPrincipal
    Frame10.Enabled = False
    Novo_FMEA2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdExcluir_ModoFalha_Click()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If txtID_ModoFalha.Text = 0 Then
    USMsgBox ("Informe o modo de falha antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir este modo de falha?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "DELETE from QualidadePPAP_FMEA_Modofalha WHERE ID = " & txtID_ModoFalha
    Conexao.Execute "DELETE from QualidadePPAP_FMEA_EfeitoFalha WHERE IDmodofalha = " & txtID_ModoFalha
    USMsgBox ("Modo de falha excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/FMEA"
    Evento = "Excluir modo de falha"
    ID_documento = txtID_ModoFalha
    Documento = "FMEA: " & txtFMEA & " - Cód. interno: " & txtCodinterno
    Documento1 = "Modo de falha potencial: " & txtModo_falha
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCamposModoFalha
    Lista3.ListItems.Clear
    ProcCarregaListaModofalha
    procCarregalistaPrincipal
    Frame8.Enabled = False
    chkAcoes.Enabled = False
    Novo_FMEA1 = False
End If

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
        txtdescricao = ""
        txtRevProduto = ""
        cmbReferencia.Clear
        txtidproduto = TBProduto!Codproduto
        txtCodinterno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
        txtRevProduto = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
        txtdescricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select N_Referencia from item_aplicacoes where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            cmbReferencia.AddItem ""
            Do While TBItem.EOF = False
                cmbReferencia.AddItem TBItem!N_referencia
                TBItem.MoveNext
            Loop
            TBItem.MoveFirst
            cmbReferencia = TBItem!N_referencia
        End If
        TBItem.Close
    Else
        USMsgBox ("Não foi encontrado nenhum produto com este código interno."), vbExclamation, "CAPRIND v5.0"
        txtidproduto = 0
        txtdescricao = ""
        txtRevProduto = ""
        cmbReferencia.Clear
    End If
    TBProduto.Close
Else
    txtidproduto = 0
    txtdescricao = ""
    txtRevProduto = ""
    cmbReferencia.Clear
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFornecedor_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False
FrmCompras_localizafornecedor.Show 1

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

Private Sub cmdImprimir_Efeitofalha_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimir_Fases_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImprimir_ModoFalha_Click()
On Error GoTo tratar_erro

ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarProduto_cliente_Click()
On Error GoTo tratar_erro

frmQualidadePPAP_LocalizarProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNovo_Efeitofalha_Click()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposEfeitoFalha
Frame10.Enabled = True
Novo_FMEA2 = True
txtEfeito_potencial.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNovo_ModoFalha_Click()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposModoFalha
Frame8.Enabled = True
chkAcoes.Enabled = True
Novo_FMEA1 = True
txtModo_falha.SetFocus

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

Private Sub cmdrevisao_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtId.Text = 0 Then
    USMsgBox ("Informe o FMEA antes de criar a revisão."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente criar uma revisão do FMEA " & txtFMEA.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If txtDataRev <> "" Then
        USMsgBox ("Não é permitido revisar o FMEA revisado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Contador = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from QualidadePPAP_FMEA where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcRevisar
        USMsgBox ("FMEA revisado com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
    TBAbrir.Close
    Lista.ListItems.Clear
    ProcCarregaLista
    '==================================
    Modulo = "Qualidade/FMEA"
    Evento = "Revisar"
    ID_documento = txtId
    Documento = "FMEA: " & txtFMEA & " - Cód. interno: " & txtCodinterno
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSair_Efeitofalha_Click()
On Error GoTo tratar_erro

ProcSair

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSair_Fases_Click()
On Error GoTo tratar_erro

ProcSair

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSair_ModoFalha_Click()
On Error GoTo tratar_erro

ProcSair

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSalvar_Efeitofalha_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame10.Enabled = False Then
    ProcVerificaSalvar
    cmdNovo_Efeitofalha.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * from QualidadePPAP_FMEA_EfeitoFalha where ID = " & txtID_Efeitofalha.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IdFMEA = txtId
    TBGravar!idfases = txtIDFase
    TBGravar!idModoFalha = txtID_ModoFalha
End If
ProcEnviadadosEfeitoFalha
TBGravar.Update
ProcCarregaListaEfeitofalha
procCarregalistaPrincipal
If Novo_FMEA2 = True Then
    USMsgBox ("Novo efeito de falha cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo efeito de falha"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar efeito de falha"
    If CodigoLista2 <> 0 And Lista4.ListItems.Count <> 0 Then
        Lista4.SelectedItem = Lista4.ListItems(CodigoLista2)
        Lista4.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/FMEA"
ID_documento = txtID_Efeitofalha
Documento = "FMEA: " & txtFMEA & " - Cód. interno: " & txtCodinterno
Documento1 = "Efeito potencial de falha: " & txtEfeito_potencial
ProcGravaEvento
'==================================
Frame8.Enabled = False
Novo_FMEA1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSalvar_ModoFalha_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame8.Enabled = False Then
    ProcVerificaSalvar
    cmdNovo_ModoFalha.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * from QualidadePPAP_FMEA_ModoFalha where ID = " & txtID_ModoFalha.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!IdFMEA = txtId
    TBGravar!idfases = txtIDFase
End If
ProcEnviadadosModoFalha
TBGravar.Update
ProcCarregaListaModofalha
procCarregalistaPrincipal
If Novo_FMEA1 = True Then
    USMsgBox ("Novo modo de falha cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo modo de falha"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar modo de falha"
    If CodigoLista1 <> 0 And Lista3.ListItems.Count <> 0 Then
        Lista3.SelectedItem = Lista3.ListItems(CodigoLista1)
        Lista3.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/FMEA"
ID_documento = txtID_ModoFalha
Documento = "FMEA: " & txtFMEA & " - Cód. interno: " & txtCodinterno
Documento1 = "Modo de falha potencial: " & txtModo_falha
ProcGravaEvento
'==================================
Frame8.Enabled = False
Novo_FMEA1 = False

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
            Case vbKeyF7: cmdCopiar_Click
            Case vbKeyF8: cmdrevisao_Click
            Case vbKeyEscape: imgSair_Click
        End Select
    Case 1:
        Select Case SSTab2.Tab
            Case 0:
                Select Case KeyCode
                    Case vbKeyF5: cmdImprimir_Fases_Click
                    Case vbKeyEscape: cmdSair_Fases_Click
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: cmdNovo_ModoFalha_Click
                    Case vbKeyF3: cmdSalvar_ModoFalha_Click
                    Case vbKeyF4: cmdExcluir_ModoFalha_Click
                    Case vbKeyF5: cmdImprimir_ModoFalha_Click
                    Case vbKeyEscape: cmdSair_ModoFalha_Click
                End Select
            Case 2:
                Select Case KeyCode
                    Case vbKeyInsert: cmdNovo_Efeitofalha_Click
                    Case vbKeyF3: cmdSalvar_Efeitofalha_Click
                    Case vbKeyF4: cmdExcluir_Efeitofalha_Click
                    Case vbKeyF5: cmdImprimir_Efeitofalha_Click
                    Case vbKeyEscape: cmdSair_Efeitofalha_Click
                End Select
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Formulario = "Qualidade/PPAP/FMEA"
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

frmQualidadePPAP_FMEA_Localizar.Show 1

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
Novo_FMEA = True
ProcLimparTudo
frmQualidadePPAP_FMEA_LocalizarProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame8.Enabled = False
Frame6.Enabled = False
Frame10.Enabled = False
ProcLimpaCamposFases
ProcLimpaCamposModoFalha
ProcLimpaCamposEfeitoFalha
Novo_FMEA1 = False
Novo_FMEA2 = False

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
    USMsgBox ("Informe o FMEA antes de excluir."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If

If USMsgBox("Deseja realmente excluir o nº " & txtFMEA.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "Update QualidadePPAP_FMEA Set DataRevisao = NULL, IDRevisao = 0 where idRevisao = " & IIf(txtId = "", 0, txtId)
    Conexao.Execute "DELETE from QualidadePPAP_FMEA WHERE ID = " & txtId
    Conexao.Execute "DELETE from QualidadePPAP_FMEA_Fases WHERE IDFMEA = " & txtId
    Conexao.Execute "DELETE from QualidadePPAP_FMEA_ModoFalha WHERE IDFMEA = " & txtId
    Conexao.Execute "DELETE from QualidadePPAP_FMEA_EfeitoFalha WHERE IDFMEA = " & txtId
    USMsgBox ("FMEA excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/FMEA"
    Evento = "Excluir"
    ID_documento = txtId
    Documento = "FMEA: " & txtFMEA & " - Cód. interno: " & txtCodinterno
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos
    Lista.ListItems.Clear
    ProcCarregaLista
    Frame2.Enabled = False
    ProcLimparTudo
    Novo_FMEA = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Plano de controle"
ProcLimpaVariaveisPrincipais

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
Qualidade_PPAP_FMEA = True
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

ProcSair

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
If txtFMEA = "" Then
    NomeCampo = "o número do FMEA"
    ProcVerificaAcao
    txtFMEA.SetFocus
    Exit Sub
End If
If txtidproduto = 0 Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    txtCodinterno.SetFocus
    Exit Sub
End If
If txtDataRev <> "" Then
    USMsgBox ("Não é permitido alterar o FMEA revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido2 = False
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * from QualidadePPAP_FMEA where ID = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
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
If Novo_FMEA = True Then
    USMsgBox ("Novo FMEA cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
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
Modulo = "Qualidade/FMEA"
ID_documento = txtId
Documento = "FMEA: " & txtFMEA & " - Cód. interno: " & txtCodinterno
Documento1 = ""
ProcGravaEvento
'==================================
Novo_FMEA = False
Caption = "Qualidade - PPAP - FMEA (" & txtFMEA & " - Rev.: " & txtRev & ")"

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
TBAbrir.Open "Select * from qualidadePPAP_FMEA where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
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
TBAbrir.Open "select * from QualidadePPAP_FMEA_Fases where id = " & Lista2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposFases
    ProcPuxadadosFases
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista3.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from QualidadePPAP_FMEA_ModoFalha where id = " & Lista3.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposModoFalha
    ProcPuxadadosModoFalha
    CodigoLista1 = Lista3.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista4_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista4.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from QualidadePPAP_FMEA_EfeitoFalha where id = " & Lista4.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposEfeitoFalha
    ProcPuxadadosEfeitoFalha
    Novo_FMEA2 = False
    CodigoLista2 = Lista4.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId.Text = "0" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        cmdNovo.SetFocus
    Case 1:
        If Novo_FMEA = True Then
            USMsgBox ("Salve o FMEA antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            CmdSalvar.SetFocus
            Exit Sub
        End If
        ProcLimpaCamposFases
        ProcCarregalistaFases
        procCarregalistaPrincipal
        SSTab2.Tab = 0
        cmdImprimir_Fases.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtId.Text = IIf(IsNull(TBAbrir!ID), "", TBAbrir!ID)
txtFMEA.Text = IIf(IsNull(TBAbrir!FMEA), "", TBAbrir!FMEA)
txtRev.Text = IIf(IsNull(TBAbrir!Rev), "", TBAbrir!Rev)
txtDataemissao.Text = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtDataRev.Text = IIf(IsNull(TBAbrir!DataRevisao), "", Format(TBAbrir!DataRevisao, "dd/mm/yy"))
txtResp.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtResponsabilidade.Text = IIf(IsNull(TBAbrir!Responsabilidade), "", TBAbrir!Responsabilidade)
txtidproduto.Text = IIf(IsNull(TBAbrir!IDProduto), "0", TBAbrir!IDProduto)

txtIDcliente = IIf(IsNull(TBAbrir!IDCliente), "0", TBAbrir!IDCliente)
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes where IDCliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    txtCliente = IIf(IsNull(TBClientes!NomeRazao), "", TBClientes!NomeRazao)
End If

txtIDforn = IIf(IsNull(TBAbrir!IDforn), "0", TBAbrir!IDforn)
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Compras_fornecedores where IDCliente = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    txtFornecedor = IIf(IsNull(TBClientes!Nome_Razao), "", TBClientes!Nome_Razao)
End If
TBClientes.Close

If txtidproduto <> "" Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where codproduto = " & txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        txtCodinterno = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
        txtRevProduto = IIf(IsNull(TBItem!RevDesenho), "", TBItem!RevDesenho)
        txtdescricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto & " order by n_referencia", Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                cmbReferencia.AddItem IIf(IsNull(TBCiclo!N_referencia), "", TBCiclo!N_referencia)
                TBCiclo.MoveNext
            Loop
        End If
        TBCiclo.Close
        If IsNull(TBAbrir!N_referencia) = False And TBAbrir!N_referencia <> "" Then cmbReferencia = TBAbrir!N_referencia
    End If
    TBItem.Close
End If

txtEquipe = IIf(IsNull(TBAbrir!Equipe), "", TBAbrir!Equipe)
txtAprovado = IIf(IsNull(TBAbrir!Aprovado), "", TBAbrir!Aprovado)
txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
txtDatachave = IIf(IsNull(TBAbrir!DataChave), "__/__/____", Format(TBAbrir!DataChave, "dd/mm/yyyy"))
Caption = "Qualidade - PPAP - FMEA (" & TBAbrir!FMEA & " - Rev.: " & TBAbrir!Rev & ")"
Frame2.Enabled = True
Novo_FMEA = False
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

If SQL_FMEA = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open SQL_FMEA, Conexao, adOpenKeyset, adLockOptimistic
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
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!FMEA), "", TBLISTA!FMEA)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Rev), "", TBLISTA!Rev)
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "select * from projproduto where codproduto = " & TBLISTA!IDProduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .Item(.Count).SubItems(3) = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
            End If
            TBItem.Close
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
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

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from qualidadePPAP_FMEA order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("ID = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId.Text = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from qualidadePPAP_FMEA where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
    Else
        USMsgBox ("Fim dos cadastros de FMEA."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_FMEA = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from qualidadePPAP_FMEA order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("ID = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId.Text = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from QualidadePPAP_FMEA where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
    Else
        USMsgBox ("Fim dos cadastros de FMEA."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_FMEA = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRevisar()
On Error GoTo tratar_erro

Contador = 0
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from qualidadePPAP_FMEA", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
Contador = TBAbrir!Rev
Contador = Contador + 1
txtRev = Contador
TBGravar!Rev = Contador
TBGravar!Data = TBAbrir!Data
TBGravar!Responsavel = pubUsuario
TBGravar!FMEA = TBAbrir!FMEA
TBGravar!IDProduto = TBAbrir!IDProduto
TBGravar!N_referencia = TBAbrir!N_referencia
TBGravar!datacod = TBAbrir!datacod
TBGravar!Responsabilidade = TBAbrir!Responsabilidade
TBGravar!IDCliente = TBAbrir!IDCliente
TBGravar!IDforn = TBAbrir!IDforn
TBGravar!Equipe = TBAbrir!Equipe
TBGravar!Aprovado = TBAbrir!Aprovado
TBGravar!Obs = TBAbrir!Obs
TBGravar!DataChave = TBAbrir!DataChave
TBAbrir!DataRevisao = Date
TBGravar.Update
txtId = TBGravar!ID
TBAbrir!IDRevisao = TBGravar!ID

Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "select * from qualidadePPAP_FMEA_FASES where IDFMEA = " & TBAbrir!ID, Conexao, adOpenKeyset, adLockOptimistic
Do While TBCiclo.EOF = False
    Set TBExecucao = CreateObject("adodb.recordset")
    TBExecucao.Open "select * from qualidadePPAP_FMEA_FASES", Conexao, adOpenKeyset, adLockOptimistic
    TBExecucao.AddNew
    TBExecucao!IdFMEA = TBGravar!ID
    TBExecucao!IDPlano = TBCiclo!IDPlano
    TBExecucao!idDimensao = TBCiclo!idDimensao
    
    TBExecucao.Update
    Set TBCarteira = CreateObject("adodb.recordset")
    TBCarteira.Open "select * from qualidadePPAP_FMEA_ModoFalha where IDfases = " & TBCiclo!ID, Conexao, adOpenKeyset, adLockOptimistic
    Do While TBCarteira.EOF = False
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select * from qualidadePPAP_FMEA_ModoFalha", Conexao, adOpenKeyset, adLockOptimistic
        TBCFOP.AddNew
        TBCFOP!IdFMEA = TBGravar!ID
        TBCFOP!idfases = TBExecucao!ID
        TBCFOP!ModoFalha = TBCarteira!ModoFalha
        TBCFOP!ControlePrevencao = TBCarteira!ControlePrevencao
        TBCFOP!Ocorrencia = TBCarteira!Ocorrencia
        TBCFOP!ControleDeteccao = TBCarteira!ControleDeteccao
        TBCFOP!Deteccao = TBCarteira!Deteccao
        TBCFOP!AcoesRecomendadas = TBCarteira!AcoesRecomendadas
        TBCFOP!RespConclusao = TBCarteira!RespConclusao
        TBCFOP!Sever = TBCarteira!Sever
        TBCFOP!NPR = TBCarteira!NPR
        If TBCarteira!chkAcoes = True Then
            TBCFOP!AcoesImplementacoes = TBCarteira!AcoesImplementacoes
            TBCFOP!NPR_acoes = TBCarteira!NPR_acoes
            TBCFOP!Ocorrencia_acoes = TBCarteira!Ocorrencia_acoes
            TBCFOP!Deteccao_acoes = TBCarteira!Deteccao_acoes
            TBCFOP!Sever_acoes = TBCarteira!Sever_acoes
            TBCFOP!chkAcoes = True
        End If
        
        TBCFOP.Update
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "select * from qualidadePPAP_FMEA_EfeitoFalha where IDModoFalha = " & TBCarteira!ID, Conexao, adOpenKeyset, adLockOptimistic
        Do While TBItem.EOF = False
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "select * from qualidadePPAP_FMEA_EfeitoFalha", Conexao, adOpenKeyset, adLockOptimistic
            TBProduto.AddNew
            TBProduto!IdFMEA = TBGravar!ID
            TBProduto!idfases = TBExecucao!ID
            TBProduto!idModoFalha = TBCFOP!ID
            TBProduto!EfeitoPotencial = TBItem!EfeitoPotencial
            TBProduto!Class = TBItem!Class
            TBProduto!CausaPotencial = TBItem!CausaPotencial
            TBProduto.Update
            TBProduto.Close
            TBItem.MoveNext
        Loop
        TBItem.Close
        
        TBCFOP.Close
        TBCarteira.MoveNext
    Loop
    TBCarteira.Close
    TBExecucao.Close
    TBCiclo.MoveNext
Loop
TBCiclo.Close

TBGravar.Close
TBAbrir.Update

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = 0 Then
    USMsgBox ("Informe o FMEA antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    cmdLocalizar.SetFocus
    Exit Sub
End If
NomeRel = "CQ_PPAP_FMEA.rpt"
ProcImprimirRel "{QualidadePPAP_FMEA.ID} = " & txtId, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalistaFases()
On Error GoTo tratar_erro

Lista2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from qualidadePPAP_FMEA_Fases where idFMEA = " & txtId & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
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
            Set TBplano = CreateObject("adodb.recordset")
            TBplano.Open "Select * from plano where idplano = " & TBLISTA!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
            If TBplano.EOF = False Then
                .Item(.Count).SubItems(1) = IIf(IsNull(TBplano!Fase), "", TBplano!Fase)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBplano!Grupo_op), "", TBplano!Grupo_op)
            End If
            TBplano.Close
            Set TBplanomedicao = CreateObject("adodb.recordset")
            TBplanomedicao.Open "select * from Planodimensao where IDdimensao = " & TBLISTA!idDimensao, Conexao, adOpenKeyset, adLockOptimistic
            If TBplanomedicao.EOF = False Then
                Texto = IIf(IsNull(TBplanomedicao!Tipo), "", TBplanomedicao!Tipo)
                Texto1 = IIf(IsNull(TBplanomedicao!dimdesejada), "", Format(TBplanomedicao!dimdesejada, "###,##0.0000"))
                .Item(.Count).SubItems(3) = Texto & " - " & Texto1
            End If
            TBplanomedicao.Close
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
            TBFIltro.Open "Select * from qualidadePPAP_FMEA_Fases where IDFMEA = " & txtId & " and Iddimensao = " & TBAbrir!idDimensao, Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = True Then TBFIltro.AddNew
            TBFIltro!IdFMEA = txtId
            TBFIltro!IDPlano = TBAbrir!IDPlano
            TBFIltro!idDimensao = TBAbrir!idDimensao
            
            Familiatext = ""
            Set TBInstrumentos = CreateObject("adodb.recordset")
            TBInstrumentos.Open "Select * from Planodimensao_instrumentos where ID_dimensao = " & TBAbrir!idDimensao & " order by Familia", Conexao, adOpenKeyset, adLockOptimistic
            If TBInstrumentos.EOF = False Then
                Do While TBInstrumentos.EOF = False
                    If IsNull(TBAbrir!Freq) = False And TBAbrir!Freq <> "" Then
                        If Familiatext <> "" Then Familiatext = Familiatext & " / " & TBInstrumentos!Familia Else Familiatext = TBAbrir!Freq & " - " & TBInstrumentos!Familia
                    Else
                        If Familiatext <> "" Then Familiatext = Familiatext & " / " & TBInstrumentos!Familia Else Familiatext = TBInstrumentos!Familia
                    End If
                    TBInstrumentos.MoveNext
                Loop
            End If
            TBInstrumentos.Close
            TBFIltro.Update
            
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select * from QualidadePPAP_FMEA_ModoFalha where IDFMEA = " & txtId & " and IDFases = " & TBFIltro!ID & " and ControleDeteccao = '" & Familiatext & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = True Then TBCFOP.AddNew
            TBCFOP!IdFMEA = txtId
            TBCFOP!idfases = TBFIltro!ID
            TBCFOP!ControleDeteccao = Familiatext
            TBCFOP.Update
            TBCFOP.Close
            
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

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtIDFase.Text = "0" Then
    SSTab2.Tab = 0
    Exit Sub
End If
Select Case SSTab2.Tab
    Case 0:
        ProcCarregalistaFases
        procCarregalistaPrincipal
        cmdImprimir_Fases.SetFocus
    Case 1:
        ProcLimpaCamposModoFalha
        ProcCarregaListaModofalha
        procCarregalistaPrincipal
        cmdNovo_ModoFalha.SetFocus
    Case 2:
        If txtID_ModoFalha = "0" Then
            SSTab2.Tab = 1
            USMsgBox ("Informe o modo de falha antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        If Novo_FMEA1 = True Then
            SSTab2.Tab = 1
            USMsgBox ("Salve o modo de falha antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            cmdSalvar_ModoFalha.SetFocus
            Exit Sub
        End If
        ProcLimpaCamposEfeitoFalha
        ProcCarregaListaEfeitofalha
        procCarregalistaPrincipal
        cmdNovo_Efeitofalha.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadadosModoFalha()
On Error GoTo tratar_erro

TBGravar!ModoFalha = txtModo_falha.Text
TBGravar!Sever = IIf(txtSever.Text = "", Null, txtSever.Text)
TBGravar!ControlePrevencao = txtControle_prevencao
TBGravar!Ocorrencia = IIf(txtOcorrencia = "", Null, txtOcorrencia)
TBGravar!ControleDeteccao = txtControle_deteccao
TBGravar!Deteccao = IIf(txtDeteccao.Text = "", Null, txtDeteccao.Text)
TBGravar!NPR = IIf(txtNPR.Text = "", Null, txtNPR.Text)
TBGravar!AcoesRecomendadas = txtAcoes_recomendadas.Text
TBGravar!RespConclusao = txtResp_conclusao.Text

If chkAcoes.Value = 1 Then TBGravar!chkAcoes = True Else TBGravar!chkAcoes = False
TBGravar!AcoesImplementacoes = txtAcoes_implementacoes.Text
TBGravar!Sever_acoes = IIf(txtSeveridade_acoes = "", Null, txtSeveridade_acoes)
TBGravar!Ocorrencia_acoes = IIf(txtOcorrencia_acoes = "", Null, txtOcorrencia_acoes)
TBGravar!Deteccao_acoes = IIf(txtDeteccao_acoes = "", Null, txtDeteccao_acoes)
TBGravar!NPR_acoes = IIf(txtNPR_acoes = "", Null, txtNPR_acoes)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadadosEfeitoFalha()
On Error GoTo tratar_erro

TBGravar!EfeitoPotencial = txtEfeito_potencial.Text
TBGravar!Class = IIf(txtClass = "", Null, txtClass)
TBGravar!CausaPotencial = txtCausa_potencial

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaModofalha()
On Error GoTo tratar_erro

Lista3.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from QualidadePPAP_FMEA_ModoFalha where idFMEA = " & txtId & " order by IDFases, ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista3.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!ModoFalha), "", TBLISTA!ModoFalha)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Sever), "", TBLISTA!Sever)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!ControlePrevencao), "", TBLISTA!ControlePrevencao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ocorrencia), "", TBLISTA!Ocorrencia)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!ControleDeteccao), "", TBLISTA!ControleDeteccao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Deteccao), "", TBLISTA!Deteccao)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!NPR), "", TBLISTA!NPR)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!AcoesRecomendadas), "", TBLISTA!AcoesRecomendadas)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!RespConclusao), "", TBLISTA!RespConclusao)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!AcoesImplementacoes), "", TBLISTA!AcoesImplementacoes)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Sever_acoes), "", TBLISTA!Sever_acoes)
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Ocorrencia_acoes), "", TBLISTA!Ocorrencia_acoes)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Deteccao_acoes), "", TBLISTA!Deteccao_acoes)
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!NPR_acoes), "", TBLISTA!NPR_acoes)
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

Sub ProcCarregaListaEfeitofalha()
On Error GoTo tratar_erro

Lista4.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from QualidadePPAP_FMEA_EfeitoFalha where idModofalha = " & txtID_ModoFalha, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista4.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!EfeitoPotencial), "", TBLISTA!EfeitoPotencial)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Class), "", TBLISTA!Class)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!CausaPotencial), "", TBLISTA!CausaPotencial)
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

Sub ProcSair()
On Error GoTo tratar_erro

If Novo_FMEA = True Then
    If USMsgBox("O FMEA ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        CmdSalvar_Click
        If Novo_FMEA = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_FMEA1 = True Then
    If USMsgBox("O modo de falha ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        cmdSalvar_ModoFalha_Click
        If Novo_FMEA1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_FMEA2 = True Then
    If USMsgBox("O modo de falha ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        cmdSalvar_Efeitofalha_Click
        If Novo_FMEA2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_FMEA = False
Novo_FMEA1 = False
Novo_FMEA2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtClass_LostFocus()
On Error GoTo tratar_erro

If txtClass <> "" Then
    VerifNumero = txtClass
    ProcVerificaNumero
    If VerifNumero = False Then
        txtClass = ""
        txtClass.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDeteccao_acoes_LostFocus()
On Error GoTo tratar_erro

If txtDeteccao_acoes <> "" Then
    VerifNumero = txtDeteccao_acoes
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDeteccao_acoes = ""
        txtDeteccao_acoes.SetFocus
        Exit Sub
    End If
End If
procCalculaNPR

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDeteccao_LostFocus()
On Error GoTo tratar_erro

If txtDeteccao.Text <> "" Then
    VerifNumero = txtDeteccao
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDeteccao = ""
        txtDeteccao.SetFocus
        Exit Sub
    End If
End If
procCalculaNPR

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOcorrencia_acoes_LostFocus()
On Error GoTo tratar_erro

If txtOcorrencia_acoes <> "" Then
    VerifNumero = txtOcorrencia_acoes
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOcorrencia_acoes = ""
        txtOcorrencia_acoes.SetFocus
        Exit Sub
    End If
End If
procCalculaNPR

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOcorrencia_LostFocus()
On Error GoTo tratar_erro

If txtOcorrencia.Text <> "" Then
    VerifNumero = txtOcorrencia
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOcorrencia = ""
        txtOcorrencia.SetFocus
        Exit Sub
    End If
End If
procCalculaNPR

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtsever_LostFocus()
On Error GoTo tratar_erro

If txtSever.Text <> "" Then
    VerifNumero = txtSever
    ProcVerificaNumero
    If VerifNumero = False Then
        txtSever = ""
        txtSever.SetFocus
        Exit Sub
    End If
End If
procCalculaNPR

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procCarregalistaPrincipal()
On Error GoTo tratar_erro

ListaFMEA.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from QualidadePPAP_FMEA_modoFalha where idFMEA = " & txtId & " order by IDFases, ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaFMEA.ListItems
            Qtde = 0
            Qtd = 0
            qt = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from qualidadePPAP_FMEA_fases where id = " & TBLISTA!idfases, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Set TBplano = CreateObject("adodb.recordset")
                TBplano.Open "Select * from plano where idplano = " & TBAbrir!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
                If TBplano.EOF = False Then
                    .Add , , IIf(IsNull(TBplano!Fase), "", TBplano!Fase)
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBplano!Grupo_op), "", TBplano!Grupo_op)
                End If
                TBplano.Close
                Set TBplanomedicao = CreateObject("adodb.recordset")
                TBplanomedicao.Open "select * from Planodimensao where IDdimensao = " & IIf(IsNull(TBAbrir!idDimensao), 0, TBAbrir!idDimensao), Conexao, adOpenKeyset, adLockOptimistic
                If TBplanomedicao.EOF = False Then
                    Texto = IIf(IsNull(TBplanomedicao!Tipo), "", TBplanomedicao!Tipo)
                    Texto1 = IIf(IsNull(TBplanomedicao!dimdesejada), "", Format(TBplanomedicao!dimdesejada, "###,##0.0000"))
                    .Item(.Count).SubItems(2) = Texto & " - " & Texto1
                End If
                TBplanomedicao.Close
            End If
            TBAbrir.Close
            
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!ModoFalha), "", TBLISTA!ModoFalha)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Sever), "", TBLISTA!Sever)
            Qtde = IIf(IsNull(TBLISTA!Sever), "0", TBLISTA!Sever)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!ControlePrevencao), "", TBLISTA!ControlePrevencao)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Ocorrencia), "", TBLISTA!Ocorrencia)
            qt = IIf(IsNull(TBLISTA!Ocorrencia), "0", TBLISTA!Ocorrencia)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!ControleDeteccao), "", TBLISTA!ControleDeteccao)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Deteccao), "", TBLISTA!Deteccao)
            Qtd = IIf(IsNull(TBLISTA!Deteccao), "0", TBLISTA!Deteccao)
            .Item(.Count).SubItems(12) = qt * Qtd * Qtde
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!AcoesRecomendadas), "", TBLISTA!AcoesRecomendadas)
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!RespConclusao), "", TBLISTA!RespConclusao)
            .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!AcoesImplementacoes), "", TBLISTA!AcoesImplementacoes)
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!Sever_acoes), "", TBLISTA!Sever_acoes)
            .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA!Ocorrencia_acoes), "", TBLISTA!Ocorrencia_acoes)
            .Item(.Count).SubItems(18) = IIf(IsNull(TBLISTA!Deteccao_acoes), "", TBLISTA!Deteccao_acoes)
            .Item(.Count).SubItems(19) = IIf(IsNull(TBLISTA!NPR_acoes), "", TBLISTA!NPR_acoes)

            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from qualidadePPAP_FMEA_efeitoFalha where idModoFalha = " & TBLISTA!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!EfeitoPotencial), "", TBAbrir!EfeitoPotencial)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Class), "", TBAbrir!Class)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!CausaPotencial), "", TBAbrir!CausaPotencial)
            End If
            TBAbrir.Close
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

Sub procCalculaNPR()
On Error GoTo tratar_erro

Qtde = 0
Qtd = 0
qt = 0
If txtSever <> "" And txtOcorrencia <> "" And txtDeteccao <> "" Then
    Qtde = txtSever
    Qtd = txtOcorrencia
    qt = txtDeteccao
    txtNPR = Qtde * Qtd * qt
Else
    txtNPR = ""
End If

Qtde = 0
Qtd = 0
qt = 0
If txtSeveridade_acoes <> "" And txtOcorrencia_acoes <> "" And txtDeteccao_acoes <> "" Then
    Qtde = txtSeveridade_acoes
    Qtd = txtOcorrencia_acoes
    qt = txtDeteccao_acoes
    txtNPR_acoes = Qtde * Qtd * qt
Else
    txtNPR_acoes = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSeveridade_acoes_LostFocus()
On Error GoTo tratar_erro

If txtSeveridade_acoes.Text <> "" Then
    VerifNumero = txtSeveridade_acoes
    ProcVerificaNumero
    If VerifNumero = False Then
        txtSeveridade_acoes = ""
        txtSeveridade_acoes.SetFocus
        Exit Sub
    End If
End If
procCalculaNPR

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
