VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlano_producao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Plano da produção"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15360
   ControlBox      =   0   'False
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   58
      Top             =   9720
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   17754
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
      TabCaption(0)   =   "Plano da produção"
      TabPicture(0)   =   "frmPlano_producao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Lista de OS's"
      TabPicture(1)   =   "frmPlano_producao.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameOS"
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(2)=   "Lista_OS"
      Tab(1).ControlCount=   3
      Begin VB.Frame FrameOS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   4335
         Left            =   -74925
         TabIndex        =   65
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_descricao_posto_OS 
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
            Left            =   7125
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do posto de trabalho."
            Top             =   390
            Width           =   6630
         End
         Begin VB.TextBox Txt_posto 
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
            Left            =   5490
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Posto de trabalho."
            Top             =   390
            Width           =   1620
         End
         Begin VB.TextBox Txt_fase 
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
            Left            =   2715
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Fase."
            Top             =   390
            Width           =   645
         End
         Begin VB.TextBox Txt_tempo_total 
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
            Height          =   315
            Left            =   4500
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Tempo total."
            Top             =   390
            Width           =   980
         End
         Begin VB.TextBox Txt_versao 
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Versão."
            Top             =   390
            Width           =   660
         End
         Begin VB.TextBox Txt_cod_ref 
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
            Left            =   2370
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   1560
            Width           =   1905
         End
         Begin VB.TextBox txt_cod_interno 
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
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   1560
            Width           =   1725
         End
         Begin VB.TextBox Txt_rev 
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   1560
            Width           =   435
         End
         Begin VB.TextBox Txt_cliente 
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
            Left            =   4155
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Cliente."
            Top             =   960
            Width           =   10860
         End
         Begin VB.TextBox Txt_ID_cliente 
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
            Left            =   3325
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Código do cliente."
            Top             =   960
            Width           =   805
         End
         Begin VB.TextBox Txt_qtde 
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
            Left            =   3375
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade."
            Top             =   390
            Width           =   1110
         End
         Begin VB.TextBox Txt_OS 
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
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Número da OS."
            Top             =   390
            Width           =   1395
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Empenhos (Pedidos)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1365
            Left            =   0
            TabIndex        =   66
            Top             =   2970
            Width           =   15195
            Begin MSComctlLib.ListView Lista_pedidos 
               Height          =   1005
               Left            =   180
               TabIndex        =   36
               Top             =   210
               Width           =   14835
               _ExtentX        =   26167
               _ExtentY        =   1773
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
                  Text            =   "ID"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Object.Tag             =   "N"
                  Text            =   "Cód. carteira"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Object.Tag             =   "N"
                  Text            =   "Ped. interno"
                  Object.Width           =   1676
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Object.Tag             =   "N"
                  Text            =   "Rev."
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Object.Tag             =   "T"
                  Text            =   "Cliente"
                  Object.Width           =   6345
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Object.Tag             =   "T"
                  Text            =   "Cód. interno"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   6
                  Object.Tag             =   "T"
                  Text            =   "Rev."
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Object.Tag             =   "T"
                  Text            =   "Cod. de ref."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Object.Tag             =   "T"
                  Text            =   "Descrição"
                  Object.Width           =   6345
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   9
                  Object.Tag             =   "N"
                  Text            =   "Qtde. vend."
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   10
                  Object.Tag             =   "D"
                  Text            =   "Prazo final"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Object.Tag             =   "T"
                  Text            =   "Tipo"
                  Object.Width           =   0
               EndProperty
            End
         End
         Begin VB.TextBox Txt_ordem 
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
            Left            =   2215
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Número da ordem."
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Txt_status 
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
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   960
            Width           =   2025
         End
         Begin VB.CommandButton Cmd_localizar_OS 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1590
            Picture         =   "frmPlano_producao.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Localizar OS's."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_prazo 
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
            Left            =   13770
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Prazo final."
            Top             =   390
            Width           =   1245
         End
         Begin VB.TextBox Txt_descricao 
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
            Left            =   4290
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   1560
            Width           =   10725
         End
         Begin VB.TextBox Txt_obs_ordem 
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
            Height          =   795
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Observações."
            Top             =   2130
            Width           =   14835
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   9307
            TabIndex        =   86
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   2572
            TabIndex        =   85
            Top             =   1350
            Width           =   1500
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   1965
            TabIndex        =   84
            Top             =   1350
            Width           =   345
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   9338
            TabIndex        =   83
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   3645
            TabIndex        =   82
            Top             =   750
            Width           =   165
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ordem"
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
            Left            =   2522
            TabIndex        =   81
            Top             =   750
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Versão"
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
            Left            =   2123
            TabIndex        =   80
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   2865
            TabIndex        =   79
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prazo final"
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
            Index           =   6
            Left            =   14017
            TabIndex        =   78
            Top             =   150
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição do posto de trabalho"
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
            Index           =   5
            Left            =   9323
            TabIndex        =   77
            Top             =   150
            Width           =   2235
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Index           =   4
            Left            =   5663
            TabIndex        =   76
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tempo total"
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
            Left            =   4563
            TabIndex        =   75
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quantidade"
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
            Left            =   3510
            TabIndex        =   74
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cód. interno"
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
            Left            =   592
            TabIndex        =   70
            Top             =   1350
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "OS*"
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
            Index           =   1
            Left            =   712
            TabIndex        =   69
            Top             =   180
            Width           =   330
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Status"
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
            Left            =   990
            TabIndex        =   68
            Top             =   750
            Width           =   465
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observação"
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
            Left            =   7155
            TabIndex        =   67
            Top             =   1920
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   60
         Top             =   9090
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
            ItemData        =   "frmPlano_producao.frx":013A
            Left            =   7020
            List            =   "frmPlano_producao.frx":0144
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   180
            Width           =   1965
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
            TabIndex        =   12
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
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
            Left            =   2880
            TabIndex        =   10
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   16
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao.frx":015C
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
            TabIndex        =   15
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao.frx":3900
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
            TabIndex        =   13
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
            TabIndex        =   14
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao.frx":7409
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
            TabIndex        =   17
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao.frx":B4F8
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
         Begin VB.Label Label21 
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
            Index           =   1
            Left            =   5670
            TabIndex        =   64
            Top             =   240
            Width           =   1260
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
            TabIndex        =   63
            Top             =   240
            Width           =   1095
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
            TabIndex        =   62
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               registros por página"
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
            TabIndex        =   61
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   39
         Top             =   1200
         Width           =   11820
         Begin VB.TextBox txtIDContato 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1800
            MaxLength       =   60
            MouseIcon       =   "frmPlano_producao.frx":ED84
            MousePointer    =   99  'Custom
            TabIndex        =   44
            ToolTipText     =   "Digite o nome para contato."
            Top             =   240
            Visible         =   0   'False
            Width           =   950
         End
         Begin VB.TextBox txtNomeContato 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmPlano_producao.frx":F08E
            MousePointer    =   99  'Custom
            TabIndex        =   43
            ToolTipText     =   "Nome do contato."
            Top             =   240
            Width           =   9855
         End
         Begin VB.TextBox txtdepartamento 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmPlano_producao.frx":F398
            MousePointer    =   99  'Custom
            TabIndex        =   42
            ToolTipText     =   "Departamento do contato."
            Top             =   630
            Width           =   9855
         End
         Begin VB.TextBox txttelcontato 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1770
            MaxLength       =   40
            MouseIcon       =   "frmPlano_producao.frx":F6A2
            MousePointer    =   99  'Custom
            TabIndex        =   41
            ToolTipText     =   "Ramal do contato."
            Top             =   1020
            Width           =   9855
         End
         Begin VB.TextBox TxtEmail_Contato 
            Enabled         =   0   'False
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
            Left            =   1770
            MouseIcon       =   "frmPlano_producao.frx":F9AC
            MousePointer    =   99  'Custom
            TabIndex        =   40
            ToolTipText     =   "E-mail do cliente."
            Top             =   1440
            Width           =   9855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
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
            Left            =   600
            TabIndex        =   48
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do contato:"
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
            Left            =   405
            TabIndex        =   47
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ramal:"
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
            Index           =   13
            Left            =   1200
            TabIndex        =   46
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
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
            Index           =   16
            Left            =   1215
            TabIndex        =   45
            Top             =   1478
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1965
         Left            =   75
         TabIndex        =   49
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_descricao_posto 
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
            Left            =   11580
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do posto de trabalho."
            Top             =   390
            Width           =   3435
         End
         Begin VB.ComboBox Cmb_posto 
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
            Left            =   9690
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Posto de trabalho."
            Top             =   390
            Width           =   1875
         End
         Begin VB.TextBox Txt_ID 
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
            Height          =   315
            Left            =   180
            TabIndex        =   0
            ToolTipText     =   "ID."
            Top             =   390
            Width           =   855
         End
         Begin VB.TextBox txtRespValidacao 
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
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   2355
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
            Left            =   5550
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   1755
         End
         Begin VB.TextBox Txt_data 
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
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   825
         End
         Begin VB.TextBox Txt_responsavel 
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
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2355
         End
         Begin VB.TextBox Txt_numero_plano 
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
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Número do plano de produção."
            Top             =   390
            Width           =   1275
         End
         Begin VB.TextBox Txt_obs 
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
            Height          =   825
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            ToolTipText     =   "Observações do plano."
            Top             =   990
            Width           =   14835
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição do posto de trabalho"
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
            Left            =   12180
            TabIndex        =   73
            Top             =   180
            Width           =   2235
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Posto de trabalho*"
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
            Left            =   9945
            TabIndex        =   72
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "ID"
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
            Left            =   570
            TabIndex        =   71
            Top             =   180
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   3900
            TabIndex        =   56
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   2580
            TabIndex        =   55
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   7507
            TabIndex        =   54
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   5700
            TabIndex        =   53
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000001&
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -11580
            TabIndex        =   52
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº plano"
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
            Left            =   1335
            TabIndex        =   51
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observação"
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
            Left            =   7155
            TabIndex        =   50
            Top             =   780
            Width           =   870
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   57
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   12
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   36
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   42
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   84
         ButtonTop3      =   2
         ButtonWidth3    =   44
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   130
         ButtonTop4      =   2
         ButtonWidth4    =   45
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   177
         ButtonTop5      =   2
         ButtonWidth5    =   60
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   239
         ButtonTop6      =   2
         ButtonWidth6    =   55
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   296
         ButtonTop7      =   2
         ButtonWidth7    =   55
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Validação"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Validar/Cancelar validação (F9)"
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   353
         ButtonTop8      =   2
         ButtonWidth8    =   62
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonAlignment9=   2
         ButtonType9     =   1
         ButtonStyle9    =   -1
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState9    =   -1
         ButtonLeft9     =   417
         ButtonTop9      =   4
         ButtonWidth9    =   2
         ButtonHeight9   =   54
         ButtonCaption10 =   "Ajuda"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Ajuda (F1)"
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   421
         ButtonTop10     =   2
         ButtonWidth10   =   41
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Sair"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Sair (Esc)"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   464
         ButtonTop11     =   2
         ButtonWidth11   =   30
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   5
         ButtonLeft12    =   496
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12090
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmPlano_producao.frx":FCB6
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   59
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   10
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Salvar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Salvar (F3)"
         ButtonKey2      =   "3"
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
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   44
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   86
         ButtonTop3      =   2
         ButtonWidth3    =   45
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
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
         ButtonLeft4     =   133
         ButtonTop4      =   2
         ButtonWidth4    =   60
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Registro anterior."
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
         ButtonLeft5     =   195
         ButtonTop5      =   2
         ButtonWidth5    =   55
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo registro."
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
         ButtonLeft6     =   252
         ButtonTop6      =   2
         ButtonWidth6    =   55
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonAlignment7=   2
         ButtonType7     =   1
         ButtonStyle7    =   -1
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState7    =   -1
         ButtonLeft7     =   309
         ButtonTop7      =   4
         ButtonWidth7    =   2
         ButtonHeight7   =   54
         ButtonCaption8  =   "Ajuda"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Ajuda (F1)"
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   313
         ButtonTop8      =   2
         ButtonWidth8    =   41
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Sair"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Sair (Esc)"
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   356
         ButtonTop9      =   2
         ButtonWidth9    =   30
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   5
         ButtonLeft10    =   388
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         ButtonUseMaskColor10=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   8970
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmPlano_producao.frx":168B0
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   5775
         Left            =   75
         TabIndex        =   9
         Top             =   3300
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10186
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
            Text            =   "ID"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Nº plano"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   7588
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Posto de trab."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Descrição do posto"
            Object.Width           =   7588
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Validado"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_OS 
         Height          =   4035
         Left            =   -74940
         TabIndex        =   37
         Top             =   5670
         Width           =   15195
         _ExtentX        =   26802
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
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "OS"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Versão"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Fase"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "D"
            Text            =   "Tempo total"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Posto de trab."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "D"
            Text            =   "Pr. final"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   4315
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "Cód. de ref."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   4315
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPlano_producao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_plano_prod     As Boolean 'OK
Dim Novo_plano_prod1       As Boolean 'OK
Public StrSql_plano_prod   As String 'OK
Dim TBLISTA_plano_prod     As ADODB.Recordset 'OK

Private Sub ProcAjuda()
On Error GoTo tratar_erro

'FunAbrirVideoWeb ("http://www.youtube.com/watch?v=i46JnPbSe98&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=36&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If Txt_numero_plano = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Planotexto from ProducaoFases_OS order by Planotexto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Planotexto = '" & Txt_numero_plano & "'")
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        Txt_numero_plano = TBLISTA!Planotexto
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from ProducaoFases_OS where Planotexto = '" & Txt_numero_plano & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaDadosPlano
        ProcCarregaListaOS
    Else
        USMsgBox ("Fim dos cadastros de plano da produção."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_plano_prod1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If Txt_numero_plano = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Planotexto from ProducaoFases_OS order by Planotexto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Planotexto = '" & Txt_numero_plano & "'")
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        Txt_numero_plano = TBLISTA!Planotexto
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from ProducaoFases_OS where Planotexto = '" & Txt_numero_plano & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaDadosPlano
        ProcCarregaListaOS
    Else
        USMsgBox ("Fim dos cadastros de plano da produção."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_plano_prod1 = False

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
        .ButtonState(8) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(8) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_posto_Click()
On Error GoTo tratar_erro

Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select Descricao from cadmaquinas where maquina = '" & Cmb_posto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    Txt_descricao_posto = TBMaquinas!Descricao
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_OS_Click()
On Error GoTo tratar_erro
  
frmPlano_producao_localizar_OS.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoOS()
On Error GoTo tratar_erro

If FunVerifValidacaoRegistro("criar nova", txtDtValidacao, "plano", "OS", True) = False Then Exit Sub
ProcLimpaCamposOS True
Novo_plano_prod1 = True
FrameOS.Enabled = True
Cmd_localizar_OS_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosOS(NovaOS As Boolean)
On Error GoTo tratar_erro

ProcLimpaCamposOS False
If Txt_OS <> "" Then
    If NovaOS = True Then TextoFiltro = " and and OS.Maquina = '" & Cmb_posto & "' and P.status <> 'Cancelada' and P.DtValidacao IS NOT NULL and (OS.ID_apontamento = " & Txt_ID & " or OS.ID_apontamento IS NULL and OS.Pronto = 'NÃO' or OS.ID_apontamento = 0 and OS.Pronto = 'NÃO')" Else TextoFiltro = ""
    Set TBOrdem = CreateObject("adodb.recordset")
    TBOrdem.Open "Select F.Versao, OS.Fase, OS.Quantidade, OS.TempoTotalLote, OS.Maquina, CM.Descricao, OS.Prazofinal, OS.Status, P.Ordem, P.IDcliente, P.Cliente, P.desenho, P.Revitem, P.N_Referencia, P.produto, OS.Obs from ((Ordemservico OS LEFT JOIN Fases F ON F.IDfase = OS.IDfase) INNER JOIN CadMaquinas CM ON CM.Maquina = OS.Maquina) INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.Idproducao = " & Txt_OS & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdem.EOF = False Then
        Txt_versao = IIf(IsNull(TBOrdem!versao), "", TBOrdem!versao)
        Txt_fase = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
        Txt_qtde = IIf(IsNull(TBOrdem!quantidade), "", TBOrdem!quantidade)
        Txt_tempo_total = IIf(IsNull(TBOrdem!TempoTotalLote), "00:00:00", TBOrdem!TempoTotalLote)
        Txt_posto = IIf(IsNull(TBOrdem!maquina), "", TBOrdem!maquina)
        Txt_descricao_posto_OS = IIf(IsNull(TBOrdem!Descricao), "", TBOrdem!Descricao)
        Txt_prazo = IIf(IsNull(TBOrdem!PrazoFinal), "", Format(TBOrdem!PrazoFinal, "dd/mm/yy"))
        Txt_status = IIf(IsNull(TBOrdem!status), "", TBOrdem!status)
        Txt_ordem = IIf(IsNull(TBOrdem!Ordem), "", TBOrdem!Ordem)
        Txt_ID_cliente = IIf(IsNull(TBOrdem!IDCliente), "", TBOrdem!IDCliente)
        Txt_cliente = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
        Txt_cod_interno = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
        txt_rev = IIf(IsNull(TBOrdem!Revitem), "", TBOrdem!Revitem)
        Txt_cod_ref = IIf(IsNull(TBOrdem!N_referencia), "", TBOrdem!N_referencia)
        Txt_descricao = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
        Txt_obs_ordem = IIf(IsNull(TBOrdem!Obs), "", TBOrdem!Obs)
        
        ProcCarregaListaPedidos
    End If
    TBOrdem.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_plano_prod.AbsolutePage <> 2 Then
    If TBLISTA_plano_prod.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_plano_prod.PageCount - 1)
    Else
        TBLISTA_plano_prod.AbsolutePage = TBLISTA_plano_prod.AbsolutePage - 2
        ProcExibePagina (TBLISTA_plano_prod.AbsolutePage)
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
    TBLISTA_plano_prod.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_plano_prod.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_plano_prod.AbsolutePage = 1
ProcExibePagina (TBLISTA_plano_prod.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_plano_prod.AbsolutePage <> -3 Then
    If TBLISTA_plano_prod.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_plano_prod.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_plano_prod.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_plano_prod.AbsolutePage = TBLISTA_plano_prod.PageCount
ProcExibePagina (TBLISTA_plano_prod.AbsolutePage)

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
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF9: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, "PCP/Plano da produção"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoOS
            Case vbKeyF3: ProcSalvarOS
            Case vbKeyF4: ProcExcluirOS
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Sub ProcCarregaDadosPlano()
On Error GoTo tratar_erro

Txt_ID = TBCompras!ID
Txt_numero_plano = IIf(IsNull(TBCompras!Planotexto), "", TBCompras!Planotexto)
Txt_data = IIf(IsNull(TBCompras!Data), "", (Format(TBCompras!Data, "dd/mm/yy")))
Txt_responsavel = IIf(IsNull(TBCompras!Responsavel), "", (TBCompras!Responsavel))
If IsNull(TBCompras!Posto) = False And TBCompras!Posto <> "" Then Cmb_posto = TBCompras!Posto
txtDtValidacao = IIf(IsNull(TBCompras!DtValidacao), "", (TBCompras!DtValidacao))
txtRespValidacao = IIf(IsNull(TBCompras!RespValidacao), "", TBCompras!RespValidacao)
Txt_obs.Text = IIf(IsNull(TBCompras!Observacao), "", TBCompras!Observacao)
Caption = "PCP - Plano da produção - (Plano : " & IIf(IsNull(TBCompras!Planotexto), "", TBCompras!Planotexto) & ")"
Frame1.Enabled = True
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaOS()
On Error GoTo tratar_erro

Lista_OS.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select OS.IDproducao, F.Versao, OS.Fase, OS.Quantidade, OS.TempoTotalLote, OS.Maquina, OS.Prazofinal, P.Ordem, P.IDcliente, P.Cliente, P.desenho, P.Revitem, P.N_Referencia, P.produto, OS.Obs from ((Ordemservico OS INNER JOIN Fases F ON F.IDfase = OS.IDfase) INNER JOIN CadMaquinas CM ON CM.Maquina = OS.Maquina) INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.ID_apontamento = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_OS.ListItems
            .Add , , TBLISTA!IDProducao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!versao), "", TBLISTA!versao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Fase), "", TBLISTA!Fase)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!quantidade), "", TBLISTA!quantidade)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!TempoTotalLote), "00:00:00", TBLISTA!TempoTotalLote)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Revitem), "", TBLISTA!Revitem)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Produto), "", TBLISTA!Produto)
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

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = ""
Txt_numero_plano = ""
Txt_data = Format(Date, "dd/mm/yy")
Txt_responsavel = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
ProcCarregaComboPostoTrab Cmb_posto, "Bloqueado = 'False'", False, False
Txt_descricao_posto = ""
Txt_obs = ""
CodigoLista = 0
Caption = "PCP - Plano da produção"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposOS(LimparOS As Boolean)
On Error GoTo tratar_erro

If LimparOS = True Then Txt_OS = ""
Txt_versao = ""
Txt_fase = ""
Txt_qtde = ""
Txt_tempo_total = ""
Txt_posto = ""
Txt_descricao_posto_OS = ""
Txt_prazo = ""
Txt_status = ""
Txt_ordem = ""
Txt_ID_cliente = ""
Txt_cliente = ""
Txt_cod_interno = ""
txt_rev = ""
Txt_cod_ref = ""
Txt_descricao = ""
Txt_obs_ordem = ""
Lista_pedidos.ListItems.Clear
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarOS()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "plano", "a OS", True) = False Then Exit Sub

If FrameOS.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"

If Txt_status = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    Txt_OS.SetFocus
    Exit Sub
End If

Conexao.Execute "Update Ordemservico Set ID_apontamento = " & Txt_ID & " where IDproducao = " & Txt_OS
ProcCarregaListaOS

If Novo_plano_prod1 = True Then
    USMsgBox ("Nova OS cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova OS"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar OS"
    With Lista_OS
        If CodigoLista1 <> 0 And .ListItems.Count <> 0 Then
            .SelectedItem = .ListItems(CodigoLista1)
            .SetFocus
        End If
    End With
End If
Novo_plano_prod1 = False
'==================================
Modulo = "PCP/Plano da produção"
ID_documento = Txt_OS
Documento = "Nº plano: " & Txt_numero_plano
Documento1 = "OS: " & Txt_OS
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirOS()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_OS
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) OS('s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBOS = CreateObject("adodb.recordset")
            TBOS.Open "Select Quantidade, IDFase, TTLPREVS, TempoTotalLote from Ordemservico where Ordem = " & .ListItems.Item(InitFor).ListSubItems(7) & " and Fase = '" & .ListItems.Item(InitFor).ListSubItems(2) & "' and ID_apontamento IS NULL", Conexao, adOpenKeyset, adLockOptimistic
            If TBOS.EOF = False Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Quantidade from Ordemservico where IDproducao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBOS!quantidade = TBOS!quantidade + TBFI!quantidade
                    Set TBFases = CreateObject("adodb.recordset")
                    TBFases.Open "Select TESegundos, TPSegundos from Fases where IDFase = " & TBOS!IDFase, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFases.EOF = False Then
                        DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * TBOS!quantidade) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
                        TBOS!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
                        TBOS!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
                    End If
                    TBFases.Close
                    TBOS.Update
                    TBFI.Delete
                End If
                TBFI.Close
            Else
                Conexao.Execute "Update Ordemservico Set ID_apontamento = NULL where IDproducao = " & .ListItems.Item(InitFor)
            End If
            TBOS.Close
            '==================================
            Modulo = "PCP/Plano da produção"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº plano: " & Txt_numero_plano
            Documento1 = "OS: " & .ListItems.Item(InitFor)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) OS('s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("OS('s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposOS True
    ProcCarregaListaOS
    FrameOS.Enabled = False
    Novo_plano_prod1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 12, True
ProcCarregaToolBar2 Me, 15195, 10, True

Formulario = "PCP/Plano da produção"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Cmb_opcao_lista = "Validação"


ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "PCP/Plano da produção"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmPlano_producao_localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir ests(s) plano(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "Update Ordemservico Set ID_apontamento = NULL where ID_apontamento = " & .ListItems.Item(InitFor)
            Conexao.Execute "DELETE from ProducaoFases_OS where ID = " & .ListItems(InitFor)
                        
            '==================================
            Modulo = "PCP/Plano da produção"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº plano: " & .ListItems(InitFor).SubItems(1)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) plano(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Plano(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcLimparTudo
    ProcCarregaLista (1)
    Novo_plano_prod = False
    Frame1.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro

If Txt_numero_plano.Text <> "" Then
    NomeRel = "Pcp_plano da producao.rpt"
    ProcImprimirRel "{ProducaoFases_OS.ID} = " & Txt_ID, ""
Else
    USMsgBox ("Informe o plano antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
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
Frame1.Enabled = True
Cmb_posto.SetFocus
Novo_plano_prod = True
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

FrameOS.Enabled = False
ProcLimpaCamposOS True
Lista_pedidos.ListItems.Clear
Lista_OS.ListItems.Clear
Novo_plano_prod1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_plano_prod = True Then
    If USMsgBox("O plano ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_plano_prod = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_plano_prod1 = True Then
    If USMsgBox("A OS ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarOS
        If Novo_plano_prod1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_plano_prod = False
Novo_plano_prod1 = False
Unload Me

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
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_posto = "" Then
    NomeCampo = "o posto de trabalho"
    ProcVerificaAcao
    Cmb_posto.SetFocus
    Exit Sub
End If
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from ProducaoFases_OS where ID = " & IIf(Txt_ID = "", 0, Txt_ID), Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then
    TBCompras.AddNew
    TBCompras!Planotexto = FunCriarNovoNumero
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "o mesmo", "plano", True) = False Then Exit Sub
End If
TBCompras!Data = IIf(Txt_data = "", Date, Txt_data)
TBCompras!Responsavel = IIf(Txt_responsavel = "", pubUsuario, Txt_responsavel)
TBCompras!Posto = Cmb_posto
TBCompras!Observacao = IIf(Txt_obs = "", Null, Txt_obs)
TBCompras.Update
ProcCarregaDadosPlano
TBCompras.Close

If Novo_plano_prod = True Then
    USMsgBox ("Novo plano cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_plano_prod = "Select PFOS.*, CM.Descricao from ProducaoFases_OS PFOS INNER JOIN CadMaquinas CM ON CM.Maquina = PFOS.Posto where PFOS.ID = " & Txt_ID
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
Modulo = "PCP/Plano da produção"
ID_documento = Txt_ID
Documento = "Nº plano: " & Txt_numero_plano
Documento1 = ""
ProcGravaEvento
'==================================
Novo_plano_prod = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_OS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "OS" Then
    With Lista_OS
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("ProducaoFases_OS", "ID = " & Txt_ID, True) = False Then GoTo Proximo
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

Private Sub Lista_OS_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_OS
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("ProducaoFases_OS", "ID = " & Txt_ID, "plano da produção", "OS", "excluir esta", True, True) = False Then
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

Private Sub Lista_OS_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcLimpaCamposOS True
Txt_OS = Lista_OS.SelectedItem
ProcCarregaDadosOS False
FrameOS.Enabled = True
CodigoLista1 = Lista_OS.SelectedItem.index
Novo_plano_prod1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_pedidos, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "ID" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidadoSemMsg("ProducaoFases_OS", "ID = " & .ListItems.Item(InitFor), True) = False Then GoTo Proximo
                Else
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select ID from ProducaoFases_OS where ID = " & .ListItems(InitFor) & " and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Set TBProducaoFases = CreateObject("adodb.recordset")
                        TBProducaoFases.Open "Select OS.IDproducao from Ordemservico OS LEFT JOIN ProducaoFases PF ON PF.OS = OS.IDproducao where OS.ID_apontamento = " & .ListItems.Item(InitFor) & " and (OS.Pronto = 'SIM' or PF.IDProducao IS NOT NULL)", Conexao, adOpenKeyset, adLockOptimistic
                        If TBProducaoFases.EOF = False Then
                            TBFI.Close
                            TBProducaoFases.Close
                            GoTo Proximo
                        End If
                        TBProducaoFases.Close
                    End If
                    TBFI.Close
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
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir" Then
                If FunVerificaRegistroValidado("ProducaoFases_OS", "ID = " & .ListItems.Item(InitFor), "mesmo", "plano da produção", "excluir este", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            Else
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ID from ProducaoFases_OS where ID = " & .ListItems(InitFor) & " and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Set TBProducaoFases = CreateObject("adodb.recordset")
                    TBProducaoFases.Open "Select OS.IDproducao from Ordemservico OS LEFT JOIN ProducaoFases PF ON PF.OS = OS.IDproducao where OS.ID_apontamento = " & .ListItems.Item(InitFor) & " and (OS.Pronto = 'SIM' or PF.IDProducao IS NOT NULL)", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProducaoFases.EOF = False Then
                        USMsgBox ("Não é permitido cancelar a validação deste plano, pois exite(m) OS('s) que já foi(ram) apontada(s)."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                    End If
                    TBProducaoFases.Close
                End If
                TBFI.Close
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
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from ProducaoFases_OS where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDadosPlano
    CodigoLista = Lista.SelectedItem.index
    Novo_plano_prod = False
End If
TBCompras.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If Txt_numero_plano = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        If Lista.Visible = True Then Lista.SetFocus
    Case 1:
        If Novo_plano_prod = True Then
            USMsgBox ("Salve o plano antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            Exit Sub
        End If
        Lista_OS.SetFocus
        ProcCarregaListaOS
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_OS_Change()
On Error GoTo tratar_erro

ProcLimpaCamposOS False
If Txt_OS <> "" Then
    VerifNumero = Txt_OS
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_OS = ""
        Txt_OS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCriarNovoNumero() As String
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Planotexto from ProducaoFases_OS where Year (Data) = '" & Year(Date) & "' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Numero = Left(TBAbrir!Planotexto, Len(TBAbrir!Planotexto) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBAbrir.Close

a = Numero
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: FunCriarNovoNumero = "PLP-0000" & Numero & "/" & Ano
    Case 2: FunCriarNovoNumero = "PLP-000" & Numero & "/" & Ano
    Case 3: FunCriarNovoNumero = "PLP-00" & Numero & "/" & Ano
    Case 4: FunCriarNovoNumero = "PLP-0" & Numero & "/" & Ano
    Case 5: FunCriarNovoNumero = "PLP-" & Numero & "/" & Ano
End Select

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_plano_prod = "" Then Exit Sub
Set TBLISTA_plano_prod = CreateObject("adodb.recordset")
TBLISTA_plano_prod.Open StrSql_plano_prod, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_plano_prod.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_plano_prod.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_plano_prod.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_plano_prod.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_plano_prod.RecordCount - IIf(Pagina > 1, (TBLISTA_plano_prod.PageSize * (Pagina - 1)), 0), TBLISTA_plano_prod.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_plano_prod.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_plano_prod!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_plano_prod!Planotexto), "", TBLISTA_plano_prod!Planotexto)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_plano_prod!Data), "", Format(TBLISTA_plano_prod!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_plano_prod!Responsavel), "", TBLISTA_plano_prod!Responsavel)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_plano_prod!Posto), "", TBLISTA_plano_prod!Posto)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_plano_prod!Descricao), "", TBLISTA_plano_prod!Descricao)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_plano_prod!DtValidacao), "Não", "Sim")
    End With
    TBLISTA_plano_prod.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_plano_prod.RecordCount
If TBLISTA_plano_prod.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_plano_prod.PageCount
ElseIf TBLISTA_plano_prod.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_plano_prod.PageCount & " de: " & TBLISTA_plano_prod.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_plano_prod.AbsolutePage - 1 & " de: " & TBLISTA_plano_prod.PageCount
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

Sub ProcCarregaListaPedidos()
On Error GoTo tratar_erro

Lista_pedidos.ListItems.Clear
If Txt_ordem = "" Or Txt_ordem = "0" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VC.*, VP.Ncotacao, VP.Revisao, VP.Cliente, PP.ID FROM (vendas_proposta VP INNER JOIN vendas_carteira VC ON VP.cotacao = VC.cotacao) INNER JOIN Producao_pedidos PP on VC.Codigo = PP.IDCarteira where PP.Ordem = " & Txt_ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_pedidos.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!CODIGO
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ncotacao), "", TBLISTA!Ncotacao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Rev_codinterno), "", TBLISTA!Rev_codinterno)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!descricao_tecnica), "", Trim(TBLISTA!descricao_tecnica))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.0000"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcValidarRegistros Lista, "PCP/Plano da produção"
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoOS
    Case 2: ProcSalvarOS
    Case 3: ProcExcluirOS
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
