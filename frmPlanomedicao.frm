VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlanomedicao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Controle de medição"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   330
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlanomedicao.frx":0000
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
      ScreenHeight    =   1080
      ScreenWidth     =   2560
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
      TabIndex        =   107
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
      SearchText      =   ""
      Value           =   0
   End
   Begin VB.ComboBox cmbcodref 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1900
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Código de referência."
      Top             =   2310
      Width           =   2325
   End
   Begin VB.ComboBox cmbNivel 
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
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmPlanomedicao.frx":0442
      Left            =   11400
      List            =   "frmPlanomedicao.frx":045E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      ToolTipText     =   "Nível."
      Top             =   2940
      Width           =   1020
   End
   Begin MSComctlLib.ListView ListaControle 
      Height          =   4485
      Left            =   75
      TabIndex        =   25
      Top             =   4590
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7911
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
         Text            =   "P. controle"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Aprovado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Restrição"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   6536
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Fase"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Versão"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Posto de trab."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "N° de rastreab."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Validado"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5295
      Left            =   75
      TabIndex        =   50
      Top             =   4410
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   9340
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Tipo da dimensão"
         Object.Width           =   5653
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Dim. indicada"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Dim. superior"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Dim. inferior"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Tol. sup."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Tol. inf."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Max. encontrada"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Min. encontrada"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Aprovado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Restrição"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Vista"
         Object.Width           =   1058
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_doc 
      Height          =   6105
      Left            =   75
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10769
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. da peça"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Caminho"
         Object.Width           =   23460
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Controle de medição"
      TabPicture(0)   =   "frmPlanomedicao.frx":0481
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USImageList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Dimensões"
      TabPicture(1)   =   "frmPlanomedicao.frx":049D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtnumero"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "USImageList2"
      Tab(1).Control(3)=   "USToolBar2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Documentos"
      TabPicture(2)   =   "frmPlanomedicao.frx":04B9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame14"
      Tab(2).Control(1)=   "txtID_doc"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CommonDialog1"
      Tab(2).Control(3)=   "USToolBar3"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2265
         Left            =   -74925
         TabIndex        =   117
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_cod_peca 
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
            Left            =   3000
            TabIndex        =   53
            ToolTipText     =   "Caminho."
            Top             =   390
            Width           =   1425
         End
         Begin VB.TextBox txtResponsavel_doc 
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
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   1935
         End
         Begin VB.TextBox txtData_doc 
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
            MaxLength       =   25
            TabIndex        =   51
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   855
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14040
            Picture         =   "frmPlanomedicao.frx":04D5
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Localizar arquivo (F2)"
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_caminho 
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
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "Caminho."
            Top             =   390
            Width           =   9585
         End
         Begin VB.TextBox Txt_obs_doc 
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
            Height          =   1095
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            ToolTipText     =   "Observação."
            Top             =   1020
            Width           =   14835
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmPlanomedicao.frx":05D7
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Visualizar arquivo."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frmPlanomedicao.frx":0B99
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Limpar caminho."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código da peça"
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
            Left            =   3157
            TabIndex        =   123
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Label42 
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
            Index           =   2
            Left            =   1560
            TabIndex        =   122
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho do arquivo"
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
            Left            =   8520
            TabIndex        =   121
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label Label42 
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
            Index           =   0
            Left            =   450
            TabIndex        =   119
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   7155
            TabIndex        =   118
            Top             =   810
            Width           =   945
         End
      End
      Begin VB.TextBox txtID_doc 
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
         Left            =   -70425
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   115
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   4110
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   110
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
            ItemData        =   "frmPlanomedicao.frx":0CD7
            Left            =   7170
            List            =   "frmPlanomedicao.frx":0CE1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   210
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
            TabIndex        =   27
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
            Left            =   2940
            TabIndex        =   26
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   31
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanomedicao.frx":0CF9
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
            TabIndex        =   30
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanomedicao.frx":449D
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
            TabIndex        =   28
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
            TabIndex        =   29
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanomedicao.frx":7FA6
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
            TabIndex        =   32
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanomedicao.frx":C095
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   5820
            TabIndex        =   133
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label18 
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
            Index           =   1
            Left            =   3570
            TabIndex        =   124
            Top             =   240
            Width           =   1440
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
            TabIndex        =   113
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
            TabIndex        =   112
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label30 
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
            Left            =   2250
            TabIndex        =   111
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox txtnumero 
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
         Left            =   -73830
         Locked          =   -1  'True
         MaxLength       =   20
         MouseIcon       =   "frmPlanomedicao.frx":F921
         MousePointer    =   99  'Custom
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   5070
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   77
         Top             =   1200
         Width           =   11820
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
            MouseIcon       =   "frmPlanomedicao.frx":FC2B
            MousePointer    =   99  'Custom
            TabIndex        =   79
            ToolTipText     =   "Departamento do contato."
            Top             =   630
            Width           =   9855
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
            MouseIcon       =   "frmPlanomedicao.frx":FF35
            MousePointer    =   99  'Custom
            TabIndex        =   78
            ToolTipText     =   "Nome do contato."
            Top             =   240
            Width           =   9855
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
            TabIndex        =   83
            Top             =   1478
            Width           =   480
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
            TabIndex        =   82
            Top             =   1080
            Width           =   495
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
            TabIndex        =   81
            Top             =   300
            Width           =   1290
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
            TabIndex        =   80
            Top             =   690
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   -74925
         TabIndex        =   73
         Top             =   360
         Width           =   11745
         Begin VB.CommandButton cmdagregar_lista 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   810
            MouseIcon       =   "frmPlanomedicao.frx":1023F
            MousePointer    =   99  'Custom
            Picture         =   "frmPlanomedicao.frx":10391
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Salvar."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdnovo_lista 
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
            Height          =   615
            Left            =   180
            MouseIcon       =   "frmPlanomedicao.frx":10B6A
            MousePointer    =   99  'Custom
            Picture         =   "frmPlanomedicao.frx":10CBC
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Novo."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdExcluir_lista 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   1440
            MouseIcon       =   "frmPlanomedicao.frx":111E2
            MousePointer    =   99  'Custom
            Picture         =   "frmPlanomedicao.frx":11334
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Alterar status para cancelado/requisitado."
            Top             =   180
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   3255
         Left            =   75
         TabIndex        =   60
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtinspetor 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   130
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   3015
         End
         Begin VB.TextBox txtDtValidacao 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5550
            Locked          =   -1  'True
            TabIndex        =   127
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   2025
         End
         Begin VB.TextBox txtRespValidacao 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7590
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   126
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   3015
         End
         Begin VB.CommandButton cmdRepRet 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14730
            Picture         =   "frmPlanomedicao.frx":11B83
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Criar RNC."
            Top             =   300
            Width           =   315
         End
         Begin VB.TextBox Txt_ID_RNC 
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
            Left            =   13470
            Locked          =   -1  'True
            TabIndex        =   120
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "ID RNC."
            Top             =   1620
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.TextBox Txt_posto_trab 
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
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Posto de trabalho."
            Top             =   1620
            Width           =   1640
         End
         Begin VB.TextBox txtVersao 
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
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Versão da fase."
            Top             =   1620
            Width           =   720
         End
         Begin VB.TextBox Txt_qtde_liberada 
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
            Left            =   7365
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade de peças liberadas."
            Top             =   1620
            Width           =   1305
         End
         Begin VB.TextBox Txt_qtde_lote 
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
            Left            =   6375
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade do lote."
            Top             =   1620
            Width           =   975
         End
         Begin VB.CommandButton cmdRNC 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmPlanomedicao.frx":11C65
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Criar RNC."
            Top             =   1620
            Width           =   315
         End
         Begin VB.TextBox TxtRNC 
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
            Left            =   13470
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "N° da RNC."
            Top             =   1620
            Width           =   1215
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reposição"
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
            Height          =   525
            Left            =   12030
            TabIndex        =   100
            Top             =   180
            Width           =   1335
            Begin VB.CheckBox Checklaudorepnao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "NÃO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   252
               Left            =   660
               TabIndex        =   5
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox Checklaudorepsim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SIM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   252
               Left            =   90
               TabIndex        =   4
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.TextBox txtQtde_amt 
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
            Left            =   12360
            TabIndex        =   21
            ToolTipText     =   "Quantidade amostra."
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aprovado"
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
            Height          =   525
            Left            =   10680
            TabIndex        =   62
            Top             =   180
            Width           =   1335
            Begin VB.CheckBox Checklaudoapronao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "NÃO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   252
               Left            =   660
               TabIndex        =   3
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox Checklaudoaprosim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SIM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   252
               Left            =   90
               TabIndex        =   2
               Top             =   225
               Width           =   612
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Retrabalho"
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
            Height          =   525
            Left            =   13380
            TabIndex        =   61
            Top             =   180
            Width           =   1335
            Begin VB.CheckBox Checklaudorestsim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SIM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   252
               Left            =   90
               TabIndex        =   6
               Top             =   240
               Width           =   555
            End
            Begin VB.CheckBox Checklaudorestnao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "NÃO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   252
               Left            =   660
               TabIndex        =   7
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.TextBox txtdesenho 
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
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   990
            Width           =   1635
         End
         Begin VB.TextBox txtdata 
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
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   1125
         End
         Begin VB.TextBox txtPm 
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
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Número do controle de medição."
            Top             =   375
            Width           =   1185
         End
         Begin VB.TextBox Txtpeca 
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
            Left            =   4605
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Nº de rastreabilidade."
            Top             =   1620
            Width           =   1755
         End
         Begin VB.TextBox txtdescricao 
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
            Left            =   4170
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   990
            Width           =   10845
         End
         Begin VB.TextBox txtfase 
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
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Fase."
            Top             =   1620
            Width           =   720
         End
         Begin VB.TextBox txtQtd 
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
            Left            =   8685
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Saldo."
            Top             =   1620
            Width           =   1305
         End
         Begin VB.TextBox txtQuant_liber 
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
            Left            =   10005
            MaxLength       =   30
            TabIndex        =   19
            ToolTipText     =   "Quantidade de peças encontradas."
            Top             =   1620
            Width           =   1305
         End
         Begin VB.TextBox txtgrupo 
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
            Left            =   3300
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Grupo/operação."
            Top             =   1620
            Width           =   1290
         End
         Begin VB.TextBox txtobservacaolaudo 
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
            Height          =   885
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            ToolTipText     =   "Observações para aprovação com restrição."
            Top             =   2220
            Width           =   14835
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   3570
            TabIndex        =   131
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora da validação"
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
            Left            =   5715
            TabIndex        =   129
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   8107
            TabIndex        =   128
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label31 
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
            Left            =   1833
            TabIndex        =   114
            Top             =   1410
            Width           =   1275
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1013
            TabIndex        =   106
            Top             =   1410
            Width           =   525
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liberadas"
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
            Left            =   7672
            TabIndex        =   103
            Top             =   1410
            Width           =   690
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
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
            Left            =   6705
            TabIndex        =   102
            Top             =   1410
            Width           =   315
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº RNC"
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
            Left            =   13807
            TabIndex        =   101
            Top             =   1410
            Width           =   540
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amostra"
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
            Left            =   12607
            TabIndex        =   99
            Top             =   1410
            Width           =   600
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nível"
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
            Left            =   11640
            TabIndex        =   98
            Top             =   1410
            Width           =   345
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de referência"
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
            Left            =   2250
            TabIndex        =   97
            Top             =   800
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno*"
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
            Index           =   0
            Left            =   375
            TabIndex        =   72
            Top             =   795
            Width           =   1335
         End
         Begin VB.Label Label9 
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
            Left            =   1770
            TabIndex        =   71
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P. controle"
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
            Index           =   0
            Left            =   322
            TabIndex        =   70
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de rastreabilidade*"
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
            Left            =   4680
            TabIndex        =   69
            Top             =   1410
            Width           =   1605
         End
         Begin VB.Label Label4 
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
            Index           =   0
            Left            =   9247
            TabIndex        =   68
            Top             =   800
            Width           =   690
         End
         Begin VB.Label Label16 
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
            Left            =   323
            TabIndex        =   67
            Top             =   1410
            Width           =   435
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Left            =   9142
            TabIndex        =   66
            Top             =   1410
            Width           =   390
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Encontradas"
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
            Left            =   10207
            TabIndex        =   65
            Top             =   1410
            Width           =   900
         End
         Begin VB.Label Label22 
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
            Left            =   3563
            TabIndex        =   64
            Top             =   1410
            Width           =   765
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações para aprovação com restrição"
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
            Left            =   6045
            TabIndex        =   63
            Top             =   2010
            Width           =   3135
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
         Height          =   3075
         Left            =   -74925
         TabIndex        =   84
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtdim_sup 
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
            Left            =   7110
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Dimensão superior."
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox txtDim_inf 
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
            Left            =   8430
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Dimenção inferior"
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox txtobservacaodim 
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
            Height          =   1320
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            ToolTipText     =   "Observações."
            Top             =   1620
            Width           =   5055
         End
         Begin VB.TextBox txtMin_enc 
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
            Left            =   13710
            MaxLength       =   30
            TabIndex        =   40
            ToolTipText     =   "Medida mínima encontrada."
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox Txtfrequencia 
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
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Frequência de medição."
            Top             =   990
            Width           =   6765
         End
         Begin VB.TextBox txtdesejada 
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
            Left            =   5790
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Dimensão indicada."
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox txttolinf 
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
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Tolerância inferior."
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox txttolsup 
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
            Left            =   9745
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Tolerância superior."
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox txtencontrada 
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
            Left            =   12390
            MaxLength       =   30
            TabIndex        =   39
            ToolTipText     =   "Medida máxima encontrada."
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox txtresponsavel 
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
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Responsável."
            Top             =   990
            Width           =   4785
         End
         Begin VB.TextBox txttipo 
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
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Tipo da dimensão."
            Top             =   375
            Width           =   5595
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Restrição"
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
            Height          =   525
            Left            =   13440
            TabIndex        =   86
            Top             =   795
            Width           =   1575
            Begin VB.CheckBox Checkdimrestsim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SIM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   180
               TabIndex        =   45
               Top             =   210
               Width           =   585
            End
            Begin VB.CheckBox Checkdimrestnao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "NÃO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   252
               Left            =   810
               TabIndex        =   46
               Top             =   210
               Width           =   732
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aprovado"
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
            Height          =   525
            Left            =   11850
            TabIndex        =   85
            Top             =   795
            Width           =   1575
            Begin VB.CheckBox Checkdimapronao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "NÃO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   252
               Left            =   810
               TabIndex        =   44
               Top             =   225
               Width           =   615
            End
            Begin VB.CheckBox Checkdimaprosim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SIM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   252
               Left            =   180
               TabIndex        =   43
               Top             =   225
               Width           =   732
            End
         End
         Begin MSComctlLib.ListView lista_familia 
            Height          =   1530
            Left            =   5340
            TabIndex        =   48
            Top             =   1410
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   2699
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Família do instrumento"
               Object.Width           =   16431
            EndProperty
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dim. inferior"
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
            Left            =   8647
            TabIndex        =   105
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dim. superior"
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
            TabIndex        =   104
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label24 
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
            Left            =   2235
            TabIndex        =   95
            Top             =   1410
            Width           =   945
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min. enc."
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
            Left            =   14032
            TabIndex        =   94
            Top             =   180
            Width           =   660
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frequência de medição"
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
            Left            =   2737
            TabIndex        =   93
            Top             =   800
            Width           =   1650
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tol. inf."
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
            Left            =   11445
            TabIndex        =   92
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dim. indicada"
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
            Left            =   5970
            TabIndex        =   91
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo da dimensão"
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
            Left            =   2355
            TabIndex        =   90
            Top             =   180
            Width           =   1245
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tol. sup."
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
            Left            =   10082
            TabIndex        =   89
            Top             =   180
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. enc."
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
            Left            =   12682
            TabIndex        =   88
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label6 
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
            Left            =   8895
            TabIndex        =   87
            Top             =   800
            Width           =   915
         End
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9570
         Top             =   780
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmPlanomedicao.frx":11D47
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   108
         Top             =   330
         Visible         =   0   'False
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   14
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
         ButtonCaption8  =   "Laudo"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Emitir laudo final (F7)"
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
         ButtonWidth8    =   42
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "RNC"
         ButtonEnabled9  =   0   'False
         ButtonToolTipText9=   "Abrir lista de RNC"
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
         ButtonLeft9     =   397
         ButtonTop9      =   2
         ButtonWidth9    =   30
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Validação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Validar/Cancelar validação  (F9)"
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
         ButtonLeft10    =   429
         ButtonTop10     =   2
         ButtonWidth10   =   62
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonAlignment11=   2
         ButtonType11    =   1
         ButtonStyle11   =   -1
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   -1
         ButtonLeft11    =   493
         ButtonTop11     =   4
         ButtonWidth11   =   2
         ButtonHeight11  =   54
         ButtonCaption12 =   "Ajuda"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Ajuda (F1)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   497
         ButtonTop12     =   2
         ButtonWidth12   =   41
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Sair"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Sair (Esc)"
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   540
         ButtonTop13     =   2
         ButtonWidth13   =   30
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   5
         ButtonLeft14    =   572
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         ButtonUseMaskColor14=   0   'False
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -61860
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   109
         Top             =   330
         Visible         =   0   'False
         Width           =   15195
         _ExtentX        =   26802
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   44
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Excluir"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Excluir (F4)"
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
         ButtonLeft2     =   48
         ButtonTop2      =   2
         ButtonWidth2    =   45
         ButtonHeight2   =   21
         ButtonCaption3  =   "Relatório"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Relatório (F5)"
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
         ButtonLeft3     =   95
         ButtonTop3      =   2
         ButtonWidth3    =   60
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Anterior"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Registro anterior."
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
         ButtonLeft4     =   157
         ButtonTop4      =   2
         ButtonWidth4    =   55
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Próximo"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Próximo registro."
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
         ButtonLeft5     =   214
         ButtonTop5      =   2
         ButtonWidth5    =   55
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Instrumentos"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Instumentos da dimensão (F7)"
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
         ButtonLeft6     =   271
         ButtonTop6      =   2
         ButtonWidth6    =   86
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Restrição"
         ButtonEnabled7  =   0   'False
         ButtonToolTipText7=   "Aprovar medição com restrição (F8)"
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
         ButtonLeft7     =   359
         ButtonTop7      =   2
         ButtonWidth7    =   62
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Dim. peça"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Dimensão por peça (F10)"
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
         ButtonLeft8     =   423
         ButtonTop8      =   2
         ButtonWidth8    =   63
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Atualizar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft9     =   488
         ButtonTop9      =   2
         ButtonWidth9    =   59
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
         ButtonLeft10    =   549
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   553
         ButtonTop11     =   2
         ButtonWidth11   =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   596
         ButtonTop12     =   2
         ButtonWidth12   =   30
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   5
         ButtonLeft13    =   628
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -69675
         Top             =   3930
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   116
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
            Name            =   "Tahoma"
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
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   9660
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
         End
      End
   End
End
Attribute VB_Name = "frmPlanomedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Aprovado                  As String  'OK
Dim MaxEncontrada             As Double  'OK
Dim MinEncontrada             As Double  'OK
Dim MediaEncontrada           As Double  'OK
Public Novo_Controle          As Boolean 'OK
Public Novo_Controle1         As Boolean 'OK
Public StrSql_Controlemedicao As String  'OK
Public Gravar_dimensao        As Boolean 'OK
Dim TBLISTA_Controle_Medicao     As ADODB.Recordset 'OK

Private Sub Checkdimapronao_Click()
On Error GoTo tratar_erro

If Checkdimapronao.Value = 1 Then Checkdimaprosim.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checkdimaprosim_Click()
On Error GoTo tratar_erro

If Checkdimaprosim.Value = 1 Then Checkdimapronao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checkdimrestnao_Click()
On Error GoTo tratar_erro

If Checkdimrestnao.Value = 1 Then Checkdimrestsim.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checkdimrestsim_Click()
On Error GoTo tratar_erro

If Checkdimrestsim.Value = 1 Then Checkdimrestnao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checklaudoapronao_Click()
On Error GoTo tratar_erro

If Checklaudoapronao.Value = 1 Then Checklaudoaprosim.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checklaudoaprosim_Click()
On Error GoTo tratar_erro

If Checklaudoaprosim.Value = 1 Then Checklaudoapronao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checklaudorepnao_Click()
On Error GoTo tratar_erro

If Checklaudorepnao.Value = 1 Then Checklaudorepsim.Value = 0
If Checklaudorestnao.Value = 1 Then
    Txt_ID_RNC = 0
    txtRNC = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checklaudorepsim_Click()
On Error GoTo tratar_erro

If Checklaudorepsim.Value = 1 Then Checklaudorepnao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checklaudorestnao_Click()
On Error GoTo tratar_erro

If Checklaudorestnao.Value = 1 Then Checklaudorestsim.Value = 0
If Checklaudorepnao.Value = 1 Then
    Txt_ID_RNC = 0
    txtRNC = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Checklaudorestsim_Click()
On Error GoTo tratar_erro

If Checklaudorestsim.Value = 1 Then Checklaudorestnao.Value = 0

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

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbNivel_Click()
On Error GoTo tratar_erro

txtQtde_amt = FunCalculaAmostragem(cmbNivel, IIf(txtQuant_liber = "", 0, txtQuant_liber))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcInstrumentos()
On Error GoTo tratar_erro

Acao = "cadatrar/visualizar os instrumentos da dimensão"
If txtNumero = "" Then
    NomeCampo = "a dimensão"
    ProcVerificaAcao
    Lista.SetFocus
    Exit Sub
End If
frmPlanomedicao_instrumentos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPeca()
On Error GoTo tratar_erro

Acao = "cadatrar/visualizar as dimensões por peça"
If txtNumero = "" Then
    NomeCampo = "a dimensão"
    ProcVerificaAcao
    Lista.SetFocus
    Exit Sub
End If
frmPlanomedicao_peca.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtPm.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from medicao order by IdPlano", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idplano = " & txtPm)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtPm.Text = TBLISTA!IDPlano
        Desenho = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from medicao where IdPlano = " & txtPm.Text, Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaDados
        ProcCarregaLista
        ProcCarregaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros de controle de medição."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Controle = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaControle(Pagina As Integer)
On Error GoTo tratar_erro

ListaControle.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_Controlemedicao = "" Then Exit Sub
Set TBLISTA_Controle_Medicao = CreateObject("adodb.recordset")
TBLISTA_Controle_Medicao.Open StrSql_Controlemedicao, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Controle_Medicao.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaControle.ListItems.Clear
TBLISTA_Controle_Medicao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Controle_Medicao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Controle_Medicao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Controle_Medicao.RecordCount - IIf(Pagina > 1, (TBLISTA_Controle_Medicao.PageSize * (Pagina - 1)), 0), TBLISTA_Controle_Medicao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Controle_Medicao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaControle.ListItems
        .Add , , TBLISTA_Controle_Medicao!IDPlano
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Controle_Medicao!Data), "", Format(TBLISTA_Controle_Medicao!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Controle_Medicao!Inspetor), "", TBLISTA_Controle_Medicao!Inspetor)
        If TBLISTA_Controle_Medicao!laudofinal = True Then .Item(.Count).SubItems(3) = "Sim" Else .Item(.Count).SubItems(3) = "Não"
        If TBLISTA_Controle_Medicao!restricao = True Then .Item(.Count).SubItems(4) = "Sim" Else .Item(.Count).SubItems(4) = "Não"
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Controle_Medicao!Desenho), "", TBLISTA_Controle_Medicao!Desenho)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Controle_Medicao!Descricao), "", TBLISTA_Controle_Medicao!Descricao)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Controle_Medicao!Fase), "", TBLISTA_Controle_Medicao!Fase)
        
        If IsNull(TBLISTA_Controle_Medicao!Fase) = False And TBLISTA_Controle_Medicao!Fase <> "" And TBLISTA_Controle_Medicao!Fase <> "0" Then
            Set TBFases = CreateObject("adodb.recordset")
            TBFases.Open "Select Versao FROM Fases where IDfase = " & TBLISTA_Controle_Medicao!IDFase, Conexao, adOpenKeyset, adLockOptimistic
            If TBFases.EOF = False Then
                .Item(.Count).SubItems(8) = IIf(IsNull(TBFases!versao), "", TBFases!versao)
            End If
            Set TBFases = CreateObject("adodb.recordset")
            TBFases.Open "Select Maquina FROM Ordemservico where IDProducao = " & TBLISTA_Controle_Medicao!ID_inspecionado, Conexao, adOpenKeyset, adLockOptimistic
            If TBFases.EOF = False Then
                .Item(.Count).SubItems(9) = IIf(IsNull(TBFases!maquina), "", TBFases!maquina)
            End If
            TBFases.Close
        End If
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Controle_Medicao!Peca), "", TBLISTA_Controle_Medicao!Peca)
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Controle_Medicao!RespValidacao), "", TBLISTA_Controle_Medicao!RespValidacao)
    End With
    TBLISTA_Controle_Medicao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Controle_Medicao.RecordCount
If TBLISTA_Controle_Medicao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Controle_Medicao.PageCount
ElseIf TBLISTA_Controle_Medicao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Controle_Medicao.PageCount & " de: " & TBLISTA_Controle_Medicao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Controle_Medicao.AbsolutePage - 1 & " de: " & TBLISTA_Controle_Medicao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarDim()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If NumeroME = 0 Then
    NomeCampo = "a dimensão na lista"
    ProcVerificaAcao
    Lista.SetFocus
    Exit Sub
End If
If txtencontrada = "" Then
    NomeCampo = "a dimensão máxima encontrada"
    ProcVerificaAcao
    txtencontrada.SetFocus
    Exit Sub
End If
If txtMin_enc = "" Then
    NomeCampo = "a dimensão mínima encontrada"
    ProcVerificaAcao
    txtMin_enc.SetFocus
    Exit Sub
End If
ProcCalculaMedicao

Set TBplanolaudo = CreateObject("adodb.recordset")
TBplanolaudo.Open "Select * from medicao where idplano = " & txtPm, Conexao, adOpenKeyset, adLockOptimistic
If TBplanolaudo!RespValidacao <> "Null" Or TBplanolaudo!RespValidacao <> "" Then
    MsgBox ("Plano de medição validado, não é possivel salvar!"), vbInformation + vbOKOnly
    TBplanolaudo.Close
    Exit Sub
End If

Set TBplanolaudo = CreateObject("adodb.recordset")
TBplanolaudo.Open "Select * from medicaodimensao where idmedicao = " & NumeroME, Conexao, adOpenKeyset, adLockOptimistic



'==================================
Modulo = "Qualidade/Controle de medição"
Evento = "Salvar medição"
ID_documento = txtPm
Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
Documento1 = "Tipo da dimensão: " & txttipo & " - Dimensão indicada: " & txtdesejada
ProcGravaEvento
'==================================
TBplanolaudo!max_enc = txtencontrada.Text
TBplanolaudo!min_enc = txtMin_enc
TBplanolaudo!Freq = Txtfrequencia
If Resultado = "Aprovado" Then
    TBplanolaudo!laudodim = "Sim"
    Checkdimaprosim.Value = 1
    TBplanolaudo!restricao = "Não"
    Checkdimrestnao.Value = 1
    USMsgBox ("Medição aprovada com sucesso."), vbInformation, "CAPRIND v5.0"
Else
    TBplanolaudo!laudodim = "Não"
    Checkdimapronao.Value = 1
    TBplanolaudo!restricao = "Não"
    Checkdimrestnao.Value = 1
    USMsgBox ("Medição reprovada com sucesso."), vbInformation, "CAPRIND v5.0"
End If
TBplanolaudo!Observacao = Trim(txtobservacaodim)
TBplanolaudo.Update
ProcCarregaLista
If CodigoLista1 <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista1)
    Lista.SetFocus
End If

ProcAtualizarAprovadoControle

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizarAprovadoControle()
On Error GoTo tratar_erro

Set TBplanolaudo = CreateObject("adodb.recordset")
TBplanolaudo.Open "Select * from medicaodimensao where IdPlano = " & txtPm & " and LaudoDim IS NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBplanolaudo.EOF = False Then
    TextoFiltro = "LaudoFinal = Null"
    Checklaudoaprosim.Value = 0
    Checklaudoapronao.Value = 0
Else
    Set TBplanolaudo = CreateObject("adodb.recordset")
    TBplanolaudo.Open "Select * from medicaodimensao where IdPlano = " & txtPm & " and LaudoDim = 'Não'", Conexao, adOpenKeyset, adLockOptimistic
    If TBplanolaudo.EOF = False Then
        TextoFiltro = "LaudoFinal = 'False'"
        Checklaudoaprosim.Value = 0
        Checklaudoapronao.Value = 1
        
        ProcAtualizaStatusDimInsp 2
    Else
        TextoFiltro = "LaudoFinal = 'True'"
        Checklaudoaprosim.Value = 1
        Checklaudoapronao.Value = 0
        
        ProcAtualizaStatusDimInsp 1
    End If
End If
TBplanolaudo.Close
Conexao.Execute "Update Medicao Set " & TextoFiltro & " where IdPlano = " & txtPm

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_doc()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame14.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txt_Caminho = "" Then
    NomeCampo = "o caminho"
    ProcVerificaAcao
    cmdImportar.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Medicao_documentos WHERE ID = " & txtID_doc, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadados_doc
TBGravar.Update
txtID_doc = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Doc
If Novo_Controle1 = True Then
    USMsgBox ("Novo documento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo documento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar documento"
    If CodigoLista2 <> 0 And Lista_doc.ListItems.Count <> 0 Then
        Lista_doc.SelectedItem = Lista_doc.ListItems(CodigoLista2)
        Lista_doc.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Controle de medição"
ID_documento = txtID_doc
Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
If Txt_cod_peca <> "" Then Documento1 = "Cód. da peça: " & Txt_cod_peca & " - Caminho: " & txt_Caminho Else Documento1 = "Caminho: " & txt_Caminho
ProcGravaEvento
'==================================
Novo_Controle1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_doc()
On Error GoTo tratar_erro

TBGravar!ID_PM = txtPm
TBGravar!Data = IIf(txtData_doc = "", Date, txtData_doc)
TBGravar!Responsavel = IIf(txtResponsavel_doc = "", pubUsuario, txtResponsavel_doc)
TBGravar!Codigo_peca = Txt_cod_peca
TBGravar!caminho = txt_Caminho
TBGravar!Obs = Trim(Txt_obs_doc)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaMedicao()
On Error GoTo tratar_erro

Desejada = Format(txtdesejada, "###,##0.0000")
TolSup = Format(Desejada + IIf(txttolsup = "", 0, txttolsup), "###,##0.0000")
TolInf = Format(Desejada + IIf(txttolinf = "", 0, txttolinf), "###,##0.0000")
MaxEncontrada = Format(IIf(txtencontrada = "", 0, txtencontrada), "###,##0.0000")
MinEncontrada = Format(IIf(txtMin_enc = "", 0, txtMin_enc), "###,##0.0000")
If MinEncontrada >= TolInf And MinEncontrada <= TolSup And MaxEncontrada >= TolInf And MaxEncontrada <= TolSup Then Resultado = "Aprovado" Else Resultado = "Reprovado"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLaudo()
On Error GoTo tratar_erro

Acao = "cadastrar/visualizar o laudo final"
If txtPm = "" Then
    NomeCampo = "o controle de medição"
    ProcVerificaAcao
    Exit Sub
End If
frmPlanomedicao_laudo.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRNC()
On Error GoTo tratar_erro

Acao = "abrir a lista de RNC"
If txtPm = "" Then
    NomeCampo = "o controle de medição"
    ProcVerificaAcao
    Exit Sub
End If
frmPlanomedicao_ListaRNC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

txt_Caminho = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txt_Caminho <> "" Then ProcAbrirArquivo txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Controle_Medicao.AbsolutePage <> 2 Then
    If TBLISTA_Controle_Medicao.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Controle_Medicao.PageCount - 1)
    Else
        TBLISTA_Controle_Medicao.AbsolutePage = TBLISTA_Controle_Medicao.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Controle_Medicao.AbsolutePage)
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
    TBLISTA_Controle_Medicao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Controle_Medicao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Controle_Medicao.AbsolutePage = 1
ProcExibePagina (TBLISTA_Controle_Medicao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Controle_Medicao.AbsolutePage <> -3 Then
    If TBLISTA_Controle_Medicao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Controle_Medicao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Controle_Medicao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Controle_Medicao.AbsolutePage = TBLISTA_Controle_Medicao.PageCount
ProcExibePagina (TBLISTA_Controle_Medicao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtPm.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from medicao order by IdPlano", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idplano =" & txtPm)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtPm.Text = TBLISTA!IDPlano
        Desenho = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from medicao where IdPlano = " & txtPm.Text, Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaDados
        ProcCarregaLista
        ProcCarregaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros de controle de medição."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Controle = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarRestricao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente aprovar esta(s) dimensão(ões) com restrição?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "UPDATE Medicaodimensao Set laudodim = 'Sim', restricao = 'Sim', Observacao = '" & Trim(txtobservacaodim) & "' where idmedicao = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Controle de medição"
            Evento = "Aprovar dimensão com restrição"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
            Documento1 = "Tipo da dimensão: " & .ListItems(InitFor).ListSubItems(2) & " - Dimensão indicada: " & .ListItems(InitFor).ListSubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) dimensão(ões) antes de aprovar com restrição."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Dimensão(ões) aprovada(s) com restrição com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
    Novo_Controle = False
    
    ProcAtualizarAprovadoControle
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdRepRet_Click()
On Error GoTo tratar_erro

ProcRepRet

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdRNC_Click()
On Error GoTo tratar_erro

RNC_Inspecao_Recebimento = False
RNC_Controle_Medicao = True
RNC_Nao_Conformidade = False
RNC_Solicitacao_Desvio = False
frmQualidade_RNC.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRepRet()
On Error GoTo tratar_erro

Acao = "visualizar a(s) ordem(ns) de reposição e a(s) Os(s) de retrabalho"
If txtPm = "" Then
    NomeCampo = "o controle de medição"
    ProcVerificaAcao
    Exit Sub
End If
frmPlanomedicao_ListaRepRet.Show 1

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
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcGravar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcLaudo
            Case vbKeyF8: ProcRepRet
            Case vbKeyF9: ProcRNC
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: ProcSalvarDim
            Case vbKeyF4: ProcExcluirDim
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcInstrumentos
            Case vbKeyF8: ProcSalvarRestricao
            Case vbKeyF10: ProcPeca
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_doc
            Case vbKeyF2: cmdImportar_Click
            Case vbKeyF3: procSalvar_doc
            Case vbKeyF4: procExcluir_doc
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBplanolaudo = CreateObject("adodb.recordset")
TBplanolaudo.Open "Select * from medicaodimensao where idplano = " & txtPm.Text & " order by indice, idmedicao", Conexao, adOpenKeyset, adLockOptimistic
If TBplanolaudo.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBplanolaudo.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBplanolaudo.EOF = False
        With Lista.ListItems.Add(, , TBplanolaudo!IDMedicao)
            .SubItems(1) = IIf(IsNull(TBplanolaudo!indice), "", TBplanolaudo!indice)
            .SubItems(2) = IIf(IsNull(TBplanolaudo!Tipo), "", TBplanolaudo!Tipo)
            .SubItems(3) = Format(TBplanolaudo!dimdesejada, "###,##0.0000")
            .SubItems(4) = Format(TBplanolaudo!Dim_superior, "###,##0.0000")
            .SubItems(5) = Format(TBplanolaudo!Dim_inferior, "###,##0.0000")
            .SubItems(6) = Format(TBplanolaudo!TolSup, "###,##0.0000")
            .SubItems(7) = Format(TBplanolaudo!TolInf, "###,##0.0000")
            .SubItems(8) = Format(TBplanolaudo!max_enc, "###,##0.0000")
            .SubItems(9) = Format(TBplanolaudo!min_enc, "###,##0.0000")
            
            If TBplanolaudo!laudodim = "Sim" Then .SubItems(10) = "Sim" Else .SubItems(10) = "Não"
            If TBplanolaudo!restricao = "Sim" Then .SubItems(11) = "Sim" Else .SubItems(11) = "Não"
            .SubItems(12) = IIf(IsNull(TBplanolaudo!Vista), "", TBplanolaudo!Vista)
            If TBplanolaudo!laudodim = "Não" Or IsNull(TBplanolaudo!laudodim) = True Then
                .ForeColor = vbRed
                .ListSubItems(1).ForeColor = vbRed
                .ListSubItems(2).ForeColor = vbRed
                .ListSubItems(3).ForeColor = vbRed
                .ListSubItems(4).ForeColor = vbRed
                .ListSubItems(5).ForeColor = vbRed
                .ListSubItems(6).ForeColor = vbRed
                .ListSubItems(7).ForeColor = vbRed
                .ListSubItems(8).ForeColor = vbRed
                .ListSubItems(9).ForeColor = vbRed
                .ListSubItems(10).ForeColor = vbRed
                .ListSubItems(11).ForeColor = vbRed
                .ListSubItems(12).ForeColor = vbRed
            End If
        End With
        TBplanolaudo.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBplanolaudo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Familia()
On Error GoTo tratar_erro

lista_familia.ListItems.Clear
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from Medicaodimensao_Familia where ID_dimensao = " & txtNumero.Text & " order by familia", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBFamilia.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBFamilia.EOF = False
        With lista_familia.ListItems
            .Add , , TBFamilia!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBFamilia!Familia), "", TBFamilia!Familia)
        End With
        TBFamilia.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBFamilia.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista_Doc()
On Error GoTo tratar_erro

Lista_doc.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from Medicao_documentos where ID_PM = " & txtPm & " order by Codigo_peca", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_doc.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codigo_peca), "", TBLISTA!Codigo_peca)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!caminho), "", TBLISTA!caminho)
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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 14, True
ProcCarregaToolBar2 Me, 15225, 13, True
ProcCarregaToolBar3 Me, 15225, 10, True

Formulario = "Qualidade/Controle de medição"
Direitos

Cmb_opcao_lista.ListIndex = 1

If Inspecaorecebimento_AnexarPlano = True Then
    Proclimparplano
    With frmCompras_recebimento
        txtdesenho.Text = .txtNomenclatura.Text
        txtdescricao.Text = .txtEspecificacoes.Text
        Txtpeca.Text = .Txt_lote
        txtFase.Text = 0
        txtQTD = Format(QTLOTE, "###,##0.0000")
        cmbNivel = .cmbNivel
        txtQuant_liber = .Txtenc
        txtQtde_amt = .Txtamostra
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from medicao where idlista = " & .ListProdReceb.SelectedItem.ListSubItems(5) & " and desenho = '" & txtdesenho & "' and peca = '" & .Txt_lote & "' and id_inspecionado = " & .ListProdReceb.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtPm.Text = TBAbrir!IDPlano
            ProcCarregaDados
            ProcCarregaLista
        Else
            ProcNovaMedicao
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from medicao where IDplano = " & txtPm, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Novo_Controle = True
                ProcCarregaDados
                Novo_Controle = True
                Gravar_dimensao = True
                ProcGravaDim
            End If
        End If
    End With
    Frame1.Enabled = True
    cmbNivel.Enabled = True
End If
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Controle de medição"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362CM" Then
    If USMsgBox("Deseja realmente atualizar os instrumentos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select idmedicao, instutilizado from Medicaodimensao order by idmedicao", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                Set TBInstrumentos = CreateObject("adodb.recordset")
                TBInstrumentos.Open "Select * from Medicaodimensao_instrumentos where idmedicao = " & TBAbrir!IDMedicao & " and instutilizado = '" & TBAbrir!instutilizado & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBInstrumentos.EOF = True Then TBInstrumentos.AddNew
                TBInstrumentos!IDMedicao = TBAbrir!IDMedicao
                TBInstrumentos!instutilizado = TBAbrir!instutilizado
                TBInstrumentos.Update
                TBInstrumentos.Close
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Controle de medição"
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

Private Sub ProcImprimir()
On Error GoTo tratar_erro

Acao = "abrir o menu de impressão"
If txtPm = "" Then
    Acao = "o controle de medição"
    ProcVerificaAcao
    Exit Sub
End If
frmPlanomedicao_menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro
  
frmPlanomedicao_Localizar.Show 1

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
Proclimparplano
ProcLimparTudo
frmPlanomedicao_Novo.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame2.Enabled = False
Frame14.Enabled = False
Proclimparmedida
Proclimpacampos_doc
Lista.ListItems.Clear
Lista_doc.ListItems.Clear
Novo_Controle1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Controle = True Then
    If USMsgBox("O controle de medição ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar
        If Novo_Controle = True Then Exit Sub Else Unload Me
    Else
        If txtPm <> "" Then ProcExcluir1 txtPm
    End If
End If
If Novo_Controle1 = True Then
    If USMsgBox("O documento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_doc
        If Novo_Controle1 = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Controle = False
Novo_Controle1 = False
Inspecaorecebimento_AnexarPlano = False
Unload Me

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
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
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

Private Sub Lista_doc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    Contador = 0
    With Lista_doc
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_doc, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_doc_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_doc.ListItems.Count = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select * from Medicao_documentos where id = " & Lista_doc.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    Proclimpacampos_doc
    ProcPuxadados_Doc
    CodigoLista2 = Lista_doc.SelectedItem.index
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxadados_Doc()
On Error GoTo tratar_erro

txtID_doc = TBMaterial!ID
txtData_doc = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
txtResponsavel_doc = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
Txt_cod_peca = IIf(IsNull(TBMaterial!Codigo_peca), "", TBMaterial!Codigo_peca)
txt_Caminho = IIf(IsNull(TBMaterial!caminho), "", TBMaterial!caminho)
Txt_obs_doc = IIf(IsNull(TBMaterial!Obs), "", TBMaterial!Obs)
Novo_Controle1 = False
Frame14.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Proclimparmedida
NumeroME = Lista.SelectedItem
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from medicaodimensao where idmedicao = " & NumeroME, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    txttipo.Text = IIf(IsNull(TBOrdem!Tipo), "", TBOrdem!Tipo)
    txtdesejada.Text = IIf(IsNull(TBOrdem!dimdesejada), "", Format(TBOrdem!dimdesejada, "###,##0.0000"))
    txtencontrada.Text = IIf(IsNull(TBOrdem!max_enc), "", Format(TBOrdem!max_enc, "###,##0.0000"))
    txtMin_enc = IIf(IsNull(TBOrdem!min_enc), "", Format(TBOrdem!min_enc, "###,##0.0000"))
    txttolsup.Text = IIf(IsNull(TBOrdem!TolSup), "", Format(TBOrdem!TolSup, "###,##0.0000"))
    txttolinf.Text = IIf(IsNull(TBOrdem!TolInf), "", Format(TBOrdem!TolInf, "###,##0.0000"))
    txtNumero.Text = IIf(IsNull(TBOrdem!IDMedicao), "", TBOrdem!IDMedicao)
    txtResponsavel.Text = pubUsuario
    If TBOrdem!laudodim = "Sim" Then
        Checkdimaprosim.Value = 1
        Checkdimapronao.Value = 0
    Else
        Checkdimaprosim.Value = 0
        Checkdimapronao.Value = 1
    End If
    If TBOrdem!restricao = "Sim" Then
        Checkdimrestsim.Value = 1
        Checkdimrestnao.Value = 0
    Else
        Checkdimrestsim.Value = 0
        Checkdimrestnao.Value = 1
    End If
    txtobservacaodim.Text = IIf(IsNull(TBOrdem!Observacao), "", TBOrdem!Observacao)
    Txtfrequencia = IIf(IsNull(TBOrdem!Freq), "", TBOrdem!Freq)
    txtdim_sup = IIf(IsNull(TBOrdem!Dim_superior), "", Format(TBOrdem!Dim_superior, "###,##0.0000"))
    txtDim_inf = IIf(IsNull(TBOrdem!Dim_inferior), "", Format(TBOrdem!Dim_inferior, "###,##0.0000"))
    ProcCarregaLista_Familia
    CodigoLista1 = Lista.SelectedItem.index
    Frame2.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravar()
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
If txtdesenho.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    Exit Sub
End If
If Txtpeca.Text = "" Then
    NomeCampo = "o número de rastreabilidade"
    ProcVerificaAcao
    Txtpeca.SetFocus
    Exit Sub
End If
valor = IIf(txtQuant_liber = "", 0, txtQuant_liber)
If valor <= 0 Then
    NomeCampo = "a quantidade de peças liberadas"
    ProcVerificaAcao
    txtQuant_liber.SetFocus
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Medicaodimensao where idPlano = " & txtPm, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Checklaudoaprosim.Value = 0 And Checklaudoapronao.Value = 0 Then
        USMsgBox "Informe se o controle de mediçao foi aprovado ou não.", vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If Checklaudorestsim.Value = 0 And Checklaudorestnao.Value = 0 Then
        USMsgBox "Informe se o controle de medição tem retrabalho ou não.", vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If Checklaudorepsim.Value = 0 And Checklaudorepnao.Value = 0 Then
        USMsgBox "Informe se o controle de medição tem reposição ou não.", vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End If
TBAbrir.Close
If cmbNivel = "" Then
    NomeCampo = "o nível"
    ProcVerificaAcao
    cmbNivel.SetFocus
    Exit Sub
End If
valor = IIf(txtQtde_amt = "", 0, txtQtde_amt)
If valor <= 0 Then
    NomeCampo = "a quantidade de amostra"
    ProcVerificaAcao
    txtQtde_amt.SetFocus
    Exit Sub
End If
Set TBplanomedicao = CreateObject("adodb.recordset")
TBplanomedicao.Open "Select * from medicao where IdPlano = " & txtPm, Conexao, adOpenKeyset, adLockOptimistic
If TBplanomedicao.EOF = True Then TBplanomedicao.AddNew
If Gravar_dimensao = True Then ProcGravaDim
TBplanomedicao!Data = txtData
TBplanomedicao!Inspetor = txtinspetor.Text
TBplanomedicao!Desenho = txtdesenho.Text
If Checklaudoaprosim.Value = 1 Then
    TBplanomedicao!laudofinal = True
    Aprovado = "SIM"
Else
    TBplanomedicao!laudofinal = False
    Aprovado = "NÃO"
End If

'Reposição
If Checklaudorepsim.Value = 1 Then
    TBplanomedicao!reposicao = True
    If USMsgBox("Deseja emitir uma nova ordem de reposicao?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem1:
        QuantSolicitado = 0
        reposicao = InputBox("Favor informar a quantidade para reposição.")
        If IsNumeric(reposicao) = True Then
            QuantSolicitado = reposicao
        Else
            USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
            GoTo Mensagem1
        End If
        If QuantSolicitado <> 0 Then ProcCriarOrdemReposicao
    End If
Else
    Set TBOrdem = CreateObject("adodb.recordset")
    TBOrdem.Open "Select * from producao where Idplano = " & txtPm & " and status <> 'Aberta'", Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdem.EOF = False Then
        USMsgBox ("Não é permitido alterar este controle de medição para reposição = NÃO, pois o mesmo está sendo utilizado no módulo PCP/Gerenciamento de ordem."), vbExclamation, "CAPRIND v5.0"
        TBOrdem.Close
        Exit Sub
    End If
    TBOrdem.Close
    TBplanomedicao!reposicao = False
    Set TBOrdem = CreateObject("adodb.recordset")
    TBOrdem.Open "Select * from producao where idplano = " & txtPm, Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdem.EOF = False Then
        Conexao.Execute "DELETE from Producaomaterial where Ordem = " & TBOrdem!Ordem
        Conexao.Execute "DELETE from Producao where Ordem = " & TBOrdem!Ordem
        Conexao.Execute "DELETE from Ordemservico_maq_utilizadas WHERE Ordem = " & TBOrdem!Ordem
        Conexao.Execute "DELETE from OSH from Ordemservico_HoraUtilizadaporDia OSH INNER JOIN Ordemservico OS ON OS.IDProducao = OSH.OS Where OS.Ordem = " & TBOrdem!Ordem
        Conexao.Execute "DELETE from OrdemServico where Ordem = " & TBOrdem!Ordem
    End If
End If

'Retrabalho
If Checklaudorestsim.Value = 1 Then
    TBplanomedicao!restricao = True
    If USMsgBox("Deseja emitir uma nova OS de retrabalho?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem:
        QuantSolicitado = 0
        Retrabalho = InputBox("Favor informar a quantidade para retrabalho.")
        If IsNumeric(Retrabalho) = True Then
            QuantSolicitado = Retrabalho
        Else
            USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
            GoTo Mensagem
        End If
        If QuantSolicitado <> 0 Then ProcCriarOSRetrabalho
    End If
ElseIf IsNumeric(Txtpeca) = True Then
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from producao where Ordem = " & Txtpeca & " and Ap_backup = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            NomeTabelaAp = "ProducaoFases_Backup"
        Else
            NomeTabelaAp = "ProducaoFases"
        End If
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * FROM Ordemservico INNER JOIN " & NomeTabelaAp & " ON Ordemservico.idproducao = " & NomeTabelaAp & ".OS where Ordemservico.ID_controle = " & txtPm & " and Ordemservico.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            USMsgBox ("Não é permitido alterar este controle de medição para retrabalho = NÃO, pois o mesmo está sendo utilizado no módulo PCP/Gerenciamento de ordem."), vbExclamation, "CAPRIND v5.0"
            TBOrdem.Close
            Exit Sub
        End If
        TBOrdem.Close
        TBplanomedicao!restricao = False
        Conexao.Execute "DELETE from OrdemServico where ID_controle = " & txtPm & " and retrabalho = 'True'"
End If

TBplanomedicao!Fase = IIf(txtFase = "", Null, txtFase)
TBplanomedicao!Quant_liberada = txtQuant_liber
TBplanomedicao!Peca = Txtpeca.Text
TBplanomedicao!Descricao = txtdescricao.Text
TBplanomedicao!Observacao = txtobservacaolaudo
TBplanomedicao!qtde_amostra = txtQtde_amt
TBplanomedicao!Nivel = cmbNivel
TBplanomedicao!ID_RNC = Txt_ID_RNC
TBplanomedicao.Update
If Novo_Controle = True Then
    USMsgBox ("Novo controle de medição cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_Controlemedicao = "Select * from Medicao where idplano = " & txtPm
    ProcCarregaListaControle (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaListaControle (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And ListaControle.ListItems.Count <> 0 Then
        ListaControle.SelectedItem = ListaControle.ListItems(CodigoLista)
        ListaControle.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Controle de medição"
ID_documento = txtPm
Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
Documento1 = ""
ProcGravaEvento
'==================================

'If IsNull(TBplanomedicao!IDFase) = False And TBplanomedicao!IDFase <> "" Then
'    Set TBFases = CreateObject("adodb.recordset")
'    TBFases.Open "Select * FROM CadMaquinas INNER JOIN ordemservico ON CadMaquinas.Maquina = ordemservico.Maquina where ordemservico.IDproducao = " & TBplanomedicao!ID_inspecionado & " and CadMaquinas.Insp_final = 'True'", Conexao, adOpenKeyset, adLockOptimistic
'    If TBFases.EOF = False Then
'        ProcGravarEstoque TBplanomedicao!ID_inspecionado
'    End If
'End If
TBplanomedicao.Close

If Checklaudoaprosim.Value = 1 Then ProcAtualizaStatusDimInsp 1 Else ProcAtualizaStatusDimInsp 2

Frame2.Enabled = True
Inspecaorecebimento_AnexarPlano = False
Novo_Controle = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaStatusDimInsp(TextoFiltro As Integer)
On Error GoTo tratar_erro

Conexao.Execute "Update CR Set CR.Dimensional = " & TextoFiltro & " from Compras_recebimento CR INNER JOIN Medicao M on M.IDlista = CR.IDestoque and M.id_inspecionado = CR.ID where M.idplano = " & txtPm

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarOSRetrabalho()
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from Ordemservico where IDProducao = " & TBplanomedicao!ID_inspecionado, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    'Busca dados das fases do processo
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select * from fases where idfase = " & TBOrdem!IDFase, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Ordemservico where ID_controle = " & txtPm & " and Retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!ID_controle = txtPm
        TBGravar!Ordem = Txtpeca
        TBGravar!IDFase = TBFases!IDFase
        TBGravar!IDPlano = TBOrdem!IDPlano
        TBGravar!quantidade = QuantSolicitado
        TBGravar!Fase = TBFases!Fase
        TBGravar!maquina = TBOrdem!maquina
        DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * QuantSolicitado) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
        TBGravar!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
        TBGravar!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
        TBGravar!Pronto = "NÃO"
        TBGravar!Preparacao = IIf(IsNull(TBFases!Preparacao), "00:00:00", TBFases!Preparacao)
        TBGravar!Execucao = IIf(IsNull(TBFases!Execucao), "00:00:00", TBFases!Execucao)
        TBGravar!IDPROCESSO = TBFases!IDPROCESSO
        TBGravar!PrazoFinal = TBOrdem!PrazoFinal
        TBGravar!descfase = TBFases!Descricao
        TBGravar!TempoPreparacao = TBFases!TempoPreparacao
        TBGravar!TempoExecucao = TBFases!TempoExecucao
        TBGravar!OSControlada = TBOrdem!OSControlada
        TBGravar!Processo_controlado = TBOrdem!Processo_controlado
        TBGravar!custos = TBOrdem!custos
        
        If TBFases!pecahora = True Then
            TBGravar!Pcshora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        Else
            If IsNull(TBGravar!Execucao) = False And TBGravar!Execucao <> "00:00:00" Then
                ElapsedTime (TBGravar!Execucao)
                TBGravar!Pcshora = 3600 / s
            End If
        End If
        TBGravar!pc_te = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        
        TBGravar!status = "Aguardando"
        TBGravar!Retrabalho = True
        
        'Verifica custo previsto da os
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from cadmaquinas where maquina = '" & TBOrdem!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            TotalFaseSeg = 0
            CustoFase = 0
            CustohoraSeg = 0
            TotalPreparacaoSeg = 0
            CustopreparacaoSeg = 0
            
            'Verifica custo de execucao por segundos * custo da hora maquina
            CustohoraSeg = TBMaquinas!PrecoHora / 3600
            ElapsedTime (TBOrdem!Execucao)
            TotalFaseSeg = s
            CustoFase = CustohoraSeg * TotalFaseSeg
            
            'Verifica custo de preparacao por segundos * custo da hora maquina
            If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then CustohoraSeg = TBMaquinas!PrecoHora_Setup / 3600
            ElapsedTime (TBOrdem!Preparacao)
            TotalPreparacaoSeg = s
            CustopreparacaoSeg = CustohoraSeg * TotalPreparacaoSeg
            
            If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then TBGravar!Valor_hs_prep = TBMaquinas!PrecoHora_Setup Else TBGravar!Valor_hs_prep = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
            TBGravar!Valor_hs_exec = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
            TBGravar!CPPECA = Format(CustoFase + (CustopreparacaoSeg / QuantSolicitado), "###,##0.0000000000")
            TBGravar!CPLOTE = Format(TBGravar!CPPECA * QuantSolicitado, "###,##0.00")
        End If
        TBMaquinas.Close
        
        TBGravar.Update
        TBGravar.Close
    End If
    TBFases.Close
End If
TBOrdem.Close

OF = Txtpeca
ProcAcertaOS Txt_qtde_lote, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarOrdemReposicao()
On Error GoTo tratar_erro

Mensagem:
    Familiatext = InputBox("Favor informar a versão do processo.")
    If Familiatext = "" Then
        Familiatext = "A"
    Else
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select fases.* from (fases INNER JOIN Processos ON Fases.IDProcesso = Processos.IDProcesso) INNER JOIN Projproduto ON Projproduto.Codproduto = Processos.Codproduto where Projproduto.Desenho = '" & txtdesenho & "' and Fases.Versao = '" & Familiatext & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = True Then
            USMsgBox ("Não foi encontrado nenhuma fase com esta versão para o código " & txtdesenho & "."), vbExclamation, "CAPRIND v5.0"
            TBFI.Close
            GoTo Mensagem
        End If
        TBFI.Close
    End If

    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from producao where Ordem = " & Txtpeca, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Ordem from producao order by Ordem desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Contador = TBAbrir!Ordem + 1
        Else
            Contador = 1
        End If
        ProcBuscaProcesso
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from Producao where Idplano = " & txtPm, Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = True Then
            TBOrdem.AddNew
            TBOrdem!status = "Aberta"
            TBOrdem!Saldo = True
            TBOrdem!Data_cadastro = Date
            TBOrdem!Responsavel = pubUsuario
            TBOrdem!Impof = 0
            TBOrdem!Escopo = False
        Else
            Contador = TBAbrir!Ordem
        End If
        ProcEnviaDadosOrdem
        ProcCriarRequisicao Familiatext
        ProcCriarOrdemServico Familiatext
        ProcAcertaOS TBOrdem!Quant, False
        TBOrdem.Update
        TBOrdem.Close
    End If
    TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBuscaProcesso()
On Error GoTo tratar_erro

IDPROCESSO = 0
Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select PR.IDprocesso from Processos PR INNER JOIN projproduto PROD ON PROD.Codproduto = PR.Codproduto where PROD.desenho = '" & txtdesenho & "' and PR.tipo <> 'C' and PR.Bloqueado = 'False' and PR.Revisao = (Select MAX(PR1.Revisao) from Processos PR1 where PR1.NProcesso = PR.NProcesso)", Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    IDPROCESSO = TBProcessos!IDPROCESSO
End If
TBProcessos.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosOrdem()
On Error GoTo tratar_erro

TBOrdem!ID_empresa = TBFI!ID_empresa
TBOrdem!IDPROCESSO = IDPROCESSO
TBOrdem!Ordem = Contador
TBOrdem!Quant = QuantSolicitado
TBOrdem!PrazoEntrega = TBFI!PrazoEntrega
Data = TBFI!PrazoEntrega
TBOrdem!Data = Date
TBOrdem!status = "Aberta"
TBOrdem!Desenho = TBFI!Desenho
TBOrdem!Revitem = TBFI!Revitem
If TBFI!N_referencia <> "" Then TBOrdem!N_referencia = TBFI!N_referencia
TBOrdem!Produto = TBFI!Produto
If TBFI!Cliente <> "" Then TBOrdem!Cliente = TBFI!Cliente
TBOrdem!Responsavel = pubUsuario
TBOrdem!pronta = "NÃO"
TBOrdem!IMPREQ = TBFI!IMPREQ
TBOrdem!Tipo = TBFI!Tipo
TBOrdem!Consignacao = TBFI!Consignacao
TBOrdem!OSControlada = TBFI!OSControlada
TBOrdem!Processo_controlado = TBFI!Processo_controlado

TBOrdem!reposicao = True
TBOrdem!IDPlano = txtPm
TBOrdem.Update
OF = TBOrdem!Ordem

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from producao_pedidos where Ordem = " & Txtpeca, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from producao_pedidos where Ordem = " & Contador & " and IDCarteira = " & TBAbrir!IDcarteira, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then TBGravar.AddNew
        TBGravar!Ordem = Contador
        TBGravar!IDcarteira = TBAbrir!IDcarteira
        TBGravar.Update
        
        Conexao.Execute "Update vendas_carteira Set Tem_ordem = 'True' where Codigo = " & TBAbrir!IDcarteira
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarRequisicao(Versao_fase As String)
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from producao where Ordem = " & Txtpeca & " and consignacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where desenho = '" & txtdesenho.Text & "' and DtValidacaoConj IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "Select PC.*, P.PBruto, P.SubTipoItem from projconjunto PC INNER JOIN projproduto P ON P.Desenho = PC.Desenho where PC.codproduto = " & TBItem!Codproduto & " and PC.Versao = '" & Versao_fase & "' and P.bloqueado = 'False' order by PC.Posicao, PC.codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBProcessos.EOF = False Then
            Do While TBProcessos.EOF = False
                Set TBMaterial = CreateObject("adodb.recordset")
                TBMaterial.Open "Select * from producaomaterial where Ordem = " & Contador & " and Codigo = '" & TBProcessos!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaterial.EOF = True Then TBMaterial.AddNew
                TBMaterial!Posicao = TBProcessos!Posicao
                TBMaterial!quantidade = TBProcessos!quantidade * QuantSolicitado
                TBMaterial!Unidade = TBProcessos!Unidade
                TBMaterial!CODIGO = TBProcessos!Desenho
                TBMaterial!Descricao = TBProcessos!Descricao
                TBMaterial!Ordem = Contador
                
                TBMaterial!PesoMetro = TBProcessos!PesoMetro
                TBMaterial!pesounidade = TBProcessos!Peso
                TBMaterial!PesoTotal = TBProcessos!PesoTotal * QuantSolicitado
                TBMaterial!Dimensao = TBProcessos!Dimensoes
                Peso = TBProcessos!quantidade
                If TBProcessos!Un_Kg <> "N/a" And TBProcessos!Un_Kg <> "" And (TBProcessos!Unidade = "KG" Or TBProcessos!Unidade = "MT" Or TBProcessos!Unidade = "MM" Or TBProcessos!Unidade = "M³") Then
                    Select Case TBProcessos!Unidade
                        Case "KG": Peso = TBProcessos!PesoTotal
                        Case "MT": Peso = (TBProcessos!Dimensoes / 1000) * TBProcessos!quantidade
                        Case "MM": Peso = TBProcessos!Dimensoes * TBProcessos!quantidade
                        Case "M³": Peso = TBProcessos!PesoTotal
                    End Select
                End If
                
 '===============================================================================================
 ' Acrescentar código abaixo
 '===============================================================================================
            If TBProcessos!Unidade = "M³" Then
                TBMaterial!Requisitado = Peso * txtQuantidade
                TBMaterial!DimensaoTotal = TBProcessos!Dimensoes
                TBMaterial!Total_pc = TBProcessos!quantidade
            Else
                TBMaterial!Requisitado = Format(Peso * txtQuantidade, "###,##0.0000")
                If TBProcessos!Un_Kg = "Mt²" Then TBMaterial!DimensaoTotal = ((TBProcessos!Dimensoes / 1000) / 1000) * TBMaterial!quantidade Else TBMaterial!DimensaoTotal = (TBProcessos!Dimensoes / 1000) * TBMaterial!quantidade
             If TBProcessos!Unidade = "KG" Or TBProcessos!SubTipoItem = 1 Or TBProcessos!SubTipoItem = 2 Or TBProcessos!SubTipoItem = 3 Then
                If TBProcessos!Unidade = "KG" And (TBProcessos!Un_Kg = "Mt²" Or TBProcessos!Un_Kg = "Mt/L") Then
                    If IsNull(TBProcessos!PBruto) = False And TBProcessos!PBruto > 0 And TBProcessos!PBruto <> "" Then TBMaterial!Total_pc = Format(TBMaterial!Requisitado / TBProcessos!PBruto, "###,##0.0000") Else TBMaterial!Total_pc = Null
                Else
                    If TBProcessos!Unidade = "PÇ" Or TBProcessos!Unidade = "PC" Or TBProcessos!Unidade = "UN" Or TBProcessos!Unidade = "CJ" Then TBMaterial!Total_pc = TBMaterial!Requisitado Else TBMaterial!Total_pc = Null
                End If
            End If
           End If
'=======================================================================================================
      
                TBMaterial!versao = TBProcessos!Versao_desenho
                TBMaterial!Saida = "NÃO"
                TBMaterial.Update
                TBProcessos.MoveNext
            Loop
            TBMaterial.Close
        End If
        TBProcessos.Close
    End If
    TBItem.Close
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarOrdemServico(Versao_fase As String)
On Error GoTo tratar_erro

TotalUtilizado = "00:00:00"
'Busca dados das fases do processo
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select F.* from fases F INNER JOIN Processos P ON P.IDProcesso = F.IDProcesso where F.idprocesso = " & IDPROCESSO & " AND F.versao = '" & Versao_fase & "' and P.DtValidacao IS NOT NULL order by F.fase", Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    Do While TBFases.EOF = False
        Set TBProducaoFases = CreateObject("adodb.recordset")
        TBProducaoFases.Open "Select * from ordemservico where Ordem = " & Contador & " and IdFase = " & TBFases!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBProducaoFases.EOF = True Then
            TBProducaoFases.AddNew
            If TBFases!Nao_aponta = True Then
                TBProducaoFases!Pronto = "SIM"
                TBProducaoFases!status = "Concluída"
                TBProducaoFases!DataConclusao = Date
            Else
                TBProducaoFases!Pronto = "NÃO"
                TBProducaoFases!status = "Aguardando"
            End If
        End If
        TBProducaoFases!Fase = TBFases!Fase
        TBProducaoFases!Rev_Fase = IIf(IsNull(TBFases!Revisao), 0, TBFases!Revisao)
        TBProducaoFases!Grupo_op = IIf(IsNull(TBFases!Grupo_op), "", TBFases!Grupo_op)
        TBProducaoFases!IDFase = TBFases!IDFase
        TBProducaoFases!IDPlano = FunVerifIDPlano(TBFases!IDFase)
        
        TBProducaoFases!maquina = TBFases!maquina
        
        DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * QuantSolicitado) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
        TBProducaoFases!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
        TBProducaoFases!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
        
        'Verifica se a maquina agrega custos/eficiencia na ordem
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from cadmaquinas where maquina = '" & TBFases!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            If TBMaquinas!custos = True Then TBProducaoFases!custos = True Else TBProducaoFases!custos = False
            If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then TBProducaoFases!Valor_hs_prep = TBMaquinas!PrecoHora_Setup Else TBProducaoFases!Valor_hs_prep = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
            TBProducaoFases!Valor_hs_exec = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
        End If
        TBMaquinas.Close
        
        TBProducaoFases!IDPROCESSO = TBFases!IDPROCESSO
        TBProducaoFases!Ordem = Contador
        TBProducaoFases!quantidade = QuantSolicitado
        TBProducaoFases!pecahora = TBFases!pecahora
        If TBFases!pecahora = True Then
            TBProducaoFases!Pcshora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        Else
            If IsNull(TBFases!Execucao) = False And TBFases!Execucao <> "00:00:00" Then
                ElapsedTime (TBFases!Execucao)
                TBProducaoFases!Pcshora = 3600 / s
            End If
        End If
        TBProducaoFases!pc_te = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        TBProducaoFases!Preparacao = IIf(IsNull(TBFases!Preparacao), "00:00:00", TBFases!Preparacao)
        TBProducaoFases!Execucao = IIf(IsNull(TBFases!Execucao), "00:00:00", TBFases!Execucao)
        TBProducaoFases!TempoPreparacao = TBFases!TempoPreparacao
        TBProducaoFases!TempoExecucao = TBFases!TempoExecucao
        TBProducaoFases!descfase = TBFases!Descricao
        If IsNull(TBFases!TESegundos) = True Or TBFases!TESegundos = "" Then
            ElapsedTime (TBProducaoFases!Execucao)
            TBProducaoFases!TESegundos = s
        Else
            TBProducaoFases!TESegundos = TBFases!TESegundos
        End If
        
        TBProducaoFases.Update
        TBFases.MoveNext
    Loop
End If

'Prazo final da OS
ProcDefinirPrazosOS Contador, Data, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'Sub ProcAcertaOS()
'On Error GoTo tratar_erro
'
'TotalFaseSeg = 0
'TotalFaseSegPrep = 0
'CustoOrdem = 0
'TotalOrdem = 0
'PcHora = 0
'Set TBOS = CreateObject("adodb.recordset")
'TBOS.Open "Select * from ordemServico where Ordem = " & OF & " order by idproducao", Conexao, adOpenKeyset, adLockOptimistic
'If TBOS.EOF = False Then
'    TBOS.MoveFirst
'    Do While TBOS.EOF = False
'        TOTALPECA = 0
'        TotalOS = 0
'
'        Set TBFases = CreateObject("adodb.recordset")
'        TBFases.Open "Select * FROM FASES WHERE IDFASE = " & TBOS!IDFase, Conexao, adOpenKeyset, adLockOptimistic
'        If TBFases.EOF = False Then
'            PcHora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
'
'            'Tempo total por peça
'            ElapsedTime (IIf(IsNull(TBFases!Execucao), 0, TBFases!Execucao))
'            If PcHora <> 0 Then TOTALPECA = TOTALPECA + (S / PcHora)
'
'            'Tempo total do lote
'            If PcHora <> 0 Then TotalOS = S / PcHora Else TotalOS = 0
'            ElapsedTime (IIf(IsNull(TBFases!Preparacao), 0, TBFases!Preparacao))
'            TotalOS = (TotalOS * TBOS!Quantidade) + S
'
'            'Custo total do lote
'            CustoOS = (TBFases!Custo * TBOS!Quantidade) + IIf(IsNull(TBFases!CustoPrep), 0, TBFases!CustoPrep)
'            CustoOrdem = CustoOrdem + CustoOS
'
'            TBOS!pecahora = TBFases!pecahora
'            If TBFases!pecahora = True Then
'                TBOS!Pcshora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
'            Else
'                If IsNull(TBOS!Execucao) = False And TBOS!Execucao <> "00:00:00" Then
'                    ElapsedTime (TBOS!Execucao)
'                    TBOS!Pcshora = 3600 / S
'                End If
'            End If
'            'Tempo total por peça
'            TBOS!TempoExecucao = TOTALPECA
'            TBOS!TempoExecucao = FormataTempo(TBOS!TempoExecucao)
'
'            TBOS!TTLPREVS = TotalOS 'Tempo total do lote previsto em segundos
'            TBOS!TempoTotalLote = FormataTempo(TBOS!TTLPREVS) 'Tempo total do lote previsto
'
'            'Custo por peça
'            If TBOS!Quantidade <> 0 Then TBOS!CPPECA = Format(TBFases!Custo + (IIf(IsNull(TBFases!CustoPrep), 0, TBFases!CustoPrep) / TBOS!Quantidade), "###,##0.0000000000") Else TBOS!CPPECA = Format(TBFases!Custo + IIf(IsNull(TBFases!CustoPrep), 0, TBFases!CustoPrep), "###,##0.0000000000")
'
'            'Custo do lote
'            TBOS!CPLOTE = Format(CustoOS, "###,##0.00")
'
'            TotalOrdem = TotalOrdem + TotalOS
'            TBOS.Update
'        End If
'        TBFases.Close
'        TBOS.MoveNext
'    Loop
'End If
'TBOS.Close
'
'Set TBAbrir = CreateObject("adodb.recordset")
'TBAbrir.Open "Select * from producao where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
'If TBAbrir.EOF = False Then
'    QuantSolicitado = TBAbrir!Quant
'
'    'Custo por peça
'    If Int(QuantSolicitado) <> 0 Then TBAbrir!cpp = CustoOrdem / Int(QuantSolicitado) Else TBAbrir!cpp = CustoOrdem
'    'Custo do lote
'    TBAbrir!CTTPrev = CustoOrdem
'    'Tempo total por peça
'    If TotalOrdem <> 0 Then
'        TBAbrir!TPP = TotalOrdem / Int(QuantSolicitado)
'        TBAbrir!TPP = FormataTempo(TBAbrir!TPP)
'    Else
'        TBAbrir!TPP = "00:00:00"
'    End If
'    'Tempo total do lote
'    TBAbrir!TTTPrev = TotalOrdem
'    TBAbrir!TTTPrev = FormataTempo(TBAbrir!TTTPrev)
'    'Tempo total do lote em segundos
'    TBAbrir!TTTPREVSegundos = TotalOrdem
'    TBAbrir.Update
'End If
'TBAbrir.Close
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Sub ProcVerificaQtdeLiberar()
On Error GoTo tratar_erro

If Txtpeca = "" Then Exit Sub
quantidade = 0
Qtde = 0
QTLOTE = 0

If TBAbrir!Tipo_inspecao = 1 Then
    'Pedido de compra, programação e consignação
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select Enc from Compras_recebimento where Id = " & TBAbrir!ID_inspecionado, Conexao, adOpenKeyset, adLockOptimistic
    If TBPedido.EOF = False Then
        Qtde = TBPedido!Enc
        Txt_qtde_lote = Format(TBPedido!Enc, "###,##0.0000")
    End If
    TBPedido.Close
    TextoFiltro = "IDlista = " & TBAbrir!IDlista & " and desenho = '" & TBAbrir!Desenho & "' and peca = '" & TBAbrir!Peca & "'"
Else
    'Ordem
    Set TBOrdem = CreateObject("adodb.recordset")
    TBOrdem.Open "Select Quantidade, QTOK from ordemservico where Idproducao = " & TBAbrir!ID_inspecionado, Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdem.EOF = False Then
        If IsNull(TBOrdem!QTOK) = False And TBOrdem!QTOK >= TBOrdem!quantidade Then Qtde = TBOrdem!QTOK Else Qtde = TBOrdem!quantidade
    End If
    TBOrdem.Close
    Txt_qtde_lote = Format(Qtde, "###,##0.0000")
    TextoFiltro = "ID_inspecionado = " & TBAbrir!ID_inspecionado
End If

If Novo_Controle = True Then TextoFiltro1 = "and IDplano <> " & txtPm

'Verifica qtde liberada
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Sum(Quant_liberada) as quantidade from Medicao where " & TextoFiltro & " " & TextoFiltro1, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    quantidade = IIf(IsNull(TBFI!quantidade), 0, TBFI!quantidade)
End If
TBFI.Close

Qtde = Format(Qtde - quantidade, "###,##0.0000")
If Qtde >= 0 Then txtQTD = Format(Qtde, "###,##0.0000") Else txtQTD = "0,0000"
Txt_qtde_liberada = Format(quantidade, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarEstoque(OS As Long)
On Error GoTo tratar_erro

Qtd_Prog = IIf(txtQuant_liber = "", 0, txtQuant_liber)
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select Qtdestoque, Id_inspecionado, laudofinal from medicao where idplano = " & txtPm.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = False Then
    If Aprovado = "SIM" And Qtd_Prog <> 0 Then
        qtdeliberada = Qtd_Prog
        quantestoque = IIf(IsNull(TBplano!QtdEstoque), 0, TBplano!QtdEstoque)
        If qtdeliberada <> quantestoque Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select * from estoque_controle where Lote = '" & Txtpeca & "' and desenho = '" & txtdesenho & "' and (status = 'ENTRADA_ORDEM' or status = 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                ProcCriarEntradaEstoque OS
                ProcGravarQtdeEstoque
            Else
                USMsgBox ("Não foi encontrado o cadastrado deste produto no módulo de engenharia, favor verificar."), vbExclamation, "CAPRIND v5.0"
                TBProduto.Close
                Exit Sub
            End If
            TBProduto.Close
            TBplano!QtdEstoque = txtQuant_liber.Text
            TBplano.Update
        End If
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from estoque_movimentacao where lote = '" & Txtpeca.Text & "' and Ordem = " & txtPm.Text, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_controle where IDEstoque = " & TBAbrir!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                ProcAlterarEntradaEstoqueExcluir txtPm, Txtpeca, TBplano!ID_inspecionado, TBplano!laudofinal, txtQuant_liber
                If qtdeliberada = 0 Then TBEstoque.Delete
            End If
            TBEstoque.Close
            Conexao.Execute "DELETE from estoque_movimentacao where lote = '" & Txtpeca.Text & "' and Ordem = " & txtPm.Text
            ProcGravarQtdeEstoque
            '==================================
            Modulo = "Qualidade/Controle de medição"
            Evento = "Excluir entrada no estoque"
            ID_documento = txtPm
            Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
        TBAbrir.Close
    End If
End If
TBplano.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaQtdeLiberada(OS As Long)
On Error GoTo tratar_erro

qtdeliberar = 0
qtdeliberada = 0
quantestoque = 0
If txtFase <> "" And txtFase <> "0" Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from producao where Ordem = " & Txtpeca.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        
        'Verifica se a ordem é controlada para verif. a qtde a liberar
        If TBproducao!OSControlada = False And TBproducao!Processo_controlado = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select QTOK from ordemservico where IDproducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If IsNull(TBAbrir!QTOK) = False And TBAbrir!QTOK <> "" And TBAbrir!QTOK <> "0" Then qtdeliberar = TBAbrir!QTOK Else qtdeliberar = TBproducao!Quant
            End If
            TBAbrir.Close
        End If
    End If
    TBproducao.Close
Else
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select compras_pedido_lista.* from compras_pedido inner join compras_pedido_lista on compras_pedido.idpedido = compras_pedido_lista.idpedido where compras_pedido.pedido = '" & Txtpeca & "' and compras_pedido_lista.desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBPedido.EOF = False Then
        qtdeliberar = TBproducao!Quant_Comp
    End If
    TBPedido.Close
End If

Qtd_Prog = txtQuant_liber.Text
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Qtdestoque) as qtdeliberada from medicao where idplano <> " & txtPm.Text & " and peca = '" & Txtpeca.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
    qtdeliberada = qtdeliberada + Qtd_Prog
Else
    qtdeliberada = Qtd_Prog
End If
TBAbrir.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarEntradaEstoque(OS As Long)
On Error GoTo tratar_erro

quantidade = 0
Qtde = 0
Permitido = False
If TBEstoque.EOF = True Then
    If USMsgBox("Deseja enviar " & Format(txtQuant_liber.Text, "###,##0.0000") & " " & TBProduto!Unidade & "(s) do produto '" & TBProduto!Desenho & "' ao estoque?, caso escolha não, você poderá fazer a movimentação manualmente.", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Permitido = True
        frmPlanoMedicao_LA.Show 1
        TBEstoque.AddNew
        If frmPlanoMedicao_LA.LA <> "" Then TBEstoque!local_armaz = frmPlanoMedicao_LA.LA
        USMsgBox ("Produto acrescentado ao estoque com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Nova entrada no estoque"
    End If
Else
    frmPlanoMedicao_LA.LA = IIf(IsNull(TBEstoque!local_armaz), "", TBEstoque!local_armaz)
    Permitido = True
    Evento = "Alterar entrada no estoque"
End If
If Permitido = True Then
    ProcEnviaDadosEstoqueControle
    ProcVerificaQtdeLiberada OS
    If qtdeliberada < qtdeliberar Then TBEstoque!status = "ENTRADA_ORDEM_PARCIAL" Else TBEstoque!status = "ENTRADA_ORDEM"
    TBEstoque.Update
    Ordem = OS
    ProcEnviaDadosEstoqueMovimentacao
    ProcGravarQtdeEstoque
    '==================================
    Modulo = "Qualidade/Controle de medição"
    ID_documento = txtPm
    Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosEstoqueControle()
On Error GoTo tratar_erro

TBEstoque!LOTE = Txtpeca.Text
TBEstoque!Certificado = 0
TBEstoque!Corrida = 0
TBEstoque!Desenho = txtdesenho.Text

'Atualiza valor do produto no estoque
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Producao where Ordem = " & Txtpeca, Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    TBEstoque!ID_Cliente = TBproducao!IDCliente
    TBEstoque!Cliente = TBproducao!Cliente
    If TBproducao!Consignacao = True Then TBEstoque!Consignacao = True
    TBEstoque!ID_empresa = TBproducao!ID_empresa
    TBEstoque!Ref = TBproducao!N_referencia
    
                                       'ORDEM            QTDE. PREVISTA                                      QTDE. OK                                                    QT. PROD.(OK+NC)                                                                                                     CUSTO LOTE                                              CUSTO PEÇA                                      CUSTO TERCEIROS                                             CUSTO MATERIAL                                                CUSTO OUTRAS                                                  ORDEM CONSIGNADA
    ValorTotal = FunCalculaValorUnitOrdem(TBproducao!Ordem, IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant), IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd), IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd) + IIf(IsNull(TBproducao!QuantNC), 0, TBproducao!QuantNC), IIf(IsNull(TBproducao!CTTReal), 0, TBproducao!CTTReal), IIf(IsNull(TBproducao!CPR), 0, TBproducao!CPR), IIf(IsNull(TBproducao!CTServico), 0, TBproducao!CTServico), IIf(IsNull(TBproducao!CTMaterial), 0, TBproducao!CTMaterial), IIf(IsNull(TBproducao!CTOutras), 0, TBproducao!CTOutras), TBproducao!Consignacao)
End If
TBproducao.Close

quantestoque = txtQuant_liber
TBEstoque!valor_unitario = Format(ValorTotal, "###,##0.0000000000")
TBEstoque!Valor_total = Format(quantestoque * ValorTotal, "###,##0.00")

TBEstoque!Descricao = txtdescricao.Text
TBEstoque!Un = TBProduto!Unidade
TBEstoque!Data = Format(Date, "dd/mm/yy")
TBEstoque!Responsavel = txtinspetor.Text
TBEstoque!Classe = TBProduto!Classe
TBEstoque!descricaotecnica = TBProduto!descricaotecnica
TBEstoque!peso_unit = TBProduto!peso_metro
TBEstoque!imagem = TBProduto!imagem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosEstoqueMovimentacao()
On Error GoTo tratar_erro

Qtd_Prog = txtQuant_liber
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from estoque_movimentacao where idestoque = " & TBEstoque!IDEstoque & " and lote = '" & Txtpeca.Text & "' and Ordem = " & txtPm.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    TBAbrir.AddNew
    TBAbrir!Data = Date
    TBAbrir!Destino = "Interno"
    TBAbrir!Terceiros = False
End If
TBAbrir!Entrada = Qtd_Prog

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Movimentar_estoque_pc from (Empresa E INNER JOIN producao P on P.ID_Empresa = E.codigo) INNER JOIN Ordemservico O on P.ordem = O.Ordem where E.Movimentar_estoque_pc = 'True' and O.IDproducao = " & Ordem & "", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    TBAbrir!Entrada_PC = Qtd_Prog
Else
    TBAbrir!Entrada_PC = Null
End If
TBFI.Close

TBAbrir!IDEstoque = TBEstoque!IDEstoque
TBAbrir!Operacao = TBEstoque!status
TBAbrir!Desenho = TBProduto!Desenho
TBAbrir!Descricao = txtdescricao

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from producao where Ordem = " & Txtpeca.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    TBAbrir!Cliente = TBproducao!Cliente
End If
TBproducao.Close

TBAbrir!Documento = Txtpeca.Text
TBAbrir!LOTE = Txtpeca.Text
TBAbrir!Responsavel = txtinspetor.Text
TBAbrir!DtEmissao = txtData.Text
TBAbrir!Ordem = txtPm.Text

'Atualiza valor
ValorTotal = TBEstoque!valor_unitario
quantestoque = txtQuant_liber
TBAbrir!VlrUnit = Format(ValorTotal, "###,##0.0000000000")
TBAbrir!vlrTotal = Format(quantestoque * ValorTotal, "###,##0.00")

TBAbrir.Update
TBAbrir.Close

ProcEmpenhaProdutoeAtualQtdeEntEmpOrdem Txtpeca, txtdesenho, Qtd_Prog, TBEstoque!IDEstoque

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarQtdeEstoque()
On Error GoTo tratar_erro

If frmPlanoMedicao_LA.LA <> "" Then TextoFiltro = "and EC.local_armaz = '" & frmPlanoMedicao_LA.LA & "'" Else TextoFiltro = ""
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open "Select Sum(EM.Saida) as QtdeSaida, Sum(EM.Entrada) as Qtd_Prog, Sum(ISNULL(EM.Saida_PC, 0)) as QtdeSaidaPC, Sum(ISNULL(EM.Entrada_PC, 0)) as Qtd_ProgPC FROM estoque_movimentacao EM INNER JOIN estoque_controle EC ON EM.IDEstoque = EC.IDEstoque where EC.Desenho = '" & txtdesenho & "' and EC.lote = '" & Txtpeca.Text & "' " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Qtd_Prog = IIf(IsNull(TBOS!Qtd_Prog), 0, TBOS!Qtd_Prog)
    Qtd_ProgPC = IIf(IsNull(TBOS!Qtd_ProgPC), 0, TBOS!Qtd_ProgPC)
    QtdeSaida = IIf(IsNull(TBOS!QtdeSaida), 0, TBOS!QtdeSaida)
    QtdeSaidaPC = IIf(IsNull(TBOS!QtdeSaidaPC), 0, TBOS!QtdeSaidaPC)
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from estoque_controle EC where Desenho = '" & txtdesenho & "' and lote = '" & Txtpeca.Text & "' " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        TBFI!estoque_real = Format(Qtd_Prog - QtdeSaida, "###,##0.0000")
        TBFI!estoque_real_PC = Format(Qtd_ProgPC - QtdeSaidaPC, "###,##0.0000")
        TBFI!estoque_venda = TBFI!estoque_real
        TBFI!Qtde = TBFI!estoque_real
        TBFI!Valor_total = Format(TBFI!valor_unitario * TBFI!estoque_real, "###,##0.0000000000")
        TBFI.Update
    End If
    TBFI.Close
End If
TBOS.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravaDim()
On Error GoTo tratar_erro

If txtFase <> "" And txtFase <> "0" Then
    INNERJOINTEXTO = "(Plano P INNER JOIN Fases F ON F.IDFase = P.IDFase)"
    TextoFiltro = "and P.fase = " & txtFase & " and F.Versao = '" & txtVersao & "'"
Else
    INNERJOINTEXTO = "Plano P"
    TextoFiltro = ""
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select P.IDplano, PD.* from " & INNERJOINTEXTO & " INNER JOIN planodimensao PD ON PD.IDPlano = P.IDPlano where P.desenho = '" & txtdesenho & "' " & TextoFiltro & " order by PD.indice", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from medicaodimensao", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!IDPlano = txtPm.Text
        TBGravar!Vista = IIf(IsNull(TBAbrir!Vista), "", TBAbrir!Vista)
        TBGravar!indice = IIf(IsNull(TBAbrir!indice), "", TBAbrir!indice)
        TBGravar!Tipo = IIf(IsNull(TBAbrir!Tipo), "", TBAbrir!Tipo)
        TBGravar!dimdesejada = TBAbrir!dimdesejada
        TBGravar!TolSup = TBAbrir!TolSup
        TBGravar!TolInf = TBAbrir!TolInf
        TBGravar!Dim_superior = TBAbrir!Dim_superior
        TBGravar!Dim_inferior = TBAbrir!Dim_inferior
        TBGravar!Responsavel = pubUsuario
        TBGravar!Freq = IIf(IsNull(TBAbrir!Freq), "", TBAbrir!Freq)
        
        TBGravar.Update
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from Planodimensao_instrumentos where ID_dimensao = " & TBAbrir!idDimensao & " order by Familia", Conexao, adOpenKeyset, adLockOptimistic
        Do While TBFamilia.EOF = False
            Set TBExecucao = CreateObject("adodb.recordset")
            TBExecucao.Open "select * from Medicaodimensao_Familia", Conexao, adOpenKeyset, adLockOptimistic
            TBExecucao.AddNew
            TBExecucao!id_dimensao = TBGravar!IDMedicao
            TBExecucao!Familia = TBFamilia!Familia
            TBExecucao.Update
            TBExecucao.Close
            TBFamilia.MoveNext
        Loop
        TBFamilia.Close
        TBGravar.Close
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

ProcCarregaLista
Gravar_dimensao = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBAbrir!Fase) = False And TBAbrir!Fase <> "" And TBAbrir!Fase <> "0" Then TextoFase = " - Fase : " & TBAbrir!Fase & ")" Else TextoFase = ")"
Caption = "Qualidade - Controle de medição - (Controle : " & TBAbrir!IDPlano & " - Cód. interno : " & TBAbrir!Desenho & TextoFase
txtPm.Text = TBAbrir!IDPlano
txtData.Text = Format(TBAbrir!Data, "dd/mm/yy")
txtinspetor.Text = IIf(IsNull(TBAbrir!Inspetor), "", TBAbrir!Inspetor)
If TBAbrir!laudofinal = True Then Checklaudoaprosim.Value = 1 Else Checklaudoapronao.Value = 1
If TBAbrir!reposicao = True Then Checklaudorepsim.Value = 1 Else Checklaudorepnao.Value = 1
If TBAbrir!restricao = True Then Checklaudorestsim.Value = 1 Else Checklaudorestnao.Value = 1
txtdesenho.Text = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
If IsNull(TBAbrir!Fase) = False And TBAbrir!Fase <> "" And TBAbrir!Fase <> "0" Then
    txtFase.Text = TBAbrir!Fase
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select Versao FROM Fases where IDfase = " & TBAbrir!IDFase, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        txtVersao = IIf(IsNull(TBFases!versao), "", TBFases!versao)
    End If
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select Grupo_op, Maquina FROM Ordemservico where IDProducao = " & TBAbrir!ID_inspecionado, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        txtGrupo = IIf(IsNull(TBFases!Grupo_op), "", TBFases!Grupo_op)
        Txt_posto_trab = IIf(IsNull(TBFases!maquina), "", TBFases!maquina)
    End If
    TBFases.Close
End If
If IsNull(TBAbrir!Nivel) = False Then cmbNivel = TBAbrir!Nivel
txtQtde_amt = IIf(IsNull(TBAbrir!qtde_amostra), "", Format(TBAbrir!qtde_amostra, "###,##0.0000"))
Txtpeca.Text = IIf(IsNull(TBAbrir!Peca), "", TBAbrir!Peca)
txtdescricao.Text = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
txtQuant_liber = IIf(IsNull(TBAbrir!Quant_liberada), "0,0000", Format(TBAbrir!Quant_liberada, "###,##0.0000"))
txtobservacaolaudo = IIf(IsNull(TBAbrir!Observacao), "", TBAbrir!Observacao)
txtDtValidacao.Text = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
txtRespValidacao.Text = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)

Txt_ID_RNC = IIf(IsNull(TBAbrir!ID_RNC), 0, TBAbrir!ID_RNC)
Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select ID_texto, Seq FROM CQ_RNC where ID = " & Txt_ID_RNC, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    txtRNC = IIf(IsNull(TBCompras_Pedido!Seq), TBCompras_Pedido!id_texto, TBCompras_Pedido!id_texto & "/" & IIf(TBCompras_Pedido!Seq < 10, "0" & TBCompras_Pedido!Seq, TBCompras_Pedido!Seq))
End If
TBCompras_Pedido.Close

ProcVerificaQtdeLiberar
Frame1.Enabled = True
cmbcodref.Enabled = True
cmbNivel.Enabled = True
Novo_Controle = False
Gravar_dimensao = False
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovaMedicao()
On Error GoTo tratar_erro

If Inspecaorecebimento_AnexarPlano = False Then
    With frmPlanomedicao_Novo
        Set TBplano = CreateObject("adodb.recordset")
        TBplano.Open "Select * from plano where desenho = '" & .ListView1.SelectedItem.ListSubItems(8) & "' and IDfase = " & .ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBplano.EOF = False Then
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Medicao", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Desenho = IIf(IsNull(TBplano!Desenho), "", TBplano!Desenho)
            TBGravar!Peca = .ListView1.SelectedItem.ListSubItems(1)
            If IsNull(TBplano!Descricao) = False Then TBGravar!Descricao = TBplano!Descricao Else TBGravar!Descricao = .ListView1.SelectedItem.ListSubItems(9)
            TBGravar!Inspetor = pubUsuario
            TBGravar!Data = Date
            TBGravar!Fase = IIf(IsNull(TBplano!Fase), "", TBplano!Fase)
            TBGravar!IDFase = IIf(IsNull(TBplano!IDFase), "", TBplano!IDFase)
            If IsNull(TBplano!Nivel) = False And TBplano!Nivel <> "" Then TBGravar!Nivel = TBplano!Nivel
            TBGravar!ID_inspecionado = .ListView1.SelectedItem.ListSubItems(2)
            TBGravar!Tipo_inspecao = 2
            TBGravar.Update
            txtPm = TBGravar!IDPlano
            TBGravar.Close
            txtGrupo.Text = .ListView1.SelectedItem.ListSubItems(7)
            Txt_posto_trab = .ListView1.SelectedItem.ListSubItems(5)
        Else
            USMsgBox ("Não foi encontrado nehum plano de inspeção cadastrado para este produto."), vbExclamation, "CAPRIND v5.0"
            Permitido = False
            TBplano.Close
            Exit Sub
        End If
        TBplano.Close
    End With
Else
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Medicao", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    TBGravar!Desenho = txtdesenho
    TBGravar!Peca = Txtpeca
    TBGravar!Descricao = txtdescricao
    TBGravar!Inspetor = pubUsuario
    TBGravar!Data = Date
    TBGravar!Fase = 0
    TBGravar!Quant_liberada = txtQuant_liber
    TBGravar!qtde_amostra = txtQtde_amt
    TBGravar!Nivel = cmbNivel
    With frmCompras_recebimento.ListProdReceb
        TBGravar!IDlista = .SelectedItem.ListSubItems(5)
        TBGravar!ID_inspecionado = .SelectedItem
    End With
    TBGravar!Tipo_inspecao = 1
    TBGravar.Update
    txtPm = TBGravar!IDPlano
    TBGravar.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_doc()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Proclimpacampos_doc
Novo_Controle1 = True
Frame14.Enabled = True
cmdImportar_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Proclimparplano()
On Error GoTo tratar_erro
  
txtPm.Text = ""
txtdesenho.Text = ""
txtFase.Text = ""
txtGrupo.Text = ""
Txt_posto_trab = ""
txtdescricao.Text = ""
txtinspetor = pubUsuario
Txtpeca.Text = ""
txtData.Text = Format(Date, "dd/mm/yy")
Checklaudoaprosim.Value = 0
Checklaudoapronao.Value = 0
Checklaudorestsim.Value = 0
Checklaudorestnao.Value = 0
Checklaudorepsim.Value = 0
Checklaudorepnao.Value = 0
txtobservacaolaudo.Text = ""
Txt_qtde_lote = ""
Txt_qtde_liberada = ""
txtQuant_liber = "0,0000"
txtQTD.Text = ""
txtQtde_amt = ""
cmbNivel.ListIndex = -1
Txt_ID_RNC = 0
txtRNC = ""
txtRespValidacao = ""
txtDtValidacao = ""
CodigoLista = 0
Caption = "Qualidade - Controle de medição"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Proclimparmedida()
On Error GoTo tratar_erro

txtNumero.Text = ""
txttipo.Text = ""
txtdesejada.Text = ""
txtencontrada = ""
txttolsup.Text = ""
txttolinf.Text = ""
Txtfrequencia.Text = ""
txtResponsavel = pubUsuario
Checkdimaprosim.Value = 0
Checkdimapronao.Value = 0
Checkdimrestsim.Value = 0
Checkdimrestnao.Value = 0
txtobservacaodim.Text = ""
txtMin_enc = ""
CodigoLista1 = 0
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proclimpacampos_doc()
On Error GoTo tratar_erro

txtID_doc = 0
txtData_doc = Format(Date, "dd/mm/yy")
txtResponsavel_doc = pubUsuario
Txt_cod_peca = ""
txt_Caminho = ""
Txt_obs_doc = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Cmb_opcao_lista.Text <> "Excluir" Then
    MsgBox ("Selecione a opção excluir em operação da lista!"), vbInformation + vbOKOnly
    Exit Sub
End If

Permitido = False
With ListaControle
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) controle(s) de medição?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            ProcExcluir1 .ListItems(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) controle(s) de medição antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Controle(s) de medição excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimparplano
    ProcCarregaListaControle (1)
    Frame1.Enabled = False
    Lista.ListItems.Clear
    cmbNivel.Enabled = False
    Novo_Controle = False
    Gravar_dimensao = False
    ProcLimparTudo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir1(IDcontrole As Long)
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Medicao where IdPlano = " & IDcontrole, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    If TBFIltro!Tipo_inspecao = 2 Then
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select * FROM CadMaquinas INNER JOIN ordemservico ON CadMaquinas.Maquina = ordemservico.Maquina where ordemservico.Idproducao = " & TBFIltro!ID_inspecionado & " and CadMaquinas.Insp_final = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from estoque_movimentacao where lote = '" & TBFIltro!Peca & "' and Ordem = " & TBFIltro!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select * from estoque_controle where IDEstoque = " & TBAbrir!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    ProcAlterarEntradaEstoqueExcluir TBFIltro!IDPlano, TBFIltro!Peca, TBFIltro!Fase, TBFIltro!laudofinal, TBFIltro!Quant_liberada
                    If qtdeliberada = 0 Then TBEstoque.Delete
                End If
                TBEstoque.Close
                Conexao.Execute "DELETE from estoque_movimentacao where lote = '" & TBFIltro!Peca & "' and Ordem = " & TBFIltro!IDPlano
                ProcGravarQtdeEstoque
                '==================================
                Modulo = "Qualidade/Controle de medição"
                Evento = "Excluir entrada no estoque"
                ID_documento = IDcontrole
                Documento = "Nº de rastreabilidade: " & TBFIltro!Peca & " - Cód. interno: " & TBFIltro!Desenho & " - Fase: " & TBFIltro!Fase
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
        End If
        TBFases.Close
        
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from producao where Idplano = " & TBFIltro!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            Conexao.Execute "DELETE from Producaomaterial where Ordem = " & TBOrdem!Ordem
            Conexao.Execute "DELETE from Producao where Ordem = " & TBOrdem!Ordem
            Conexao.Execute "DELETE from Ordemservico_maq_utilizadas WHERE Ordem = " & TBOrdem!Ordem
            Conexao.Execute "DELETE from OSH from Ordemservico_HoraUtilizadaporDia OSH INNER JOIN Ordemservico OS ON OS.IDProducao = OSH.OS Where OS.Ordem = " & TBOrdem!Ordem
            Conexao.Execute "DELETE from OrdemServico where Ordem = " & TBOrdem!Ordem
        End If
        TBOrdem.Close
        Conexao.Execute "DELETE from OrdemServico where ID_controle = " & TBFIltro!IDPlano & " and retrabalho = 'True'"
    End If
    
    ProcAtualizaStatusDimInsp 0
    
    If IsNull(TBFIltro!ID_RNC) = False And TBFIltro!IDPlano <> "" Then Conexao.Execute "DELETE from CQ_RNC where ID = " & TBFIltro!ID_RNC
    Conexao.Execute "DELETE from MF from Medicaodimensao_Familia MF INNER JOIN Medicaodimensao M ON M.idmedicao = MF.ID_dimensao where M.IdPlano = " & TBFIltro!IDPlano
    Conexao.Execute "DELETE from MI from Medicaodimensao_instrumentos MI INNER JOIN Medicaodimensao M ON M.idmedicao = MI.Idmedicao where M.IdPlano = " & TBFIltro!IDPlano
    Conexao.Execute "DELETE from MP from Medicaodimensao_peca MP INNER JOIN Medicaodimensao M ON M.idmedicao = MP.IDdimensao where M.IdPlano = " & TBFIltro!IDPlano
    Conexao.Execute "DELETE from MEDICAOdimensao where idplano = " & TBFIltro!IDPlano
    Conexao.Execute "DELETE from MEDICAO WHERE idplano = " & TBFIltro!IDPlano
    
    '==================================
    Modulo = "Qualidade/Controle de medição"
    Evento = "Excluir"
    ID_documento = IDcontrole
    Documento = "Nº de rastreabilidade: " & TBFIltro!Peca & " - Cód. interno: " & TBFIltro!Desenho & " - Fase: " & TBFIltro!Fase
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirDim()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) deimesão(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from Medicaodimensao_Familia where ID_dimensao = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Medicaodimensao_instrumentos where idmedicao = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Medicaodimensao_peca where IDdimensao = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Medicaodimensao where idmedicao = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Controle de medição"
            Evento = "Excluir medição"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
            Documento1 = "Tipo da dimensão: " & .ListItems(InitFor).ListSubItems(2) & " - Dimensão indicada: " & .ListItems(InitFor).ListSubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) dimensão(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Dimensão(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimparmedida
    ProcCarregaLista
    Frame2.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_doc()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) documento(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Medicao_documentos where Id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Controle de medição"
            Evento = "Excluir documento"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº de rastreabilidade: " & Txtpeca & " - Cód. interno: " & txtdesenho & " - Fase: " & txtFase
            If .ListItems(InitFor).ListSubItems(1) <> "" Then Documento1 = "Cód. da peça: " & .ListItems(InitFor).ListSubItems(1) & " - Caminho: " & .ListItems(InitFor).ListSubItems(2) Else Documento1 = "Caminho: " & .ListItems(InitFor).ListSubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) documentos(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Documentos(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimpacampos_doc
    ProcCarregaLista_Doc
    Frame14.Enabled = False
    Novo_Controle1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlterarEntradaEstoqueExcluir(IDPlano As Long, Ordem As Long, OS As Long, Aprovado As Boolean, Qtde_encontrada As Double)
On Error GoTo tratar_erro

qtdeliberar = 0
qtdeliberada = 0
quantestoque = 0
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from producao where Ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select QTOK from ordemservico where Idproducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!QTOK) = False And TBAbrir!QTOK >= TBproducao!Quant Then qtdeliberar = TBAbrir!QTOK Else qtdeliberar = TBproducao!Quant
    End If
    TBAbrir.Close
End If
TBproducao.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Qtdestoque) as qtdeliberada from medicao where peca = '" & Ordem & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
End If
qtdeliberada = qtdeliberada - Qtde_encontrada
If Aprovado = False Or Qtde_encontrada = 0 Then Conexao.Execute "Update medicao Set QtdEstoque = 0 where IDPlano = " & IDPlano

If qtdeliberada < qtdeliberar Then TBEstoque!status = "ENTRADA_ORDEM_PARCIAL" Else TBEstoque!status = "ENTRADA_ORDEM"
TBEstoque.Update

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaControle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "P. controle" Then
    With ListaControle
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from Medicao where IdPlano = " & .ListItems(InitFor) & " and Tipo_inspecao = 2", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    ProcVerificaRegistroUtilizadoSemMsg "producao", "Idplano = " & TBFIltro!IDPlano & " and status <> 'Aberta'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select * from producao where Ordem = " & TBFIltro!Peca & " and Ap_backup = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then NomeTabelaAp = "ProducaoFases_Backup" Else NomeTabelaAp = "ProducaoFases"
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select * FROM Ordemservico INNER JOIN " & NomeTabelaAp & " ON Ordemservico.idproducao = " & NomeTabelaAp & ".OS where Ordemservico.ID_controle = " & TBFIltro!IDPlano & " and Ordemservico.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then
                        .ListItems.Item(InitFor).Checked = False
                        TBOrdem.Close
                        GoTo Proximo
                    End If
                    TBOrdem.Close
                End If
                TBFIltro.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaControle, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaControle_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaControle
    For InitFor = 1 To .ListItems.Count
        
        If .ListItems.Item(InitFor).Checked = True Then
        
            'Se opção da lista = Excluir
            If Cmb_opcao_lista.Text = "Excluir" Then
                'Verifica se o registro selecionado esta validado
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from Medicao where IdPlano = " & .ListItems(InitFor) & " and respValidacao <> 'Null'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    MsgBox ("Não é permitido excluir plano de medição ja validado!"), vbInformation + vbOKOnly
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                TBFIltro.Close
                
                'Verifica se o registro esta sendo utilizado em outro modulo
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from Medicao where IdPlano = " & .ListItems(InitFor) & " and Tipo_inspecao = 2", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    Mensagem = "Não é permitido excluir este controle de medição, pois o mesmo está sendo utilizado no módulo"
                    ProcVerificaRegistroUtilizado "producao", "Idplano = " & TBFIltro!IDPlano & " and status <> 'Aberta'", "PCP/Gerenciamento de ordem"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select * from producao where Ordem = " & TBFIltro!Peca & " and Ap_backup = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then NomeTabelaAp = "ProducaoFases_Backup" Else NomeTabelaAp = "ProducaoFases"
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select * FROM Ordemservico INNER JOIN " & NomeTabelaAp & " ON Ordemservico.idproducao = " & NomeTabelaAp & ".OS where Ordemservico.ID_controle = " & TBFIltro!IDPlano & " and Ordemservico.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then
                        USMsgBox ("Não é permitido excluir este controle de medição, pois o mesmo está sendo utilizado no módulo PCP/Gerenciamento de ordem"), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        TBOrdem.Close
                        Exit Sub
                    End If
                    TBOrdem.Close
                End If
                TBFIltro.Close
            End If
            
     
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaControle_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaControle.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from medicao where idplano = " & ListaControle.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Proclimparplano
    ProcCarregaDados
    CodigoLista = ListaControle.SelectedItem.index
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtPm = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        cmbcodref.Visible = True
        cmbNivel.Visible = True
        ListaControle.Visible = True
        Lista.Visible = False
        Lista_doc.Visible = False
        If ListaControle.Visible = True Then ListaControle.SetFocus
    Case 1:
        ListaControle.Visible = False
        Lista.Visible = True
        Lista_doc.Visible = Fase
        cmbcodref.Visible = False
        cmbNivel.Visible = False
'        With USToolBar2
'            Set TBCompras = CreateObject("adodb.recordset")
'            TBCompras.Open "Select * from Compras_pedido where Pedido = '" & txtPeca & "'", Conexao, adOpenKeyset, adLockOptimistic
'            If TBCompras.EOF = False Then .ButtonState(7) = 5 Else .ButtonState(7) = 0
'            TBCompras.Close
'            .Refresh
'        End With
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcCarregaLista
        Lista.SetFocus
    Case 2:
        ListaControle.Visible = False
        Lista.Visible = False
        Lista_doc.Visible = True
        cmbcodref.Visible = False
        cmbNivel.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcCarregaLista_Doc
        Lista_doc.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Controle = True Then
    USMsgBox ("Salve o controle de medição antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procValidar()
On Error GoTo tratar_erro

If Cmb_opcao_lista.Text <> "Validação" Then
    MsgBox ("Selecione a opção validar em operação da lista!"), vbInformation + vbOKOnly
    Exit Sub
End If
    
Permitido = False
With ListaControle
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente Validar/Desvalidar este(s) controle(s) de medição?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from Medicao where IdPlano = " & .ListItems(InitFor) & " ", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    
                    'Valida ou desvalida o resgisto
                    If IsNull(TBFIltro!RespValidacao) = True Then
                        TBFIltro!DtValidacao = Now
                        TBFIltro!RespValidacao = pubUsuario
                    Else
                        TBFIltro!DtValidacao = Null
                        TBFIltro!RespValidacao = Null
                    End If
                    TBFIltro.Update
                End If
                TBFIltro.Close
            
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) controle(s) de medição para realizar a operação de validação."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Operação de validação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimparplano
    ProcCarregaListaControle (1)
    Frame1.Enabled = False
    Lista.ListItems.Clear
    cmbNivel.Enabled = False
    Novo_Controle = False
    Gravar_dimensao = False
    ProcLimparTudo
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

ProcCarregaComboCodRef cmbcodref, "P.desenho = '" & txtdesenho.Text & "'", 0, "", False, True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtencontrada_Change()
On Error GoTo tratar_erro

If txtencontrada.Text <> "" Then
    VerifNumero = txtencontrada.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtencontrada.Text = ""
        txtencontrada.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtencontrada_Click()
On Error GoTo tratar_erro

If txtencontrada = "0,0000" Then txtencontrada = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtencontrada_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtencontrada

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtencontrada_LostFocus()
On Error GoTo tratar_erro

txtencontrada.Text = Format(txtencontrada.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMin_enc_Change()
On Error GoTo tratar_erro

If txtMin_enc.Text <> "" Then
    VerifNumero = txtMin_enc.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtMin_enc.Text = ""
        txtMin_enc.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMin_enc_Click()
On Error GoTo tratar_erro

If txtMin_enc = "0,0000" Then txtMin_enc = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMin_enc_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtMin_enc

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMin_enc_LostFocus()
On Error GoTo tratar_erro

txtMin_enc.Text = Format(txtMin_enc.Text, "###,##0.0000")
    
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

Private Sub txtQtde_amt_Change()
On Error GoTo tratar_erro

If txtQtde_amt.Text <> "" Then
    VerifNumero = txtQtde_amt.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_amt.Text = ""
        txtQtde_amt.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_amt_LostFocus()
On Error GoTo tratar_erro

txtQtde_amt.Text = Format(txtQtde_amt.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuant_liber_Change()
On Error GoTo tratar_erro

If txtQuant_liber.Text <> "" Then
    VerifNumero = txtQuant_liber.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuant_liber.Text = ""
        txtQuant_liber.SetFocus
        Exit Sub
    End If
End If
txtQtde_amt = FunCalculaAmostragem(cmbNivel, IIf(txtQuant_liber = "", 0, txtQuant_liber))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuant_liber_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQuant_liber

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuant_liber_LostFocus()
On Error GoTo tratar_erro

txtQuant_liber.Text = Format(txtQuant_liber.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txttolinf_Change()
On Error GoTo tratar_erro

If txttolinf.Text <> "" Then
    VerifNumero = txttolinf.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txttolinf.Text = ""
        txttolinf.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txttolinf_LostFocus()
On Error GoTo tratar_erro

txttolinf.Text = Format(txttolinf.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txttolsup_Change()
On Error GoTo tratar_erro

If txttolsup.Text <> "" Then
    VerifNumero = txttolsup.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txttolsup.Text = ""
        txttolsup.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txttolsup_LostFocus()
On Error GoTo tratar_erro

txttolsup.Text = Format(txttolsup.Text, "###,##0.0000")
    
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
    Case 3: ProcGravar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcLaudo
    'Case 9: ProcRepRet
    Case 9: ProcRNC
    Case 10: procValidar
    'Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvarDim
    Case 2: ProcExcluirDim
    Case 3: ProcImprimir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    Case 6: ProcInstrumentos
    Case 7: ProcSalvarRestricao
    Case 8: ProcPeca
    Case 9: ProcAtualizar
    'Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procNovo_doc
    Case 2: procSalvar_doc
    Case 3: procExcluir_doc
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
