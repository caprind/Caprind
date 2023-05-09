VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlanoinspecao 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Qualidade - Plano de inspeção"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlanoinspecao.frx":0000
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
      TabIndex        =   71
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
   Begin MSComctlLib.ListView Lista1 
      Height          =   6630
      Left            =   80
      TabIndex        =   33
      Top             =   3075
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11695
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
      NumItems        =   11
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
         Text            =   "Carac. núm."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Tipo da dimensão"
         Object.Width           =   5124
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Dimensão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Dim. superior"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Dim. inferior"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Tol. superior"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Tol. inferior"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Freq. de medição"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Vista"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView Lista2 
      Height          =   7215
      Left            =   80
      TabIndex        =   35
      Top             =   2490
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12726
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   25585
      EndProperty
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
      Left            =   2860
      Locked          =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Código de referência."
      Top             =   2310
      Width           =   2400
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6315
      Left            =   80
      TabIndex        =   12
      Top             =   2760
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11139
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Plano"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   1058
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
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   7929
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Fase"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Versão"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Grupo/op."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Validado"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Frame framemed 
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
      Height          =   1455
      Left            =   75
      TabIndex        =   52
      Top             =   1605
      Visible         =   0   'False
      Width           =   15195
      Begin VB.CheckBox Chk_relatorio_pcp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Relatório do PCP"
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
         Left            =   13500
         TabIndex        =   25
         Top             =   450
         Width           =   1485
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
         Left            =   1180
         MaxLength       =   50
         TabIndex        =   22
         ToolTipText     =   "Característica número."
         Top             =   390
         Width           =   2835
      End
      Begin VB.ComboBox cmbtipomed 
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
         ItemData        =   "frmPlanoinspecao.frx":0442
         Left            =   4035
         List            =   "frmPlanoinspecao.frx":0444
         Sorted          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Tipo da dimensão."
         Top             =   390
         Width           =   9045
      End
      Begin VB.TextBox txtIndice 
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
         MaxLength       =   5
         TabIndex        =   21
         ToolTipText     =   "Índice."
         Top             =   390
         Width           =   990
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
         Left            =   3265
         MaxLength       =   20
         TabIndex        =   28
         ToolTipText     =   "Dimenção inferior."
         Top             =   990
         Width           =   1530
      End
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
         Left            =   1720
         MaxLength       =   20
         TabIndex        =   27
         ToolTipText     =   "Dimensão superior."
         Top             =   990
         Width           =   1530
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
         Left            =   4810
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Tolerância superior."
         Top             =   990
         Width           =   1530
      End
      Begin VB.TextBox txtFrequencia 
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
         Left            =   7905
         MaxLength       =   20
         TabIndex        =   31
         ToolTipText     =   "Frequencia de medição."
         Top             =   990
         Width           =   5505
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
         Left            =   6355
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Tolerância inferior."
         Top             =   990
         Width           =   1530
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
         Left            =   180
         MaxLength       =   20
         TabIndex        =   26
         ToolTipText     =   "Dimensão."
         Top             =   990
         Width           =   1530
      End
      Begin VB.CommandButton cmdnovotipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13080
         Picture         =   "frmPlanoinspecao.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cadatrar/localizar tipo da dimensão."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_vista 
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
         Left            =   13425
         MaxLength       =   3
         TabIndex        =   32
         ToolTipText     =   "Vista."
         Top             =   990
         Width           =   1560
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Característica número"
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
         Left            =   1810
         TabIndex        =   62
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vista"
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
         Left            =   14033
         TabIndex        =   61
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo da dimensão*"
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
         Left            =   7890
         TabIndex        =   60
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         Left            =   2010
         TabIndex        =   59
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Índice"
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
         Left            =   458
         TabIndex        =   58
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frequencia de medição"
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
         Left            =   9832
         TabIndex        =   57
         Top             =   780
         Width           =   1650
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tol. superior*"
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
         Left            =   5080
         TabIndex        =   56
         Top             =   780
         Width           =   990
      End
      Begin VB.Label Label10 
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
         Left            =   3600
         TabIndex        =   55
         Top             =   780
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensão*"
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
         Left            =   555
         TabIndex        =   54
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tol. inferior*"
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
         Left            =   6663
         TabIndex        =   53
         Top             =   780
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
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
      TabCaption(0)   =   "Plano de inspeção"
      TabPicture(0)   =   "frmPlanoinspecao.frx":0548
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Dimensões"
      TabPicture(1)   =   "frmPlanoinspecao.frx":0564
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   10065
         Left            =   -74985
         TabIndex        =   37
         Top             =   300
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   17754
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Dimensões"
         TabPicture(0)   =   "frmPlanoinspecao.frx":0580
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Txt_ID"
         Tab(0).Control(1)=   "USToolBar2"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Famílias de instrumentos"
         TabPicture(1)   =   "frmPlanoinspecao.frx":059C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "USToolBar3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Txt_ID1"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.TextBox Txt_ID1 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   480
            MaxLength       =   20
            MouseIcon       =   "frmPlanoinspecao.frx":05B8
            MousePointer    =   99  'Custom
            TabIndex        =   39
            Text            =   "0"
            ToolTipText     =   "Dimensão."
            Top             =   4620
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox Txt_ID 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   -74520
            MaxLength       =   20
            MouseIcon       =   "frmPlanoinspecao.frx":08C2
            MousePointer    =   99  'Custom
            TabIndex        =   38
            Text            =   "0"
            ToolTipText     =   "Dimensão."
            Top             =   3810
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Frame Frame5 
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
            Height          =   855
            Left            =   75
            TabIndex        =   40
            Top             =   1320
            Width           =   15195
            Begin VB.ComboBox Cmb_familia 
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
               Left            =   180
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   34
               ToolTipText     =   "Família do instrumento."
               Top             =   390
               Width           =   14835
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Família do instrumento*"
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
               Left            =   6750
               TabIndex        =   41
               Top             =   180
               Width           =   1695
            End
         End
         Begin DrawSuite2022.USToolBar USToolBar2 
            Height          =   975
            Left            =   -74925
            TabIndex        =   72
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
               Left            =   12330
               Top             =   120
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmPlanoinspecao.frx":0BCC
               Count           =   1
            End
         End
         Begin DrawSuite2022.USToolBar USToolBar3 
            Height          =   975
            Left            =   75
            TabIndex        =   73
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
            Begin DrawSuite2022.USImageList USImageList3 
               Left            =   12330
               Top             =   120
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmPlanoinspecao.frx":5FB0
               Count           =   1
            End
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
         Height          =   1455
         Left            =   75
         TabIndex        =   42
         Top             =   1290
         Width           =   15195
         Begin VB.TextBox txtRespValidacao_prod 
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
            Left            =   12420
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação do plano da produto."
            Top             =   375
            Width           =   2595
         End
         Begin VB.TextBox txtDtValidacao_prod 
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
            Left            =   10290
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação do plano do produto."
            Top             =   375
            Width           =   2115
         End
         Begin VB.ComboBox cmbNivel 
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
            ItemData        =   "frmPlanoinspecao.frx":B251
            Left            =   13455
            List            =   "frmPlanoinspecao.frx":B26D
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   78
            ToolTipText     =   "Nível."
            Top             =   1020
            Width           =   1560
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
            Left            =   10950
            Locked          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            ToolTipText     =   "Fase do processo."
            Top             =   1020
            Width           =   900
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
            Left            =   12540
            Locked          =   -1  'True
            TabIndex        =   76
            TabStop         =   0   'False
            ToolTipText     =   "Grupo/operação do processo."
            Top             =   1020
            Width           =   900
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
            Left            =   11865
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            ToolTipText     =   "Versão da fase."
            Top             =   1020
            Width           =   660
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
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação do plano da fase."
            Top             =   375
            Width           =   2115
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
            Left            =   7650
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação do plano da fase."
            Top             =   360
            Width           =   2625
         End
         Begin VB.TextBox txtIDFase 
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
            Left            =   10020
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   74
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "ID da fase."
            Top             =   1020
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.TextBox txtdescricao 
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
            Left            =   5190
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   1020
            Width           =   5745
         End
         Begin VB.TextBox txtRev_item 
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
            Left            =   2235
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   1020
            Width           =   540
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
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   375
            Width           =   540
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
            Left            =   210
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   1020
            Width           =   2010
         End
         Begin VB.TextBox txtinspetor 
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
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   2565
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
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   840
         End
         Begin VB.TextBox txtPI 
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
            Left            =   210
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Número do plano de inspeção."
            Top             =   375
            Width           =   1305
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Resp. pela validação (prod.)"
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
            Left            =   12697
            TabIndex        =   83
            Top             =   180
            Width           =   2040
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora valid. (prod.)"
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
            Left            =   10485
            TabIndex        =   82
            Top             =   180
            Width           =   1725
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação (fase)"
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
            Left            =   5610
            TabIndex        =   80
            Top             =   180
            Width           =   1935
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Resp. pela validação (fase)"
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
            Left            =   7980
            TabIndex        =   79
            Top             =   180
            Width           =   1965
         End
         Begin VB.Label Label23 
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
            Left            =   11948
            TabIndex        =   65
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label21 
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
            Left            =   2333
            TabIndex        =   64
            Top             =   810
            Width           =   345
         End
         Begin VB.Label Label20 
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
            Left            =   1635
            TabIndex        =   63
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno"
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
            Left            =   600
            TabIndex        =   51
            Top             =   810
            Width           =   1230
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
            Left            =   7717
            TabIndex        =   50
            Top             =   810
            Width           =   690
         End
         Begin VB.Label Label7 
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
            Left            =   3765
            TabIndex        =   49
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data "
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
            Left            =   2310
            TabIndex        =   48
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Plano de insp."
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
            Left            =   262
            TabIndex        =   47
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label12 
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
            Left            =   11228
            TabIndex        =   46
            Top             =   810
            Width           =   345
         End
         Begin VB.Label Label15 
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
            Index           =   1
            Left            =   12608
            TabIndex        =   45
            Top             =   810
            Width           =   765
         End
         Begin VB.Label Label5 
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
            Left            =   3210
            TabIndex        =   44
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nível*"
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
            Left            =   14018
            TabIndex        =   43
            Top             =   810
            Width           =   435
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   66
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
            ItemData        =   "frmPlanoinspecao.frx":B290
            Left            =   6960
            List            =   "frmPlanoinspecao.frx":B29A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   180
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
            TabIndex        =   13
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
            TabIndex        =   15
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   19
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanoinspecao.frx":B2B2
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
            TabIndex        =   18
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanoinspecao.frx":EA5D
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
            TabIndex        =   16
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
            TabIndex        =   17
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanoinspecao.frx":12568
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
            TabIndex        =   20
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlanoinspecao.frx":16659
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
         Begin VB.Label Label15 
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
            Index           =   5
            Left            =   3360
            TabIndex        =   84
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
            TabIndex        =   81
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label24 
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
            Top             =   240
            Width           =   1095
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   70
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   15
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
         ButtonCaption8  =   "Copiar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Copiar (F7)"
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
         ButtonWidth8    =   44
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Revisar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Revisar (F8)"
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
         ButtonLeft9     =   399
         ButtonTop9      =   2
         ButtonWidth9    =   51
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Validação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Validar/Cancelar validação (F9)"
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
         ButtonLeft10    =   452
         ButtonTop10     =   2
         ButtonWidth10   =   53
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Atualizar"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonKey11     =   "10"
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
         ButtonLeft11    =   507
         ButtonTop11     =   2
         ButtonWidth11   =   50
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonAlignment12=   2
         ButtonType12    =   1
         ButtonStyle12   =   -1
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   -1
         ButtonLeft12    =   559
         ButtonTop12     =   4
         ButtonWidth12   =   2
         ButtonHeight12  =   54
         ButtonCaption13 =   "Ajuda"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Ajuda (F1)"
         ButtonKey13     =   "12"
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
         ButtonLeft13    =   563
         ButtonTop13     =   2
         ButtonWidth13   =   41
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonCaption14 =   "Sair"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Sair (Esc)"
         ButtonKey14     =   "13"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft14    =   606
         ButtonTop14     =   2
         ButtonWidth14   =   30
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonKey15     =   "14"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState15   =   5
         ButtonLeft15    =   638
         ButtonTop15     =   2
         ButtonWidth15   =   24
         ButtonHeight15  =   24
         ButtonUseMaskColor15=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12330
            Top             =   120
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmPlanoinspecao.frx":19EE6
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frmPlanoinspecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dimensao        As Double 'OK
Dim DimSup          As Double 'OK
Dim DimInf          As Double 'OK
Public Novo_Plano   As Boolean 'OK
Public Novo_Plano1  As Boolean 'OK
Public Novo_Plano2  As Boolean 'OK
Public StrSql_Plano_Localizar As String 'OK
Dim TBLISTA_Plano_Insp As ADODB.Recordset 'OK
Public Copiar As Boolean 'OK

Private Sub ProcExcluirFamilia()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) família(s) de instrumentos?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Planodimensao_instrumentos WHERE id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Plano de inspeção"
            Evento = "Excluir família de instrumento"
            ID_documento = .ListItems(InitFor)
            Documento = "Plano de inspeção: " & txtPI
            Documento1 = "Família do instrumento: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) famíla(s) de instrumentos antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Famíla(s) de instrumentos excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparCamposFamilia
    ProcCarregaListaFamilia
    Novo_Plano2 = False
    Frame5.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoFamilia()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Plano", "IDPlano = " & txtPI, "plano de inspeção", "família", "criar nova", True, True) = False Then Exit Sub
If txtIDFase <> "" And txtIDFase <> "0" Then
    If FunVerifiProcRevisado(txtIDFase, "criar nova família", True) = True Then Exit Sub
End If
ProcLimparCamposFamilia
ProcCarregaFamilia
Novo_Plano2 = True
Frame5.Enabled = True
Cmb_familia.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarFamilia()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame5.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_familia = "" Then
    NomeCampo = "a família do instrumento"
    ProcVerificaAcao
    Cmb_familia.SetFocus
    Exit Sub
End If
Set TBplanolaudo = CreateObject("adodb.recordset")
TBplanolaudo.Open "Select * from Planodimensao_instrumentos where ID = " & Txt_ID1, Conexao, adOpenKeyset, adLockOptimistic
If TBplanolaudo.EOF = True Then
    TBplanolaudo.AddNew
Else
    If FunVerificaRegistroValidado("Plano", "IDPlano = " & txtPI, "plano de inspeção", "família", "alterar a", True, True) = False Then Exit Sub
    If txtIDFase <> "" And txtIDFase <> "0" Then
        If FunVerifiProcRevisado(txtIDFase, "alterar a família", True) = True Then Exit Sub
    End If
End If
TBplanolaudo!id_dimensao = Txt_ID
TBplanolaudo!Familia = Cmb_familia
TBplanolaudo.Update
Txt_ID1 = TBplanolaudo!ID
ProcCarregaListaFamilia
If Novo_Plano2 = True Then
    USMsgBox ("Nova família do instrumento cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova família do instrumento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar família do instrumento"
    If CodigoLista2 <> 0 And Lista2.ListItems.Count <> 0 Then
        Lista2.SelectedItem = Lista2.ListItems(CodigoLista2)
        Lista2.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Plano de inspeção"
ID_documento = Txt_ID1
Documento = "Plano de inspeção: " & txtPI
Documento1 = "Família do instrumento: " & Cmb_familia
ProcGravaEvento
'==================================
Novo_Plano2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtPI = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from plano order by idplano", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idplano = " & txtPI)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtPI = TBLISTA!IDPlano
        Set TBplano = CreateObject("adodb.recordset")
        TBplano.Open "Select * from plano where idplano = " & txtPI, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        ProcLimparCamposDimensoes
        ProcLimparCamposFamilia
        ProcCarregaDados
        ProcCarregaListaDimensao
        ProcCarregaListaFamilia
    Else
        USMsgBox ("Fim dos cadastros de plano."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Plano1 = False
Novo_Plano2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_Plano_Localizar = "" Then Exit Sub
Set TBLISTA_Plano_Insp = CreateObject("adodb.recordset")
TBLISTA_Plano_Insp.Open StrSql_Plano_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Plano_Insp.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Plano_Insp.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Plano_Insp.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Plano_Insp.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Plano_Insp.RecordCount - IIf(Pagina > 1, (TBLISTA_Plano_Insp.PageSize * (Pagina - 1)), 0), TBLISTA_Plano_Insp.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Plano_Insp.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Plano_Insp!IDPlano
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Plano_Insp!Rev), "", TBLISTA_Plano_Insp!Rev)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Plano_Insp!Data), "", Format(TBLISTA_Plano_Insp!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Plano_Insp!Inspetor), "", Trim(TBLISTA_Plano_Insp!Inspetor))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Plano_Insp!Desenho), "", TBLISTA_Plano_Insp!Desenho)
        
        'Revisão do produto
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Revdesenho from projproduto where desenho = '" & TBLISTA_Plano_Insp!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then .Item(.Count).SubItems(5) = IIf(IsNull(TBItem!RevDesenho), "", TBItem!RevDesenho)
        
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Plano_Insp!Descricao), "", Trim(TBLISTA_Plano_Insp!Descricao))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Plano_Insp!Fase), "", TBLISTA_Plano_Insp!Fase)
        
        'Versão da fase
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select versao from fases where IDfase = " & IIf(IsNull(TBLISTA_Plano_Insp!IDFase), 0, TBLISTA_Plano_Insp!IDFase), Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then .Item(.Count).SubItems(8) = IIf(IsNull(TBItem!versao), "", TBItem!versao)
        TBItem.Close
        
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Plano_Insp!Grupo_op), "", TBLISTA_Plano_Insp!Grupo_op)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Plano_Insp!DtValidacao), "Não", "Sim")
    End With
    TBLISTA_Plano_Insp.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Plano_Insp.RecordCount
If TBLISTA_Plano_Insp.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Plano_Insp.PageCount
ElseIf TBLISTA_Plano_Insp.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Plano_Insp.PageCount & " de: " & TBLISTA_Plano_Insp.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Plano_Insp.AbsolutePage - 1 & " de: " & TBLISTA_Plano_Insp.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaDimensao()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from planodimensao where idplano = " & txtPI & " order by indice", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista1.ListItems
            .Add , , TBLISTA!idDimensao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!indice), "", TBLISTA!indice)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Numero), "", TBLISTA!Numero)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!dimdesejada), "", Format(TBLISTA!dimdesejada, "###,##0.0000"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Dim_superior), "", Format(TBLISTA!Dim_superior, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Dim_inferior), "", Format(TBLISTA!Dim_inferior, "###,##0.0000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!TolSup), "", Format(TBLISTA!TolSup, "###,##0.0000"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!TolInf), "", Format(TBLISTA!TolInf, "###,##0.0000"))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Freq), "", TBLISTA!Freq)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Vista), "", TBLISTA!Vista)
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

Private Sub ProcCarregaListaFamilia()
On Error GoTo tratar_erro

Lista2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Planodimensao_instrumentos where ID_dimensao = " & Txt_ID & " order by Familia", Conexao, adOpenKeyset, adLockOptimistic
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
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Familia), "", Trim(TBLISTA!Familia))
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

Private Sub ProcAnteriorDimensao()
On Error GoTo tratar_erro

If Txt_ID = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Planodimensao where idplano = " & txtPI & "  order by indice", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idDimensao = " & Txt_ID)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        Txt_ID = TBLISTA!idDimensao
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from Planodimensao where idDimensao = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimparCamposDimensoes
        ProcLimparCamposFamilia
        ProcCarregaDadosDimensao
        ProcCarregaListaFamilia
    Else
        USMsgBox ("Fim dos cadastros de dimensões."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If txtPI = "" Then
    NomeCampo = "o plano de inspeção"
    Acao = "copiar"
    ProcVerificaAcao
    Exit Sub
End If
Copiar = True
frmPlanoinspecao_Novo.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarDimensao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If cmbtipomed = "" Then
    NomeCampo = "o tipo da dimensão"
    ProcVerificaAcao
    cmbtipomed.SetFocus
    Exit Sub
End If
If txtdesejada = "" Then
    NomeCampo = "a dimensão"
    ProcVerificaAcao
    txtdesejada.SetFocus
    Exit Sub
End If
If txttolsup = "" Then
    NomeCampo = "a tolerância superior"
    ProcVerificaAcao
    txttolsup.SetFocus
    Exit Sub
End If
If txttolinf = "" Then
    NomeCampo = "a tolerância inferior"
    ProcVerificaAcao
    txttolinf.SetFocus
    Exit Sub
End If
Set TBplanolaudo = CreateObject("adodb.recordset")
TBplanolaudo.Open "Select * from Planodimensao where iddimensao = " & Txt_ID.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBplanolaudo.EOF = True Then
    TBplanolaudo.AddNew
Else
    If FunVerificaRegistroValidado("Plano", "IDPlano = " & txtPI, "plano de inspeção", "dimensões", "alterar as", True, True) = False Then Exit Sub
    If txtIDFase <> "" And txtIDFase <> "0" Then
        If FunVerifiProcRevisado(txtIDFase, "alterar as dimensões", True) = True Then Exit Sub
    End If
End If
TBplanolaudo!IDPlano = txtPI
TBplanolaudo!Tipo = Left(cmbtipomed.Text, 100)

If Chk_relatorio_pcp.Value = 1 Then
TBplanolaudo!PCP = True
Else
TBplanolaudo!PCP = False
End If

TBplanolaudo!dimdesejada = txtdesejada.Text
TBplanolaudo!TolSup = txttolsup.Text
TBplanolaudo!TolInf = txttolinf.Text
TBplanolaudo!Vista = Txt_vista
TBplanolaudo!indice = txtIndice
TBplanolaudo!Dim_superior = IIf(txtdim_sup = "", Null, txtdim_sup)
TBplanolaudo!Dim_inferior = IIf(txtDim_inf = "", Null, txtDim_inf)
TBplanolaudo!Numero = IIf(txtNumero = "", Null, txtNumero)
If Txtfrequencia.Text <> "" Then
    TBplanolaudo!Freq = Txtfrequencia
    TBplanolaudo!Cartacontrole = "*"
Else
    TBplanolaudo!Freq = ""
    TBplanolaudo!Cartacontrole = ""
End If
TBplanolaudo.Update
Txt_ID = TBplanolaudo!idDimensao
ProcCarregaListaDimensao
If Novo_Plano1 = True Then
    USMsgBox ("Novo plano de dimensão cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo plano de dimensão"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar plano de dimensão"
    If CodigoLista1 <> 0 And Lista1.ListItems.Count <> 0 Then
        Lista1.SelectedItem = Lista1.ListItems(CodigoLista1)
        Lista1.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Plano de inspeção"
ID_documento = Txt_ID
Documento = "Plano de inspeção: " & txtPI
Documento1 = "Tipo da dimensão: " & cmbtipomed & " - Dimensão: " & txtdesejada
ProcGravaEvento
'==================================
Novo_Plano1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoDimensao()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Plano", "IDPlano = " & txtPI, "plano de inspeção", "dimensões", "criar novas", True, True) = False Then Exit Sub
If txtIDFase <> "" And txtIDFase <> "0" Then
    If FunVerifiProcRevisado(txtIDFase, "criar novas dimensões", True) = True Then Exit Sub
End If
ProcLimparCamposDimensoes
ProcCarrega_Tipo
Novo_Plano1 = True
framemed.Enabled = True
txtIndice.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFamilia()
On Error GoTo tratar_erro

ProcCarregaComboFamilia Cmb_familia, "familia <> 'Null' and qualidade = 'True'", False

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
        .ButtonState(10) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(10) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdnovotipo_Click()
On Error GoTo tratar_erro

If cmbtipomed.Enabled = False Then Exit Sub
Qualidade_Plano = True
Faturamento = False
frmPlanoinspecao_Tipodimensao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Plano_Insp.AbsolutePage <> 2 Then
    If TBLISTA_Plano_Insp.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Plano_Insp.PageCount - 1)
    Else
        TBLISTA_Plano_Insp.AbsolutePage = TBLISTA_Plano_Insp.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Plano_Insp.AbsolutePage)
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
    TBLISTA_Plano_Insp.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Plano_Insp.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Plano_Insp.AbsolutePage = 1
ProcExibePagina (TBLISTA_Plano_Insp.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Plano_Insp.AbsolutePage <> -3 Then
    If TBLISTA_Plano_Insp.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Plano_Insp.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Plano_Insp.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Plano_Insp.AbsolutePage = TBLISTA_Plano_Insp.PageCount
ProcExibePagina (TBLISTA_Plano_Insp.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtPI = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from plano order by idplano", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.Find ("idplano = " & txtPI)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtPI = TBLISTA!IDPlano
        Set TBplano = CreateObject("adodb.recordset")
        TBplano.Open "Select * from plano where idplano = " & txtPI, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        ProcLimparCamposDimensoes
        ProcLimparCamposFamilia
        ProcCarregaDados
        ProcCarregaListaDimensao
        ProcCarregaListaFamilia
    Else
        USMsgBox ("Fim dos cadastros de plano."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Plano1 = False
Novo_Plano2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximoDimensao()
On Error GoTo tratar_erro

If Txt_ID = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Planodimensao where idplano = " & txtPI & " order by indice", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.Find ("idDimensao = " & Txt_ID)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        Txt_ID = TBLISTA!idDimensao
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from Planodimensao where idDimensao = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimparCamposDimensoes
        ProcLimparCamposFamilia
        ProcCarregaDadosDimensao
        ProcCarregaListaFamilia
    Else
        USMsgBox ("Fim dos cadastros de dimensões."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisar()
On Error GoTo tratar_erro

Acao = "revisar"
If txtPI = "" Then
    NomeCampo = "o plano de inspeção"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Plano = True Then
    USMsgBox ("Salve o plano de inspeção antes de revisar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmPlanoinspecao_revisao.Show

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
            Case vbKeyF7: ProcCopiar
            Case vbKeyF8: ProcRevisar
            Case vbKeyF9: If Cmb_opcao_lista = "Validação" Then frmPlanoinspecao_validacao.Show 1
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case SSTab2.Tab
           Case 0:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovoDimensao
                    Case vbKeyF3: ProcSalvarDimensao
                    Case vbKeyF4: ProcExcluirDimensao
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovoFamilia
                    Case vbKeyF3: ProcSalvarFamilia
                    Case vbKeyF4: ProcExcluirFamilia
                    Case vbKeyF5: ProcImprimir
                    'Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
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

ProcCarregaToolBar1 Me, 15200, 14, True
ProcCarregaToolBar2 Me, 15200, 10, True
ProcCarregaToolBar3 Me, 15200, 10, True

Formulario = "Qualidade/Plano de inspeção"
Cmb_opcao_lista = "Validação"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
With cmbtipomed
    .Clear
    Set TBInstrumentos = CreateObject("adodb.recordset")
    TBInstrumentos.Open "Select * from tipodimensao where tipo is not null order by tipo", Conexao, adOpenKeyset, adLockOptimistic
    If TBInstrumentos.EOF = False Then
        Do While TBInstrumentos.EOF = False
            .AddItem TBInstrumentos!Tipo
            TBInstrumentos.MoveNext
        Loop
    End If
    TBInstrumentos.Close
End With

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Plano de inspeção"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
  
If txtPI.Text = "" Then
    NomeCampo = "o plano de inspeção"
    Acao = "visualizar impressão"
    ProcVerificaAcao
    Exit Sub
End If
frmPlanoinspecao_Menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362P" Then frmPlanoinspecao_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmPlanoinspecao_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza tolerância inferior
            Set TBplano = CreateObject("adodb.recordset")
            TBplano.Open "Select tolinf from Planodimensao where tolinf <> 0 order by iddimensao", Conexao, adOpenKeyset, adLockOptimistic
            If TBplano.EOF = False Then
                TBplano.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBplano.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBplano.MoveFirst
                Do While TBplano.EOF = False
                    TBplano!TolInf = -TBplano!TolInf
                    TBplano.Update
                    TBplano.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            Set TBplano = CreateObject("adodb.recordset")
            TBplano.Open "Select tolinf from Medicaodimensao where tolinf <> 0 order by idmedicao", Conexao, adOpenKeyset, adLockOptimistic
            If TBplano.EOF = False Then
                TBplano.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBplano.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBplano.MoveFirst
                Do While TBplano.EOF = False
                    TBplano!TolInf = -TBplano!TolInf
                    TBplano.Update
                    TBplano.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBplano.Close
        End If
        
        If .Chk2.Value = 1 Then
            'Atualizar dados das dimensões
            Conexao.Execute "Update Planodimensao Set PCP = 'True' where PCP is null"
        End If
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Plano de inspeção"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End With
End If
            
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
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) plano(s) de inspeção?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBplano = CreateObject("adodb.recordset")
            TBplano.Open "Select * from plano where idplano = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBplano.EOF = False Then
                Conexao.Execute "DELETE from plano WHERE idplano = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from Planodimensao_instrumentos from Planodimensao_instrumentos INNER JOIN planodimensao ON Planodimensao_instrumentos.ID_dimensao = planodimensao.idDimensao Where planodimensao.idplano = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from planodimensao where idplano = " & .ListItems(InitFor)
            
                If IsNull(TBplano!IDFase) = False And TBplano!IDFase <> "0" Then
                    Conexao.Execute "UPDATE Fases Set Plano_inspecao = 'False' where IDFase = " & TBplano!IDFase
                Else
                    Conexao.Execute "UPDATE projproduto Set Plano_inspecao = 'False' where Desenho = '" & TBplano!Desenho & "'"
                End If
            
                '==================================
                Modulo = "Qualidade/Plano de inspeção"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Plano de inspeção: " & .ListItems(InitFor)
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBplano.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) plano(s) de inspeção antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Plano(s) de inspeção excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpar
    Lista.ListItems.Clear
    ProcCarregaLista (1)
    Novo_Plano = False
    Frame1.Enabled = False
    With cmbReferencia
        .Locked = True
        .TabStop = False
    End With
    ProcLimparTudo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparCamposDimensoes()
On Error GoTo tratar_erro

Txt_ID = 0
txtIndice = ""
txtNumero = ""
cmbtipomed.Clear
Chk_relatorio_pcp.Value = 1
txtdesejada.Text = ""
txtDim_inf = ""
txtdim_sup = ""
txttolsup.Text = ""
txttolinf.Text = ""
Txtfrequencia.Text = ""
Txt_vista = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparCamposFamilia()
On Error GoTo tratar_erro

Txt_ID1 = 0
Cmb_familia.ListIndex = -1
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
  
frmPlanoinspecao_Localizar.Show 1

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
ProcLimpar
ProcLimparTudo
Copiar = False
frmPlanoinspecao_Novo.Show 1
If Novo_Plano = True And Frame1.Enabled = True Then
    With cmbReferencia
        .Locked = False
        .TabStop = True
    End With
    cmbNivel.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

framemed.Enabled = False
Frame5.Enabled = False
ProcLimparCamposDimensoes
ProcLimparCamposFamilia
Lista1.ListItems.Clear
Lista2.ListItems.Clear
Novo_Plano1 = False
Novo_Plano2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Plano = True Then
    If USMsgBox("O plano de inspeção ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Plano = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Plano1 = True Then
    If USMsgBox("A dimensão ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarDimensao
        If Novo_Plano1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Plano2 = True Then
    If USMsgBox("A família do instrumento ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarFamilia
        If Novo_Plano2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Plano = False
Novo_Plano1 = False
Novo_Plano2 = False
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
If cmbNivel = "" Then
    USMsgBox ("Informe o nível antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmbNivel.SetFocus
    Exit Sub
End If
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select * from plano where idplano = " & IIf(txtPI = "", 0, txtPI), Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = True Then
    TBplano.AddNew
Else
    If FunVerificaRegistroValidado("plano", "idplano = " & txtPI, "mesmo", "o plano de inspeção", "alterar", True, True) = False Then Exit Sub
    If txtIDFase <> "" And txtIDFase <> "0" Then
        If FunVerifiProcRevisado(txtIDFase, "alterar este plano", True) = True Then Exit Sub
        Conexao.Execute "UPDATE Fases Set Plano_inspecao = 'False' where IDFase = " & TBplano!IDFase
    Else
        Conexao.Execute "UPDATE projproduto Set Plano_inspecao = 'False' where Desenho = '" & TBplano!Desenho & "'"
    End If
End If
TBplano!Rev = IIf(txtRev = "", 0, txtRev)
TBplano!Data = IIf(txtData = "", Date, txtData)
TBplano!Inspetor = IIf(txtinspetor = "", pubUsuario, txtinspetor)
TBplano!Desenho = txtdesenho.Text
TBplano!Descricao = txtdescricao.Text
TBplano!IDFase = IIf(txtIDFase = "", Null, txtIDFase)
TBplano!Fase = IIf(txtFase = "", Null, txtFase)
TBplano!Grupo_op = IIf(txtGrupo = "", Null, txtGrupo)
TBplano!Nivel = cmbNivel
TBplano.Update
txtPI = TBplano!IDPlano
TBplano.Close

Lista.ListItems.Clear
If Novo_Plano = True Then
    USMsgBox ("Novo plano de inspeção cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_Plano_Localizar = "Select * from plano where idplano = " & IIf(txtPI = "", 0, txtPI)
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
Modulo = "Qualidade/Plano de inspeção"
ID_documento = txtPI
Documento = "Plano de inspeção: " & txtPI
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Plano = False
If txtIDFase <> "" And txtIDFase <> "0" Then
    Conexao.Execute "UPDATE Fases Set Plano_inspecao = 'True' where IDFase = " & txtIDFase
Else
    Conexao.Execute "UPDATE projproduto Set Plano_inspecao = 'True' where Desenho = '" & txtdesenho & "'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Label12_DblClick()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362LIBFASE" Then
    With txtFase
        .Locked = False
        .TabStop = True
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "Plano" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBplano = CreateObject("adodb.recordset")
                TBplano.Open "Select IDPlano, IDFase, Desenho, Fase from Plano where IDPlano = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBplano.EOF = False Then
                    If Cmb_opcao_lista = "Excluir" Then
                        If FunVerificaRegistroValidadoSemMsg("Plano", "IDPlano = " & .ListItems(InitFor), True) = False Then GoTo Proximo
                        
                        If IsNull(TBplano!IDFase) = False And TBplano!IDFase <> 0 Then
                            If FunVerifiProcRevisado(TBplano!IDFase, "excluir este plano", False) = True Then GoTo Proximo
                        End If
                        
                        If IsNull(TBplano!IDFase) = False And TBplano!IDFase <> 0 Then
                            ProcVerificaRegistroUtilizadoSemMsg "Ordemservico", "IDPlano = " & TBplano!IDPlano
                            If Permitido = False Then GoTo Proximo
                        End If
                        
                        If IsNull(TBplano!IDFase) = False And TBplano!IDFase <> 0 Then Familiatext = "IDFase = " & TBplano!IDFase Else Familiatext = "Desenho = '" & TBplano!Desenho & "'"
                        ProcVerificaRegistroUtilizadoSemMsg "Medicao", Familiatext
                        If Permitido = False Then GoTo Proximo
                        
                    End If
                End If
                TBplano.Close
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
            Set TBplano = CreateObject("adodb.recordset")
            TBplano.Open "Select IDPlano, IDFase, Desenho, Fase from Plano where IDPlano = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBplano.EOF = False Then
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidado("Plano", "IDPlano = " & .ListItems(InitFor), "mesmo", "plano de inspeção", "excluir este", True, True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If IsNull(TBplano!IDFase) = False And TBplano!IDFase <> 0 Then
                        If FunVerifiProcRevisado(TBplano!IDFase, "excluir este plano", True) = True Then
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    End If
                
                    Mensagem = "Não é permitido excluir este plano de inspeção, pois o mesmo está sendo utilizado no módulo"
                
                    If IsNull(TBplano!IDFase) = False And TBplano!IDFase <> 0 Then
                        ProcVerificaRegistroUtilizado "Ordemservico", "IDPlano = " & TBplano!IDPlano, "PCP/Gerenciamento de ordem"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    End If
                    
                    If IsNull(TBplano!IDFase) = False And TBplano!IDFase <> 0 Then Familiatext = "IDFase = " & TBplano!IDFase Else Familiatext = "Desenho = '" & TBplano!Desenho & "'"
                    ProcVerificaRegistroUtilizado "Medicao", Familiatext, "Qualidade/Controle de medição"
                    If Permitido = False Then .ListItems.Item(InitFor).Checked = False
                End If
            End If
            TBplano.Close
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
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select * from Plano where IDPlano = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = False Then
    ProcLimpar
    ProcCarregaDados
    Novo_Plano = False
    Frame1.Enabled = True
    With cmbReferencia
        .Locked = False
        .TabStop = True
    End With
    CodigoLista = Lista.SelectedItem.index
End If
TBplano.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Plano", "IDPlano = " & txtPI, True) = False Then GoTo Proximo
                If txtIDFase <> "" And txtIDFase <> "0" Then
                    If FunVerifiProcRevisado(txtIDFase, "excluir estas dimensões", False) = True Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Plano", "IDPlano = " & txtPI, "plano de inspeção", "dimensões", "excluir estas", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If txtIDFase <> "" And txtIDFase <> "0" Then
                If FunVerifiProcRevisado(txtIDFase, "excluir estas dimensões", True) = True Then .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista1.ListItems.Count = 0 Then Exit Sub
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from planodimensao where iddimensao = " & Lista1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    ProcLimparCamposDimensoes
    ProcCarregaDadosDimensao
    CodigoLista1 = Lista1.SelectedItem.index
End If
TBOrdem.Close
framemed.Enabled = True
Novo_Plano1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosDimensao()
On Error GoTo tratar_erro

ProcCarregaFamilia
ProcCarrega_Tipo
NomeCampo = "o tipo da dimensão"
cmbtipomed.Text = IIf(IsNull(TBOrdem!Tipo), "", TBOrdem!Tipo)
2:
    Txt_ID.Text = TBOrdem!idDimensao
    If TBOrdem!PCP = True Then Chk_relatorio_pcp.Value = 1 Else Chk_relatorio_pcp.Value = 0
    txtdesejada.Text = IIf(IsNull(TBOrdem!dimdesejada), "", Format(TBOrdem!dimdesejada, "###,##0.0000"))
    txttolsup.Text = IIf(IsNull(TBOrdem!TolSup), "", Format(TBOrdem!TolSup, "###,##0.0000"))
    txttolinf.Text = IIf(IsNull(TBOrdem!TolInf), "", Format(TBOrdem!TolInf, "###,##0.0000"))
    Txt_vista = IIf(IsNull(TBOrdem!Vista), "", TBOrdem!Vista)
    txtIndice.Text = IIf(IsNull(TBOrdem!indice), "", TBOrdem!indice)
    txtNumero = IIf(IsNull(TBOrdem!Numero), "", TBOrdem!Numero)
    Txtfrequencia.Text = IIf(IsNull(TBOrdem!Freq), "", TBOrdem!Freq)
    txtdim_sup.Text = IIf(IsNull(TBOrdem!Dim_superior), "", Format(TBOrdem!Dim_superior, "###,##0.0000"))
    txtDim_inf.Text = IIf(IsNull(TBOrdem!Dim_inferior), "", Format(TBOrdem!Dim_inferior, "###,##0.0000"))
       
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " desta dimensão."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista2
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Plano", "IDPlano = " & txtPI, True) = False Then GoTo Proximo
                If txtIDFase <> "" And txtIDFase <> "0" Then
                    If FunVerifiProcRevisado(txtIDFase, "excluir esta família", False) = True Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista2, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Plano", "IDPlano = " & txtPI, "plano de inspeção", "família", "excluir esta", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If txtIDFase <> "" And txtIDFase <> "0" Then
                If FunVerifiProcRevisado(txtIDFase, "excluir esta família", True) = True Then .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista2.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Planodimensao_instrumentos where ID = " & Lista2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimparCamposFamilia
    Txt_ID1 = TBAbrir!ID
    NomeCampo = "a família"
    Cmb_familia = TBAbrir!Familia
2:
    CodigoLista2 = Lista2.SelectedItem.index
End If
TBAbrir.Close
Frame5.Enabled = True
Novo_Plano2 = False

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarrega_Tipo()
On Error GoTo tratar_erro
  
With cmbtipomed
    .Clear
    Set TBTipo = CreateObject("adodb.recordset")
    TBTipo.Open "Select * from tipodimensao where tipo <> 'Null' order by tipo", Conexao, adOpenKeyset, adLockOptimistic
    If TBTipo.EOF = False Then
        Do While TBTipo.EOF = False
            .AddItem TBTipo!Tipo
            TBTipo.MoveNext
        Loop
    End If
    TBTipo.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

Caption = "Qualidade - Plano de inspeção - (Plano : " & TBplano!IDPlano & " - Rev. : " & IIf(IsNull(TBplano!Rev), "0", TBplano!Rev) & " - Cód. interno : " & TBplano!Desenho & ")"
txtPI.Text = TBplano!IDPlano
txtRev.Text = IIf(IsNull(TBplano!Rev), "0", TBplano!Rev)
txtData.Text = IIf(IsNull(TBplano!Data), "", Format(TBplano!Data, "dd/mm/yy"))
txtinspetor.Text = IIf(IsNull(TBplano!Inspetor), "", TBplano!Inspetor)
txtDtValidacao = IIf(IsNull(TBplano!DtValidacao), "", TBplano!DtValidacao)
txtRespValidacao = IIf(IsNull(TBplano!RespValidacao), "", TBplano!RespValidacao)
txtdesenho.Text = IIf(IsNull(TBplano!Desenho), "", TBplano!Desenho)

'Revisão do produto
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select Revdesenho, DtValidacaoPlano, RespValidacaoPlano from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtRev_item = IIf(IsNull(TBItem!RevDesenho), "", TBItem!RevDesenho)
    txtDtValidacao_prod = IIf(IsNull(TBItem!DtValidacaoPlano), "", TBItem!DtValidacaoPlano)
    txtRespValidacao_prod = IIf(IsNull(TBItem!RespValidacaoPlano), "", TBItem!RespValidacaoPlano)
End If

txtdescricao.Text = IIf(IsNull(TBplano!Descricao), "", TBplano!Descricao)
txtIDFase = IIf(IsNull(TBplano!IDFase), 0, TBplano!IDFase)
txtFase = IIf(IsNull(TBplano!Fase), "", TBplano!Fase)
txtGrupo = IIf(IsNull(TBplano!Grupo_op), "", TBplano!Grupo_op)

'Versão da fase
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select versao from fases where IDfase = " & txtIDFase, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then txtVersao = IIf(IsNull(TBItem!versao), "", TBItem!versao)
TBItem.Close

If IsNull(TBplano!Nivel) = False And TBplano!Nivel <> "" Then cmbNivel = TBplano!Nivel
Novo_Plano = False
ProcLimparTudo
Caption = "Qualidade - Plano de inspeção (Plano : " & TBplano!IDPlano & " - Cód. interno : " & TBplano!Desenho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro
  
txtPI.Text = ""
txtRev = ""
txtData.Text = Format(Date, "dd/mm/yy")
txtinspetor.Text = pubUsuario
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtDtValidacao_prod.Text = ""
txtRespValidacao_prod.Text = ""
txtdesenho.Text = ""
txtRev_item = ""
cmbReferencia.Clear
txtdescricao.Text = ""
txtIDFase = 0
With txtFase
    .Text = ""
    .Locked = True
    .TabStop = False
End With
txtVersao = ""
txtGrupo = ""
cmbNivel.ListIndex = -1
CodigoLista = 0
Caption = "Qualidade - Plano de inspeção"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirDimensao()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) dimensão(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from planodimensao WHERE iddimensao = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Planodimensao_instrumentos WHERE id_dimensao = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Plano de inspeção"
            Evento = "Excluir dimensão"
            ID_documento = .ListItems(InitFor)
            Documento = "Plano de inspeção: " & txtPI
            Documento1 = "Tipo da dimensão: " & .ListItems(InitFor).ListSubItems(3) & " - Dimensão: " & .ListItems(InitFor).ListSubItems(4)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) dimensão(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Dimensão(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparCamposDimensoes
    ProcCarregaListaDimensao
    Novo_Plano1 = False
    framemed.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtPI = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
If SSTab1.Tab = 0 Then
    cmbReferencia.Visible = True
    Lista.Visible = True
    Lista1.Visible = False
    Lista2.Visible = False
    framemed.Visible = False
    If Lista.Visible = True Then Lista.SetFocus
Else
    cmbReferencia.Visible = False
    Lista.Visible = False
    If SSTab2.Tab = 0 Then
        Lista1.Visible = True
        framemed.Visible = True
    Else
        Lista2.Visible = True
        framemed.Visible = False
    End If
    If Novo_Plano = True Then
        USMsgBox ("Salve o plano de inspeção antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
        SSTab1.Tab = 0
        Exit Sub
    End If
    If Lista1.Visible = True Then Lista1.SetFocus Else Lista2.SetFocus
    ProcCarregaListaDimensao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab2.Tab = 0 Then
    Lista1.Visible = True
    Lista2.Visible = False
    framemed.Visible = True
    Lista1.SetFocus
    ProcCarregaListaDimensao
Else
    Lista1.Visible = False
    Lista2.Visible = True
    framemed.Visible = False
    If Txt_ID = 0 Then
        USMsgBox ("Informe a dimensão antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
        SSTab2.Tab = 0
        Exit Sub
    End If
    Lista2.SetFocus
    ProcCarregaListaFamilia
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesejada_Change()
On Error GoTo tratar_erro

If txtdesejada.Text <> "" Then
    VerifNumero = txtdesejada.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtdesejada.Text = ""
        txtdesejada.SetFocus
        Exit Sub
    End If
    ProcCalculaTolInf
    ProcCalculaTolSup
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesejada_LostFocus()
On Error GoTo tratar_erro

txtdesejada.Text = Format(txtdesejada.Text, "###,##0.0000")
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Revdesenho from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtRev_item = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
End If
TBProduto.Close

ProcCarregaComboCodRef cmbReferencia, "P.desenho = '" & txtdesenho & "'", 0, "", False, True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDim_inf_Change()
On Error GoTo tratar_erro

If txtdesejada.Text = "" Then
    ProcLimpaCamposDim
    Exit Sub
End If
If txtDim_inf.Text <> "" Then
    VerifNumero = txtDim_inf.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDim_inf.Text = ""
        txtDim_inf.SetFocus
        Exit Sub
    End If
    ProcCalculaTolInf
Else
    txttolinf.Text = ""
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTolInf()
On Error GoTo tratar_erro

Dimensao = IIf(txtdesejada = "", 0, txtdesejada)
DimInf = IIf(txtDim_inf = "", 0, txtDim_inf)
txttolinf.Text = Format(DimInf - Dimensao, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDim_inf_LostFocus()
On Error GoTo tratar_erro

txtDim_inf.Text = Format(txtDim_inf.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdim_sup_Change()
On Error GoTo tratar_erro

If txtdesejada.Text = "" Then
    ProcLimpaCamposDim
    Exit Sub
End If
If txtdim_sup.Text <> "" Then
    VerifNumero = txtdim_sup.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtdim_sup.Text = ""
        txtdim_sup.SetFocus
        Exit Sub
    End If
    ProcCalculaTolSup
Else
    txttolsup.Text = ""
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTolSup()
On Error GoTo tratar_erro

Dimensao = IIf(txtdesejada = "", 0, txtdesejada)
DimSup = IIf(txtdim_sup = "", 0, txtdim_sup)
txttolsup.Text = Format(DimSup - Dimensao, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdim_sup_LostFocus()
On Error GoTo tratar_erro

txtdim_sup.Text = Format(txtdim_sup.Text, "###,##0.0000")

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

Private Sub txtNumero_Change()
On Error GoTo tratar_erro

If txtNumero <> "" Then
    VerifNumero = txtNumero
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero = ""
        txtNumero.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposDim()
On Error GoTo tratar_erro

txtdim_sup.Text = ""
txtDim_inf.Text = ""
txttolsup.Text = ""
txttolinf.Text = ""

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
    Case 8: ProcCopiar
    Case 9: ProcRevisar
    Case 10: frmPlanoinspecao_validacao.Show 1
    Case 11: ProcAtualizar
    'Case 13: ProcAjuda
    Case 14: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoDimensao
    Case 2: ProcSalvarDimensao
    Case 3: ProcExcluirDimensao
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

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoFamilia
    Case 2: ProcSalvarFamilia
    Case 3: ProcExcluirFamilia
    Case 4: ProcImprimir
    Case 5: ProcAnteriorDimensao
    Case 6: ProcProximoDimensao
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifiProcRevisado(IDFase As Long, MsgemPadrao As String, MostrarMsgem As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifiProcRevisado = False
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select P.Nprocesso, P.Revisao from processos P INNER JOIN Fases F ON F.IDProcesso = P.IDProcesso where F.IDFase = " & IDFase, Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "Select IDprocesso from Processos where Nprocesso = '" & TBFases!Nprocesso & "' and Revisao > " & TBFases!Revisao, Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessos.EOF = False Then
        If MostrarMsgem = True Then USMsgBox ("Não é permitido " & MsgemPadrao & ", pois o processo vinculado já foi revisado."), vbExclamation, "CAPRIND v5.0"
        FunVerifiProcRevisado = True
    End If
    TBProcessos.Close
End If
TBFases.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
