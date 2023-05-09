VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlanomedicao_peca 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Controle de medição - Dimensões -  Por peça"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
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
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Operação da lista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   12945
      TabIndex        =   31
      Top             =   9480
      Width           =   2310
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
         ItemData        =   "frmPlanomedicao_peca.frx":0000
         Left            =   180
         List            =   "frmPlanomedicao_peca.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   170
         Width           =   1965
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
      Left            =   2790
      Locked          =   -1  'True
      MaxLength       =   20
      MouseIcon       =   "frmPlanomedicao_peca.frx":0022
      MousePointer    =   99  'Custom
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtid 
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
      Left            =   1950
      MaxLength       =   25
      MouseIcon       =   "frmPlanomedicao_peca.frx":032C
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Text            =   "0"
      ToolTipText     =   "Tipo da dimensão."
      Top             =   4260
      Visible         =   0   'False
      Width           =   825
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5925
      Left            =   60
      TabIndex        =   13
      Top             =   3540
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10451
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
         Object.Tag             =   "N"
         Text            =   "Cód. peça"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Tipo da dimensão"
         Object.Width           =   8916
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Dim. indicada"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Tol. sup."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Tol. inf."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Dim. encontr."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Aprovado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Restrição"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Vista"
         Object.Width           =   1147
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
      Height          =   2535
      Left            =   55
      TabIndex        =   15
      Top             =   990
      Width           =   15195
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
         Left            =   13695
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Dimensão encontrada."
         Top             =   375
         Width           =   1305
      End
      Begin VB.TextBox txtCodigo 
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código da peça."
         Top             =   375
         Width           =   1590
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
         Left            =   11820
         TabIndex        =   24
         Top             =   740
         Width           =   1575
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
            TabIndex        =   8
            Top             =   225
            Width           =   585
         End
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
            TabIndex        =   9
            Top             =   210
            Width           =   615
         End
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
         Left            =   13425
         TabIndex        =   23
         Top             =   740
         Width           =   1575
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
            TabIndex        =   11
            Top             =   210
            Width           =   732
         End
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
            Height          =   252
            Left            =   180
            TabIndex        =   10
            Top             =   210
            Width           =   585
         End
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
         Left            =   11055
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Tolerância superior."
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
         Left            =   12375
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Tolerância inferior."
         Top             =   375
         Width           =   1305
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
         Left            =   9735
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Dimensão indicada."
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
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Frequência de medição."
         Top             =   950
         Width           =   6765
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
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Tipo da dimensão."
         Top             =   375
         Width           =   7935
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
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   950
         Width           =   4785
      End
      Begin VB.TextBox txtobs 
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
         Height          =   825
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         ToolTipText     =   "Observações."
         Top             =   1560
         Width           =   14805
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dim. encontr."
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
         Left            =   13860
         TabIndex        =   27
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   420
         TabIndex        =   25
         Top             =   180
         Width           =   1110
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
         Left            =   11392
         TabIndex        =   22
         Top             =   180
         Width           =   630
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
         Left            =   9915
         TabIndex        =   21
         Top             =   180
         Width           =   945
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
         Left            =   12750
         TabIndex        =   20
         Top             =   180
         Width           =   555
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
         TabIndex        =   19
         Top             =   750
         Width           =   1650
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
         TabIndex        =   18
         Top             =   750
         Width           =   915
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
         Left            =   5130
         TabIndex        =   17
         Top             =   180
         Width           =   1245
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
         Left            =   7110
         TabIndex        =   16
         Top             =   1350
         Width           =   945
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   29
      Top             =   9615
      Width           =   12765
      _ExtentX        =   22516
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   9
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
      ButtonCaption5  =   "Restrição"
      ButtonEnabled5  =   0   'False
      ButtonToolTipText5=   "Aprovar medição com restrição (F7)"
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
      ButtonWidth5    =   62
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonAlignment6=   2
      ButtonType6     =   1
      ButtonStyle6    =   -1
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   -1
      ButtonLeft6     =   259
      ButtonTop6      =   4
      ButtonWidth6    =   2
      ButtonHeight6   =   54
      ButtonCaption7  =   "Ajuda"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Ajuda (F1)"
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
      ButtonWidth7    =   41
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Sair"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Sair (Esc)"
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
      ButtonLeft8     =   306
      ButtonTop8      =   2
      ButtonWidth8    =   30
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonKey9      =   "9"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   5
      ButtonLeft9     =   338
      ButtonTop9      =   2
      ButtonWidth9    =   24
      ButtonHeight9   =   24
      ButtonUseMaskColor9=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   6060
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmPlanomedicao_peca.frx":0636
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmPlanomedicao_peca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_dimensaoPeca As Boolean 'OK
Dim CODIGO            As Integer 'OK
Dim Codigo1           As Integer 'OK
Dim Pagina            As Integer 'OK

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Conexao.Execute "DELETE from Medicaodimensao_peca_relatorios"
NomeRel = "CQ_plano medicao_peca.rpt"
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Medicaodimensao_peca.* FROM Medicaodimensao_peca INNER JOIN Medicaodimensao ON Medicaodimensao_peca.IDdimensao = Medicaodimensao.idmedicao where Medicaodimensao.IDPlano = " & frmPlanomedicao.txtPm & " order by Medicaodimensao.idmedicao, Medicaodimensao_peca.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    TBFI.MoveLast
    If TBFI!CODIGO > 4 Then
        CODIGO = TBFI!CODIGO
        If CODIGO <= 4 Then
            Pagina = 1
        ElseIf CODIGO <= 8 Then
                Pagina = 2
            ElseIf CODIGO <= 12 Then
                    Pagina = 3
                ElseIf CODIGO <= 16 Then
                        Pagina = 4
                    ElseIf CODIGO <= 20 Then
                            Pagina = 5
                        ElseIf CODIGO <= 24 Then
                                Pagina = 6
                            ElseIf CODIGO <= 28 Then
                                    Pagina = 7
                                ElseIf CODIGO <= 32 Then
                                        Pagina = 8
        End If
        TBFI.MoveFirst
        Codigo1 = TBFI!CODIGO
        Do While CODIGO > 0
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Medicaodimensao_peca.* FROM Medicaodimensao_peca INNER JOIN Medicaodimensao ON Medicaodimensao_peca.IDdimensao = Medicaodimensao.idmedicao where Medicaodimensao.IDPlano = " & frmPlanomedicao.txtPm & " and Medicaodimensao_peca.Codigo = " & Codigo1 & " order by Medicaodimensao.idmedicao, Medicaodimensao_peca.Codigo", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Do While TBFIltro.EOF = False
                    ProcGravarRel
                    TBFIltro.MoveNext
                Loop
            End If
            TBFIltro.Close
            If CODIGO = 1 Or Codigo1 = 4 Or Codigo1 = 8 Or Codigo1 = 12 Or Codigo1 = 16 Or Codigo1 = 20 Or Codigo1 = 24 Or Codigo1 = 28 Or Codigo1 = 32 Then
                Select Case Codigo1
                    Case Is <= 4: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "01" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 1 & " de " & Pagina & "'"
                    Case Is > 4 <= 8: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "05" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 2 & " de " & Pagina & "'"
                    Case Is > 8 <= 12: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "09" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 3 & " de " & Pagina & "'"
                    Case Is > 12 <= 16: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "13" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 4 & " de " & Pagina & "'"
                    Case Is > 16 <= 20: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "17" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 5 & " de " & Pagina & "'"
                    Case Is > 20 <= 24: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "21" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 6 & " de " & Pagina & "'"
                    Case Is > 24 <= 28: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "25" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 7 & " de " & Pagina & "'"
                    Case Is > 28 <= 32: Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "29" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 8 & " de " & Pagina & "'"
                End Select
                ProcImprimirRel "", ""
                Conexao.Execute "DELETE from Medicaodimensao_peca_relatorios"
            End If
            CODIGO = CODIGO - 1
            Codigo1 = Codigo1 + 1
        Loop
    Else
        CODIGO = TBFI!CODIGO
        Pagina = 1
        TBFI.MoveFirst
        Codigo1 = TBFI!CODIGO
        Do While CODIGO > 0
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Medicaodimensao_peca.* FROM Medicaodimensao_peca INNER JOIN Medicaodimensao ON Medicaodimensao_peca.IDdimensao = Medicaodimensao.idmedicao where Medicaodimensao.IDPlano = " & frmPlanomedicao.txtPm & " and Medicaodimensao_peca.Codigo = " & Codigo1 & " order by Medicaodimensao.idmedicao, Medicaodimensao_peca.Codigo", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Do While TBFIltro.EOF = False
                    ProcGravarRel
                    TBFIltro.MoveNext
                Loop
            End If
            TBFIltro.Close
            Conexao.Execute "Update Medicaodimensao_peca_relatorios Set Numero_serie = '" & "01" & " À " & "0" & Codigo1 & "', Pagina = '" & "Página " & 1 & " de " & Pagina & "'"
            CODIGO = CODIGO - 1
            Codigo1 = Codigo1 + 1
        Loop
        ProcImprimirRel "", ""
        Conexao.Execute "DELETE from Medicaodimensao_peca_relatorios"
    End If
End If
TBFI.Close
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarRel()
On Error GoTo tratar_erro

With frmPlanomedicao
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Medicaodimensao_peca_relatorios", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from Producao where Ordem = " & .Txtpeca, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        TBGravar!IDPlano = .txtPm
        TBGravar!idDimensao = TBFIltro!idDimensao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Medicao where Data = '" & Format(.txtData, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Contador = 0
            ID = 0
            Do While TBAbrir.EOF = False
                If ID <> TBAbrir!IDPlano Then Contador = Contador + 1
                ID = TBAbrir!IDPlano
                TBAbrir.MoveNext
            Loop
        End If
        If Contador < 10 Then Contador1 = 0 & Contador Else Contador1 = Contador
       
        AnoTexto = Right(Year(.txtData), 2)
        Select Case AnoTexto
            Case "01": AnoTexto = 1
            Case "02": AnoTexto = 2
            Case "03": AnoTexto = 3
            Case "04": AnoTexto = 4
            Case "05": AnoTexto = 5
            Case "06": AnoTexto = 6
            Case "07": AnoTexto = 7
            Case "08": AnoTexto = 8
            Case "09": AnoTexto = 9
        End Select
        
        MesTexto = Month(.txtData)
        If MesTexto < 10 Then MesTexto = 0 & MesTexto
        
        DiaTexto = Day(.txtData)
        If DiaTexto < 10 Then DiaTexto = 0 & DiaTexto
        
        TBGravar!Documento = AnoTexto & MesTexto & DiaTexto & "-" & Contador1
        TBGravar!Cliente = TBproducao!Cliente
        TBGravar!Codigo_interno = TBproducao!Desenho
        TBGravar!Revisao = TBproducao!Revitem
        TBGravar!Codigo_referencia = TBproducao!N_referencia
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Rev from item_aplicacoes where n_referencia = '" & TBproducao!N_referencia & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBGravar!Revisao_referencia = TBAbrir!Rev
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select EC.Corrida from Producao_NF_Consignada PNFC INNER JOIN Estoque_controle EC ON EC.IDestoque = PNFC.IDestoque where PNFC.Ordem = " & TBproducao!Ordem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBGravar!Certificado = TBAbrir!Corrida
        End If
        TBAbrir.Close
        TBGravar!Ordem = TBproducao!Ordem
        TBGravar!quantidade = TBproducao!Quant
        TBGravar!CODIGO = Codigo1
        TBGravar!Encontrada = TBFIltro!Encontrada
    End If
    TBproducao.Close
    TBGravar.Update
    TBGravar.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) dimensão(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Medicaodimensao_peca where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Controle de medição/Dimensões/Por peça"
            Evento = "Excluir dimensão"
            ID_documento = .ListItems(InitFor)
            Documento = "Código da peça: " & .ListItems(InitFor).ListSubItems(1) & " - Tipo da dimensão: " & .ListItems(InitFor).ListSubItems(3) & " - Dimensão indicada: " & .ListItems(InitFor).ListSubItems(4)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) dimensão(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Dimensão(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
    Frame2.Enabled = False
    Novo_dimensaoPeca = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from Medicaodimensao_peca where IDdimensao = " & txtNumero & " order by codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    CODIGO = IIf(IsNull(TBAbrir!CODIGO), 0, TBAbrir!CODIGO)
    txtCodigo = CODIGO + 1
Else
    txtCodigo = 1
End If
Novo_dimensaoPeca = True
Frame2.Enabled = True
txtencontrada.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarRestricao()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente aprovar esta(s) dimensão(ões) com restrição?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "UPDATE Medicaodimensao_peca Set restricao = 'Sim', Obs = '" & Trim(txtobservacaodim) & "' where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Controle de medição/Dimensões/Por peça"
            Evento = "Aprovar dimensão com restrição"
            ID_documento = .ListItems(InitFor)
            Documento = "Código da peça: " & .ListItems(InitFor).ListSubItems(1) & " - Tipo da dimensão: " & .ListItems(InitFor).ListSubItems(3) & " - Dimensão indicada: " & .ListItems(InitFor).ListSubItems(4)
            Documento1 = ""
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
    Novo_dimensaoPeca = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_dimensaoPeca = True Then
    If USMsgBox("a dimensão ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_dimensaoPeca = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_dimensaoPeca = False
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
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtencontrada = "" Then
    NomeCampo = "a dimensão encontrada"
    ProcVerificaAcao
    txtencontrada.SetFocus
    Exit Sub
End If
ProcCalculaMedicao
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "select * from Medicaodimensao_peca where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!Responsavel = txtResponsavel
    TBGravar!Data = Date
End If
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close

ProcCarregaLista
If Novo_dimensaoPeca = True Then
    USMsgBox ("Nova dimensão cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
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
Modulo = "Qualidade/Controle de medição/Dimensões/Por peça"
ID_documento = txtId
Documento = "Código da peça: " & txtCodigo & " - Tipo da dimensão: " & txttipo & " - Dimensão indicada: " & txtdesejada
Documento1 = ""
ProcGravaEvento
'==================================
Novo_dimensaoPeca = False

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
Encontrada = Format(IIf(txtencontrada = "", 0, txtencontrada), "###,##0.0000")
If Encontrada >= TolInf And Encontrada <= TolSup Then Resultado = "Aprovado" Else Resultado = "Reprovado"

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
        .ButtonState(3) = 0
        .ButtonState(5) = 5
    Else
        .ButtonState(3) = 5
        .ButtonState(5) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: If Cmb_opcao_lista = "Restrição" Then ProcSalvarRestricao
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 9, True
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista = "Excluir"
With frmPlanomedicao
    txtNumero = .txtNumero
    txttipo = .txttipo
    txtdesejada = .txtdesejada
    txttolsup = .txttolsup
    txttolinf = .txttolinf
    Txtfrequencia = .Txtfrequencia
End With
ProcCarregaLista

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtCodigo = ""
txtencontrada = ""
txtObs = ""
txtResponsavel = pubUsuario
Checkdimaprosim.Value = 0
Checkdimapronao.Value = 0
Checkdimrestnao.Value = 0
Checkdimrestsim.Value = 0
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!CODIGO = txtCodigo
TBGravar!idDimensao = txtNumero
TBGravar!Encontrada = txtencontrada
TBGravar!Obs = Trim(txtObs)
If Resultado = "Aprovado" Then
    TBGravar!Aprovado = "Sim"
    Checkdimaprosim.Value = 1
    Checkdimapronao.Value = 0
    TBGravar!restricao = "Não"
    Checkdimrestnao.Value = 1
    Checkdimrestsim.Value = 0
Else
    TBGravar!Aprovado = "Não"
    Checkdimapronao.Value = 1
    Checkdimrestsim.Value = 0
    TBGravar!restricao = "Não"
    Checkdimrestnao.Value = 1
    Checkdimrestsim.Value = 0
End If

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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from medicaodimensao_peca where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtId = TBAbrir!ID
    txtCodigo = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
    txtencontrada.Text = IIf(IsNull(TBAbrir!Encontrada), "", Format(TBAbrir!Encontrada, "###,##0.0000"))
    txtResponsavel.Text = pubUsuario
    txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
    txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    If TBAbrir!Aprovado = "Sim" Then
        Checkdimaprosim.Value = 1
        Checkdimapronao.Value = 0
    Else
        Checkdimaprosim.Value = 0
        Checkdimapronao.Value = 1
    End If
    If TBAbrir!restricao = "Sim" Then
        Checkdimrestsim.Value = 1
        Checkdimrestnao.Value = 0
    Else
        Checkdimrestsim.Value = 0
        Checkdimrestnao.Value = 1
    End If
    Novo_dimensaoPeca = False
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close
Frame2.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodigo_LostFocus()
On Error GoTo tratar_erro

If txtCodigo.Text <> "" Then
    VerifNumero = txtCodigo.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCodigo.Text = ""
        txtCodigo.SetFocus
        Exit Sub
    End If
End If
    
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

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Medicaodimensao_peca.ID, Medicaodimensao_peca.Codigo, Medicaodimensao_peca.Encontrada, Medicaodimensao_peca.Aprovado as Aprovado1, Medicaodimensao_peca.Restricao as Restricao1, medicaodimensao.* from Medicaodimensao_peca INNER JOIN medicaodimensao ON Medicaodimensao_peca.idDimensao = medicaodimensao.idmedicao where Medicaodimensao_peca.IDdimensao = " & frmPlanomedicao.txtNumero & " order by Medicaodimensao_peca.codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!indice), "", TBLISTA!indice)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
            .Item(.Count).SubItems(4) = Format(TBLISTA!dimdesejada, "###,##0.0000")
            .Item(.Count).SubItems(5) = Format(TBLISTA!TolSup, "###,##0.0000")
            .Item(.Count).SubItems(6) = Format(TBLISTA!TolInf, "###,##0.0000")
            .Item(.Count).SubItems(7) = Format(TBLISTA!Encontrada, "###,##0.0000")
            If TBLISTA!Aprovado1 = "Sim" Then .Item(.Count).SubItems(8) = "Sim" Else .Item(.Count).SubItems(8) = "Não"
            If TBLISTA!restricao1 = "Sim" Then .Item(.Count).SubItems(9) = "Sim" Else .Item(.Count).SubItems(9) = "Não"
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: ProcImprimir
    Case 5: ProcSalvarRestricao
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
