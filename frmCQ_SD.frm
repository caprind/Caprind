VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCQ_SD 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Solicitação de desvio"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Solicitação de desvio"
      TabPicture(0)   =   "frmCQ_SD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "PBLista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "USImageList1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Dados adicionais"
      TabPicture(1)   =   "frmCQ_SD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USImageList2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Framelista"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   80
         Top             =   9090
         Width           =   15195
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
            TabIndex        =   22
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
            Left            =   3780
            TabIndex        =   21
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   26
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SD.frx":0038
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
            TabIndex        =   25
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SD.frx":37DF
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
            TabIndex        =   23
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
            TabIndex        =   24
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SD.frx":72ED
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
            TabIndex        =   27
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SD.frx":B3E2
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
         Begin VB.Label Label24 
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
            Left            =   4410
            TabIndex        =   85
            Top             =   240
            Width           =   1440
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
            TabIndex        =   83
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
            TabIndex        =   82
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label23 
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
            Left            =   3090
            TabIndex        =   81
            Top             =   240
            Width           =   645
         End
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -62790
         Top             =   570
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCQ_SD.frx":EC83
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13020
         Top             =   540
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCQ_SD.frx":13744
         Count           =   1
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   38
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
            MouseIcon       =   "frmCQ_SD.frx":19693
            MousePointer    =   99  'Custom
            TabIndex        =   42
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
            MouseIcon       =   "frmCQ_SD.frx":1999D
            MousePointer    =   99  'Custom
            TabIndex        =   41
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
            MouseIcon       =   "frmCQ_SD.frx":19CA7
            MousePointer    =   99  'Custom
            TabIndex        =   40
            ToolTipText     =   "Departamento do contato."
            Top             =   630
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
            MouseIcon       =   "frmCQ_SD.frx":19FB1
            MousePointer    =   99  'Custom
            TabIndex        =   39
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   1478
            Width           =   480
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   75
         TabIndex        =   20
         Top             =   4740
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Nº SD"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nº Ref."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "OS"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   6588
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Qtde. NC"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Frame Framelista 
         BackColor       =   &H00E0E0E0&
         Height          =   8685
         Left            =   -74925
         TabIndex        =   52
         Top             =   1320
         Width           =   15225
         Begin VB.ComboBox cmbDesvio 
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
            ItemData        =   "frmCQ_SD.frx":1A2BB
            Left            =   12300
            List            =   "frmCQ_SD.frx":1A2C5
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   32
            ToolTipText     =   "Tipo do desvio."
            Top             =   390
            Width           =   2715
         End
         Begin VB.TextBox txtDocumento 
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
            Left            =   8820
            MaxLength       =   50
            TabIndex        =   31
            ToolTipText     =   "Número do documento."
            Top             =   390
            Width           =   3465
         End
         Begin VB.TextBox txtFuncao 
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
            Left            =   4260
            MaxLength       =   50
            TabIndex        =   29
            ToolTipText     =   "Função."
            Top             =   390
            Width           =   3345
         End
         Begin VB.TextBox txtResponsavel_outros 
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
            Left            =   180
            MaxLength       =   50
            TabIndex        =   28
            ToolTipText     =   "Responsável."
            Top             =   390
            Width           =   4065
         End
         Begin VB.TextBox txtTexto4 
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
            Height          =   1635
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            ToolTipText     =   "Ordem de compra."
            Top             =   6900
            Width           =   14835
         End
         Begin VB.TextBox txtTexto2 
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
            Height          =   1665
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            ToolTipText     =   "Origem do desvio."
            Top             =   3020
            Width           =   14835
         End
         Begin VB.TextBox txtTexto3 
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
            Height          =   1665
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            ToolTipText     =   "Ação corretiva."
            Top             =   4960
            Width           =   14835
         End
         Begin VB.TextBox txtTexto1 
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
            Height          =   1665
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            ToolTipText     =   "Descrição do desvio."
            Top             =   1080
            Width           =   14835
         End
         Begin MSComCtl2.DTPicker txtData_outros 
            Height          =   315
            Left            =   7620
            TabIndex        =   30
            ToolTipText     =   "Data."
            Top             =   390
            Width           =   1185
            _ExtentX        =   2090
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo do desvio"
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
            Index           =   8
            Left            =   13140
            TabIndex        =   78
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Número do documento"
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
            Index           =   7
            Left            =   9742
            TabIndex        =   77
            Top             =   180
            Width           =   1620
         End
         Begin VB.Label Label2 
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
            Index           =   6
            Left            =   8040
            TabIndex        =   76
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Função"
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
            Left            =   5670
            TabIndex        =   75
            Top             =   180
            Width           =   525
         End
         Begin VB.Label Label2 
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
            Index           =   4
            Left            =   1755
            TabIndex        =   74
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ordem de compra"
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
            Left            =   6960
            TabIndex        =   70
            Top             =   6690
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Origem do desvio"
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
            Left            =   6975
            TabIndex        =   60
            Top             =   2820
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição do desvio"
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
            Left            =   6885
            TabIndex        =   58
            Top             =   870
            Width           =   1425
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ação corretiva"
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
            Left            =   7065
            TabIndex        =   53
            Top             =   4755
            Width           =   1050
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   3405
         Left            =   75
         TabIndex        =   47
         Top             =   1320
         Width           =   15225
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
            Left            =   12960
            Locked          =   -1  'True
            TabIndex        =   84
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "ID RNC."
            Top             =   390
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.ComboBox cmbStatus 
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
            ItemData        =   "frmCQ_SD.frx":1A2E3
            Left            =   7710
            List            =   "frmCQ_SD.frx":1A2F0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   2715
         End
         Begin VB.TextBox txtNref 
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
            Left            =   10440
            MaxLength       =   50
            TabIndex        =   4
            ToolTipText     =   "Nº de referência."
            Top             =   390
            Width           =   2505
         End
         Begin VB.TextBox txtQtdeNC 
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
            Left            =   3240
            TabIndex        =   9
            ToolTipText     =   "Quantidade não conforme."
            Top             =   990
            Width           =   1455
         End
         Begin VB.TextBox Txt_codref 
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
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   1590
            Width           =   2295
         End
         Begin VB.CommandButton cmdRNC 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmCQ_SD.frx":1A313
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Criar RNC."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtOS 
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
            Left            =   4710
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "OS."
            Top             =   990
            Width           =   1635
         End
         Begin VB.TextBox txtmaquina 
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
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Posto de trabalho."
            Top             =   990
            Width           =   5925
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
            Height          =   315
            Left            =   7740
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Grupo/op."
            Top             =   990
            Width           =   1365
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
            Height          =   315
            Left            =   6750
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Fase."
            Top             =   990
            Width           =   975
         End
         Begin VB.TextBox txtRNC 
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
            Left            =   12960
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Nº da RNC."
            Top             =   390
            Width           =   1725
         End
         Begin VB.TextBox txtOrdem 
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Ordem."
            Top             =   990
            Width           =   1575
         End
         Begin VB.CommandButton cmdOS 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   6360
            Picture         =   "frmCQ_SD.frx":1A3F5
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Localizar OS."
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox txtCliente 
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Cliente."
            Top             =   2190
            Width           =   14895
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
            Height          =   315
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   1590
            Width           =   10245
         End
         Begin VB.TextBox txtqtde 
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade."
            Top             =   990
            Width           =   1455
         End
         Begin VB.TextBox txtdesenho 
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   1590
            Width           =   2295
         End
         Begin VB.TextBox txtID 
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
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Número da solicitação de desvio."
            Top             =   390
            Width           =   1245
         End
         Begin VB.TextBox txtData 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   2460
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   5235
         End
         Begin VB.TextBox txtObs 
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
            Height          =   435
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            ToolTipText     =   "Observação."
            Top             =   2820
            Width           =   14895
         End
         Begin VB.Label Label22 
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
            Left            =   8835
            TabIndex        =   73
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº de referência"
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
            Left            =   11100
            TabIndex        =   69
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde. NC"
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
            Left            =   3630
            TabIndex        =   68
            Top             =   780
            Width           =   675
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   3028
            TabIndex        =   67
            Top             =   1380
            Width           =   1350
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "OS*"
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
            Left            =   5377
            TabIndex        =   66
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label16 
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
            Left            =   11445
            TabIndex        =   65
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   8055
            TabIndex        =   64
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label13 
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
            Left            =   7065
            TabIndex        =   63
            Top             =   780
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   13552
            TabIndex        =   62
            Top             =   180
            Width           =   540
         End
         Begin VB.Label Label6 
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
            Left            =   727
            TabIndex        =   61
            Top             =   780
            Width           =   480
         End
         Begin VB.Label Label1 
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
            Left            =   7380
            TabIndex        =   59
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde."
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
            Left            =   2287
            TabIndex        =   57
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label10 
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
            Left            =   9577
            TabIndex        =   56
            Top             =   1380
            Width           =   690
         End
         Begin VB.Label Label5 
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
            Left            =   877
            TabIndex        =   55
            Top             =   1380
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº SD"
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
            TabIndex        =   54
            Top             =   180
            Width           =   465
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
            Left            =   4620
            TabIndex        =   51
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
            Left            =   1818
            TabIndex        =   50
            Top             =   180
            Width           =   345
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
            TabIndex        =   49
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label Label21 
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
            TabIndex        =   48
            Top             =   2610
            Width           =   945
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   75
         TabIndex        =   71
         Top             =   9720
         Width           =   15225
         _ExtentX        =   26855
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
         Left            =   75
         TabIndex        =   72
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   1720
         ButtonCount     =   11
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   36
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
         ButtonLeft3     =   78
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
         ButtonLeft4     =   124
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
         ButtonLeft5     =   171
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
         ButtonLeft6     =   233
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
         ButtonLeft7     =   290
         ButtonTop7      =   2
         ButtonWidth7    =   55
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonAlignment8=   2
         ButtonType8     =   1
         ButtonStyle8    =   -1
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState8    =   -1
         ButtonLeft8     =   347
         ButtonTop8      =   4
         ButtonWidth8    =   2
         ButtonHeight8   =   54
         ButtonCaption9  =   "Ajuda"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Ajuda (F1)"
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
         ButtonLeft9     =   351
         ButtonTop9      =   2
         ButtonWidth9    =   41
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Sair"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Sair (Esc)"
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
         ButtonLeft10    =   394
         ButtonTop10     =   2
         ButtonWidth10   =   30
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   5
         ButtonLeft11    =   426
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   79
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
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
         ButtonUseMaskColor2=   0   'False
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
         ButtonLeft6     =   271
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
         ButtonLeft7     =   275
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
         ButtonLeft8     =   318
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
         ButtonLeft9     =   350
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   1170
      Top             =   150
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
End
Attribute VB_Name = "frmCQ_SD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_SD As Boolean 'OK
Public StrSql_CQ_SD_Localizar As String 'OK
Dim TBLISTA_CQ_SD As ADODB.Recordset 'OK

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CQ_SD order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos2
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_SD where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        ProcPuxadados2
    Else
        USMsgBox ("Fim dos cadastros de solicitação de desvio."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_SD = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdOS_Click()
On Error GoTo tratar_erro

frmCQ_SD_Ordem.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CQ_SD order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos2
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_SD where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        ProcPuxadados2
    Else
        USMsgBox ("Fim dos cadastros de solicitação de desvio."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_SD = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CQ_SD.AbsolutePage <> 2 Then
    If TBLISTA_CQ_SD.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_CQ_SD.PageCount - 1)
    Else
        TBLISTA_CQ_SD.AbsolutePage = TBLISTA_CQ_SD.AbsolutePage - 2
        ProcExibePagina (TBLISTA_CQ_SD.AbsolutePage)
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
    TBLISTA_CQ_SD.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_CQ_SD.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CQ_SD.AbsolutePage = 1
ProcExibePagina (TBLISTA_CQ_SD.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CQ_SD.AbsolutePage <> -3 Then
    If TBLISTA_CQ_SD.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_CQ_SD.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_CQ_SD.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CQ_SD.AbsolutePage = TBLISTA_CQ_SD.PageCount
ProcExibePagina (TBLISTA_CQ_SD.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdRNC_Click()
On Error GoTo tratar_erro

Acao = "criar a RNC"
If txtOS = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    cmdOS.SetFocus
    Exit Sub
End If
If Novo_SD = True Then
    USMsgBox ("Salve a solicitação de desvio antes de criar a RNC."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtQtdeNC = "" Or txtQtdeNC = "0,0000" Then
    USMsgBox ("Informe a quantidade não conforme antes de criar a RNC."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
RNC_Controle_Medicao = False
RNC_Inspecao_Recebimento = False
RNC_Nao_Conformidade = False
RNC_Solicitacao_Desvio = True
frmQualidade_RNC.Show

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
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: procSalvar_outros
            Case vbKeyF4: procExcluir_outros
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Caption = "Qualidade - Solicitação de desvio (Nº SD : " & TBAbrir!ID & ")"
txtId = TBAbrir!ID
txtData = IIf(IsNull(TBAbrir!Data), "", (Format(TBAbrir!Data, "dd/mm/yy")))
txtResponsavel.Text = IIf(IsNull(TBAbrir!Responsavel), "", (TBAbrir!Responsavel))
txtOrdem = IIf(IsNull(TBAbrir!Ordem), "", TBAbrir!Ordem)
txtOS = IIf(IsNull(TBAbrir!OS), "", TBAbrir!OS)

Txt_ID_RNC = IIf(IsNull(TBAbrir!ID_RNC), 0, TBAbrir!ID_RNC)
Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select ID_texto, Seq FROM CQ_RNC where ID = " & Txt_ID_RNC, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    txtRNC = IIf(IsNull(TBCompras_Pedido!Seq), TBCompras_Pedido!id_texto, TBCompras_Pedido!id_texto & "/" & IIf(TBCompras_Pedido!Seq < 10, "0" & TBCompras_Pedido!Seq, TBCompras_Pedido!Seq))
End If
TBCompras_Pedido.Close

txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
txtQtdeNC = IIf(IsNull(TBAbrir!qtdeNC), "", Format(TBAbrir!qtdeNC, "###,##0.0000"))
txtNRef = IIf(IsNull(TBAbrir!NRef), "", TBAbrir!NRef)
If IsNull(TBAbrir!status) = False Then cmbStatus = TBAbrir!status
Novo_SD = False
Frame1.Enabled = True

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select ordemservico.Ordem,ordemservico.idfase,ordemservico.fase,ordemservico.maquina,producao.desenho,producao.produto,producao.quant,producao.cliente from Producao INNER JOIN ordemservico on producao.Ordem = ordemservico.Ordem where Ordemservico.IDProducao = " & txtOS, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    txtOrdem = IIf(IsNull(TBOrdem!Ordem), "", TBOrdem!Ordem)
    txtdesenho = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
    txtdescricao = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
    txtQtde = IIf(IsNull(TBOrdem!Quant), "", Format(TBOrdem!Quant, "###,##0.0000"))
    txtCliente = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
    txtFase = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
    txtmaquina = IIf(IsNull(TBOrdem!maquina), "", TBOrdem!maquina)
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select grupo_op from Fases where idfase = " & TBOrdem!IDFase, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        txtGrupo = IIf(IsNull(TBFases!Grupo_op), "", TBFases!Grupo_op)
    End If
    TBFases.Close
End If
TBOrdem.Close

'Verifica se a SD está amarrada a outro módulo e bloqueia os botões necessários
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select ID_SD from CQ_NC_FABRICA where ID_SD = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    With txtQtdeNC
        .Locked = True
        .TabStop = False
    End With
    With cmdOS
        .Enabled = False
        .TabStop = False
    End With
Else
    With txtQtdeNC
        .Locked = False
        .TabStop = True
    End With
    With cmdOS
        .Enabled = True
        .TabStop = True
    End With
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos2()
On Error GoTo tratar_erro

txtResponsavel_outros = ""
txtFuncao = ""
txtData_outros = Date
txtDocumento = ""
cmbDesvio.ListIndex = -1
txtTexto1 = ""
txtTexto2 = ""
txtTexto3 = ""
txtTexto4 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = ""
txtOS = ""
txtResponsavel.Text = pubUsuario
txtData.Text = Format(Date, "dd/mm/yy")
cmbStatus = "Aberto"
txtOrdem = ""
txtdesenho = ""
Txt_codref = ""
txtdescricao = ""
txtQtde = ""
txtCliente = ""
Txt_ID_RNC = 0
txtRNC = ""
txtFase = ""
txtGrupo = ""
txtmaquina = ""
txtQtdeNC = ""
txtObs = ""
txtNRef = ""
CodigoLista = 0
Caption = "Qualidade - Solicitação de desvio"

If RNC_Nao_Conformidade = False Then
    With txtQtdeNC
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    With cmdOS
        .Enabled = True
        .TabStop = True
    End With
Else
    With txtQtdeNC
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With cmdOS
        .Enabled = False
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados2()
On Error GoTo tratar_erro

TBGravar!Texto1 = Trim(txtTexto1)
TBGravar!Texto2 = Trim(txtTexto2)
TBGravar!Texto3 = Trim(txtTexto3)
TBGravar!texto4 = Trim(txtTexto4)
TBGravar!Responsavel_Outros = txtResponsavel_outros
TBGravar!Funcao = txtFuncao
TBGravar!Data_Outros = txtData_outros
TBGravar!Documento = txtDocumento
TBGravar!Desvio = cmbDesvio

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_outros()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_SD WHERE ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    ProcEnviaDados2
    TBGravar.Update
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/solicitação de desvio"
    Evento = "Alterar outros"
    ID_documento = txtId
    Documento = "Nº SD: " & txtId
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_outros()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir este registro?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from CQ_SD where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        TBGravar!Texto1 = Null
        TBGravar!Texto2 = Null
        TBGravar!Texto3 = Null
        TBGravar!texto4 = Null
        TBGravar!Responsavel_Outros = Null
        TBGravar!Funcao = Null
        TBGravar!Data_Outros = Null
        TBGravar!Documento = Null
        TBGravar!Desvio = Null
        TBGravar.Update
    End If
    USMsgBox ("Registro excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/solicitação de desvio"
    Evento = "Excluir outros"
    ID_documento = txtId
    Documento = "Nº SD: " & txtId
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos2
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 11, True
ProcCarregaToolBar2 Me, 15195, 9, True
Formulario = "Qualidade/solicitação de desvio"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais

If RNC_Nao_Conformidade = False Then
    With USToolBar1
        .ButtonState(1) = 0
        .Refresh
    End With
Else
    With USToolBar1
        .ButtonState(1) = 5
        .Refresh
    End With
    'Carrega dados da SD vinculada com uma nao conformidade
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CQ_SD where ID = " & IIf(frmcqnc.txtID_SD = "", 0, frmcqnc.txtID_SD), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcLimpaCampos
        ProcPuxaDados
    Else
        cmbStatus = "Aberto"
        With frmcqnc
            txtOS = .cmbOS
            txtQtdeNC = .txtnc
        End With
        
        ProcNovo
        ProcSalvar
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_SD where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then ProcPuxaDados
        TBAbrir.Close
    End If
End If

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/solicitação de desvio"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmCQ_SD_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) solicitação(ões) de desvio?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            '==================================
            Modulo = "Qualidade/solicitação de desvio"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº SD: " & .ListItems(InitFor)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select ID_RNC from CQ_SD WHERE id = " & .ListItems(InitFor) & " and ID_RNC IS NOT NULL and ID_RNC <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then Conexao.Execute "DELETE from CQ_RNC where ID = " & TBFI!ID_RNC
            TBFI.Close
            
            Conexao.Execute "DELETE from CQ_SD where id = " & .ListItems(InitFor)
            
            'Exclui SD na nao conformidade
            Conexao.Execute "Update CQ_NC_FABRICA Set ID_SD = Null where ID_SD = " & .ListItems(InitFor)
            If RNC_Nao_Conformidade = True Then frmcqnc.txtID_SD = ""
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) solicitação(ões) de desvio antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("solicitação(ões) de desvio excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    Novo_SD = False
    Frame1.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe a solicitação de desvio antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    ImgAbrir.SetFocus
    Exit Sub
End If
NomeRel = "CQ_SD.rpt"
ProcImprimirRel "{CQ_SD.id}= " & txtId, ""

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
If RNC_Nao_Conformidade = False Then
    ProcLimpaCampos
    Frame1.Enabled = True
End If
Novo_SD = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_SD = True Then
    If USMsgBox("A solicitação de desvio ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_SD = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_SD = False
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
Acao = "salvar"
If Frame1.Enabled = False And RNC_Nao_Conformidade = False Then
    ProcVerificaSalvar
    Exit Sub
ElseIf Frame1.Enabled = False And Novo_SD = False Then
    USMsgBox ("É necessário clicar novamente no botão de criar SD no módulo de Qualidade/Não conformidade."), vbExclamation, "CAPRIND v5.0"
    Unload Me
    Exit Sub
End If

If txtOS = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    cmdOS_Click
    Exit Sub
End If
If txtNRef <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    If Novo_SD = True Then
        TBAbrir.Open "select * from cq_sd where NRef = '" & txtNRef & "'", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBAbrir.Open "select * from cq_sd where NRef = '" & txtNRef & "' and ID <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBAbrir.EOF = False Then
        USMsgBox ("Este número de referência está sendo utilizado, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtNRef.SetFocus
        Exit Sub
    End If
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_SD where ID = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
End If
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
If Novo_SD = True Then
    USMsgBox ("Nova solicitação de desvio cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_CQ_SD_Localizar = "Select * from CQ_SD where ID = " & txtId
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If ListView1.ListItems.Count <> 0 And CodigoLista <> 0 Then
        ListView1.SelectedItem = ListView1.ListItems(CodigoLista)
        ListView1.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/solicitação de desvio"
ID_documento = txtId.Text
Documento = "Nº SD: " & txtId
Documento1 = ""
ProcGravaEvento
'==================================
Novo_SD = False

'Grava SD na nao conformidade
If RNC_Nao_Conformidade = True Then
    With frmcqnc
        Conexao.Execute "Update CQ_NC_FABRICA Set ID_SD = " & txtId & " where Codigo = " & IIf(.txtidos = "", 0, .txtidos)
        .txtID_SD = txtId
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListView1, ColumnHeader
End If
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_SD where id = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = ListView1.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        ListView1.SetFocus
    Case 1:
        If Novo_SD = True Then
            SSTab1.Tab = 0
            USMsgBox ("Salve a solicitação de desvio antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        txtResponsavel_outros.SetFocus
        ProcLimpaCampos2
        ProcPuxadados2
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

If txtdesenho <> "" Then
    txtdescricao = ""
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto WHERE desenho = '" & txtdesenho.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    End If
    TBProduto.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procbloqueia()
On Error GoTo tratar_erro

Framelista.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcDesbloqueia()
On Error GoTo tratar_erro

Framelista.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
Set TBLISTA_CQ_SD = CreateObject("adodb.recordset")
TBLISTA_CQ_SD.Open StrSql_CQ_SD_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_CQ_SD.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Codproduto = 0
Dataini = 0
TotContas = 0
ListView1.ListItems.Clear
TBLISTA_CQ_SD.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_CQ_SD.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_CQ_SD.PageSize
ContadorReg = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_CQ_SD.RecordCount - IIf(Pagina > 1, (TBLISTA_CQ_SD.PageSize * (Pagina - 1)), 0), TBLISTA_CQ_SD.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_CQ_SD.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLISTA_CQ_SD!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_CQ_SD!status), "", TBLISTA_CQ_SD!status)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_CQ_SD!NRef), "", TBLISTA_CQ_SD!NRef)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_CQ_SD!Data), "", Format(TBLISTA_CQ_SD!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_CQ_SD!Responsavel), "", TBLISTA_CQ_SD!Responsavel)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_CQ_SD!Ordem), "", TBLISTA_CQ_SD!Ordem)
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select producao.desenho,producao.produto,producao.quant from Producao INNER JOIN ordemservico on producao.Ordem = ordemservico.Ordem where producao.Ordem = " & TBLISTA_CQ_SD!Ordem, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            .Item(.Count).SubItems(6) = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBItem!Produto), "", TBItem!Produto)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBItem!Quant), "0,0000", Format(TBItem!Quant, "###,##0.0000"))
        End If
        TBItem.Close
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_CQ_SD!qtdeNC), "0,0000", Format(TBLISTA_CQ_SD!qtdeNC, "###,##0.0000"))
    End With
    TBLISTA_CQ_SD.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_CQ_SD.RecordCount
If TBLISTA_CQ_SD.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_CQ_SD.PageCount
ElseIf TBLISTA_CQ_SD.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_CQ_SD.PageCount & " de: " & TBLISTA_CQ_SD.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_CQ_SD.AbsolutePage - 1 & " de: " & TBLISTA_CQ_SD.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!Ordem = IIf(txtOrdem = "", Null, txtOrdem)
TBGravar!OS = txtOS
TBGravar!ID_RNC = Txt_ID_RNC
TBGravar!Obs = txtObs
TBGravar!qtdeNC = IIf(txtQtdeNC = "", 0, txtQtdeNC)
TBGravar!NRef = txtNRef
TBGravar!status = cmbStatus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados2()
On Error GoTo tratar_erro

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from CQ_SD where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtTexto1 = IIf(IsNull(TBFornecedor!Texto1), "", TBFornecedor!Texto1)
    txtTexto2 = IIf(IsNull(TBFornecedor!Texto2), "", TBFornecedor!Texto2)
    txtTexto3 = IIf(IsNull(TBFornecedor!Texto3), "", TBFornecedor!Texto3)
    txtTexto4 = IIf(IsNull(TBFornecedor!texto4), "", TBFornecedor!texto4)
    txtResponsavel_outros = IIf(IsNull(TBFornecedor!Responsavel_Outros), "", TBFornecedor!Responsavel_Outros)
    txtFuncao = IIf(IsNull(TBFornecedor!Funcao), "", TBFornecedor!Funcao)
    txtData_outros = IIf(IsNull(TBFornecedor!Data_Outros), Date, TBFornecedor!Data_Outros)
    txtDocumento = IIf(IsNull(TBFornecedor!Documento), "", TBFornecedor!Documento)
    If IsNull(TBFornecedor!Desvio) = False And TBFornecedor!Desvio <> "" Then cmbDesvio = TBFornecedor!Desvio
End If
TBFornecedor.Close

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

Private Sub txtOS_Change()
On Error GoTo tratar_erro

If txtOS = "" Then Exit Sub
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select ordemservico.Ordem, ordemservico.idfase, ordemservico.fase, ordemservico.maquina, producao.desenho, producao.N_Referencia, producao.produto, producao.quant, producao.cliente from Producao INNER JOIN ordemservico on producao.Ordem = ordemservico.Ordem where Ordemservico.IDProducao = " & txtOS, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    txtOrdem = IIf(IsNull(TBOrdem!Ordem), "", TBOrdem!Ordem)
    txtdesenho = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
    Txt_codref = IIf(IsNull(TBOrdem!N_referencia), "", TBOrdem!N_referencia)
    txtdescricao = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
    txtQtde = IIf(IsNull(TBOrdem!Quant), "", Format(TBOrdem!Quant, "###,##0.0000"))
    txtCliente = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
    txtFase = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
    txtmaquina = IIf(IsNull(TBOrdem!maquina), "", TBOrdem!maquina)
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select grupo_op from Fases where idfase = " & TBOrdem!IDFase, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        txtGrupo = IIf(IsNull(TBFases!Grupo_op), "", TBFases!Grupo_op)
    End If
    TBFases.Close
End If
TBOrdem.Close

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

Private Sub txtQtdeNC_LostFocus()
On Error GoTo tratar_erro

If txtQtdeNC.Text <> "" Then
    VerifNumero = txtQtdeNC.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtdeNC.Text = ""
        txtQtdeNC.SetFocus
        Exit Sub
    End If
    txtQtdeNC = Format(txtQtdeNC, "###,##0.0000")
End If

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
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    'Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procSalvar_outros
    Case 2: procExcluir_outros
    Case 3: ProcImprimir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
