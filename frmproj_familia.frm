VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmproj_familia 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Famílias"
   ClientHeight    =   10035
   ClientLeft      =   225
   ClientTop       =   1425
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
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
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   58
      Top             =   9120
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
         TabIndex        =   29
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
         Left            =   2730
         TabIndex        =   27
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
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
         ItemData        =   "frmproj_familia.frx":0000
         Left            =   6990
         List            =   "frmproj_familia.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   187
         Width           =   1965
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   33
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_familia.frx":0022
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
         TabIndex        =   32
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_familia.frx":37C9
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
         TabIndex        =   30
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
         TabIndex        =   31
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_familia.frx":72D8
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
         TabIndex        =   34
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_familia.frx":B3C9
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
         Left            =   3360
         TabIndex        =   65
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
         TabIndex        =   62
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
         TabIndex        =   61
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label21 
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
         TabIndex        =   60
         Top             =   240
         Width           =   645
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
         TabIndex        =   59
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.TextBox txtId_familia 
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
      Left            =   1740
      MaxLength       =   100
      MouseIcon       =   "frmproj_familia.frx":EC57
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Text            =   "0"
      Top             =   6450
      Visible         =   0   'False
      Width           =   975
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
      Height          =   2055
      Left            =   60
      TabIndex        =   35
      Top             =   990
      Width           =   15195
      Begin VB.ComboBox cmbClassificacao_produto 
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
         ItemData        =   "frmproj_familia.frx":EF61
         Left            =   8250
         List            =   "frmproj_familia.frx":EF63
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   66
         ToolTipText     =   "Classificação do produto (tipo do item)."
         Top             =   1590
         Width           =   2745
      End
      Begin VB.CommandButton CmdCF 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14370
         Picture         =   "frmproj_familia.frx":EF65
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
         Top             =   990
         Width           =   315
      End
      Begin VB.TextBox Txt_ID_CF 
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
         Height          =   315
         Left            =   13440
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "ID CF."
         Top             =   990
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton Cmd_limpar_CF 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14700
         Picture         =   "frmproj_familia.frx":F067
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar CF."
         Top             =   990
         Width           =   315
      End
      Begin VB.TextBox txtLetra 
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
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   990
         Width           =   945
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   390
         Width           =   2025
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
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pela validação."
         Top             =   390
         Width           =   3735
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   1185
      End
      Begin VB.ComboBox Cmb_centro 
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
         Left            =   8085
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Centro de custo."
         Top             =   990
         Width           =   5355
      End
      Begin VB.CommandButton cmdgrupo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14685
         Picture         =   "frmproj_familia.frx":F1A5
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cadastrar/Localizar grupo."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtGrupo 
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
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Grupo."
         Top             =   390
         Width           =   3585
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   4255
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox txtFamilia 
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
         MaxLength       =   100
         TabIndex        =   8
         ToolTipText     =   "Família."
         Top             =   990
         Width           =   6930
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   1375
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   2865
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
         Left            =   180
         MaxLength       =   100
         TabIndex        =   13
         ToolTipText     =   "Descrição."
         Top             =   1590
         Width           =   8070
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aplicação*"
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
         Height          =   555
         Left            =   10995
         TabIndex        =   36
         Top             =   1380
         Width           =   4005
         Begin VB.CheckBox chkQualidade 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qualidade"
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
            Left            =   2880
            TabIndex        =   17
            Top             =   270
            Width           =   1035
         End
         Begin VB.CheckBox chkfabricacao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "PCP"
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
            TabIndex        =   16
            Top             =   270
            Width           =   585
         End
         Begin VB.CheckBox chkvendas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vendas"
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
            Left            =   1230
            TabIndex        =   15
            Top             =   270
            Width           =   855
         End
         Begin VB.CheckBox chkcompras 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Compras"
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
            TabIndex        =   14
            Top             =   270
            Width           =   945
         End
      End
      Begin VB.TextBox Txt_CF 
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
         Left            =   13440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Classificação fiscal."
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classificação (Bloco K)*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   71
         Left            =   8782
         TabIndex        =   67
         Top             =   1380
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM"
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
         Index           =   18
         Left            =   13732
         TabIndex        =   64
         Top             =   780
         Width           =   330
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
         Left            =   5445
         TabIndex        =   57
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
         Left            =   8190
         TabIndex        =   56
         Top             =   180
         Width           =   1980
      End
      Begin VB.Label Label2 
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
         Index           =   5
         Left            =   600
         TabIndex        =   55
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de custo"
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
         Left            =   10185
         TabIndex        =   50
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo"
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
         Left            =   12645
         TabIndex        =   43
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   4500
         TabIndex        =   41
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código*"
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
         Left            =   405
         TabIndex        =   40
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   2325
         TabIndex        =   39
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família*"
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
         Left            =   4320
         TabIndex        =   38
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição*"
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
         Left            =   3825
         TabIndex        =   37
         Top             =   1380
         Width           =   780
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   45
      Top             =   9750
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
   Begin MSComctlLib.ListView Lista 
      Height          =   4305
      Left            =   60
      TabIndex        =   26
      Top             =   4800
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7594
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   8206
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   8206
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Validada"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   44
      Top             =   0
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   33
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
      ButtonLeft2     =   37
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   75
      ButtonTop3      =   2
      ButtonWidth3    =   38
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   115
      ButtonTop4      =   2
      ButtonWidth4    =   39
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   156
      ButtonTop5      =   2
      ButtonWidth5    =   51
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   209
      ButtonTop6      =   2
      ButtonWidth6    =   47
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   258
      ButtonTop7      =   2
      ButtonWidth7    =   46
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Status"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Status (F7)"
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   306
      ButtonTop8      =   2
      ButtonWidth8    =   39
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Descrição padrão"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Cadastrar descrição padrão (F8)"
      ButtonKey9      =   "9"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft9     =   347
      ButtonTop9      =   2
      ButtonWidth9    =   91
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "De, para"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Atualizar de, para (F9)"
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
      ButtonLeft10    =   440
      ButtonTop10     =   2
      ButtonWidth10   =   50
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Validação"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Validar/Cancelar validação (F10)"
      ButtonKey11     =   "11"
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
      ButtonLeft11    =   492
      ButtonTop11     =   2
      ButtonWidth11   =   53
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
      ButtonLeft12    =   547
      ButtonTop12     =   4
      ButtonWidth12   =   2
      ButtonHeight12  =   54
      ButtonCaption13 =   "Ajuda"
      ButtonEnabled13 =   0   'False
      ButtonIconSize13=   32
      ButtonToolTipText13=   "Ajuda (F1)"
      ButtonKey13     =   "13"
      ButtonAlignment13=   2
      BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft13    =   551
      ButtonTop13     =   2
      ButtonWidth13   =   36
      ButtonHeight13  =   21
      ButtonUseMaskColor13=   0   'False
      ButtonCaption14 =   "Sair"
      ButtonEnabled14 =   0   'False
      ButtonIconSize14=   32
      ButtonToolTipText14=   "Sair (Esc)"
      ButtonKey14     =   "14"
      ButtonAlignment14=   2
      BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft14    =   589
      ButtonTop14     =   2
      ButtonWidth14   =   26
      ButtonHeight14  =   21
      ButtonUseMaskColor14=   0   'False
      ButtonEnabled15 =   0   'False
      ButtonIconSize15=   32
      ButtonKey15     =   "15"
      ButtonAlignment15=   2
      BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState15   =   5
      ButtonLeft15    =   617
      ButtonTop15     =   2
      ButtonWidth15   =   24
      ButtonHeight15  =   24
      ButtonUseMaskColor15=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11250
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmproj_familia.frx":F2A7
         Count           =   1
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vendas"
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
      Left            =   60
      TabIndex        =   51
      Top             =   3930
      Width           =   15195
      Begin VB.CommandButton Cmd_localizar_PC1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14340
         Picture         =   "frmproj_familia.frx":17759
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Localizar plano de contas."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_ID_PC1 
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
         TabIndex        =   52
         Text            =   "0"
         ToolTipText     =   "ID PC."
         Top             =   390
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Txt_descricao_PC1 
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
         MaxLength       =   255
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   12255
      End
      Begin VB.CommandButton Cmd_limpar_PC1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14670
         Picture         =   "frmproj_familia.frx":1785B
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Limpar conta contábil."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_codigo_PC1 
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
         MaxLength       =   255
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   1875
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   862
         TabIndex        =   54
         Top             =   180
         Width           =   510
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Index           =   7
         Left            =   7830
         TabIndex        =   53
         Top             =   180
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Compras"
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
      Left            =   60
      TabIndex        =   46
      Top             =   3060
      Width           =   15195
      Begin VB.CommandButton Cmd_limpar_PC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14670
         Picture         =   "frmproj_familia.frx":17999
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar conta contábil."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_descricao_PC 
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
         MaxLength       =   255
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   12255
      End
      Begin VB.TextBox Txt_ID_PC 
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
         TabIndex        =   47
         Text            =   "0"
         ToolTipText     =   "ID PC."
         Top             =   390
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton Cmd_localizar_PC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14340
         Picture         =   "frmproj_familia.frx":17AD7
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Localizar plano de contas."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_codigo_PC 
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
         MaxLength       =   255
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   1875
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Index           =   4
         Left            =   7837
         TabIndex        =   49
         Top             =   180
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   862
         TabIndex        =   48
         Top             =   180
         Width           =   510
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmproj_familia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Familia_Produto        As Boolean 'OK
Public Sql_Familia_Localizar    As String 'OK
Public FormulaRel_Familia       As String 'OK
Dim TBLISTA_Familia As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then FunAbrirVideoWeb ("http://www.youtube.com/user/procamonline?feature=watch")
'If Compras_Familia = True Then
If Vendas_Familia = True Then FunAbrirVideoWeb ("http://www.youtube.com/watch?v=P6Pvl64Z8zE&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=45&feature=plcp")
If Qualidade_Familia = True Then FunAbrirVideoWeb ("http://www.youtube.com/watch?v=Fwoxxf6mV_Y&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=44&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmproj_familia_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtid_familia = 0 Then Exit Sub
If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then
    TextoFiltro = ""
Else
    If Compras_Familia = True Then TextoFiltro = "where Compras = 'True'"
    If Vendas_Familia = True Then TextoFiltro = "where Vendas = 'True'"
    If Qualidade_Familia = True Then TextoFiltro = "where Qualidade = 'True'"
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from projfamilia " & TextoFiltro & " order by Familia", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("codigo = " & txtid_familia)
    TBAbrir.MovePrevious
    If TBAbrir.BOF = False Then
        txtid_familia = TBAbrir!CODIGO
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from projfamilia where Codigo = " & txtid_familia, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
    Else
        USMsgBox ("Fim dos cadastros de famílias."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloquear()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtLetra.Text = "" Then
    NomeCampo = "a família"
    Acao = "bloquear/desbloquear"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Familia_Produto = True Then
    USMsgBox ("Salve a família antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    imgSalvar.SetFocus
    Exit Sub
End If
frmproj_familia_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkcompras_Click()
On Error GoTo tratar_erro

ProcMostraEscondePC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkFabricacao_Click()
On Error GoTo tratar_erro

ProcMostraEscondePC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkQualidade_Click()
On Error GoTo tratar_erro

ProcMostraEscondePC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVendas_Click()
On Error GoTo tratar_erro

ProcMostraEscondePC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostraEscondePC()
On Error GoTo tratar_erro

If chkCompras.Value = 1 And chkVendas.Value = 1 Then
    Frame8.Enabled = True
    Frame11.Enabled = True
ElseIf chkCompras.Value = 1 Or chkVendas.Value = 1 Then
        If chkCompras.Value = 1 Then
            Frame8.Enabled = True
            Frame11.Enabled = False
            ProcLimpaCamposPC False, True
        Else
            Frame8.Enabled = False
            Frame11.Enabled = True
            ProcLimpaCamposPC True, False
        End If
    Else
        Frame8.Enabled = False
        Frame11.Enabled = False
        ProcLimpaCamposPC True, True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposPC(LimparCompras As Boolean, LimparVendas As Boolean)
On Error GoTo tratar_erro

If LimparCompras = True Then
    Txt_ID_PC = 0
    Txt_codigo_PC = ""
    Txt_descricao_PC = ""
End If
If LimparVendas = True Then
    Txt_ID_PC1 = 0
    Txt_codigo_PC1 = ""
    Txt_descricao_PC1 = ""
End If

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
        .ButtonState(11) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(11) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_PC_Click()
On Error GoTo tratar_erro

ProcLimpaCamposPC True, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_PC1_Click()
On Error GoTo tratar_erro

ProcLimpaCamposPC False, True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro

Plano_contas_produtos = False
Plano_contas_familias = True
Plano_centro_de_custo = False
Plano_instituicao = False
Plano_opcoesgerais = False
Plano_Faturamento = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
Plano_PCP = False
Aplic = 1
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC1_Click()
On Error GoTo tratar_erro

Plano_contas_produtos = False
Plano_contas_familias = True
Plano_centro_de_custo = False
Plano_instituicao = False
Plano_opcoesgerais = False
Plano_Faturamento = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
Plano_PCP = False
Aplic = 2
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdCF_Click()
On Error GoTo tratar_erro

Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Clientes = False
Compras_Pedido = False
Familia_NCM = True
ClassFiscal = False
frmProj_Classificacao_Fiscal.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdGrupo_Click()
On Error GoTo tratar_erro

frmproj_familia_grupo.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtid_familia = 0 Then Exit Sub
If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then
    TextoFiltro = ""
Else
    If Compras_Familia = True Then TextoFiltro = "where Compras = 'True'"
    If Vendas_Familia = True Then TextoFiltro = "where Vendas = 'True'"
    If Qualidade_Familia = True Then TextoFiltro = "where Qualidade = 'True'"
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from projfamilia " & TextoFiltro & " order by Familia", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.Find ("codigo = " & txtid_familia)
    TBAbrir.MoveNext
    If TBAbrir.EOF = False Then
        txtid_familia = TBAbrir!CODIGO
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from projfamilia where Codigo = " & txtid_familia, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
    Else
        USMsgBox ("Fim dos cadastros de famílias."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Familia.AbsolutePage <> 2 Then
    If TBLISTA_Familia.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Familia.PageCount - 1)
    Else
        TBLISTA_Familia.AbsolutePage = TBLISTA_Familia.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Familia.AbsolutePage)
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
    TBLISTA_Familia.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Familia.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Familia.AbsolutePage = 1
ProcExibePagina (TBLISTA_Familia.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Familia.AbsolutePage <> -3 Then
    If TBLISTA_Familia.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Familia.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Familia.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Familia.AbsolutePage = TBLISTA_Familia.PageCount
ProcExibePagina (TBLISTA_Familia.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcLocalizar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcBloquear
    Case vbKeyF8: ProcDescricao
    Case vbKeyF9: ProcDePara
    Case vbKeyF10:
        If Cmb_opcao_lista = "Validação" Then
            If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then Formulario = "Engenharia/Famílias"
            If Vendas_Familia = True Then Formulario = "Vendas/Famílias"
            If Compras_Familia = True Then Formulario = "Compras/Famílias"
            If Qualidade_Familia = True Then Formulario = "Qualidade/Famílias"
            ProcValidarRegistros Lista, Formulario
        End If
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar()
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
If txtLetra.Text = "" Then
    NomeCampo = "o código"
    ProcVerificaAcao
    txtLetra.SetFocus
    Exit Sub
End If

'If cmbClassificacao_produto.Text = "" Then
'    NomeCampo = "Classificação"
'    ProcVerificaAcao
'    cmbClassificacao_produto.SetFocus
'    Exit Sub
'End If

If txtfamilia.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    txtfamilia.SetFocus
    Exit Sub
End If
If txtdescricao.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
If chkVendas.Value = 0 And chkCompras.Value = 0 And chkFabricacao.Value = 0 And chkQualidade.Value = 0 Then
    NomeCampo = "a aplicação"
    ProcVerificaAcao
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Codigo from projfamilia where letra = '" & txtLetra.Text & "' and Codigo <> " & txtid_familia, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido utilizar este código, pois o mesmo já foi cadastrado."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Codigo from projfamilia where Familia = '" & txtfamilia & "' and Codigo <> " & txtid_familia, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido utilizar esta família, pois a mesma já foi cadastrada."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If

Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select * from projfamilia where Codigo = " & txtid_familia, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = True Then
    TBMaterial.AddNew
Else
    If FunVerificaRegistroValidado("projfamilia", "Codigo = " & txtid_familia, "mesma", "a família", "alterar", False, True) = False Then Exit Sub
    If txtfamilia <> TBMaterial!Familia Then
        Conexao.Execute "Update CFI Set familia = '" & txtfamilia & "' where familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update Compras_fornecedores_familia Set familia = '" & txtfamilia & "' where familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update Compras_pedido_lista Set familia = '" & txtfamilia & "' where familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update Estoque_Controle Set classe = '" & txtfamilia & "' where classe = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update Instrumentos Set Familia = '" & txtfamilia & "' where Familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update Planodimensao_instrumentos Set Familia = '" & txtfamilia & "' where Familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update projproduto Set classe = '" & txtfamilia & "' where classe = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update Requisicao_materiais_lista Set Familia = '" & txtfamilia & "' where Familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update tbl_Detalhes_Nota Set familia = '" & txtfamilia & "' where familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update vendas_carteira Set familia = '" & txtfamilia & "' where familia = '" & TBMaterial!Familia & "'"
        Conexao.Execute "Update Vendas_analise Set familia = '" & txtfamilia & "' where familia = '" & TBMaterial!Familia & "'"
        
        If chkCompras.Value = 1 Then ComprasTexto = "Compras = 'True'" Else ComprasTexto = "Compras = 'False'"
        If chkVendas.Value = 1 Then VendasTexto = "Vendas = 'True'" Else VendasTexto = "Vendas = 'False'"
        If chkFabricacao.Value = 1 Then PCPTexto = "Producao = 'True'" Else PCPTexto = "Producao = 'False'"
        If chkQualidade.Value = 1 Then QualidadeTexto = "Qualidade = 'True'" Else QualidadeTexto = "Qualidade = 'False'"
        Conexao.Execute "Update projproduto Set " & ComprasTexto & ", " & VendasTexto & ", " & PCPTexto & ", " & QualidadeTexto & " where classe = '" & TBMaterial!Familia & "'"
    End If
    
    If Cmb_centro <> "" Then
        If Cmb_centro.ItemData(Cmb_centro.ListIndex) <> IIf(IsNull(TBMaterial!ID_CC), 0, TBMaterial!ID_CC) Then
            If USMsgBox("Deseja atualizar o centro de custo em todos os produtos/serviços cadastrados nesta família?", vbYesNo, "CAPRIND v5.0") = vbYes Then Conexao.Execute "Update projproduto Set ID_CC = " & Cmb_centro.ItemData(Cmb_centro.ListIndex) & " Where Classe = '" & TBMaterial!Familia & "'"
        End If
    Else
        If IsNull(TBMaterial!ID_CC) = False And TBMaterial!ID_CC <> 0 Then
            If USMsgBox("Deseja atualizar o centro de custo em todos os produtos/serviços cadastrados nesta família?", vbYesNo, "CAPRIND v5.0") = vbYes Then Conexao.Execute "Update projproduto Set ID_CC = 0 Where Classe = '" & TBMaterial!Familia & "'"
        End If
    End If
    
    If Txt_ID_PC <> TBMaterial!ID_PC Or Txt_ID_PC1 <> TBMaterial!ID_PC1 Then
        If USMsgBox("Deseja atualizar a(s) conta(s) contábil(eis) em todos os registros não validados cadastrados nesta família?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Conexao.Execute "Update projproduto Set ID_PC = " & IIf(Frame8.Visible = True, Txt_ID_PC, 0) & ", ID_PC1 = " & IIf(Frame11.Visible = True, Txt_ID_PC1, 0) & " where Classe = '" & TBMaterial!Familia & "' and DtValidacao IS NULL"
            If Txt_ID_PC1 <> TBMaterial!ID_PC1 Then Conexao.Execute "Update CCR set CCR.ID_PC = " & IIf(Frame11.Visible = True, Txt_ID_PC1, 0) & " from (CC_realizado CCR INNER JOIN Projproduto P ON P.Codproduto = CCR.Cod_produto) INNER JOIN Projfamilia PF ON PF.Familia = P.Classe where P.Classe = '" & TBMaterial!Familia & "' and P.DtValidacao IS NULL"
        End If
    End If
End If
TBMaterial!Data = IIf(txtData = "", Date, txtData)
TBMaterial!Responsavel = IIf(txtResponsavel = "", pubUsuario, txtResponsavel)
TBMaterial!Grupo = txtGrupo.Text
TBMaterial!Letra = txtLetra.Text
TBMaterial!Familia = txtfamilia.Text
If Cmb_centro <> "" Then TBMaterial!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex) Else TBMaterial!ID_CC = 0
TBMaterial!Descricao = txtdescricao.Text
TBMaterial!ID_CF = IIf(Txt_ID_CF = "", 0, Txt_ID_CF)
If chkCompras.Value = 1 Then TBMaterial!Compras = True Else TBMaterial!Compras = False
If chkVendas.Value = 1 Then TBMaterial!Vendas = True Else TBMaterial!Vendas = False
If chkFabricacao.Value = 1 Then TBMaterial!Fabricacao = True Else TBMaterial!Fabricacao = False
If chkQualidade.Value = 1 Then TBMaterial!Qualidade = True Else TBMaterial!Qualidade = False
'===================================================
If cmbClassificacao_produto <> "" Then TBMaterial!ID_Tipo = cmbClassificacao_produto.ItemData(cmbClassificacao_produto.ListIndex) Else TBMaterial!ID_Tipo = Null

'===================================================
' Acerta o cadastro de todos os produtos da Familia
'===================================================
If cmbClassificacao_produto.Text <> "" Then
Var = cmbClassificacao_produto.ItemData(cmbClassificacao_produto.ListIndex)
Conexao.Execute "update projproduto set ID_Tipo = " & Int(Var) & " where Classe = '" & txtdescricao & "'"
'===================================================
' Acerta a movimentação de estoque
'===================================================
Conexao.Execute "update Estoque_movimentacao Set ID_Tipo = " & Int(Var) & " from Estoque_movimentacao Where Familia = '" & txtdescricao & "'"
'===================================================
End If

TBMaterial!ID_PC = Txt_ID_PC 'Compras
TBMaterial!ID_PC1 = Txt_ID_PC1 'Vendas
TBMaterial.Update
txtid_familia = TBMaterial!CODIGO
TBMaterial.Close
Lista.ListItems.Clear
If Novo_Familia_Produto = True Then
    USMsgBox ("Nova família cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Familia_Localizar = "Select * from projfamilia where Letra = '" & txtLetra.Text & "'"
    ProcAtualizalista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista1 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista1)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = Formulario
ID_documento = txtid_familia
Documento = "Código: " & txtLetra.Text
Documento1 = ""
ProcGravaEvento
'==================================
With txtLetra
    .Locked = True
    .TabStop = False
End With
Novo_Familia_Produto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If Sql_Familia_Localizar = "" Then Exit Sub
Set TBLISTA_Familia = CreateObject("adodb.recordset")
TBLISTA_Familia.Open Sql_Familia_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Familia.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Familia.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Familia.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Familia.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Familia.RecordCount - IIf(Pagina > 1, (TBLISTA_Familia.PageSize * (Pagina - 1)), 0), TBLISTA_Familia.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Familia.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Familia!CODIGO
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Familia!Data), "", Format(TBLISTA_Familia!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Familia!Responsavel), "", TBLISTA_Familia!Responsavel)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Familia!Letra), "", TBLISTA_Familia!Letra)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Familia!Familia), "", TBLISTA_Familia!Familia)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Familia!Descricao), "", TBLISTA_Familia!Descricao)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Familia!DtValidacao) = False, "Sim", "Não")
    End With
    TBLISTA_Familia.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Familia.RecordCount
If TBLISTA_Familia.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Familia.PageCount
ElseIf TBLISTA_Familia.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Familia.PageCount & " de: " & TBLISTA_Familia.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Familia.AbsolutePage - 1 & " de: " & TBLISTA_Familia.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtid_familia = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
txtStatus = "Liberado"
txtGrupo = ""
txtLetra.Text = ""
txtfamilia.Text = ""
ProcCarregaComboSetor Cmb_centro, "Setor IS NOT NULL and DtBloq IS NULL and (Consolidacao = 'False' or Consolidacao is null)", "", False, True, False, "", True, False
txtdescricao.Text = ""
Txt_ID_CF = ""
Txt_CF = ""

If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then
    chkCompras.Value = 0
    chkVendas.Value = 0
    chkFabricacao.Value = 0
    chkQualidade.Value = 0
End If
If Compras_Familia = True Then
    chkFabricacao.Value = 0
    chkQualidade.Value = 0
    chkQualidade.Enabled = True
End If
If Vendas_Familia = True Then
    chkCompras.Value = 0
    chkFabricacao.Value = 0
End If
If Qualidade_Familia = True Then
    chkCompras.Value = 0
    chkFabricacao.Value = 0
End If

ProcLimpaCamposPC True, True
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then
    Caption = "Engenharia - Famílias"
    Formulario = "Engenharia/Famílias"
End If
If Compras_Familia = True Then
    Caption = "Compras - Famílias"
    Formulario = "Compras/Famílias"
End If
If Vendas_Familia = True Then
    Caption = "Vendas - Famílias"
    Formulario = "Vendas/Famílias"
End If
If Qualidade_Familia = True Then
    Caption = "Qualidade - Famílias"
    Formulario = "Qualidade/Famílias"
End If
Direitos
ProcLimpaVariaveisPrincipais
Formulario_familia = Formulario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
NomeRel = "Engenharia_familia.rpt"
ProcImprimirRel FormulaRel_Familia, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDescricao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "cadastrar a descrição padrão"
If txtLetra.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Familia_Produto = True Then
    USMsgBox ("Salve a família antes de cadastrar as descrições padrão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmproj_familia_descpadrao.Show 1

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
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) família(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Letra from Projfamilia WHERE Codigo = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = Formulario
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Código: " & TBFI!Letra
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE FROM Projfamilia WHERE Codigo = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) família(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Família(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Lista.ListItems.Clear
    ProcAtualizalista (1)
    Frame2.Enabled = False
    Novo_Familia_Produto = False
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
Novo_Familia_Produto = True
Frame2.Enabled = True
With txtLetra
    .Locked = False
    .TabStop = True
End With
cmdGrupo.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Familia_Produto = True Then
    If USMsgBox("A família ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        If Novo_Familia_Produto = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Familia_Produto = False
Unload Me
Formulario_familia = ""

Exit Sub
tratar_erro:
    USMsgBox ("Decrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

If Compras_Familia = True Then
    Caption = "Compras - Famílias (Código : " & IIf(IsNull(TBFamilia!Letra), "", TBFamilia!Letra) & ")"
ElseIf Vendas_Familia = True Then
        Caption = "Vendas - Famílias (Código : " & IIf(IsNull(TBFamilia!Letra), "", TBFamilia!Letra) & ")"
    ElseIf Qualidade_Familia = True Then
            Caption = "Qualidade - Famílias (Código : " & IIf(IsNull(TBFamilia!Letra), "", TBFamilia!Letra) & ")"
        Else
            Caption = "Engenharia - Famílias (Código : " & IIf(IsNull(TBFamilia!Letra), "", TBFamilia!Letra) & ")"
End If

txtid_familia = TBFamilia!CODIGO
txtData = IIf(IsNull(TBFamilia!Data), "", Format(TBFamilia!Data, "dd/mm/yy"))
txtResponsavel.Text = IIf(IsNull(TBFamilia!Responsavel), "", TBFamilia!Responsavel)
txtStatus = IIf(TBFamilia!Bloqueado = True, "Bloqueado", "Liberado")
txtDtValidacao = IIf(IsNull(TBFamilia!DtValidacao), "", TBFamilia!DtValidacao)
txtRespValidacao = IIf(IsNull(TBFamilia!RespValidacao), "", TBFamilia!RespValidacao)
txtGrupo.Text = IIf(IsNull(TBFamilia!Grupo), "", TBFamilia!Grupo)

With txtLetra
    .Text = IIf(IsNull(TBFamilia!Letra), "", TBFamilia!Letra)
    .Locked = True
    .TabStop = False
End With

txtfamilia.Text = IIf(IsNull(TBFamilia!Familia), "", TBFamilia!Familia)
txtdescricao.Text = IIf(IsNull(TBFamilia!Descricao), "", TBFamilia!Descricao)

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Idclass, IDIntClasse  from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBFamilia!ID_CF), 0, TBFamilia!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Txt_ID_CF = TBFI!Idclass
    Txt_CF = IIf(IsNull(TBFI!IDIntClasse), "", TBFI!IDIntClasse)
End If

'=========================================================
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from projproduto_Tipo where ID = " & IIf(IsNull(TBFamilia!ID_Tipo), 0, TBFamilia!ID_Tipo), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then cmbClassificacao_produto = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Descricao), "", TBFI!Descricao)
TBFI.Close
'==========================================================

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Usuarios_setor.* from Usuarios_setor where ID = " & IIf(IsNull(TBFamilia!ID_CC), 0, TBFamilia!ID_CC), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
        If IsNull(TBFI!DtBloq) = False Then
            Cmb_centro.AddItem TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
        End If
        Cmb_centro = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    Else
        If IsNull(TBFI!DtBloq) = False Then
            Cmb_centro.AddItem IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
        End If
        Cmb_centro = IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    End If
End If
TBFI.Close

If TBFamilia!Compras = True Then chkCompras.Value = 1 Else chkCompras.Value = 0
If TBFamilia!Vendas = True Then chkVendas.Value = 1 Else chkVendas.Value = 0
If TBFamilia!Fabricacao = True Then chkFabricacao.Value = 1 Else chkFabricacao.Value = 0
If TBFamilia!Qualidade = True Then chkQualidade.Value = 1 Else chkQualidade.Value = 0

Txt_ID_PC = IIf(IsNull(TBFamilia!ID_PC), 0, TBFamilia!ID_PC)
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM tbl_familia where int_codfamilia = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_codigo_PC = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
    Txt_descricao_PC = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
End If

Txt_ID_PC1 = IIf(IsNull(TBFamilia!ID_PC1), 0, TBFamilia!ID_PC1)
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM tbl_familia where int_codfamilia = " & Txt_ID_PC1, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_codigo_PC1 = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
    Txt_descricao_PC1 = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
End If
TBAbrir.Close

Frame2.Enabled = True
Novo_Familia_Produto = False

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
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidadoSemMsg("projfamilia", "Codigo = " & .ListItems(InitFor), True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    
                    ProcVerificaRegistroUtilizadoSemMsg "CFI", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Compras_fornecedores_familia", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido_lista", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Estoque_Controle", "classe = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Instrumentos", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Planodimensao_instrumentos", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "projproduto", "Classe = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Requisicao_materiais_lista", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "vendas_carteira", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
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
                If FunVerificaRegistroValidado("projfamilia", "Codigo = " & .ListItems(InitFor), "mesma", "a família", "excluir", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                Mensagem = "Não é permitido excluir esta família, pois a mesma está sendo utilizada no módulo"
                ProcVerificaRegistroUtilizado "CFI", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Estoque/Almoxarifado"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Compras_fornecedores_familia", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Compras/Fornecedores ou Vendas/Clientes"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Compras_pedido_lista", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Compras"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Estoque_Controle", "classe = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Estoque/Movimentação"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Instrumentos", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Qualidade/Instrumentos"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Planodimensao_instrumentos", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Qualidade/Plano de inspeção"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "projproduto", "Classe = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Engenharia/Produtos e serviços"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Requisicao_materiais_lista", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Estoque/Requisição de materiais"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_Detalhes_Nota", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Faturamento/Nota fiscal"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "vendas_carteira", "familia = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Vendas"
                If Permitido = False Then .ListItems.Item(InitFor).Checked = False
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
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from projfamilia where Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista1 = Lista.SelectedItem.index
End If
TBFamilia.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 14, True

If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then
    Caption = "Engenharia - Famílias"
    Formulario = "Engenharia/Famílias"
End If
If Vendas_Familia = True Then
    Caption = "Vendas - Famílias"
    Formulario = "Vendas/Famílias"
    With chkVendas
        .Value = 1
        .Enabled = False
    End With
End If
If Compras_Familia = True Then
    Caption = "Compras - Famílias"
    Formulario = "Compras/Famílias"
    With chkCompras
        .Value = 1
        .Enabled = False
    End With
End If
If Qualidade_Familia = True Then
    Caption = "Qualidade - Famílias"
    Formulario = "Qualidade/Famílias"
    With chkQualidade
        .Value = 1
        .Enabled = False
    End With
End If
Direitos
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista = "Validação"

Formulario_familia = Formulario

With cmbClassificacao_produto
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Projproduto_Tipo order by codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!CODIGO & " - " & TBCarregarCombo!Descricao
            .ItemData(.NewIndex) = TBCarregarCombo!ID
            TBCarregarCombo.MoveNext
        Loop
    End If
    TBCarregarCombo.Close
End With

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDePara()
On Error GoTo tratar_erro

frmproj_familia_de_para.Show 1

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
    Case 8: ProcBloquear
    Case 9: ProcDescricao
    Case 10: ProcDePara
    Case 11:
        If Compras_Familia = False And Vendas_Familia = False And Qualidade_Familia = False Then Formulario = "Engenharia/Famílias"
        If Vendas_Familia = True Then Formulario = "Vendas/Famílias"
        If Compras_Familia = True Then Formulario = "Compras/Famílias"
        If Qualidade_Familia = True Then Formulario = "Qualidade/Famílias"
        ProcValidarRegistros Lista, Formulario
    Case 13: ProcAjuda
    Case 14: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
