VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmproj_conjunto 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Conjuntos"
   ClientHeight    =   10035
   ClientLeft      =   1740
   ClientTop       =   1425
   ClientWidth     =   15270
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1395
      Left            =   55
      TabIndex        =   49
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtComprimento 
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
         Left            =   9660
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   930
         Width           =   1425
      End
      Begin VB.TextBox txtLargura 
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
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   930
         Width           =   1425
      End
      Begin VB.CommandButton Cmd_localizar_prod 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         Picture         =   "frmproj_conjunto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Localizar produtos."
         Top             =   390
         Width           =   315
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
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pela validação."
         Top             =   930
         Width           =   2415
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
         Left            =   11100
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   930
         Width           =   1485
      End
      Begin VB.TextBox txtFamilia 
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   930
         Width           =   6045
      End
      Begin VB.TextBox txtTipo 
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
         MaxLength       =   255
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Tipo."
         Top             =   930
         Width           =   1965
      End
      Begin VB.TextBox txtref 
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
         Left            =   2550
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código de referência."
         Top             =   390
         Width           =   2115
      End
      Begin VB.TextBox txtdesenhoproduto 
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
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   1965
      End
      Begin VB.TextBox txtdescricaoproduto 
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   10335
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comprimento/mm"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   9750
         TabIndex        =   82
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Largura/mm"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   8505
         TabIndex        =   81
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável pela validação"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12817
         TabIndex        =   68
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data/hora validação"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   11115
         TabIndex        =   67
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4942
         TabIndex        =   66
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1012
         TabIndex        =   65
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
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
         Left            =   2842
         TabIndex        =   52
         Top             =   180
         Width           =   1530
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   9435
         TabIndex        =   51
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno*"
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
         Left            =   600
         TabIndex        =   50
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   76
      Top             =   2340
      Width           =   2715
      Begin VB.ComboBox cmbVersao_pesquisar 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmproj_conjunto.frx":0102
         Left            =   1950
         List            =   "frmproj_conjunto.frx":0104
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Versão."
         Top             =   90
         Width           =   795
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa por versão :"
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
         Left            =   60
         TabIndex        =   78
         Top             =   60
         Width           =   1800
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "                                                           "
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
      Height          =   1035
      Left            =   55
      TabIndex        =   73
      Top             =   2400
      Width           =   15195
      Begin VB.TextBox Txt_ID_desc_versao 
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
         Left            =   810
         TabIndex        =   74
         Text            =   "0"
         ToolTipText     =   "ID descrição da versão."
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox Txt_descricao_versao 
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
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Descrição da versão."
         Top             =   390
         Width           =   14355
      End
      Begin DrawSuite2022.USButton Cmd_salvar_desc_versao 
         Height          =   465
         Left            =   14550
         TabIndex        =   10
         ToolTipText     =   "Salvar descrição da versão (F6)."
         Top             =   390
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   820
         DibPicture      =   "frmproj_conjunto.frx":0106
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
         Caption         =   "Descrição da versão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6630
         TabIndex        =   75
         Top             =   180
         Width           =   1455
      End
   End
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
   Begin VB.TextBox Txt_cod_produto 
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "Código do produto."
      Top             =   7080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   59
      Top             =   9120
      Width           =   15195
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
         TabIndex        =   30
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
         TabIndex        =   31
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   35
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_conjunto.frx":0898
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
         TabIndex        =   34
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_conjunto.frx":403C
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
         TabIndex        =   32
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
         TabIndex        =   33
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_conjunto.frx":7B45
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
         TabIndex        =   36
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_conjunto.frx":BC34
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4410
         TabIndex        =   77
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3090
         TabIndex        =   62
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   61
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Txt_ID 
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
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "ID."
      Top             =   7050
      Visible         =   0   'False
      Width           =   435
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   56
      Top             =   0
      Width           =   15210
      _ExtentX        =   26829
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
      ButtonToolTipText2=   "Localizar conjunto (F2)"
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
      ButtonCaption6  =   "Copiar"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Copiar (F7)"
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
      ButtonWidth6    =   39
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Versão"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Criar versão (F8)"
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
      ButtonLeft7     =   250
      ButtonTop7      =   2
      ButtonWidth7    =   41
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   293
      ButtonTop8      =   2
      ButtonWidth8    =   53
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonAlignment9=   2
      ButtonType9     =   1
      ButtonStyle9    =   -1
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   -1
      ButtonLeft9     =   348
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft10    =   352
      ButtonTop10     =   2
      ButtonWidth10   =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft11    =   390
      ButtonTop11     =   2
      ButtonWidth11   =   26
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7230
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmproj_conjunto.frx":F4C0
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3495
      Left            =   60
      TabIndex        =   29
      Top             =   5595
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6165
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Pos."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   19941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Quant."
         Object.Width           =   2117
      EndProperty
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
      ForeColor       =   &H00000000&
      Height          =   2145
      Left            =   55
      TabIndex        =   39
      Top             =   3435
      Width           =   15195
      Begin VB.TextBox txtLargura2 
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
         Left            =   5822
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   990
         Width           =   1435
      End
      Begin VB.TextBox txtComprimento2 
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
         Left            =   7275
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   990
         Width           =   1485
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
         Height          =   1005
         Left            =   8790
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         ToolTipText     =   "Observações."
         Top             =   1005
         Width           =   6210
      End
      Begin VB.ComboBox Cmb_part_number_fabricante 
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
         ItemData        =   "frmproj_conjunto.frx":1624C
         Left            =   5700
         List            =   "frmproj_conjunto.frx":1624E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Part number do fabricante."
         Top             =   370
         Width           =   2310
      End
      Begin VB.TextBox Txt_percenual_perda 
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
         Left            =   7845
         MaxLength       =   50
         TabIndex        =   27
         Text            =   "0,0000"
         ToolTipText     =   "Percentual de perda."
         Top             =   1680
         Width           =   915
      End
      Begin VB.CommandButton cmdPesoBruto 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1530
         Picture         =   "frmproj_conjunto.frx":16250
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Carregar peso bruto do produto principal."
         Top             =   1680
         Width           =   315
      End
      Begin VB.TextBox Txt_posicao 
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
         Left            =   890
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Posição."
         Top             =   370
         Width           =   600
      End
      Begin VB.ComboBox cmbVersao 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmproj_conjunto.frx":16332
         Left            =   180
         List            =   "frmproj_conjunto.frx":16384
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Versão."
         Top             =   370
         Width           =   705
      End
      Begin VB.CommandButton Cmd_localizar_prod_item 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3285
         Picture         =   "frmproj_conjunto.frx":163D6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Localizar produtos."
         Top             =   370
         Width           =   315
      End
      Begin VB.TextBox cmbcodigo 
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
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   370
         Width           =   1770
      End
      Begin VB.TextBox cmbfamilia 
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
         MaxLength       =   255
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   990
         Width           =   5625
      End
      Begin VB.ComboBox cmbunkg 
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
         ItemData        =   "frmproj_conjunto.frx":164D8
         Left            =   2130
         List            =   "frmproj_conjunto.frx":164E8
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Unidade por kilograma."
         Top             =   1680
         Width           =   1035
      End
      Begin VB.ComboBox cmbcodref 
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
         ItemData        =   "frmproj_conjunto.frx":16500
         Left            =   3615
         List            =   "frmproj_conjunto.frx":16502
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Código de referência."
         Top             =   370
         Width           =   2070
      End
      Begin VB.TextBox txtpesototal 
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
         Left            =   6765
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Peso total."
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox txtkgpc 
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
         Left            =   4275
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Peso por peça."
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox txtun 
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
         Left            =   14565
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   370
         Width           =   435
      End
      Begin VB.TextBox txtdimensao 
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
         Left            =   3195
         MaxLength       =   30
         TabIndex        =   23
         Text            =   "0,00000"
         ToolTipText     =   "Dimensão a ser utilizada por peça."
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox txtpeso 
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
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Quilograma por unidade."
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtquantidade 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   5355
         MaxLength       =   50
         TabIndex        =   25
         Text            =   "0,0000000000"
         ToolTipText     =   "Quantidade."
         Top             =   1680
         Width           =   1395
      End
      Begin VB.TextBox txtdescricao 
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
         Left            =   8010
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   370
         Width           =   6540
      End
      Begin VB.TextBox txtVT 
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
         Left            =   7845
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   1680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox cmbVU 
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
         Left            =   6765
         MaxLength       =   50
         TabIndex        =   37
         Text            =   "0,0000"
         ToolTipText     =   "Valor unitário."
         Top             =   1680
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comprimento/mm"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   7395
         TabIndex        =   86
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Largura/mm"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   6112
         TabIndex        =   85
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   11348
         TabIndex        =   71
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Part number"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6413
         TabIndex        =   70
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% perda"
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
         Left            =   7935
         TabIndex        =   69
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pos.*"
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
         Left            =   965
         TabIndex        =   64
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versão*"
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
         Left            =   187
         TabIndex        =   57
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Un/Kg*"
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
         Left            =   2332
         TabIndex        =   55
         Top             =   1470
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3975
         TabIndex        =   54
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2752
         TabIndex        =   53
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso total"
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
         Left            =   6870
         TabIndex        =   48
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/pç"
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
         Left            =   4560
         TabIndex        =   47
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dim./mm*"
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
         Left            =   3285
         TabIndex        =   46
         Top             =   1470
         Width           =   885
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1935
         TabIndex        =   45
         Top             =   1770
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/unidade*"
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
         Left            =   307
         TabIndex        =   44
         Top             =   1470
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Un."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   14655
         TabIndex        =   43
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Quant.*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   5677
         TabIndex        =   42
         Top             =   1470
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   1778
         TabIndex        =   41
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   10868
         TabIndex        =   40
         Top             =   180
         Width           =   825
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   72
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
End
Attribute VB_Name = "frmproj_conjunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Conjunto As Boolean 'OK
Dim TBLISTA_Conjunto As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=kv2oTudH53k&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=51&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

cmbcodigo = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
txtUN.Text = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
cmbfamilia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
txtLargura2 = FunFormataCasasDecimais(2, IIf(IsNull(TBProduto!Largura), 0, TBProduto!Largura))
txtComprimento2 = FunFormataCasasDecimais(2, IIf(IsNull(TBProduto!Comprimento), 0, TBProduto!Comprimento))
txtpeso.Text = FunFormataCasasDecimais(10, IIf(IsNull(TBProduto!peso_metro), 0, TBProduto!peso_metro))
If IsNull(TBProduto!Un_Kg) = False Then cmbunkg = TBProduto!Un_Kg
If TBProduto!PCusto <> "" And TBProduto!PCusto <> 0 Then cmbVU = FunFormataCasasDecimais(10, TBProduto!PCusto) Else cmbVU = 0
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbcodigo_Change()
On Error GoTo tratar_erro

If cmbcodigo = "" Then Exit Sub
ProcCarregaComboCodRef cmbcodref, "P.desenho = '" & cmbcodigo & "'", 0, "", False, True
ProcCarregaComboPartNumberFab Cmb_part_number_fabricante, "P.Desenho = '" & cmbcodigo & "' and PF.Part_number IS NOT NULL"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbunkg_Click()
On Error GoTo tratar_erro

If cmbunkg = "Mt²" Then
    If txtUN = "MT" Then Label24.Caption = "Area/mt*" Else Label24.Caption = "Area/mm*"
Else
    Label24.Caption = "Dim./mm*"
End If

txtQuantidade = "1,000000"
Qtde = IIf(txtComprimento = "", 0, txtComprimento)
Qtd = IIf(txtComprimento2 = "", 0, txtComprimento2)
If (cmbunkg = "Mt²" Or cmbunkg = "Mt/L") And Qtde > 0 And Qtd > 0 Then
    Qtde = IIf(txtLargura = "", 0, txtLargura)
    Qtd = IIf(txtLargura2 = "", 0, txtLargura2)
    If cmbunkg = "Mt²" And (Qtde <= 0 Or Qtd <= 0) Then Exit Sub

    txtQuantidade = Format(FunCalcularQtdeUnKg(IIf(txtLargura = "", 0, txtLargura), txtComprimento, IIf(txtLargura2 = "", 0, txtLargura2), txtComprimento2, cmbunkg), "###,##0.000000")
End If

ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_Click()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select Posicao from ProjConjunto where codproduto = " & Txt_cod_produto & " and Versao = '" & cmbVersao_pesquisar & "' and Desenho = '" & cmbcodigo & "' and Versao_desenho = '" & cmbVersao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = True Then

    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Posicao from ProjConjunto where codproduto = " & Txt_cod_produto & " and Versao = '" & cmbVersao_pesquisar & "' and Posicao IS NOT NULL order by Posicao desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Txt_posicao = TBFIltro!Posicao + 1
    Else
        Txt_posicao = 1
    End If
Else
    Txt_posicao = TBFIltro!Posicao
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_pesquisar_Click()
On Error GoTo tratar_erro

ProcLimpaCamposItem
ProcLimpaCamposDescVersao
Frame6.Enabled = False
ProcAtualizalista (1)
If Lista.ListItems.Count <> 0 Then
    Frame6.Enabled = True
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Projconjunto_desc_versao where Codproduto = " & Txt_cod_produto & " and Versao = '" & cmbVersao_pesquisar & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_ID_desc_versao = TBAbrir!ID
        Txt_descricao_versao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
        txtDtValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
        txtRespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVU_Change()
On Error GoTo tratar_erro

If cmbVU.Text <> "" Then
    VerifNumero = cmbVU.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        cmbVU.Text = ""
        cmbVU.SetFocus
        Exit Sub
    End If
End If
ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVU_LostFocus()
On Error GoTo tratar_erro

cmbVU = Format(cmbVU, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro
  
frmproj_conjunto_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_prod_Click()
On Error GoTo tratar_erro

ProcLocalizar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_prod_item_Click()
On Error GoTo tratar_erro

'.Caption = "Localizar componentes do conjunto"
frmproj_conjunto_localizaritem.Show 1
Frame1.Refresh

cmbVersao.Text = "A"

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
If txtdesenhoproduto = "" Then
    USMsgBox ("Informe o código interno do produto que deseja cadastrar a estrutura."), vbExclamation, "CAPRIND v5.0"
    Cmd_localizar_prod_Click
    Exit Sub
End If
ProcLimpaCamposItem
Cmd_localizar_prod_item_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_desc_versao_Click()
On Error GoTo tratar_erro

If Frame6.Enabled = False Then Exit Sub
If FunVerificaRegistroValidado("Projconjunto_desc_versao", "ID = " & Txt_ID_desc_versao, "versão da estrutura", "deste registro", "cadastrar descrição da versão", False, True) = False Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Projconjunto_desc_versao where ID = " & Txt_ID_desc_versao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Codproduto = Txt_cod_produto
TBAbrir!versao = cmbVersao_pesquisar
TBAbrir!Descricao = Txt_descricao_versao
TBAbrir.Update
Txt_ID_desc_versao = TBAbrir!ID
TBAbrir.Close
USMsgBox ("Descrição da versão cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Engenharia/Conjuntos"
Evento = "Cadastrar descrição da versão"
ID_documento = Txt_ID_desc_versao
Documento = "Cód. interno: " & txtdesenhoproduto & " - Versão: " & cmbVersao_pesquisar
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Conjunto.AbsolutePage <> 2 Then
    If TBLISTA_Conjunto.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Conjunto.PageCount - 1)
    Else
        TBLISTA_Conjunto.AbsolutePage = TBLISTA_Conjunto.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Conjunto.AbsolutePage)
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
    TBLISTA_Conjunto.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Conjunto.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Conjunto.AbsolutePage = 1
ProcExibePagina (TBLISTA_Conjunto.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Conjunto.AbsolutePage <> -3 Then
    If TBLISTA_Conjunto.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Conjunto.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Conjunto.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Conjunto.AbsolutePage = TBLISTA_Conjunto.PageCount
ProcExibePagina (TBLISTA_Conjunto.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPesoBruto_Click()
On Error GoTo tratar_erro

If Txt_ID = "" Or txtdesenhoproduto = "" Then Exit Sub

If USMsgBox("Deseja realmente carregar o peso bruto do produto principal?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select PBruto from projproduto where desenho = '" & txtdesenhoproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtpeso = IIf(IsNull(TBProduto!PBruto), "0,00000", Format(TBProduto!PBruto, "###,##0.0000000000"))
    End If
    TBProduto.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 11, True

If Compras = True Then Caption = "Compras - Conjuntos"
If Vendas = True Then Caption = "Vendas - Conjuntos"
Formulario = "Engenharia/Conjuntos"
Direitos

If FunLiberaCamposEstrutura = True Then
    With txtpeso
        .Locked = False
        .TabStop = True
    End With
    With cmbunkg
        .Locked = False
        .TabStop = True
    End With
End If

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Engenharia/Conjuntos"
Direitos
If FunLiberaCamposEstrutura = True Then
    With txtpeso
        .Locked = False
        .TabStop = True
    End With
    With cmbunkg
        .Locked = False
        .TabStop = True
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
  
frmproj_conjunto_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
  
If Novo_Conjunto = True Then
    If USMsgBox("O produto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Conjunto = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Conjunto = False
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
If Txt_cod_produto = 0 Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    imgLocalizar.SetFocus
    Exit Sub
End If
If cmbVersao.Text = "" Then
    NomeCampo = "a versão"
    ProcVerificaAcao
    cmbVersao.SetFocus
    Exit Sub
End If
If Txt_posicao = "" Then
    NomeCampo = "a posição"
    ProcVerificaAcao
    Txt_posicao.SetFocus
    Exit Sub
End If
If cmbcodigo.Text = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Cmd_localizar_prod_item_Click
    Exit Sub
End If
If txtpeso.Text = "" Then
    NomeCampo = "o peso"
    ProcVerificaAcao
    txtpeso.SetFocus
    Exit Sub
End If
If cmbunkg.Text = "" Then
    NomeCampo = "a unidade do kilograma"
    ProcVerificaAcao
    cmbunkg.SetFocus
    Exit Sub
End If
If txtdimensao.Text = "" Then
    NomeCampo = "a dimensão"
    ProcVerificaAcao
    txtdimensao.SetFocus
    Exit Sub
End If
valor = IIf(txtQuantidade = "", 0, txtQuantidade)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQuantidade.SetFocus
    Exit Sub
End If
'Verifica tipo do produto para ver se é obrigatório
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select P.Codproduto from Projproduto P INNER JOIN projproduto_Tipo PT ON PT.ID = P.ID_Tipo where P.Desenho = '" & txtdesenhoproduto & "' and (PT.Codigo = '03' or PT.Codigo = '04')", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select P.Codproduto from Projproduto P INNER JOIN projproduto_Tipo PT ON PT.ID = P.ID_Tipo where P.Desenho = '" & cmbcodigo & "' and (PT.Codigo = '00' or PT.Codigo = '01' or PT.Codigo = '02' or PT.Codigo = '03' or PT.Codigo = '04' or PT.Codigo = '05' or PT.Codigo = '10')", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If Txt_percenual_perda = "" Then
            NomeCampo = "o percentual de perda"
            ProcVerificaAcao
            Txt_percenual_perda.SetFocus
            TBFI.Close
            Exit Sub
        End If
    End If
End If
TBFI.Close

If FunVerificaRegistroValidado("Projconjunto_desc_versao", "ID = " & Txt_ID_desc_versao, "mesma", "esta versão da estrutura", "salvar", False, True) = False Then Exit Sub

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from ProjConjunto where Codigo = " & IIf(Txt_ID = "", 0, Txt_ID), Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then
    TBProduto.AddNew
    TextoFiltroPos = "Posicao = Posicao + 1 where Posicao >= " & Txt_posicao
Else
    TextoFiltroPos = ""
    If Txt_posicao < TBProduto!Posicao Then
         TextoFiltroPos = "Posicao = Posicao + 1 where Posicao >= " & Txt_posicao & " and Posicao < " & TBProduto!Posicao
    ElseIf Txt_posicao > TBProduto!Posicao Then
            TextoFiltroPos = "Posicao = Posicao - 1 where Posicao > " & TBProduto!Posicao & " and Posicao <= " & Txt_posicao
    End If
End If
If TextoFiltroPos <> "" Then Conexao.Execute "Update ProjConjunto Set " & TextoFiltroPos & " and Posicao IS NOT NULL and Codproduto = " & Txt_cod_produto & " and Versao = '" & cmbVersao_pesquisar & "'"
    
TBProduto!versao = IIf(cmbVersao_pesquisar = "", "A", cmbVersao_pesquisar)
TBProduto!Posicao = Txt_posicao
TBProduto!Desenho = cmbcodigo.Text
TBProduto!Versao_desenho = cmbVersao

If Cmb_part_number_fabricante <> "" Then
    TBProduto!ID_partnumber_fabricante = Cmb_part_number_fabricante.ItemData(Cmb_part_number_fabricante.ListIndex)
End If

TBProduto!Codproduto = Txt_cod_produto.Text
TBProduto!Descricao = txtdescricao.Text
TBProduto!quantidade = txtQuantidade.Text
TBProduto!Dimensoes = txtdimensao.Text
TBProduto!Peso = txtkgpc.Text
TBProduto!PesoMetro = IIf(txtpeso.Text = "", 0, txtpeso)
TBProduto!PesoTotal = txtPesototal.Text
TBProduto!Percentual_perda = IIf(Txt_percenual_perda = "", 0, Txt_percenual_perda)
TBProduto!Unidade = txtUN.Text
TBProduto!Un_Kg = cmbunkg.Text
TBProduto!valor = cmbVU.Text
TBProduto!ValorTotal = txtVT.Text
TBProduto!Obs = Txt_obs
TBProduto.Update
Txt_ID = TBProduto!CODIGO
TBProduto.Close
ProcCarregaVersao cmbVersao_pesquisar
If Novo_Conjunto = True Then
    USMsgBox ("Novo registro agregado na estrutura com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    ProcAtualizalista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Engenharia/Conjuntos"
ID_documento = Txt_cod_produto
Documento = "Cód. interno: " & txtdesenhoproduto.Text
Documento1 = "Cód. interno: " & cmbcodigo.Text
ProcGravaEvento
'==================================
ProcLimpaCamposItem
Novo_Conjunto = False

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
                If USMsgBox("Deseja realmente excluir este(s) registro(s) da estrutura?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Desenho, Posicao, codproduto, Versao from ProjConjunto where Codigo = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Engenharia/Conjuntos"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Cód. interno: " & txtdesenhoproduto.Text
                Documento1 = "Cód. interno: " & TBFI!Desenho
                ProcGravaEvento
                '==================================
                Conexao.Execute "Update ProjConjunto Set Posicao = Posicao - 1 where Posicao > " & TBFI!Posicao & " and Posicao IS NOT NULL and codproduto = " & TBFI!Codproduto & " and Versao = '" & TBFI!versao & "'"
                Conexao.Execute "DELETE from ProjConjunto where Codigo = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) registro(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Registro(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposItem
    ProcCarregaVersao cmbVersao_pesquisar
    ProcAtualizalista (1)
    Novo_Conjunto = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF6: Cmd_salvar_desc_versao_Click
    Case vbKeyF7: ProcCopiar
    Case vbKeyF8: ProcCriarVersao
    Case vbKeyF9:
        Formulario = "Engenharia/Conjuntos"
        frmValidar.Show
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

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
If Txt_cod_produto = 0 Or cmbVersao_pesquisar = "" Then Exit Sub
Set TBLISTA_Conjunto = CreateObject("adodb.recordset")
TBLISTA_Conjunto.Open "Select * from ProjConjunto where codproduto = " & Txt_cod_produto & " and versao = '" & cmbVersao_pesquisar & "' order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Conjunto.EOF = False Then
    ProcExibePagina (Pagina)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Conjunto.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Conjunto.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Conjunto.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Conjunto.RecordCount - IIf(Pagina > 1, (TBLISTA_Conjunto.PageSize * (Pagina - 1)), 0), TBLISTA_Conjunto.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Conjunto.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Conjunto!CODIGO
        
        Select Case Len(TBLISTA_Conjunto!Posicao)
            Case 1: PosicaoTexto = "00" & TBLISTA_Conjunto!Posicao
            Case 2: PosicaoTexto = "0" & TBLISTA_Conjunto!Posicao
            Case 3: PosicaoTexto = TBLISTA_Conjunto!Posicao
        End Select
        .Item(.Count).SubItems(1) = PosicaoTexto
        
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Conjunto!Desenho), "", TBLISTA_Conjunto!Desenho)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Conjunto!Descricao), "", TBLISTA_Conjunto!Descricao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Conjunto!quantidade), "0,000000", Format(TBLISTA_Conjunto!quantidade, "###,##0.000000"))
    End With
    TBLISTA_Conjunto.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Conjunto.RecordCount
If TBLISTA_Conjunto.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Conjunto.PageCount
ElseIf TBLISTA_Conjunto.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Conjunto.PageCount & " de: " & TBLISTA_Conjunto.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Conjunto.AbsolutePage - 1 & " de: " & TBLISTA_Conjunto.PageCount
End If


1:

Exit Sub
tratar_erro:
    If Err.Number = 365 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_cod_produto = 0
txtdesenhoproduto = ""
txtRef = ""
txtDescricaoProduto.Text = ""
txtmaodeobra = "0,00"
txtMaterial = "0,00"
txtOutros = "0,00"
txtimpostos = "0,00"
txtValorTotal = "0,00000"
txtComprimento = ""
txtLargura = ""
Frame6.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposItem()
On Error GoTo tratar_erro

Txt_ID = 0
cmbVersao.ListIndex = -1
Txt_posicao = ""
cmbcodigo = ""
cmbcodref.ListIndex = -1
Cmb_part_number_fabricante.ListIndex = -1
txtUN.Text = ""
txtdescricao.Text = ""
cmbfamilia = ""
txtpeso.Text = "0,00000"
cmbunkg.ListIndex = -1
txtdimensao.Text = "0,00000"
txtkgpc.Text = "0,00000"
txtQuantidade.Text = "0,0000000000"
txtPesototal.Text = "0,00000"
Txt_percenual_perda = "0,0000"
txtVT.Text = "0,00"
cmbVU.Text = "0,00000"
txtComprimento2 = ""
txtLargura2 = ""
Txt_obs = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtdesenhoproduto.Text = "" Then
    USMsgBox ("Informe o produto antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "Engenharia_conjuntos_estrutura.rpt"
ProcImprimirRel "{projproduto.desenho}='" & txtdesenhoproduto.Text & "' and {Projconjunto.Versao} = '" & cmbVersao_pesquisar & "'", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtdesenhoproduto.Text = "" Then
    USMsgBox ("Informe o produto antes de copiar a estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar esse conjunto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Sit_REG = 2
    frmprocessos_Novo.Show 1
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
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Projconjunto_desc_versao", "ID = " & Txt_ID_desc_versao, True) = False Then GoTo Proximo
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
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select Codproduto from ProjConjunto where Codigo = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                If FunVerificaRegistroValidado("Projconjunto_desc_versao", "ID = " & Txt_ID_desc_versao, "mesma", "esta versão da estrutura", "excluir", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                End If
            End If
            TBLISTA.Close
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
ProcLimpaCamposItem
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PC.*, P.classe, P.Largura, P.Comprimento from ProjConjunto PC INNER JOIN Projproduto P ON PC.Desenho = P.Desenho where PC.Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_ID = TBLISTA!CODIGO
    If IsNull(TBLISTA!Versao_desenho) = False And TBLISTA!Versao_desenho <> "" Then cmbVersao = TBLISTA!Versao_desenho
    Txt_posicao = IIf(IsNull(TBLISTA!Posicao), "", TBLISTA!Posicao)
    cmbcodigo.Text = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
    If IsNull(TBLISTA!ID_partnumber_fabricante) = False Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select Part_number from Projproduto_fabricante where ID = " & TBLISTA!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then Cmb_part_number_fabricante = TBFIltro!Part_number
        TBFIltro.Close
    End If
    txtUN.Text = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
    txtdescricao.Text = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
    txtdimensao.Text = IIf(IsNull(TBLISTA!Dimensoes), "0,00000", Format(TBLISTA!Dimensoes, "###,##0.0000000000"))
    If IsNull(TBLISTA!Classe) = False And TBLISTA!Classe <> "" Then cmbfamilia = TBLISTA!Classe
    txtkgpc.Text = IIf(IsNull(TBLISTA!Peso), "0,00000", Format(TBLISTA!Peso, "###,##0.0000000000"))
    txtpeso.Text = IIf(IsNull(TBLISTA!PesoMetro), "0,00000", Format(TBLISTA!PesoMetro, "###,##0.0000000000"))
    txtPesototal.Text = IIf(IsNull(TBLISTA!PesoTotal), "0,00000", Format(TBLISTA!PesoTotal, "###,##0.0000000000"))
    Txt_percenual_perda = IIf(IsNull(TBLISTA!Percentual_perda), "0,0000", Format(TBLISTA!Percentual_perda, "###,##0.0000"))
    cmbVU.Text = IIf(IsNull(TBLISTA!valor), "0,00000", Format(TBLISTA!valor, "###,##0.0000000000"))
    txtVT.Text = IIf(IsNull(TBLISTA!ValorTotal), "0,00", Format(TBLISTA!ValorTotal, "###,##0.00"))
    txtLargura2 = IIf(IsNull(TBLISTA!Largura), "", Format(TBLISTA!Largura, "###,##0.00"))
    txtComprimento2 = IIf(IsNull(TBLISTA!Comprimento), "", Format(TBLISTA!Comprimento, "###,##0.00"))
    If IsNull(TBLISTA!Un_Kg) = False Then cmbunkg.Text = TBLISTA!Un_Kg
    
    'Quantidade precisa ficar depois da Un_KG pois existe um calculo de un_kg que altera a quantidade
    txtQuantidade.Text = IIf(IsNull(TBLISTA!quantidade), "0,000000", Format(TBLISTA!quantidade, "###,##0.000000"))
    Txt_obs = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
End If
TBLISTA.Close
CodigoLista = Lista.SelectedItem.index
Novo_Conjunto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Txt_percenual_perda_Change()
On Error GoTo tratar_erro

If Txt_percenual_perda <> "" Then
    VerifNumero = Txt_percenual_perda
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_percenual_perda = ""
        Txt_percenual_perda.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_percenual_perda_LostFocus()
On Error GoTo tratar_erro

Txt_percenual_perda = Format(Txt_percenual_perda, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_posicao_Change()
On Error GoTo tratar_erro

If Txt_posicao <> "" Then
    VerifNumero = Txt_posicao
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_posicao = ""
        Txt_posicao.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_posicao_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_posicao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdimensao_Change()
On Error GoTo tratar_erro

If txtdimensao.Text <> "" Then
    VerifNumero = txtdimensao.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtdimensao.Text = ""
        txtdimensao.SetFocus
        Exit Sub
    End If
End If
ProcCalculaPeso
ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaPeso()
On Error GoTo tratar_erro

If txtpeso.Text <> "" And cmbunkg.Text <> "" And txtdimensao.Text <> "" And txtQuantidade.Text <> "" Then
    If txtdimensao.Text = "" Then txtdimensao.Text = Format(0, "###,##0.0000000000")
    If cmbunkg.Text = "Mt/L" Then txtkgpc.Text = Format(txtpeso.Text / 1000 * txtdimensao, "###,##0.0000000000")
    If cmbunkg.Text = "Pç" Then txtkgpc.Text = Format(txtpeso.Text, "###,##0.0000000000")
    If cmbunkg.Text = "Mt²" Then txtkgpc.Text = Format(((txtdimensao * txtpeso) / 1000) / 1000, "###,##0.0000000000")
    If cmbunkg.Text = "N/a" Then txtkgpc.Text = Format(0, "###,##0.0000000000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcalculaPesoTotal()
On Error GoTo tratar_erro

If txtkgpc.Text <> "" And txtQuantidade <> "" Then
    txtPesototal = Format(txtkgpc.Text * txtQuantidade.Text, "###,##0.0000000000")
Else
    txtPesototal = "0,00000"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaValor()
On Error GoTo tratar_erro

ProcCalculaPeso
ProcalculaPesoTotal
If cmbVU.Text <> "" And txtQuantidade.Text <> "" And txtdimensao.Text <> "" Then
    Select Case txtUN
        Case "KG": txtVT = Format(cmbVU * txtPesototal, "###,##0.00")
        Case "MM": txtVT = Format((cmbVU * txtdimensao) * txtQuantidade, "###,##0.00")
        Case "MT": txtVT = Format(((cmbVU / 1000) * txtdimensao) * txtQuantidade, "###,##0.00")
    End Select
    If txtUN <> "KG" And txtUN <> "MM" And txtUN <> "MT" Then txtVT = Format(cmbVU * txtQuantidade, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdimensao_LostFocus()
On Error GoTo tratar_erro

txtdimensao.Text = Format(txtdimensao.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtkgpc_Change()
On Error GoTo tratar_erro

ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtkgpc_LostFocus()
On Error GoTo tratar_erro

If txtkgpc.Text <> "" Then
    VerifNumero = txtkgpc.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtkgpc.Text = ""
        txtkgpc.SetFocus
        Exit Sub
    End If
    txtkgpc.Text = Format(txtkgpc.Text, "###,##0.0000000000")
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

Private Sub txtpeso_Change()
On Error GoTo tratar_erro

If txtpeso.Text <> "" Then
    VerifNumero = txtpeso.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtpeso.Text = ""
        txtpeso.SetFocus
        Exit Sub
    End If
End If
ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpeso_LostFocus()
On Error GoTo tratar_erro

txtpeso.Text = Format(txtpeso.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpesototal_Change()
On Error GoTo tratar_erro

ProcVerificaValor

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
ProcVerificaValor
txtVT.Text = Format(txtVT.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procatualizadados(Desenho As String)
On Error GoTo tratar_erro
  
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto, Desenho, Descricao, SubTipoItem, classe, Largura, Comprimento from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Txt_cod_produto = TBProduto!Codproduto
    txtdesenhoproduto.Text = TBProduto!Desenho
    txtDescricaoProduto.Text = TBProduto!Descricao
    txtLargura = IIf(IsNull(TBProduto!Largura), "", Format(TBProduto!Largura, "###,##0.00"))
    txtComprimento = IIf(IsNull(TBProduto!Comprimento), "", Format(TBProduto!Comprimento, "###,##0.00"))
    Select Case TBProduto!SubTipoItem
        Case 0: txttipo = "Matéria-prima"
        Case 1: txttipo = "Produto Final"
        Case 2: txttipo = "Subconjunto"
        Case 3: txttipo = "Componente"
        Case 4: txttipo = "Outros"
    End Select
    txtfamilia = TBProduto!Classe
    
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select N_Referencia from item_aplicacoes where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        txtRef.Text = TBItem!N_referencia
    End If
    TBItem.Close
    Frame6.Enabled = True
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantidade_LostFocus()
On Error GoTo tratar_erro

txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtun_Change()
On Error GoTo tratar_erro

If txtUN.Text = "M³" Then
    Label4.Caption = txtUN.Text & " x Unidade*"
    txtpeso.ToolTipText = txtUN.Text & " x Unidade"
    Label5.Caption = "Un x " & txtUN.Text & "*"
    cmbunkg.ToolTipText = "Un x " & txtUN.Text
    Label6.Caption = txtUN.Text & " x PC*"
    txtkgpc.ToolTipText = txtUN.Text & " x PC"
    Label9.Caption = txtUN.Text & " Total"
    txtPesototal.ToolTipText = txtUN.Text & " Total"
Else
    Label4.Caption = "Kg/unidade*"
    txtpeso.ToolTipText = "Kg/unidade"
    Label5.Caption = "Un/Kg*"
    cmbunkg.ToolTipText = "Un/Kg"
    Label6.Caption = "Kg/pç"
    txtkgpc.ToolTipText = "Kg/pç"
    Label9.Caption = "Peso total"
    txtPesototal.ToolTipText = "Peso total"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcCopiar
    Case 7: ProcCriarVersao
    Case 8:
            Formulario = "Engenharia/Conjuntos"
            frmValidar.Show
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaVersao(versao As String)
On Error GoTo tratar_erro

With cmbVersao_pesquisar
    .Clear
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select versao from projconjunto WHERE codProduto = " & Txt_cod_produto & " and versao is not null and versao <> N'' GROUP BY versao", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Do While TBItem.EOF = False
            .AddItem TBItem!versao
            TBItem.MoveNext
        Loop
        If versao = "" Then
1:
            TBItem.MoveFirst
            .Text = TBItem!versao
        Else
            .Text = versao
        End If
    End If
End With

Exit Sub
tratar_erro:
    If Err.Number = 383 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarVersao()
On Error GoTo tratar_erro

Acao = "criar a versão"
If txtdesenhoproduto = "" Then
    NomeCampo = "o registro"
    ProcVerificaAcao
    Exit Sub
End If
If USMsgBox("Deseja realmente criar nova versão para esse conjunto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Engenharia_Conjuntos = True
    frmproj_conjunto_criar_versao.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposDescVersao()
On Error GoTo tratar_erro

Txt_ID_desc_versao = 0
Txt_descricao_versao = ""
txtDtValidacao = ""
txtRespValidacao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
