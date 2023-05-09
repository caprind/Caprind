VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmproj_produto 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Produtos e serviços"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15375
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
   Icon            =   "frmproj_produto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USToolBar USToolBar2 
      Height          =   975
      Left            =   90
      TabIndex        =   182
      Top             =   330
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   8
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
      ButtonCaption2  =   "Relatório"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Relatório (F5)"
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
      ButtonLeft2     =   42
      ButtonTop2      =   2
      ButtonWidth2    =   51
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Anterior"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Registro anterior."
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
      ButtonLeft3     =   95
      ButtonTop3      =   2
      ButtonWidth3    =   47
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Próximo"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Próximo registro."
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
      ButtonLeft4     =   144
      ButtonTop4      =   2
      ButtonWidth4    =   46
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonAlignment5=   2
      ButtonType5     =   1
      ButtonStyle5    =   -1
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   -1
      ButtonLeft5     =   192
      ButtonTop5      =   4
      ButtonWidth5    =   2
      ButtonHeight5   =   54
      ButtonCaption6  =   "Ajuda"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Ajuda (F1)"
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
      ButtonLeft6     =   196
      ButtonTop6      =   2
      ButtonWidth6    =   36
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Sair"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Sair (Esc)"
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
      ButtonLeft7     =   234
      ButtonTop7      =   2
      ButtonWidth7    =   26
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
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
      ButtonState8    =   5
      ButtonLeft8     =   262
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
      ButtonUseMaskColor8=   0   'False
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   12480
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmproj_produto.frx":000C
         Count           =   1
      End
   End
   Begin VB.ComboBox cmbFabricante 
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
      Left            =   260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   41
      ToolTipText     =   "Fabricante."
      Top             =   1710
      Visible         =   0   'False
      Width           =   9315
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Compras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2550
      TabIndex        =   204
      Top             =   1350
      Visible         =   0   'False
      Width           =   6345
      Begin VB.TextBox Txt_ID_CFOP 
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "ID da CFOP."
         Top             =   390
         Width           =   525
      End
      Begin VB.TextBox Txt_natureza_operacao 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   86
         TabStop         =   0   'False
         ToolTipText     =   "Descrição da natureza da operação."
         Top             =   390
         Width           =   3960
      End
      Begin VB.CommandButton Cmd_localizar_CFOP 
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
         Left            =   5520
         Picture         =   "frmproj_produto.frx":3D90
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Localizar CFOP."
         Top             =   390
         Width           =   315
      End
      Begin VB.CommandButton Cmd_limpar_CFOP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5850
         Picture         =   "frmproj_produto.frx":3E92
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Limpar CFOP."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_CFOP 
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
         Left            =   690
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Natureza da operação."
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID          CFOP                                   Natureza da operação"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   205
         Top             =   180
         Width           =   4200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   75
      TabIndex        =   183
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
         ItemData        =   "frmproj_produto.frx":3FD0
         Left            =   6720
         List            =   "frmproj_produto.frx":3FE3
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Width           =   2235
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
         TabIndex        =   33
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
         TabIndex        =   35
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   39
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_produto.frx":4035
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
         TabIndex        =   38
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_produto.frx":77D9
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
         TabIndex        =   36
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
         TabIndex        =   37
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_produto.frx":B2E2
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
         TabIndex        =   40
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmproj_produto.frx":F3D1
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   222
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   35
         Left            =   5370
         TabIndex        =   199
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   2040
         TabIndex        =   191
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   185
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   184
         Top             =   240
         Width           =   1275
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   180
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
   Begin VB.ComboBox cmbCliente 
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
      Left            =   4290
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   113
      ToolTipText     =   "Cliente."
      Top             =   2430
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ComboBox cmbGrupo 
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
      ItemData        =   "frmproj_produto.frx":12C5D
      Left            =   250
      List            =   "frmproj_produto.frx":12C5F
      Locked          =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   110
      TabStop         =   0   'False
      ToolTipText     =   "Grupo de clientes."
      Top             =   2430
      Visible         =   0   'False
      Width           =   3140
   End
   Begin VB.TextBox txtdescricaoproduto 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   270
      MaxLength       =   255
      TabIndex        =   28
      ToolTipText     =   "Descrição."
      Top             =   3600
      Width           =   6465
   End
   Begin VB.ComboBox cmbun 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmproj_produto.frx":12C61
      Left            =   7110
      List            =   "frmproj_produto.frx":12C63
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   29
      ToolTipText     =   "Unidade de estoque."
      Top             =   3600
      Width           =   825
   End
   Begin VB.ComboBox Cmb_un_com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmproj_produto.frx":12C65
      Left            =   7950
      List            =   "frmproj_produto.frx":12C67
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   30
      ToolTipText     =   "Unidade comercial."
      Top             =   3600
      Width           =   825
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   315
      ItemData        =   "frmproj_produto.frx":12C69
      Left            =   9510
      List            =   "frmproj_produto.frx":12C6B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   2970
      Width           =   5235
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4305
      Left            =   75
      TabIndex        =   32
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
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
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
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Part number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3186
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "D"
         Text            =   "Dt. rev."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Dimensão"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Dureza"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "NCM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Vlr. custo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Vlr. cons."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Vlr. reve."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "Validado"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vendas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8910
      TabIndex        =   206
      Top             =   1350
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton Cmd_limpar_CFOP1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5880
         Picture         =   "frmproj_produto.frx":12C6D
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Limpar CFOP."
         Top             =   390
         Width           =   315
      End
      Begin VB.CommandButton Cmd_localizar_CFOP1 
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
         Left            =   5550
         Picture         =   "frmproj_produto.frx":12DAB
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Localizar CFOP."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_natureza_operacao1 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   91
         TabStop         =   0   'False
         ToolTipText     =   "Descrição da natureza da operação."
         Top             =   390
         Width           =   3960
      End
      Begin VB.TextBox Txt_ID_CFOP1 
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         ToolTipText     =   "ID da CFOP."
         Top             =   390
         Width           =   525
      End
      Begin VB.TextBox Txt_CFOP1 
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
         Left            =   690
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   90
         TabStop         =   0   'False
         ToolTipText     =   "Natureza da operação."
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID          CFOP                                   Natureza da operação"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   207
         Top             =   180
         Width           =   4200
      End
   End
   Begin VB.ComboBox cmbFornecedor 
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
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   103
      ToolTipText     =   "Fornecedor."
      Top             =   1710
      Visible         =   0   'False
      Width           =   10485
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   172
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   8
      Tab             =   2
      TabsPerRow      =   8
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
      TabCaption(0)   =   "Dados principais"
      TabPicture(0)   =   "frmproj_produto.frx":12EAD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "CommonDialog1"
      Tab(0).Control(2)=   "txtcodproduto"
      Tab(0).Control(3)=   "USToolBar1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Dados adicionais"
      TabPicture(1)   =   "frmproj_produto.frx":12EC9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIdFabricante"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Impostos"
      TabPicture(2)   =   "frmproj_produto.frx":12EE5
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame16"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame18"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame21"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame22"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame23"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Valores e descontos"
      TabPicture(3)   =   "frmproj_produto.frx":12F01
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Clientes e forn."
      TabPicture(4)   =   "frmproj_produto.frx":12F1D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame15"
      Tab(4).Control(1)=   "txtIDGrupo"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Conta contábil"
      TabPicture(5)   =   "frmproj_produto.frx":12F39
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame19"
      Tab(5).Control(1)=   "Frame11"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Imagem"
      TabPicture(6)   =   "frmproj_produto.frx":12F55
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame10"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Documentos"
      TabPicture(7)   =   "frmproj_produto.frx":12F71
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame14"
      Tab(7).Control(1)=   "txtID_doc"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Lista_doc"
      Tab(7).Control(3)=   "PBLista1"
      Tab(7).Control(4)=   "USToolBar3"
      Tab(7).ControlCount=   5
      Begin VB.Frame Frame23 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aliquota INSS"
         Height          =   675
         Left            =   10920
         TabIndex        =   266
         Top             =   2190
         Width           =   1335
         Begin VB.TextBox Txt_INSS 
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
            Left            =   90
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   267
            TabStop         =   0   'False
            ToolTipText     =   "Alíquota de INSS."
            Top             =   270
            Width           =   1110
         End
      End
      Begin VB.Frame Frame22 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gênero do item (Bloco K)***"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   675
         Left            =   75
         TabIndex        =   264
         Top             =   2190
         Width           =   7575
         Begin VB.ComboBox cmbGenero 
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
            ItemData        =   "frmproj_produto.frx":12F8D
            Left            =   120
            List            =   "frmproj_produto.frx":12F8F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   265
            ToolTipText     =   "Gênero do item."
            Top             =   270
            Width           =   7350
         End
      End
      Begin VB.Frame Frame21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor unid tributada"
         Height          =   675
         Left            =   13590
         TabIndex        =   262
         Top             =   2190
         Width           =   1695
         Begin VB.TextBox txtvuTrib 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   263
            Text            =   "0,0000"
            ToolTipText     =   "Vlr. da venda para consumo."
            Top             =   270
            Width           =   1395
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unid tributada"
         Height          =   675
         Left            =   12270
         TabIndex        =   260
         Top             =   2190
         Width           =   1305
         Begin VB.ComboBox cmbuTrib 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmproj_produto.frx":12F91
            Left            =   240
            List            =   "frmproj_produto.frx":12F93
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   261
            ToolTipText     =   "Unidade de estoque."
            Top             =   270
            Width           =   825
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   181
         Top             =   330
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1720
         ButtonCount     =   20
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
         ButtonCaption9  =   "Referência"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Cadastro de códigos de referência (F8)"
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
         ButtonWidth9    =   60
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Similar"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Cadastro de produtos similares."
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
         ButtonLeft10    =   409
         ButtonTop10     =   2
         ButtonWidth10   =   38
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Conjunto"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Cadastro de conjuntos (F9)"
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
         ButtonLeft11    =   449
         ButtonTop11     =   2
         ButtonWidth11   =   52
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Copiar"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Copiar (F10)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   503
         ButtonTop12     =   2
         ButtonWidth12   =   39
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Revisar"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Revisar (F11)"
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
         ButtonLeft13    =   544
         ButtonTop13     =   2
         ButtonWidth13   =   44
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonCaption14 =   "Validação"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Validar/Cancelar validação."
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
         ButtonLeft14    =   590
         ButtonTop14     =   2
         ButtonWidth14   =   53
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonCaption15 =   "Atualizar"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Utilizado pelo administrador do sistema"
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
         ButtonLeft15    =   645
         ButtonTop15     =   2
         ButtonWidth15   =   50
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonCaption16 =   "Sincronizar"
         ButtonEnabled16 =   0   'False
         ButtonIconSize16=   32
         ButtonToolTipText16=   "Sincronizar produtos na nuvem"
         ButtonKey16     =   "16"
         ButtonAlignment16=   2
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft16    =   697
         ButtonTop16     =   2
         ButtonWidth16   =   60
         ButtonHeight16  =   21
         ButtonUseMaskColor16=   0   'False
         ButtonEnabled17 =   0   'False
         ButtonIconSize17=   32
         ButtonAlignment17=   2
         ButtonType17    =   1
         ButtonStyle17   =   -1
         BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState17   =   -1
         ButtonLeft17    =   759
         ButtonTop17     =   4
         ButtonWidth17   =   2
         ButtonHeight17  =   54
         ButtonCaption18 =   "Ajuda"
         ButtonEnabled18 =   0   'False
         ButtonIconSize18=   32
         ButtonToolTipText18=   "Ajuda (F1)"
         ButtonKey18     =   "17"
         ButtonAlignment18=   2
         BeginProperty ButtonFont18 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft18    =   763
         ButtonTop18     =   2
         ButtonWidth18   =   36
         ButtonHeight18  =   21
         ButtonUseMaskColor18=   0   'False
         ButtonCaption19 =   "Sair"
         ButtonEnabled19 =   0   'False
         ButtonIconSize19=   32
         ButtonToolTipText19=   "Sair (ESC)"
         ButtonKey19     =   "18"
         ButtonAlignment19=   2
         BeginProperty ButtonFont19 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft19    =   801
         ButtonTop19     =   2
         ButtonWidth19   =   26
         ButtonHeight19  =   21
         ButtonUseMaskColor19=   0   'False
         ButtonEnabled20 =   0   'False
         ButtonIconSize20=   32
         ButtonKey20     =   "19"
         ButtonAlignment20=   2
         BeginProperty ButtonFont20 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState20   =   5
         ButtonLeft20    =   829
         ButtonTop20     =   2
         ButtonWidth20   =   24
         ButtonHeight20  =   24
         ButtonUseMaskColor20=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12630
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmproj_produto.frx":12F95
            Count           =   1
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0E0&
         Height          =   675
         Left            =   7665
         TabIndex        =   166
         Top             =   2190
         Width           =   3240
         Begin VB.CheckBox Chk_credito_ICMS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Não credita ICMS (NF de entrada)? "
            Height          =   195
            Left            =   180
            TabIndex        =   95
            Top             =   165
            Width           =   2835
         End
         Begin VB.CheckBox Chk_servico_executado_cliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Serviço executado no cliente?"
            Height          =   225
            Left            =   180
            TabIndex        =   94
            Top             =   390
            Width           =   2445
         End
      End
      Begin VB.TextBox txtIdFabricante 
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
         Left            =   -72270
         Locked          =   -1  'True
         MaxLength       =   50
         MouseIcon       =   "frmproj_produto.frx":1EAAB
         MousePointer    =   99  'Custom
         TabIndex        =   165
         TabStop         =   0   'False
         ToolTipText     =   "Id do cliente."
         Top             =   8250
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2265
         Left            =   -74925
         TabIndex        =   211
         Top             =   1320
         Width           =   15195
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
            TabIndex        =   134
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
            TabIndex        =   133
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   855
         End
         Begin VB.CommandButton Cmd_localizar_doc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frmproj_produto.frx":1EDB5
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Localizar arquivo (F2)"
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_caminho_doc 
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
            Locked          =   -1  'True
            TabIndex        =   135
            TabStop         =   0   'False
            ToolTipText     =   "Caminho."
            Top             =   390
            Width           =   11385
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
            TabIndex        =   138
            ToolTipText     =   "Observação."
            Top             =   1020
            Width           =   14835
         End
         Begin VB.CommandButton Cmd_visualizar_doc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmproj_produto.frx":1EEB7
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "Visualizar arquivo."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"frmproj_produto.frx":1F479
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   450
            TabIndex        =   213
            Top             =   180
            Width           =   9120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   7155
            TabIndex        =   212
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
         Left            =   -70335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   210
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   4110
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtIDGrupo 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Left            =   -74340
         MaxLength       =   50
         MouseIcon       =   "frmproj_produto.frx":1F52D
         MousePointer    =   99  'Custom
         TabIndex        =   203
         Text            =   "0"
         ToolTipText     =   "Tipo."
         Top             =   4500
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -74925
         TabIndex        =   151
         Top             =   1320
         Width           =   15195
         Begin VB.CheckBox chkBloquear_valor 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bloquear valores"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13260
            TabIndex        =   101
            Top             =   330
            Width           =   1515
         End
         Begin VB.TextBox txtPConsumo 
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
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   99
            Text            =   "0,0000"
            ToolTipText     =   "Vlr. da venda para consumo."
            Top             =   390
            Width           =   2115
         End
         Begin VB.TextBox TxtPCusto 
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
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   97
            Text            =   "0,0000"
            ToolTipText     =   "Valor de custo."
            Top             =   390
            Width           =   1965
         End
         Begin VB.TextBox TxtPRevenda 
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
            Left            =   6775
            MaxLength       =   50
            TabIndex        =   100
            Text            =   "0,0000"
            ToolTipText     =   "Vlr. da venda para revenda."
            Top             =   390
            Width           =   2115
         End
         Begin VB.TextBox txtmarglucro 
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
            TabIndex        =   96
            Text            =   "0,00"
            ToolTipText     =   "Margem de lucro em cima do preço de custo."
            Top             =   390
            Width           =   1965
         End
         Begin VB.CommandButton cmdcalc_peso 
            BackColor       =   &H00C0C0C0&
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
            Left            =   4140
            Picture         =   "frmproj_produto.frx":1F837
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Abrir calculadora para cálculo de peso."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor de custo"
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
            Left            =   2745
            TabIndex        =   159
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Valor revendedor"
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
            Index           =   5
            Left            =   7095
            TabIndex        =   158
            Top             =   180
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor consumidor final"
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
            Index           =   4
            Left            =   4710
            TabIndex        =   157
            Top             =   180
            Width           =   1875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marg. lucro"
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
            Left            =   690
            TabIndex        =   156
            Top             =   180
            Width           =   945
         End
      End
      Begin VB.TextBox txtcodproduto 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -66435
         MaxLength       =   50
         MouseIcon       =   "frmproj_produto.frx":1FAA0
         MousePointer    =   99  'Custom
         TabIndex        =   140
         ToolTipText     =   "Numero do desenho / nomenclatura do produto"
         Top             =   5340
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -65460
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -74925
         TabIndex        =   168
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmproj_produto.frx":1FDAA
            Style           =   1  'Graphical
            TabIndex        =   132
            ToolTipText     =   "Visualizar arquivo."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14340
            Picture         =   "frmproj_produto.frx":2036C
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "Limpar caminho."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton cmdImportar 
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
            Left            =   14010
            Picture         =   "frmproj_produto.frx":204AA
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Localizar imagem (F2)"
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   129
            TabStop         =   0   'False
            ToolTipText     =   "Caminho."
            Top             =   390
            Width           =   13815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caminho da imagem"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   6375
            TabIndex        =   169
            Top             =   180
            Width           =   1425
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compras"
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
         TabIndex        =   170
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_limpar_PC 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmproj_produto.frx":205AC
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Limpar conta contábil."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton Cmd_localizar_PC 
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
            Left            =   14340
            Picture         =   "frmproj_produto.frx":206EA
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Localizar plano de contas."
            Top             =   390
            Width           =   315
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
            MouseIcon       =   "frmproj_produto.frx":207EC
            MousePointer    =   99  'Custom
            TabIndex        =   178
            Text            =   "0"
            ToolTipText     =   "ID PC."
            Top             =   390
            Visible         =   0   'False
            Width           =   765
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
            TabIndex        =   122
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   12255
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
            TabIndex        =   121
            TabStop         =   0   'False
            ToolTipText     =   "Código."
            Top             =   390
            Width           =   1875
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   862
            TabIndex        =   179
            Top             =   180
            Width           =   510
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   39
            Left            =   7837
            TabIndex        =   171
            Top             =   180
            Width           =   720
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vendas"
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
         TabIndex        =   187
         Top             =   2160
         Width           =   15195
         Begin VB.CommandButton Cmd_limpar_PC1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmproj_produto.frx":20AF6
            Style           =   1  'Graphical
            TabIndex        =   128
            ToolTipText     =   "Limpar conta contábil."
            Top             =   390
            Width           =   315
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
            TabIndex        =   126
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   12255
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
            TabIndex        =   188
            Text            =   "0"
            ToolTipText     =   "ID PC."
            Top             =   390
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton Cmd_localizar_PC1 
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
            Left            =   14340
            Picture         =   "frmproj_produto.frx":20C34
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Localizar plano de contas."
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
            TabIndex        =   125
            TabStop         =   0   'False
            ToolTipText     =   "Código."
            Top             =   390
            Width           =   1875
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   42
            Left            =   7837
            TabIndex        =   190
            Top             =   180
            Width           =   720
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   41
            Left            =   862
            TabIndex        =   189
            Top             =   180
            Width           =   510
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   3465
         Left            =   -74910
         TabIndex        =   141
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtDiasAvisoVencimento 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8490
            MaxLength       =   255
            TabIndex        =   26
            ToolTipText     =   "Dias para aviso do vencimento da validade do item."
            Top             =   1650
            Width           =   915
         End
         Begin VB.ComboBox cmbClassificacao_produto 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000080&
            Height          =   315
            ItemData        =   "frmproj_produto.frx":20D36
            Left            =   8700
            List            =   "frmproj_produto.frx":20D38
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   253
            ToolTipText     =   "Classificação do produto (tipo do item)."
            Top             =   2280
            Width           =   3105
         End
         Begin VB.Frame Frame17 
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
            Height          =   645
            Left            =   6660
            TabIndex        =   248
            Top             =   2760
            Width           =   8415
            Begin VB.CheckBox chkPerecivel 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Perecível"
               Height          =   315
               Left            =   7410
               TabIndex        =   256
               Top             =   210
               Width           =   945
            End
            Begin VB.CheckBox chkRastreavel 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Rastreavel"
               Height          =   315
               Left            =   6303
               TabIndex        =   255
               Top             =   210
               Width           =   1095
            End
            Begin VB.CheckBox chkEstoque 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Não controla estoque"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   105
               TabIndex        =   252
               Top             =   270
               Width           =   1845
            End
            Begin VB.CheckBox chkProcesso 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Tem processo"
               Enabled         =   0   'False
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   1962
               TabIndex        =   251
               Top             =   270
               Width           =   1305
            End
            Begin VB.CheckBox Chk_insp_recebimento 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inspeção recebimento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   3279
               TabIndex        =   250
               Top             =   270
               Width           =   1905
            End
            Begin VB.CheckBox Chk_tem_plano 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Tem plano"
               Enabled         =   0   'False
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   5196
               TabIndex        =   249
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.TextBox txtDias_antecipacao 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7560
            MaxLength       =   255
            TabIndex        =   25
            ToolTipText     =   "Dias para antecipação da produção."
            Top             =   1650
            Width           =   915
         End
         Begin VB.TextBox txtResponsavelAlteracao 
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
            Left            =   6450
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox txtDataAlteracao 
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
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   990
            Width           =   1755
         End
         Begin VB.TextBox txtDtValidacao 
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
            Left            =   9870
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   990
            Width           =   1755
         End
         Begin VB.TextBox txtRespValidacao 
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
            Left            =   11640
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   990
            Width           =   3375
         End
         Begin VB.TextBox txtstatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   11820
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   2280
            Width           =   3225
         End
         Begin VB.TextBox txtdata 
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
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   990
            Width           =   795
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
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   990
            Width           =   3675
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cód. interno"
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
            Height          =   585
            Left            =   180
            TabIndex        =   194
            Top             =   150
            Width           =   2535
            Begin VB.OptionButton Optautomatico 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Automático"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   0
               Top             =   270
               Width           =   1155
            End
            Begin VB.OptionButton Optmanual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Manual"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1410
               TabIndex        =   1
               Top             =   270
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo"
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
            Height          =   585
            Left            =   7110
            TabIndex        =   193
            Top             =   150
            Width           =   7905
            Begin VB.OptionButton Opt5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Serviço"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6900
               TabIndex        =   11
               Top             =   270
               Width           =   825
            End
            Begin VB.OptionButton Opt4 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Outros"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5850
               TabIndex        =   10
               Top             =   270
               Width           =   915
            End
            Begin VB.OptionButton Opt0 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Matéria-prima"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4320
               TabIndex        =   9
               Top             =   270
               Width           =   1305
            End
            Begin VB.OptionButton opt3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Componente"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2940
               TabIndex        =   8
               Top             =   270
               Width           =   1215
            End
            Begin VB.OptionButton Opt2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Subconjunto"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1560
               TabIndex        =   7
               Top             =   270
               Width           =   1215
            End
            Begin VB.OptionButton Opt1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Produto final"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   6
               Top             =   270
               Width           =   1245
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aplicação"
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
            Height          =   585
            Left            =   2760
            TabIndex        =   192
            Top             =   150
            Width           =   4305
            Begin VB.CheckBox Chk_vendas 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Vendas"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   2
               Top             =   270
               Width           =   825
            End
            Begin VB.CheckBox Chk_compras 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Compras"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1170
               TabIndex        =   3
               Top             =   270
               Width           =   975
            End
            Begin VB.CheckBox Chk_qualidade 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Qualidade"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3060
               TabIndex        =   5
               Top             =   270
               Width           =   1035
            End
            Begin VB.CheckBox Chk_PCP 
               BackColor       =   &H00E0E0E0&
               Caption         =   "PCP"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2280
               TabIndex        =   4
               Top             =   270
               Width           =   645
            End
         End
         Begin VB.TextBox Txt_data_rev 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5055
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Data da revisão."
            Top             =   1650
            Width           =   975
         End
         Begin VB.TextBox txtdesenhoproduto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            ToolTipText     =   "Código interno."
            Top             =   1650
            Width           =   2175
         End
         Begin VB.TextBox txtespecificacao 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            ToolTipText     =   "Descrição comercial."
            Top             =   2850
            Width           =   6435
         End
         Begin VB.TextBox txtreferencia 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2370
            MaxLength       =   50
            TabIndex        =   20
            ToolTipText     =   "Código de referência."
            Top             =   1650
            Width           =   2175
         End
         Begin VB.TextBox txtrevdesproduto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Revisão."
            Top             =   1650
            Width           =   480
         End
         Begin VB.TextBox txtleadtime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6900
            MaxLength       =   255
            TabIndex        =   24
            ToolTipText     =   "Lead time."
            Top             =   1650
            Width           =   645
         End
         Begin MSMask.MaskEdBox Txt_cod_serv 
            Height          =   315
            Left            =   6045
            TabIndex        =   23
            ToolTipText     =   "Código do serviço conforme tabela de Serviços da Lei Complementar 116/2003."
            Top             =   1650
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.##"
            PromptChar      =   "_"
         End
         Begin DrawSuite2022.USButton cmddescricao 
            Height          =   315
            Left            =   6660
            TabIndex        =   246
            ToolTipText     =   "Filtrar por descrição"
            Top             =   2280
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            DibPicture      =   "frmproj_produto.frx":20D3A
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            Theme           =   1
         End
         Begin DrawSuite2022.USButton cmdfamilia 
            Height          =   315
            Left            =   14670
            TabIndex        =   247
            ToolTipText     =   "Filtrar por familia"
            Top             =   1650
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            DibPicture      =   "frmproj_produto.frx":2438A
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            Theme           =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dias venc."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   74
            Left            =   8595
            TabIndex        =   259
            Top             =   1440
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Classificação (Bloco K) | Tipo do item"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   71
            Left            =   8895
            TabIndex        =   254
            Top             =   2070
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dias ante."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   55
            Left            =   7665
            TabIndex        =   228
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. serv."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   54
            Left            =   6075
            TabIndex        =   227
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. revisão"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   5130
            TabIndex        =   226
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   13155
            TabIndex        =   225
            Top             =   2070
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un com."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   7950
            TabIndex        =   224
            Top             =   2070
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un est."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   51
            Left            =   7140
            TabIndex        =   223
            Top             =   2070
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora alteração"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   4830
            TabIndex        =   209
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável pela ultima alteração"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   6930
            TabIndex        =   208
            Top             =   780
            Width           =   2445
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Família"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   11790
            TabIndex        =   200
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   10020
            TabIndex        =   198
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável pela validação"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   12337
            TabIndex        =   197
            Top             =   780
            Width           =   1980
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   390
            TabIndex        =   196
            Top             =   780
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   2355
            TabIndex        =   195
            Top             =   780
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno                Código de referência"
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
            Left            =   630
            TabIndex        =   146
            Top             =   1440
            Width           =   3690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   3090
            TabIndex        =   145
            Top             =   2070
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição comercial"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   2730
            TabIndex        =   144
            Top             =   2640
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   4658
            TabIndex        =   143
            Top             =   1440
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L. time"
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
            Index           =   16
            Left            =   6930
            TabIndex        =   142
            Top             =   1440
            Width           =   570
         End
      End
      Begin MSComctlLib.ListView Lista_doc 
         Height          =   6105
         Left            =   -74925
         TabIndex        =   139
         Top             =   3600
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Caminho"
            Object.Width           =   25576
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista1 
         Height          =   255
         Left            =   -74940
         TabIndex        =   214
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
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   215
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
            Left            =   12720
            Top             =   270
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmproj_produto.frx":279DA
            Count           =   1
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5265
         Left            =   -74925
         TabIndex        =   147
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton cmdKg_un 
            BackColor       =   &H00C0C0C0&
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
            Left            =   13380
            Picture         =   "frmproj_produto.frx":2CDBE
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Calcular Kg/unidade."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_part_number_fabricante 
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
            Left            =   9510
            MaxLength       =   50
            TabIndex        =   42
            ToolTipText     =   "Part number do fabricante."
            Top             =   390
            Width           =   2085
         End
         Begin VB.CommandButton cmdGravacao_padrao 
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
            Height          =   585
            Left            =   14680
            Picture         =   "frmproj_produto.frx":2D027
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Localizar gravação padrão."
            Top             =   3630
            Width           =   315
         End
         Begin VB.CommandButton cmdEmbalagem_padrao 
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
            Height          =   585
            Left            =   9750
            Picture         =   "frmproj_produto.frx":2D129
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Localizar embalagem padrão."
            Top             =   3630
            Width           =   315
         End
         Begin VB.CommandButton cmdInspecao_padrao 
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
            Height          =   585
            Left            =   4770
            Picture         =   "frmproj_produto.frx":2D22B
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Localizar inspeção padrão."
            Top             =   3630
            Width           =   315
         End
         Begin VB.TextBox txtGravacao 
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
            Height          =   585
            Left            =   10185
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
            TabStop         =   0   'False
            ToolTipText     =   "Gravação."
            Top             =   3630
            Width           =   4455
         End
         Begin VB.TextBox txtEmbalagem 
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
            Height          =   585
            Left            =   5190
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            TabStop         =   0   'False
            ToolTipText     =   "Embalagem."
            Top             =   3630
            Width           =   4515
         End
         Begin VB.TextBox txtInspecao 
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
            Height          =   585
            Left            =   175
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "Inspeção."
            Top             =   3630
            Width           =   4545
         End
         Begin VB.TextBox Txt_cod_serv_NFSe 
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
            Left            =   12645
            MaxLength       =   6
            TabIndex        =   78
            ToolTipText     =   "Código do serviço utilizado na NFSe."
            Top             =   2380
            Width           =   2355
         End
         Begin VB.TextBox Txt_GTIN 
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
            Left            =   11205
            MaxLength       =   14
            TabIndex        =   71
            ToolTipText     =   "GTIN (Global Trade Item Number) do produto, antigo código EAN ou código de barras."
            Top             =   2380
            Width           =   1425
         End
         Begin VB.TextBox Txt_qtde_embalagem 
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
            Left            =   2475
            MaxLength       =   255
            TabIndex        =   50
            Text            =   "0,000"
            ToolTipText     =   "Quantidade por embalagem."
            Top             =   1050
            Width           =   1140
         End
         Begin VB.TextBox txtQtde_LoteMinimo 
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
            Left            =   4780
            MaxLength       =   10
            TabIndex        =   52
            Text            =   "0,00"
            ToolTipText     =   "Quantidade lote mínimo."
            Top             =   1050
            Width           =   1140
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
            Left            =   10450
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   57
            ToolTipText     =   "Centro de custo."
            Top             =   1050
            Width           =   4560
         End
         Begin VB.TextBox txtFiname 
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
            MaxLength       =   30
            TabIndex        =   64
            ToolTipText     =   "Código FINAME."
            Top             =   2380
            Width           =   1680
         End
         Begin VB.TextBox Txt_skip_lote 
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
            Left            =   10275
            MaxLength       =   50
            TabIndex        =   70
            ToolTipText     =   "Skip lote."
            Top             =   2380
            Width           =   920
         End
         Begin VB.Frame Frame20 
            BackColor       =   &H00E0E0E0&
            Caption         =   "IMDS"
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
            Height          =   675
            Left            =   3690
            TabIndex        =   177
            Top             =   1440
            Width           =   1320
            Begin VB.CheckBox chkSubmetido 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Submetido"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   61
               Top             =   300
               Width           =   1080
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Classificação"
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
            Height          =   675
            Left            =   180
            TabIndex        =   176
            Top             =   1440
            Width           =   3495
            Begin VB.OptionButton chkimportacao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Importado"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1140
               TabIndex        =   59
               Top             =   300
               Width           =   1065
            End
            Begin VB.OptionButton chknacional 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Nacional"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   58
               Top             =   300
               Width           =   915
            End
            Begin VB.CheckBox chkexportacao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Exportação"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2250
               MouseIcon       =   "frmproj_produto.frx":2D32D
               TabIndex        =   60
               Top             =   300
               Width           =   1125
            End
         End
         Begin VB.TextBox txtPesoBruto 
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
            MaxLength       =   255
            TabIndex        =   48
            Text            =   "0,000000"
            ToolTipText     =   "Peso bruto."
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Frame FramePPAP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "PPAP"
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
            Height          =   675
            Left            =   5040
            TabIndex        =   173
            Top             =   1440
            Width           =   3330
            Begin VB.TextBox txtPPAP_Rev 
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
               Left            =   720
               MaxLength       =   10
               TabIndex        =   62
               ToolTipText     =   "Revisão do PPAP."
               Top             =   270
               Width           =   570
            End
            Begin MSMask.MaskEdBox txtPPAP_Datarev 
               Height          =   315
               Left            =   1830
               TabIndex        =   63
               ToolTipText     =   "Data da revisão do PPAP."
               Top             =   270
               Width           =   1020
               _ExtentX        =   1799
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dt. :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   34
               Left            =   1425
               TabIndex        =   175
               Top             =   270
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rev. :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   33
               Left            =   180
               TabIndex        =   174
               Top             =   270
               Width           =   450
            End
            Begin VB.Image imgCalendario2 
               Height          =   360
               Left            =   2835
               Picture         =   "frmproj_produto.frx":2D637
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   240
               Width           =   330
            End
         End
         Begin VB.TextBox Txt_Dt_ordem 
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
            Left            =   9325
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Data da última ordem."
            Top             =   1050
            Width           =   1110
         End
         Begin VB.TextBox Txt_observacoes 
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
            Height          =   585
            Left            =   175
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   79
            ToolTipText     =   "Observações."
            Top             =   4530
            Width           =   14820
         End
         Begin VB.TextBox txtCor 
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
            Left            =   1874
            MaxLength       =   50
            TabIndex        =   65
            ToolTipText     =   "Cor."
            Top             =   2380
            Width           =   2075
         End
         Begin VB.TextBox txtDureza 
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
            Left            =   8705
            MaxLength       =   50
            TabIndex        =   69
            ToolTipText     =   "Duzera."
            Top             =   2380
            Width           =   1560
         End
         Begin VB.TextBox txtEspessura 
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
            Left            =   3965
            MaxLength       =   30
            TabIndex        =   66
            ToolTipText     =   "Espessura (mm)."
            Top             =   2380
            Width           =   1560
         End
         Begin VB.TextBox txtLargura 
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
            Left            =   5540
            MaxLength       =   30
            TabIndex        =   67
            ToolTipText     =   "Largura (mm)."
            Top             =   2380
            Width           =   1560
         End
         Begin VB.TextBox txtComprimento 
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
            Left            =   7115
            MaxLength       =   30
            TabIndex        =   68
            ToolTipText     =   "Comprimento (mm)."
            Top             =   2380
            Width           =   1580
         End
         Begin VB.CommandButton cmdexcluirfabr 
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
            Left            =   11940
            Picture         =   "frmproj_produto.frx":2DABA
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Excluir fabricante."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtDtVenda 
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
            Left            =   8200
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Data da última venda."
            Top             =   1050
            Width           =   1110
         End
         Begin VB.TextBox txtDtCompra 
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
            Left            =   7085
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "Data da última compra."
            Top             =   1050
            Width           =   1110
         End
         Begin VB.TextBox txtEstReal 
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
            Left            =   5935
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque real."
            Top             =   1050
            Width           =   1140
         End
         Begin VB.TextBox txtEstMinimo 
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
            Left            =   3625
            MaxLength       =   255
            TabIndex        =   51
            Text            =   "0,000"
            ToolTipText     =   "Estoque mínimo."
            Top             =   1050
            Width           =   1140
         End
         Begin VB.TextBox txtPesoLiquido 
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
            Left            =   1320
            MaxLength       =   255
            TabIndex        =   49
            Text            =   "0,000000"
            ToolTipText     =   "Peso líquido."
            Top             =   1050
            Width           =   1140
         End
         Begin VB.TextBox txtpeso 
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
            MaxLength       =   50
            TabIndex        =   45
            ToolTipText     =   "Peso unitário."
            Top             =   390
            Width           =   1005
         End
         Begin VB.ComboBox cmbunkg 
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
            ItemData        =   "frmproj_produto.frx":2DBF8
            Left            =   13995
            List            =   "frmproj_produto.frx":2DC08
            Style           =   2  'Dropdown List
            TabIndex        =   47
            ToolTipText     =   "Unidade por kg."
            Top             =   390
            Width           =   1000
         End
         Begin VB.CommandButton cmdLocalizarFab 
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
            Left            =   11610
            Picture         =   "frmproj_produto.frx":2DC20
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Localizar fabricante."
            Top             =   390
            Width           =   315
         End
         Begin DrawSuite2022.USButton btnAplicacao 
            Height          =   525
            Left            =   14070
            TabIndex        =   257
            ToolTipText     =   "Aplicação do item"
            Top             =   2850
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   926
            DibPicture      =   "frmproj_produto.frx":2DD22
            Caption         =   "Aplicação"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   7
            Theme           =   5
         End
         Begin DrawSuite2022.USButton cmdDescricao_comercial 
            Height          =   525
            Left            =   13110
            TabIndex        =   258
            ToolTipText     =   "Aplicação do item"
            Top             =   2850
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   926
            DibPicture      =   "frmproj_produto.frx":31372
            Caption         =   "Descrição"
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
            PicAlign        =   7
            Theme           =   4
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. serv. NFSe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   73
            Left            =   13230
            TabIndex        =   245
            Top             =   2190
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GTIN"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   11730
            TabIndex        =   244
            Top             =   2190
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skip lote"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   70
            Left            =   10435
            TabIndex        =   243
            Top             =   2190
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dureza"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   69
            Left            =   9230
            TabIndex        =   242
            Top             =   2190
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comprimento / mm"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   68
            Left            =   7238
            TabIndex        =   241
            Top             =   2190
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Largura / mm"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   67
            Left            =   5848
            TabIndex        =   240
            Top             =   2190
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Espessura / mm"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   66
            Left            =   4183
            TabIndex        =   239
            Top             =   2190
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cor"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   65
            Left            =   2784
            TabIndex        =   238
            Top             =   2190
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Centro de custo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   64
            Left            =   12153
            TabIndex        =   237
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última ordem"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   63
            Left            =   9415
            TabIndex        =   236
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última venda"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   62
            Left            =   8290
            TabIndex        =   235
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última compra"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   61
            Left            =   7138
            TabIndex        =   234
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. real"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   60
            Left            =   6205
            TabIndex        =   233
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote mínimo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   59
            Left            =   4930
            TabIndex        =   232
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. mínimo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   58
            Left            =   3790
            TabIndex        =   231
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. p/ emb."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   57
            Left            =   2535
            TabIndex        =   230
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso líquido"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   56
            Left            =   1470
            TabIndex        =   229
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part number"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   10110
            TabIndex        =   220
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gravação"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   12060
            TabIndex        =   218
            Top             =   3420
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Embalagem"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   45
            Left            =   7035
            TabIndex        =   217
            Top             =   3420
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inspeção"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   44
            Left            =   2115
            TabIndex        =   216
            Top             =   3420
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. FINAME"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   540
            TabIndex        =   186
            Top             =   2190
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   7110
            TabIndex        =   167
            Top             =   4320
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso bruto"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   390
            TabIndex        =   155
            Top             =   840
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kg/Un"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   12645
            TabIndex        =   154
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un/Kg"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   14278
            TabIndex        =   153
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   13815
            TabIndex        =   152
            Top             =   465
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante(es) - marca"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   4140
            TabIndex        =   149
            Top             =   180
            Width           =   1635
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74925
         TabIndex        =   160
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtLeadTime_forn 
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
            Left            =   12510
            MaxLength       =   50
            TabIndex        =   106
            ToolTipText     =   "Lead time de compra."
            Top             =   390
            Width           =   675
         End
         Begin VB.CheckBox Chk_cliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cliente(s)"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6360
            TabIndex        =   111
            Top             =   900
            Value           =   1  'Checked
            Width           =   1005
         End
         Begin VB.CheckBox Chk_grupo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Grupo(s)"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1230
            TabIndex        =   109
            Top             =   900
            Width           =   945
         End
         Begin VB.TextBox Txt_ID_CF_cliente 
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
            Left            =   13170
            TabIndex        =   202
            TabStop         =   0   'False
            ToolTipText     =   "ID CF."
            Top             =   1110
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox Txt_CF_cliente 
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
            Left            =   13170
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   118
            TabStop         =   0   'False
            ToolTipText     =   "Classificação fiscal."
            Top             =   1110
            Width           =   1155
         End
         Begin VB.CommandButton Cmd_CF_cliente 
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
            Left            =   14340
            Picture         =   "frmproj_produto.frx":4F477
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
            Top             =   1110
            Width           =   315
         End
         Begin VB.CommandButton cmdFornecedor 
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
            Left            =   11760
            Picture         =   "frmproj_produto.frx":4F579
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Filtrar por fornecedor."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton cmdCliente 
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
            Left            =   9510
            Picture         =   "frmproj_produto.frx":4F994
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Filtrar por cliente."
            Top             =   1110
            Width           =   315
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
            Left            =   3330
            MaxLength       =   50
            TabIndex        =   112
            ToolTipText     =   "Id do cliente."
            Top             =   1110
            Width           =   870
         End
         Begin VB.TextBox txtIdFornecedor 
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
            MaxLength       =   50
            TabIndex        =   102
            TabStop         =   0   'False
            ToolTipText     =   "Id do fornecedor."
            Top             =   390
            Width           =   1045
         End
         Begin VB.CommandButton cmdExcluir_cliente 
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
            Left            =   14685
            Picture         =   "frmproj_produto.frx":4FDAF
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Excluir cliente."
            Top             =   1110
            Width           =   315
         End
         Begin VB.CommandButton cmdLocalizar_cliente 
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
            Left            =   9840
            Picture         =   "frmproj_produto.frx":4FEED
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Localizar cliente."
            Top             =   1110
            Width           =   315
         End
         Begin VB.TextBox txtRevenda_forn 
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
            Left            =   11670
            MaxLength       =   50
            TabIndex        =   117
            Text            =   "0,0000"
            ToolTipText     =   "Vlr. da venda para revenda."
            Top             =   1110
            Width           =   1485
         End
         Begin VB.TextBox txtConsumo_forn 
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
            Left            =   10260
            MaxLength       =   50
            TabIndex        =   116
            Text            =   "0,0000"
            ToolTipText     =   "Vlr. da venda para consumo."
            Top             =   1110
            Width           =   1395
         End
         Begin VB.TextBox txtPcusto_forn 
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
            Left            =   13200
            MaxLength       =   50
            TabIndex        =   107
            Text            =   "0,0000"
            ToolTipText     =   "Vlr. de custo."
            Top             =   390
            Width           =   1485
         End
         Begin VB.CommandButton cmdexcluirforn 
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
            Left            =   14685
            Picture         =   "frmproj_produto.frx":4FFEF
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Excluir fornecedor."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton cmdLocalizarForn 
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
            Left            =   12090
            Picture         =   "frmproj_produto.frx":5012D
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Localizar fornecedor."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L. time"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   47
            Left            =   12607
            TabIndex        =   219
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NCM"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   13582
            TabIndex        =   201
            Top             =   900
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. de consumo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   10395
            TabIndex        =   164
            Top             =   900
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. de revenda"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   11865
            TabIndex        =   163
            Top             =   900
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. de custo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   13485
            TabIndex        =   162
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor(es)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   6060
            TabIndex        =   161
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   148
         Top             =   1350
         Width           =   2475
         Begin VB.CommandButton Cmd_limpar_CF 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1980
            Picture         =   "frmproj_produto.frx":5022F
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Limpar CF."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_ID_CF 
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
            TabIndex        =   80
            TabStop         =   0   'False
            ToolTipText     =   "ID da NCM."
            Top             =   390
            Width           =   525
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
            Left            =   720
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   81
            TabStop         =   0   'False
            ToolTipText     =   "Classificação fiscal."
            Top             =   390
            Width           =   915
         End
         Begin VB.CommandButton CmdCF 
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
            Left            =   1650
            Picture         =   "frmproj_produto.frx":5036D
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   360
            TabIndex        =   221
            Top             =   180
            Width           =   165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NCM"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   1012
            TabIndex        =   150
            Top             =   180
            Width           =   330
         End
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
      FormWidthDT     =   15495
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15375
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frmproj_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Produto As Boolean 'OK
Public Novo_Produto1 As Boolean 'OK
Public Sql_Produto  As String 'OK
Public FormulaRel_Produto  As String 'OK
Public Copiar_Produto As Boolean 'OK
Dim Numero_Abas As Integer 'OK
Dim TBLISTA_Produto As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

If Engenharia_Produtos = True Then FunAbrirVideoWeb ("http://www.youtube.com/watch?v=qD1pO2pfxPw&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=41&feature=plcp")
If Compras_Produtos = True Then FunAbrirVideoWeb ("http://www.youtube.com/watch?v=_jKL3DT2Umc&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=37&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

'ProcCorrigeForm
'SSTab1_Click (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnAplicacao_Click()
On Error GoTo tratar_erro

If txtdesenhoproduto.Text <> "" Then
Desenho = txtdesenhoproduto.Text
DesenhoProduto = txtDescricaoProduto.Text

frmProj_produto_aplicacao.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_cliente_Click()
On Error GoTo tratar_erro

If Chk_cliente.Value = 1 Then
    Chk_grupo.Value = 0
    With txtIDcliente
        .Locked = False
        .TabStop = True
    End With
    With cmbcliente
        .Locked = False
        .TabStop = True
    End With
    cmdCliente.Enabled = True
    cmdLocalizar_cliente.Enabled = True
Else
    With txtIDcliente
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With cmbcliente
        .ListIndex = -1
        .Locked = True
        .TabStop = False
    End With
    cmdCliente.Enabled = False
    cmdLocalizar_cliente.Enabled = False
    txtConsumo_forn = ""
    txtRevenda_forn = ""
    Txt_CF_cliente = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_grupo_Click()
On Error GoTo tratar_erro

With cmbGrupo
    If Chk_grupo.Value = 1 Then
        ProcCarregaComboGrupoCliente cmbGrupo, True
        Chk_cliente.Value = 0
        .Locked = False
        .TabStop = True
    Else
        .ListIndex = -1
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_insp_recebimento_Click()
On Error GoTo tratar_erro

With Txt_skip_lote
    If Chk_insp_recebimento.Value = 1 Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
        .Text = ""
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_PCP_Click()
On Error GoTo tratar_erro

With Opt4
    If Chk_PCP.Value = 1 Then
        .Value = False
        .Enabled = False
    Else
        .Enabled = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_servico_executado_cliente_Click()
On Error GoTo tratar_erro

With Txt_INSS
    If Chk_servico_executado_cliente.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select Compras, Vendas, Fabricacao, Qualidade from projfamilia where Familia = '" & cmbfamilia & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    If TBFamilia!Compras = True Then Chk_compras.Enabled = True Else Chk_compras.Enabled = False
    If TBFamilia!Vendas = True Then Chk_vendas.Enabled = True Else Chk_vendas.Enabled = False
    If TBFamilia!Fabricacao = True Then Chk_PCP.Enabled = True Else Chk_PCP.Enabled = False
    If TBFamilia!Qualidade = True Then Chk_qualidade.Enabled = True Else Chk_qualidade.Enabled = False
End If
TBFamilia.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbunkg_Click()
On Error GoTo tratar_erro

ProcCalculaPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_CF_cliente_Click()
On Error GoTo tratar_erro

Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Clientes = False
Compras_Pedido = False
Familia_NCM = False
ClassFiscal = False
Sit_REG = 1
frmProj_Classificacao_Fiscal.Show 1

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

Private Sub Cmd_limpar_CF_Click()
On Error GoTo tratar_erro

Txt_ID_CF = ""
Txt_CF = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_CFOP_Click()
On Error GoTo tratar_erro

Txt_ID_CFOP = ""
Txt_CFOP = ""
Txt_natureza_operacao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_CFOP1_Click()
On Error GoTo tratar_erro

Txt_ID_CFOP1 = ""
Txt_CFOP1 = ""
Txt_natureza_operacao1 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_PC_Click()
On Error GoTo tratar_erro

Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_PC1_Click()
On Error GoTo tratar_erro

Txt_ID_PC1 = 0
Txt_codigo_PC1 = ""
Txt_descricao_PC1 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_CFOP1_Click()
On Error GoTo tratar_erro

Clientes = False
Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Compras_Pedido = False
Sit_REG = 2
frm_ListaNatureza.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_doc_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_caminho_doc = caminho

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

Private Sub Cmd_localizar_CFOP_Click()
On Error GoTo tratar_erro

Clientes = False
Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Compras_Pedido = False
Sit_REG = 1
frm_ListaNatureza.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_doc_Click()
On Error GoTo tratar_erro

If Txt_caminho_doc <> "" Then ProcAbrirArquivo Txt_caminho_doc

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEmbalagem_padrao_Click()
On Error GoTo tratar_erro

Aplic = 5
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
If Compras_Produtos = True Then
    Sit_REG = 1
ElseIf Vendas_Produtos = True Then
        Sit_REG = 2
    Else
        Sit_REG = 3
End If
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdGravacao_padrao_Click()
On Error GoTo tratar_erro

Aplic = 12
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
If Compras_Produtos = True Then
    Sit_REG = 1
ElseIf Vendas_Produtos = True Then
        Sit_REG = 2
    Else
        Sit_REG = 3
End If
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdInspecao_padrao_Click()
On Error GoTo tratar_erro

Aplic = 4
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
If Compras_Produtos = True Then
    Sit_REG = 1
ElseIf Vendas_Produtos = True Then
        Sit_REG = 2
    Else
        Sit_REG = 3
End If
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Label1_DblClick(index As Integer)
On Error GoTo tratar_erro

If index <> 1 Then Exit Sub
If InputBox("Informe a senha para liberar.") = "280362LIBCOD" Then
    Frame12.Enabled = True
    With txtdesenhoproduto
        .Locked = False
        .TabStop = True
    End With
    With txtDescricaoProduto
        .Locked = False
        .TabStop = True
    End With
    With cmbun
        .Locked = False
        .TabStop = True
    End With
    With Cmb_un_com
        .Locked = False
        .TabStop = True
    End With
    Opt1.Enabled = True
    Opt2.Enabled = True
    opt3.Enabled = True
    Opt4.Enabled = True
    Opt5.Enabled = True
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_doc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_doc
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("projproduto", "codproduto = " & txtcodproduto, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
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

Private Sub Lista_doc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "documento", "excluir este", True, True) = False Then
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

Private Sub Lista_doc_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_doc.ListItems.Count = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select * from projproduto_documentos where ID = " & Lista_doc.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    Proclimpacampos_doc
    ProcCarregaDados_Doc
    CodigoLista1 = Lista_doc.SelectedItem.index
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt0_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt1_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt2_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt3_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPT4_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaAplicacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt5_Click()
On Error GoTo tratar_erro

ProcBloqueiaLiberaAplicacao
If Opt5.Value = True Then
    Frame16.Enabled = True
Else
    Frame16.Enabled = False
    Chk_servico_executado_cliente.Value = 0
    Txt_INSS = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaLiberaAplicacao()
On Error GoTo tratar_erro

With Chk_PCP
    If Opt4.Value = True Then
        .Value = 0
        .Enabled = False
    Else
        .Enabled = True
    End If
End With

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

ProcAtualizalista (1)
With USToolBar1
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(8) = 5
        .ButtonState(14) = 5
    ElseIf Cmb_opcao_lista = "Status" Then
        .ButtonState(4) = 5
        .ButtonState(8) = 0
        .ButtonState(14) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(8) = 5
        .ButtonState(14) = 0
        If Cmb_opcao_lista = "Validação estrutura" Then
            Lista.ColumnHeaders.Item(15).Text = "Valid. est."
        ElseIf Cmb_opcao_lista = "Validação plano de inspeção" Then
            Lista.ColumnHeaders.Item(15).Text = "Val. plan."
        Else
            Lista.ColumnHeaders.Item(15).Text = "Validado"
        End If
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbcliente_Click()
On Error GoTo tratar_erro

If cmbcliente = "" Then Exit Sub
txtConsumo_forn = "0,00000"
txtRevenda_forn = "0,00000"
Txt_ID_CF_cliente = ""
Txt_CF_cliente = ""
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where Desenho = '" & txtdesenhoproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select C.IDCliente, PC.pconsumo, PC.prevenda, PC.ID_CF FROM clientes C INNER JOIN Projproduto_clientes PC ON C.IDCliente = PC.Idcliente where C.IDCliente = " & cmbcliente.ItemData(cmbcliente.ListIndex) & " and PC.Codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        txtIDcliente = TBClientes!IDCliente
        txtConsumo_forn = IIf(IsNull(TBClientes!PConsumo), "0,00000", Format(TBClientes!PConsumo, "###,##0.0000000000"))
        txtRevenda_forn = IIf(IsNull(TBClientes!PRevenda), "0,00000", Format(TBClientes!PRevenda, "###,##0.0000000000"))
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBClientes!ID_CF), 0, TBClientes!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Txt_ID_CF_cliente = TBAbrir!Idclass
            Txt_CF_cliente = IIf(IsNull(TBAbrir!IDIntClasse), "", TBAbrir!IDIntClasse)
        End If
    End If
    TBClientes.Close
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFabricante_Click()
On Error GoTo tratar_erro

If cmbFabricante = "" Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select FM.ID, PF.Part_number FROM Fabricante_marca FM INNER JOIN Projproduto_fabricante PF ON FM.ID = PF.Idfabricante where PF.Codproduto = " & txtcodproduto & " and FM.Fabricante = '" & cmbFabricante & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtIdFabricante = TBFornecedor!ID
    Txt_part_number_fabricante = IIf(IsNull(TBFornecedor!Part_number), "", TBFornecedor!Part_number)
End If
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFornecedor_Click()
On Error GoTo tratar_erro

If cmbFornecedor = "" Then Exit Sub
txtPcusto_forn = "0,00000"
txtLeadTime_forn = ""
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where Desenho = '" & txtdesenhoproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select Compras_fornecedores.IDCliente, Projproduto_fornecedor.pcusto, Projproduto_fornecedor.Leadtime FROM Compras_fornecedores INNER JOIN Projproduto_fornecedor ON Compras_fornecedores.IDCliente = Projproduto_fornecedor.Idfornecedor where Compras_fornecedores.IDCliente = " & cmbFornecedor.ItemData(cmbFornecedor.ListIndex) & " and Projproduto_fornecedor.Codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        txtIDfornecedor = TBFornecedor!IDCliente
        txtLeadTime_forn = IIf(IsNull(TBFornecedor!Leadtime), "", TBFornecedor!Leadtime)
        txtPcusto_forn = IIf(IsNull(TBFornecedor!PCusto), "0,00000", Format(TBFornecedor!PCusto, "###,##0.0000000000"))
    End If
    TBFornecedor.Close
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtcodproduto = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
If Engenharia_Produtos = True Then TBLISTA.Open "Select * from projproduto order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
If Compras_Produtos = True Then TBLISTA.Open "Select * from projproduto where compras = 'True' or consignacao = 'True' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
If Vendas_Produtos = True Then TBLISTA.Open "Select * from projproduto where vendas = 'True' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("CODPRODUTO = " & txtcodproduto)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtcodproduto = TBLISTA!Codproduto
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        USMsgBox ("Fim dos cadastros."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbun_Click()
On Error GoTo tratar_erro

If cmbun = "M³" Then
    Label1(10).Caption = cmbun & " x Un"
    txtpeso.ToolTipText = cmbun & " x Un"
    Label1(38).Caption = "Un x " & cmbun
    cmbunkg.ToolTipText = "Un x " & cmbun
Else
    Label1(10).Caption = "Kg/Un"
    txtpeso.ToolTipText = "Kg/Un"
    Label1(38).Caption = "Un/Kg"
    cmbunkg.ToolTipText = "Un/Kg"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro

Plano_contas_produtos = True
Plano_contas_familias = False
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

Plano_contas_produtos = True
Plano_contas_familias = False
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

Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) registro(s) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmproj_produto_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcalc_peso_Click()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & txtdesenhoproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = True
    Compras_Requisicao = False
    Compras_Cotacao = False
    Compras_Pedido = False
    Estoque_recebimento = False
    Vendas_Proposta = False
    Vendas_PI = False
    FrmCalculo_Peso.Show 1
End If
TBProduto.Close

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
Familia_NCM = False
ClassFiscal = False
Sit_REG = 2
frmProj_Classificacao_Fiscal.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcliente_Click()
On Error GoTo tratar_erro

If cmbcliente = "" Then
    USMsgBox ("Informe o cliente antes de filtrar."), vbExclamation, "CAPRIND v5.0"
    cmbcliente.SetFocus
    Exit Sub
End If
Sql_Produto = "Select P.* FROM Projproduto_clientes PC INNER JOIN Projproduto P ON P.codproduto = PC.codproduto where PC.IDCliente = " & cmbcliente.ItemData(cmbcliente.ListIndex) & " and P.bloqueado = 'False'"
FormulaRel_Produto = "{Projproduto_clientes.IDCliente} = " & cmbcliente.ItemData(cmbcliente.ListIndex) & " and {Projproduto.bloqueado} = False and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
ProcAtualizalista (1)

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
If Novo_Produto = True Then
    USMsgBox ("Salve o registro antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtdesenhoproduto = "" Then
    NomeCampo = "o registro"
    Acao = "copiar"
    ProcVerificaAcao
    Exit Sub
End If
frmProj_Produto_Copiar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmddescricao_Click()
On Error GoTo tratar_erro

If txtDescricaoProduto.Text = "" Then
    USMsgBox ("Informe a descrição antes de filtrar."), vbExclamation, "CAPRIND v5.0"
    txtDescricaoProduto.SetFocus
    Exit Sub
End If
If cmbfamilia <> "" Then
    Sql_Produto = "Select * from projproduto where descricao like '" & txtDescricaoProduto & "%' and classe = '" & cmbfamilia.Text & "' and bloqueado = 'False' order by desenho desc"
    FormulaRel_Produto = "{projproduto.descricao} like '" & txtDescricaoProduto & "*' and {projproduto.classe} = '" & cmbfamilia.Text & "' and {projproduto.bloqueado} = False "
Else
    Sql_Produto = "Select * from projproduto where descricao like '" & txtDescricaoProduto & "%' and bloqueado = 'False' order by desenho desc"
    FormulaRel_Produto = "{projproduto.descricao} like '" & txtDescricaoProduto & "*' and {projproduto.bloqueado} = False "
End If
ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDescricao_comercial_Click()
On Error GoTo tratar_erro

If txtdesenhoproduto = "" Then Exit Sub
Sit_REG = 1
frmproj_produto_descricao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdExcluir_cliente_Click()
On Error GoTo tratar_erro
  
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If cmbcliente.Enabled = False Then Exit Sub
If txtIDcliente.Text = "" Then Exit Sub
If USMsgBox("Deseja realmente excluir este cliente?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifValidacaoRegistro("excluir", txtDtValidacao, "registro", "este cliente", True) = False Then Exit Sub
    Conexao.Execute "DELETE from projproduto_clientes where codproduto = " & txtcodproduto.Text & " and idcliente = " & txtIDcliente
    USMsgBox ("Cliente excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Excluir cliente"
    ID_documento = txtIDcliente
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = "Cliente: " & cmbcliente
    ProcGravaEvento
    '==================================
    ProcSalvarUltimaAlteracao txtcodproduto
    txtIDcliente = ""
    cmbcliente.Clear
    txtConsumo_forn = "0,00000"
    txtRevenda_forn = "0,00000"
    Txt_ID_CF_cliente = ""
    Txt_CF_cliente = ""
    ProcCarregaDadosForn_clientes
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdexcluirfabr_Click()
On Error GoTo tratar_erro
  
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If txtIdFabricante.Text = "" Then Exit Sub
If USMsgBox("Deseja realmente excluir este fabricante?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "DELETE from projproduto_fabricante where codproduto = " & txtcodproduto.Text & " and idfabricante = " & txtIdFabricante
    USMsgBox ("Fabricante excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Excluir fabricante"
    ID_documento = txtIdFabricante
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = "Fabricante: " & cmbFabricante
    ProcGravaEvento
    '==================================
    ProcSalvarUltimaAlteracao txtcodproduto
    txtIdFabricante = ""
    cmbFabricante.Clear
    Txt_part_number_fabricante = ""
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * From projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        ProcCarregaDadosOutros
    End If
    TBProduto.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdExcluirforn_Click()
On Error GoTo tratar_erro
  
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If cmbFornecedor.Enabled = False Then Exit Sub
If txtIDfornecedor.Text = "" Then Exit Sub
If USMsgBox("Deseja realmente excluir este fornecedor?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifValidacaoRegistro("excluir", txtDtValidacao, "registro", "este fornecedor", True) = False Then Exit Sub
    Conexao.Execute "DELETE from projproduto_fornecedor where codproduto = " & txtcodproduto.Text & " and idfornecedor = " & txtIDfornecedor
    USMsgBox ("Fornecedor excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Excluir fornecedor"
    ID_documento = txtIDfornecedor
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = "Fornecedor: " & cmbFornecedor
    ProcGravaEvento
    '==================================
    ProcSalvarUltimaAlteracao txtcodproduto
    txtIDfornecedor = ""
    cmbFornecedor.Clear
    txtLeadTime_forn = ""
    txtPcusto_forn = "0,00000"
    ProcCarregaDadosForn_clientes
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdfamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia.Text = "" Then
    USMsgBox ("Informe a familia antes de filtrar."), vbExclamation, "CAPRIND v5.0"
    cmbfamilia.SetFocus
    Exit Sub
End If
Sql_Produto = "Select * from projproduto where classe = '" & cmbfamilia.Text & "' and bloqueado = 'False' order by desenho desc"
FormulaRel_Produto = "{projproduto.classe} = '" & cmbfamilia.Text & "' and {projproduto.bloqueado} = False "
ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFornecedor_Click()
On Error GoTo tratar_erro

If cmbFornecedor = "" Then
    USMsgBox ("Informe o fornecedor antes de filtrar."), vbExclamation, "CAPRIND v5.0"
    cmbFornecedor.SetFocus
    Exit Sub
End If
Sql_Produto = "Select P.* FROM Projproduto_fornecedor PF INNER JOIN Projproduto P ON P.codproduto = PF.codproduto where PF.IDfornecedor = " & cmbFornecedor.ItemData(cmbFornecedor.ListIndex) & " and P.bloqueado = 'False'"
FormulaRel_Produto = "{Projproduto_fornecedor.IDfornecedor} = " & cmbFornecedor.ItemData(cmbFornecedor.ListIndex) & " and {Projproduto.bloqueado} = False"
ProcAtualizalista (1)

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

Private Sub cmdKg_un_Click()
On Error GoTo tratar_erro

frmproj_produto_kgUN.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_cliente_Click()
On Error GoTo tratar_erro

If cmbcliente.Enabled = False Then Exit Sub
ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
frmVendas_LocalizarCliente.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarFab_Click()
On Error GoTo tratar_erro

frmFabricante_marca.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmdlocalizarforn_Click()
On Error GoTo tratar_erro

If cmbFornecedor.Enabled = False Then Exit Sub
ProcConfVariaveisLocForn False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False
FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtcodproduto = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
If Engenharia_Produtos = True Then TBLISTA.Open "Select * from projproduto order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
If Compras_Produtos = True Then TBLISTA.Open "Select * from projproduto where compras = 'True' or consignacao = 'True' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
If Vendas_Produtos = True Then TBLISTA.Open "Select * from projproduto where vendas = 'True' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.Find ("CODPRODUTO = " & txtcodproduto)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtcodproduto = TBLISTA!Codproduto
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        USMsgBox ("Fim dos cadastros."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtdesenhoproduto = "" Then
    NomeCampo = "o registro"
    Acao = "criar revisão"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Produto = True Then
    USMsgBox ("Salve o registro antes de cadastrar as revisões."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmProj_Produto_Revisao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarPC()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "a conta contábil", "salvar", True, True) = False Then Exit Sub
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Conexao.Execute "Update CC_realizado Set ID_PC = " & IIf(Frame11.Visible = True, Txt_ID_PC1, 0) & " where Cod_produto = " & txtcodproduto & " and ID_PC = " & IIf(IsNull(TBGravar!ID_PC1), 0, TBGravar!ID_PC1)
    TBGravar!ID_PC = IIf(Frame19.Visible = True, Txt_ID_PC, 0) 'Compras
    TBGravar!ID_PC1 = IIf(Frame11.Visible = True, Txt_ID_PC1, 0) 'Vendas
    TBGravar.Update
    
    USMsgBox ("Conta contábil cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Alterar conta contábil"
    ID_documento = txtcodproduto
    Documento = "Cód. interno: " & txtdesenhoproduto
    If Chk_vendas.Value = 1 And Chk_compras.Value = 1 Then
        Documento1 = "Código do plano de vendas: " & Txt_codigo_PC & " - Descrição do plano de vendas: " & Txt_descricao_PC & "Código do plano de compras: " & Txt_codigo_PC1 & " - Descrição do plano de compras: " & Txt_descricao_PC1
    Else
        If Frame19.Visible = True And Txt_ID_PC <> 0 Or Frame11.Visible = True And Txt_ID_PC1 <> 0 Then
            If Frame19.Visible = True And Txt_ID_PC <> 0 Then Documento1 = "Código do plano: " & Txt_codigo_PC & " - Descrição do plano: " & Txt_descricao_PC Else Documento1 = "Código do plano: " & Txt_codigo_PC1 & " - Descrição do plano: " & Txt_descricao_PC1
        Else
            Documento1 = ""
        End If
    End If
    ProcGravaEvento
    '==================================
    
    ProcSalvarUltimaAlteracao txtcodproduto
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
1:
End If
TBGravar.Close

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEstrutura()
On Error GoTo tratar_erro

If txtdesenhoproduto.Text = "" Then
    NomeCampo = "o registro"
    Acao = "gerar estrutura"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Produto = True Then
    USMsgBox ("Salve o registro antes de verificar a estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Opt0.Value = True Then
    USMsgBox ("Não é permitido gerar estrutura deste registro, pois o mesmo é uma matéria-prima."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmproj_conjunto
    .Show
    .Lista.ListItems.Clear
    .ProcLimpaCampos
    .ProcLimpaCamposDescVersao
    .ProcLimpaCamposItem
    .Procatualizadados (txtdesenhoproduto)
    .ProcCarregaVersao ""
    .Novo_Conjunto = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Produto.AbsolutePage <> 2 Then
    If TBLISTA_Produto.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Produto.PageCount - 1)
    Else
        TBLISTA_Produto.AbsolutePage = TBLISTA_Produto.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Produto.AbsolutePage)
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
    TBLISTA_Produto.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Produto.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Produto.AbsolutePage = 1
ProcExibePagina (TBLISTA_Produto.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Produto.AbsolutePage <> -3 Then
    If TBLISTA_Produto.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Produto.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Produto.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Produto.AbsolutePage = TBLISTA_Produto.PageCount
ProcExibePagina (TBLISTA_Produto.AbsolutePage)

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
            Case vbKeyF3: ProcGravar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF8: ProcCodRef
            Case vbKeyF9: ProcEstrutura
            Case vbKeyF10: ProcCopiar
            Case vbKeyF11: ProcRevisao
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: ProcGravarOutros
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF3: ProcGravarImpostos
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyF3: ProcGravarValoresDesc
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyF3: ProcGravarFornCli
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 5:
        Select Case KeyCode
            Case vbKeyF3: ProcGravarPC
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 6:
        Select Case KeyCode
            Case vbKeyF2: cmdImportar_Click
            Case vbKeyF3: ProcGravarImagem
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 7:
        Select Case KeyCode
            Case vbKeyF2: Cmd_localizar_doc_Click
            Case vbKeyF3: ProcGravar_doc
            Case vbKeyF4: procExcluir_doc
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

Sub ProcGravar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Produto = False And Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
'If Chk_vendas.Value = 0 And Chk_compras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0 Then
'    NomeCampo = "a aplicação"
'    ProcVerificaAcao
'    Exit Sub
'End If
If Opt5.Value = False And Opt4.Value = False And opt3.Value = False And Opt2.Value = False And Opt1.Value = False And Opt0.Value = False Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    Exit Sub
End If
If txtdesenhoproduto.Text = "" Then
    If Optmanual.Value = False Then
        ProcVerificaSalvar
        Exit Sub
    Else
        NomeCampo = "o código interno"
        ProcVerificaAcao
        txtdesenhoproduto.SetFocus
        Exit Sub
    End If
End If
If Opt5.Value = True And Txt_cod_serv <> "__.__" Then
    Txt_cod_serv.PromptInclude = False
    If Len(Txt_cod_serv.Text) < 4 Then
        Txt_cod_serv.PromptInclude = True
        USMsgBox ("Verifique se faltam dados no campo código do serviço á serem preenchidos."), vbExclamation, "CAPRIND v5.0"
        Txt_cod_serv.SetFocus
        Exit Sub
    End If
End If
If txtDescricaoProduto.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescricaoProduto.SetFocus
    Exit Sub
End If
If cmbun.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun.SetFocus
    Exit Sub
End If
If Cmb_un_com = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
If txtrevdesproduto.Text = "" Then
    NomeCampo = "a revisão"
    ProcVerificaAcao
    txtrevdesproduto.SetFocus
    Exit Sub
End If
valor = IIf(txtleadtime = "", 0, txtleadtime)
If txtleadtime = "" Or valor < 0 Then
    NomeCampo = "o lead time"
    ProcVerificaAcao
    txtleadtime.SetFocus
    Exit Sub
End If
If txtespecificacao.Text = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    txtespecificacao.SetFocus
    Exit Sub
End If
If cmbfamilia.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If Novo_Produto = True Then
        USMsgBox ("Novo registro cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
        Evento = "Novo"
        ID_documento = txtcodproduto
        Documento = "Cód. interno: " & txtdesenhoproduto
        Documento1 = ""
        ProcGravaEvento
        '==================================
    Else
        If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "mesmo", "este registro", "alterar", True, True) = False Then Exit Sub
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Codproduto from projproduto where codproduto <> " & txtcodproduto & " and Desenho = '" & txtdesenhoproduto & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If USMsgBox("Já existe um registro com este código interno, favor alterar.", vbYesNo, "CAPRIND v5.0") = vbNo Then
                txtdesenhoproduto.SetFocus
                TBAbrir.Close
                Exit Sub
            End If
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Codproduto from projproduto where codproduto <> " & txtcodproduto & " and descricao = '" & txtDescricaoProduto & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If USMsgBox("Já existe um registro com esta descrição, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                txtDescricaoProduto.SetFocus
                TBAbrir.Close
                Exit Sub
            End If
        End If
        
        If txtdesenhoproduto <> TBProduto!Desenho Or txtDescricaoProduto <> TBProduto!Descricao Or cmbun <> TBProduto!Unidade Or Cmb_un_com <> TBProduto!Unidade Or cmbfamilia <> TBProduto!Classe Then
            Conexao.Execute "Update Certificado_qualidade Set desenho = '" & txtdesenhoproduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Certificado_Quimica Set desenho = '" & txtdesenhoproduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update CFI Set codigo_produto = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', Familia = '" & cmbfamilia & "' where codigo_produto = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Compras_pedido_lista Set desenho = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', UN = '" & cmbun & "', Unidade_com = '" & Cmb_un_com & "', Familia = '" & cmbfamilia & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Compras_programa_item Set codigo = '" & txtdesenhoproduto & "' where codigo = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Compras_Programacao set Compras_Programacao.UN = '" & cmbun & "', Compras_Programacao.Unidade_com = '" & Cmb_un_com & "' from Compras_Programacao INNER JOIN Compras_programa_item on Compras_Programacao.ID_item = Compras_programa_item.ID where Compras_programa_item.codigo = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Controle_projetos_historico Set n_cod = '" & txtdesenhoproduto & "', descricao = '" & txtDescricaoProduto & "' where n_cod = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Cotacao_item Set coditem = '" & txtdesenhoproduto & "' where coditem = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Estoque_Controle Set desenho = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', UN = '" & cmbun & "', Classe = '" & cmbfamilia & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Estoque_controle_recebimento Set desenho = '" & txtdesenhoproduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Estoque_Localarmazenamento Set codinterno = '" & txtdesenhoproduto & "' where codinterno = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Estoque_movimentacao Set desenho = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', Familia = '" & cmbfamilia & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Ferramentas Set numero = '" & txtdesenhoproduto & "' where numero = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Instrumentos Set numero = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', Familia = '" & cmbfamilia & "' where numero = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Liquido_penetrante Set desenho = '" & txtdesenhoproduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Medicao Set desenho = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Medicaodimensao_instrumentos Set Instutilizado = '" & txtdesenhoproduto & "' where Instutilizado = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Plano Set desenho = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Planodimensao Set Instrumento = '" & txtdesenhoproduto & "' where Instrumento = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Producao Set desenho = '" & txtdesenhoproduto & "', Produto = '" & txtDescricaoProduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Producao Set Codigo_produto = '" & txtdesenhoproduto & "' where Codigo_produto = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Producao_NF_Consignada Set Codinterno = '" & txtdesenhoproduto & "' where Codinterno = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Producaomaterial Set Codigo = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', Unidade = '" & cmbun & "' where Codigo = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Programas Set desenho = '" & txtdesenhoproduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Projconjunto Set desenho = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', Unidade = '" & cmbun & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update QualidadePPAP Set CodInterno = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "' where CodInterno = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update QualidadePPAP_PlanoControle Set CodInterno = '" & txtdesenhoproduto & "' where CodInterno = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Requisicao_materiais_lista Set desenho = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', UN = '" & cmbun & "', Familia = '" & cmbfamilia & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update tbl_Detalhes_Nota Set int_Cod_Produto = '" & txtdesenhoproduto & "', Txt_descricao = '" & txtDescricaoProduto & "', txt_Unid = '" & cmbun & "', Unidade_com = '" & Cmb_un_com & "', Familia = '" & cmbfamilia & "' where int_Cod_Produto = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update UltraSom Set desenho = '" & txtdesenhoproduto & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update UltraSom Set Metal_adicao = '" & txtdesenhoproduto & "' where Metal_adicao = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update UltraSom Set Metal_base = '" & txtdesenhoproduto & "' where Metal_base = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Vendas_analise Set Codinterno = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', Unidade = '" & cmbun & "', Unidade_com = '" & Cmb_un_com & "', Familia = '" & cmbfamilia & "' where Codinterno = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Vendas_analise_setores Set Codinterno = '" & txtdesenhoproduto & "', Descricao = '" & txtDescricaoProduto & "', Un = '" & cmbun & "', Unidade_com = '" & Cmb_un_com & "', Familia = '" & cmbfamilia & "' where Codinterno = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update vendas_carteira Set desenho = '" & txtdesenhoproduto & "', Descricao_tecnica = '" & txtDescricaoProduto & "', Descricao = '" & txtespecificacao & "', Unidade = '" & cmbun & "', Unidade_com = '" & Cmb_un_com & "', Familia = '" & cmbfamilia & "' where desenho = '" & TBProduto!Desenho & "'"
            Conexao.Execute "Update Vendas_programa_item Set codigo = '" & txtdesenhoproduto & "' where codigo = '" & TBProduto!Desenho & "'"
        End If
        
        'Verifica se a família foi alterada
        If cmbfamilia <> TBProduto!Classe Then
            qt = Len(txtdesenhoproduto)
            If qt > 6 Then
                Set TBFamilia = CreateObject("adodb.recordset")
                TBFamilia.Open "Select * from projfamilia where familia = '" & TBProduto!Classe & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
                
                CompLetra = Len(Letra)
                Letra1 = ""
                If Opt0.Value = True Or Opt1.Value = True Or Opt4.Value = True Or Opt5.Value = True Then
                    If Left(txtdesenhoproduto, CompLetra) = Letra Then Letra1 = Left(txtdesenhoproduto, qt - 6) Else Letra1 = Right(txtdesenhoproduto, qt - 6)
                Else
                    Letra1 = Right(txtdesenhoproduto, qt - 6)
                End If
                'Verifica se o código do produto está vinculado a família selecionada
                Set TBFamilia = CreateObject("adodb.recordset")
                TBFamilia.Open "Select * from projfamilia where Familia = '" & cmbfamilia & "' and Letra = '" & Letra1 & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFamilia.EOF = False Then
                    TBProduto!CodManual = False
                Else
                    TBProduto!CodManual = True
                End If
                TBFamilia.Close
            Else
                TBProduto!CodManual = True
            End If
        Else
            If Optmanual.Value = True Then TBProduto!CodManual = True Else TBProduto!CodManual = False
        End If
        
        If TBProduto!SubTipoItem = 1 And Opt1.Value = False Or TBProduto!SubTipoItem = 2 And Opt2.Value = False Or TBProduto!SubTipoItem = 3 And opt3.Value = False Then
            If Opt1.Value = True Then Conexao.Execute "Update Producao Set Tipo = 'E' where Desenho = '" & txtdesenhoproduto & "'"
            If Opt2.Value = True Then Conexao.Execute "Update Producao Set Tipo = 'M' where Desenho = '" & txtdesenhoproduto & "'"
            If opt3.Value = True Then Conexao.Execute "Update Producao Set Tipo = 'F' where Desenho = '" & txtdesenhoproduto & "'"
        End If
        
        ProcEnviaDados
        TBProduto.Update
        
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
        Evento = "Alterar"
        ID_documento = txtcodproduto
        Documento = "Cód. interno: " & txtdesenhoproduto
        Documento1 = ""
        ProcGravaEvento
        '==================================
        
        ProcSalvarUltimaAlteracao txtcodproduto
        ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
        If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
            Lista.SelectedItem = Lista.ListItems(CodigoLista)
            Lista.SetFocus
        End If
    End If
End If
1:
    TBProduto.Close
    ProcAtualizaConjunto
    
    'Atualiza tipo do processo
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "Select tipo from processos where codproduto = " & txtcodproduto & " and tipo <> 'C'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessos.EOF = False Then
        If Opt1.Value = True Then TBProcessos!Tipo = "E" 'Produto final
        If Opt2.Value = True Then TBProcessos!Tipo = "M" 'Subconjunto
        If opt3.Value = True Then TBProcessos!Tipo = "F" 'Componente
        TBProcessos.Update
    End If
    TBProcessos.Close
    
    If Novo_Produto = True Then
        Sql_Produto = "Select * from projproduto where desenho = '" & txtdesenhoproduto.Text & "'"
        FormulaRel_Produto = "{projproduto.desenho} = '" & txtdesenhoproduto.Text & "' and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
        ProcAtualizalista (1)
    End If
    Novo_Produto = False
    With cmbfamilia
        .Locked = False
        .TabStop = True
    End With
    ProcEsconderMostrarCC

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados()
On Error GoTo tratar_erro

If Chk_vendas.Value = 1 Then TBProduto!Vendas = True Else TBProduto!Vendas = False
If Chk_compras.Value = 1 Then TBProduto!Compras = True Else TBProduto!Compras = False
If Chk_PCP.Value = 1 Then TBProduto!Producao = True Else TBProduto!Producao = False
If Chk_qualidade.Value = 1 Then TBProduto!Qualidade = True Else TBProduto!Qualidade = False
If Opt5.Value = True Then
    TBProduto!Tipo = "S"
    TBProduto!SubTipoItem = 5 'Serviço
Else
    TBProduto!Tipo = "P"
End If
If Opt4.Value = True Then TBProduto!SubTipoItem = 4 'Outros
If opt3.Value = True Then TBProduto!SubTipoItem = 3 'Componente
If Opt2.Value = True Then TBProduto!SubTipoItem = 2 'Subconjunto
If Opt1.Value = True Then TBProduto!SubTipoItem = 1 'Produto acabado
If Opt0.Value = True Then TBProduto!SubTipoItem = 0 'Matéria-prima

txtcodproduto = TBProduto!Codproduto
TBProduto!Desenho = txtdesenhoproduto.Text
TBProduto!Classe = cmbfamilia.Text

'Verifica se já existe cód. de referencia cadastrado/ Salva cod. de referencia
Reiniciar:
    If txtreferencia <> "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where n_referencia = '" & txtreferencia & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Set TBComponente = CreateObject("adodb.recordset")
            TBComponente.Open "Select * from projproduto where codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBComponente.EOF = False Then
                If TBComponente!Desenho <> txtdesenhoproduto Then
                    If USMsgBox("Este código de referência está sendo utilizado em outro(s) produto(s), deseja excluir para ser salvo no produto " & txtdesenhoproduto & "?", vbYesNo) = vbYes Then
                        If USMsgBox("Deseja realmente excluir o código de referência " & txtreferencia & " no(s) outro(s) produto(s)?", vbYesNo) = vbYes Then
                            Conexao.Execute "DELETE from item_aplicacoes where  n_referencia = '" & txtreferencia & "'"
                            TBComponente.Close
                            GoTo Referencia
                        End If
                    End If
                Else
                    TBComponente.Close
                    GoTo Referencia
                End If
            Else
                Conexao.Execute "DELETE from item_aplicacoes where codproduto = " & TBAbrir!Codproduto
                GoTo Reiniciar
            End If
            TBComponente.Close
        Else
Referencia:
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select * from item_aplicacoes where codproduto = " & TBProduto!Codproduto & " and n_referencia = '" & txtreferencia & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = True Then TBItem.AddNew
            ProcEnviaDadosRef
            TBItem.Update
            TBItem.Close
        End If
        TBAbrir.Close
    End If
    
TBProduto!Cod_servico = Txt_cod_serv
TBProduto!Descricao = Trim(txtDescricaoProduto)
TBProduto!Estoque = True
TBProduto!descricaotecnica = Trim(txtespecificacao)
TBProduto!Unidade = cmbun
TBProduto!Unidade_com = Cmb_un_com
TBProduto!RevDesenho = txtrevdesproduto
TBProduto!Leadtime = txtleadtime
TBProduto!Dias_antecipacao = IIf(txtDias_antecipacao = "", Null, txtDias_antecipacao)
TBProduto!DiasAvisoVenc = IIf(txtDiasAvisoVencimento = "", Null, txtDiasAvisoVencimento)

TBProduto!perecivel = chkPerecivel
TBProduto!rastreavel = chkRastreavel

'============================================================
    'Atualiza estoque
    If TBProduto!Estoque = True And chkEstoque.Value = 1 Or TBProduto!Estoque = False And chkEstoque.Value = 0 Then
        'No banco controla estoque na tela não controla
        If TBProduto!Estoque = True And chkEstoque.Value = 1 Then
            Conexao.Execute "Update Estoque_Controle Set Estoque_venda = 0, Estoque_real = 0, valor_unitario = 0, valor_total = 0 where Desenho = '" & txtdesenhoproduto & "'"
        Else
            Entrada = 0
            Saida = 0
            Total = 0
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_Controle where Desenho = '" & TBProduto!Desenho & "' order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_movimentacao where idEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Do While TBAbrir.EOF = False
                            
                            'Verificar se o produto é remessa e marca como não controla estoque
                            Set TBItem = CreateObject("adodb.recordset")
                            TBItem.Open "Select CPL.IDlista from Estoque_controle_recebimento ECR INNER JOIN Compras_pedido_lista CPL ON ECR.IDPedido = CPL.IDPedido and ECR.IdLista = CPL.IdLista and ECR.Desenho = CPL.Desenho where ECR.Id = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento) & " and CPL.remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBItem.EOF = False Then ControlaEstoque = False
                            TBItem.Close
                                                        
                            If TBAbrir!Operacao <> "DEVOLUCAO_ALMOXARIFADO C/ PROB." Then Entrada = Entrada + IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada)
                            Saida = Saida + IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            
                            If TBAbrir!Operacao = "ENTRADA_INVENTÁRIO" Then
                                If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                    'Verif. valor unitário no cadastro do produto
                                    Set TBFI = CreateObject("adodb.recordset")
                                    TBFI.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBFI.EOF = False Then
                                        valor = IIf(IsNull(TBFI!PCusto), 0, TBFI!PCusto)
                                    End If
                                    TBFI.Close
                                End If
                            ElseIf TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Then
                                    'Verif. valor unitário na ordem
                                    Set TBFI = CreateObject("adodb.recordset")
                                    TBFI.Open "Select * from producao where Ordem = " & TBAbrir!LOTE, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBFI.EOF = False Then
                                                                      'ORDEM      QTDE. PREVISTA                          QTDE. OK                                        QT. PROD.(OK+NC)                                                                             CUSTO LOTE                                  CUSTO PEÇA                          CUSTO TERCEIROS                                 CUSTO MATERIAL                                    CUSTO OUTRAS                                  ORDEM CONSIGNADA
                                        valor = FunCalculaValorUnitOrdem(TBFI!Ordem, IIf(IsNull(TBFI!Quant), 0, TBFI!Quant), IIf(IsNull(TBFI!QuantProd), 0, TBFI!QuantProd), IIf(IsNull(TBFI!QuantProd), 0, TBFI!QuantProd) + IIf(IsNull(TBFI!QuantNC), 0, TBFI!QuantNC), IIf(IsNull(TBFI!CTTReal), 0, TBFI!CTTReal), IIf(IsNull(TBFI!CPR), 0, TBFI!CPR), IIf(IsNull(TBFI!CTServico), 0, TBFI!CTServico), IIf(IsNull(TBFI!CTMaterial), 0, TBFI!CTMaterial), IIf(IsNull(TBFI!CTOutras), 0, TBFI!CTOutras), TBFI!Consignacao)
                                        OF = TBFI!Ordem
                                    End If
                                    TBFI.Close
                                ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Then
                                        If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                            'Verif. valor unitário no cadastro do produto
                                            Set TBFI = CreateObject("adodb.recordset")
                                            TBFI.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                            If TBFI.EOF = False Then
                                                valor = IIf(IsNull(TBFI!PCusto), 0, TBFI!PCusto)
                                            End If
                                            TBFI.Close
                                        End If
                                    ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL" Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                            TBFIltro.Open "Select IDlista, ID_empresa from Estoque_controle_recebimento where ID = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento), Conexao, adOpenKeyset, adLockOptimistic
                                            If TBFIltro.EOF = False Then
                                                'Verifica dados da NF
                                                Set TBFI = CreateObject("adodb.recordset")
                                                TBFI.Open "Select NF.ID_empresa, NFP.Int_codigo, NFP.int_Qtd, NFP.Valor_frete, NFP.Valor_seguro, NFP.Valor_acessorias, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, NFP.Total_CSLL_prod, NFP.Total_IRPJ_prod, NFP.dbl_ValorUnitario, NFP.txt_Unid, NFP.Unidade_com from (tbl_Detalhes_Nota NFP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NFP.ID_nota = NF.ID) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo where NFPP.ID_carteira = " & TBFIltro!IDlista & " and NFPP.Codinterno = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                                If TBFI.EOF = False Then
                                                    'Verifica valor do ICMS
                                                    ValorICMS = 0
                                                    Valor1 = 0
                                                    Set TBAliquota = CreateObject("adodb.recordset")
                                                    TBAliquota.Open "Select Valor_ICMS, Valor_ICMS_ST, Valor_ICMS_SN from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & TBFI!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBAliquota.EOF = False Then
                                                        If IsNull(TBAliquota!Valor_ICMS) = False And TBAliquota!Valor_ICMS <> 0 Then
                                                            ValorICMS = TBAliquota!Valor_ICMS
                                                        ElseIf IsNull(TBAliquota!Valor_ICMS_ST) = False And TBAliquota!Valor_ICMS_ST <> 0 Then
                                                                ValorICMS = TBAliquota!Valor_ICMS_ST
                                                            ElseIf IsNull(TBAliquota!Valor_ICMS_SN) = False And TBAliquota!Valor_ICMS_SN <> 0 Then
                                                                    ValorICMS = TBAliquota!Valor_ICMS_SN
                                                        End If
                                                    End If
                                                    If ValorICMS <> 0 Then Valor1 = Format(ValorICMS / TBFI!int_Qtd, "###,##0.0000000000") 'Valor unitário de ICMS
                                                    
                                                    Set TBAliquota = CreateObject("adodb.recordset")
                                                    TBAliquota.Open "Select Simples, Real from Empresa where Codigo = " & TBFIltro!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBAliquota.EOF = False Then
                                                        If TBAliquota!Simples = True Then
                                                            Valor2 = Format(TBFI!Valor_frete / TBFI!int_Qtd, "###,##0.0000000000")
                                                            ValorPagar = Format(TBFI!Valor_seguro / TBFI!int_Qtd, "###,##0.0000000000")
                                                            ValorPago = Format(TBFI!Valor_acessorias / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_PIS_Prod = Format(TBFI!Total_PIS_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_Cofins_Prod = Format(TBFI!Total_Cofins_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_CSLL_Prod = Format(TBFI!Total_CSLL_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_IRPJ_Prod = Format(TBFI!Total_IRPJ_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário - Valor desc.) + (Valor ICMS + Frete + Seguro + Desp. acessórias) + Valor PIS + Valor Cofins + Valor CSLL + Valor IRPJ)
                                                            valor = Format((IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario * FunVerificaTabelaConversaoUnidade(TBFI!txt_Unid, TBFI!Unidade_com)) - QuantsolicitadoN2) + (Valor1 + Valor2 + ValorPagar + ValorPago + Valor_PIS_Prod + Valor_Cofins_Prod + Valor_CSLL_Prod + Valor_IRPJ_Prod), "###,##0.0000000000")
                                                        ElseIf TBAliquota!Real = True Then
                                                                Valor_PIS_Prod = Format(TBFI!Total_PIS_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                                Valor_Cofins_Prod = Format(TBFI!Total_Cofins_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                                'VALOR UNITÁRIO DO ESTOQUE = Valor unitário - (Valor desc. + Valor ICMS + Valor PIS + Valor Cofins)
                                                                valor = Format(IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario * FunVerificaTabelaConversaoUnidade(TBFI!txt_Unid, TBFI!Unidade_com)) - (QuantsolicitadoN2 + Valor1 + Valor_PIS_Prod + Valor_Cofins_Prod), "###,##0.0000000000")
                                                            Else
                                                                'VALOR UNITÁRIO DO ESTOQUE = Valor unitário - (Valor desc. + Valor ICMS)
                                                                valor = Format(IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario * FunVerificaTabelaConversaoUnidade(TBFI!txt_Unid, TBFI!Unidade_com)) - (QuantsolicitadoN2 + Valor1), "###,##0.0000000000")
                                                            End If
                                                    End If
                                                Else
                                                    Set TBPedido = CreateObject("adodb.recordset")
                                                    TBPedido.Open "Select CPL.Quant_Comp, CPL.preco_unitario_desconto, CPL.vlrICMS from Compras_pedido_lista CPL INNER JOIN Compras_comercial CC ON CC.IDpedido = CPL.IDpedido where CPL.IdLista = " & TBFIltro!IDlista & " and CC.Moeda = 'REAL'", Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBPedido.EOF = False Then
                                                        If TBPedido!Quant_Comp <> 0 Then valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto) - (IIf(IsNull(TBPedido!vlrICMS), "0", TBPedido!vlrICMS) / IIf(IsNull(TBPedido!Quant_Comp), "0", TBPedido!Quant_Comp)) Else valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
                                                    End If
                                                    TBPedido.Close
                                                End If
                                                TBFI.Close
                                            End If
                                            TBFIltro.Close
                            End If
                            TBEstoque!valor_unitario = Format(valor, "###,##0.0000000000")
                            TBAbrir!VlrUnit = Format(valor, "###,##0.0000000000")
                            If IsNull(TBAbrir!Entrada) = False And TBAbrir!Entrada <> "0" Then TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada) Else TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            TBAbrir.Update
                            TBAbrir.MoveNext
                        Loop
                    End If
                    
                    Total = Entrada - Saida
                    
                    If TBEstoque!local_armaz = "" Or IsNull(TBEstoque!local_armaz) = True Then TBEstoque!local_armaz = "N/A"
                    TBEstoque!estoque_venda = Total
                    TBEstoque!estoque_real = Total
                    TBEstoque!Valor_total = Format(valor * TBEstoque!estoque_real, "###,##0.00")
                    TBEstoque.Update
                    
                    Entrada = 0
                    Saida = 0
                    Total = 0
                    TBEstoque.MoveNext
                Loop
            End If
            TBEstoque.Close
        End If
    End If
    

    If chkEstoque.Value = 1 Then TBProduto!Estoque = False Else TBProduto!Estoque = True
    If Chk_insp_recebimento.Value = 1 Then TBProduto!Insp_recebimento = True Else TBProduto!Insp_recebimento = False
    If cmbClassificacao_produto <> "" Then TBProduto!ID_Tipo = cmbClassificacao_produto.ItemData(cmbClassificacao_produto.ListIndex) Else TBProduto!ID_Tipo = Null
    
'============================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosRef()
On Error GoTo tratar_erro

TBItem!N_referencia = txtreferencia.Text
TBItem!Codproduto = txtcodproduto.Text
TBItem!Descricao = txtDescricaoProduto.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If Engenharia_Produtos = True Then Caption = "Engenharia - Produtos e serviços (Cód. interno : " & IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho) & " - Rev. : " & IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho) & ")"
If Compras_Produtos = True Then Caption = "Compras - Produtos e serviços (Cód. interno : " & IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho) & " - Rev. : " & IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho) & ")"
If Vendas_Produtos = True Then Caption = "Vendas - Produtos e serviços (Cód. interno : " & IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho) & " - Rev. : " & IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho) & ")"

txtcodproduto.Text = IIf(IsNull(TBProduto!Codproduto), "", TBProduto!Codproduto)
If TBProduto!Vendas = True Then Chk_vendas.Value = 1 Else Chk_vendas.Value = 0
If TBProduto!Compras = True Then Chk_compras.Value = 1 Else Chk_compras.Value = 0
If TBProduto!Producao = True Then Chk_PCP.Value = 1 Else Chk_PCP.Value = 0
If TBProduto!Qualidade = True Then Chk_qualidade.Value = 1 Else Chk_qualidade.Value = 0

If TBProduto!perecivel <> "" Then
If TBProduto!perecivel = True Then chkPerecivel.Value = 1 Else chkPerecivel.Value = 0
Else
chkPerecivel.Value = 0
End If

If TBProduto!rastreavel <> "" Then
If TBProduto!rastreavel = True Then chkRastreavel.Value = 1 Else chkRastreavel.Value = 0
Else
chkRastreavel.Value = 0
End If

Select Case TBProduto!SubTipoItem
    Case 4: Opt4.Value = True 'Insumo
    Case 3: opt3.Value = True 'Componente
    Case 2: Opt2.Value = True 'Subconjunto
    Case 1: Opt1.Value = True 'Produto final
    Case 0: Opt0.Value = True 'Matéria-prima
End Select

If TBProduto!CodManual = False Then
    Optautomatico.Value = True
    Optmanual.Value = False
    If TBProduto!SubTipoItem = 0 Or TBProduto!SubTipoItem = 1 Or TBProduto!SubTipoItem = 4 Or TBProduto!SubTipoItem = 5 Then
        Opt0.Enabled = True
        Opt1.Enabled = True
        Opt2.Enabled = False
        opt3.Enabled = False
        Opt4.Enabled = True
        Opt5.Enabled = True
    Else
        Opt0.Enabled = False
        Opt1.Enabled = False
        Opt2.Enabled = True
        opt3.Enabled = True
        Opt4.Enabled = False
        Opt5.Enabled = False
    End If
Else
    Optautomatico.Value = False
    Optmanual.Value = True
    Opt0.Enabled = True
    Opt1.Enabled = True
    Opt2.Enabled = True
    opt3.Enabled = True
    Opt4.Enabled = True
    Opt5.Enabled = True
End If

NomeCampo = "a família"
If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbfamilia.Text = TBProduto!Classe
2:
    NomeCampo = "a unidade do estoque"
    If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun.Text = TBProduto!Unidade
    NomeCampo = "a unidade comercial"
    If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com.Text = TBProduto!Unidade_com
3:
    Desenho = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
    txtdesenhoproduto = Desenho
    
    If TBProduto!Tipo = "S" Then Opt5.Value = True Else Opt5.Value = False
    
    With txtreferencia
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from item_aplicacoes where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            .Text = IIf(IsNull(TBItem!N_referencia), "", TBItem!N_referencia)
            .Locked = True
        Else
            .Locked = False
        End If
        TBItem.Close
    End With
    
    txtData.Text = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
    txtDescricaoProduto.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    
    txtrevdesproduto.Text = IIf(IsNull(TBProduto!RevDesenho), "", TBProduto!RevDesenho)
    If IsNull(TBProduto!RevDesenho) = False And TBProduto!RevDesenho <> "" Then
        'Verifica data da revisão
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Data from Projproduto_revisao where Codproduto = " & TBProduto!Codproduto & " and Revisao = '" & TBProduto!RevDesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Txt_data_rev = Format(TBAbrir!Data, "dd/mm/yy")
        End If
        TBAbrir.Close
    End If
    
    txtleadtime.Text = IIf(IsNull(TBProduto!Leadtime), "", TBProduto!Leadtime)
    txtDias_antecipacao = IIf(IsNull(TBProduto!Dias_antecipacao), "", TBProduto!Dias_antecipacao)
    txtespecificacao = IIf(IsNull(TBProduto!descricaotecnica), "", (TBProduto!descricaotecnica))
    txtStatus = IIf(TBProduto!Bloqueado = True, "Bloqueado", "Liberado")
    txtResponsavel = IIf(IsNull(TBProduto!Responsavel), "", (TBProduto!Responsavel))
    txtpeso = IIf(IsNull(TBProduto!peso_metro), "", (TBProduto!peso_metro))
    txtDtValidacao = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao)
    txtRespValidacao = IIf(IsNull(TBProduto!RespValidacao), "", TBProduto!RespValidacao)
    txtDataAlteracao = IIf(IsNull(TBProduto!DtAlteracao), "", TBProduto!DtAlteracao)
    txtResponsavelAlteracao = IIf(IsNull(TBProduto!RespAlteracao), "", TBProduto!RespAlteracao)
    
    txtDiasAvisoVencimento.Text = IIf(IsNull(TBProduto!DiasAvisoVenc), "", TBProduto!DiasAvisoVenc)
    
    Novo_Produto = False
    Frame2.Enabled = True
    Frame12.Enabled = False
    
    'Verifica se o registro já foi utilizado e bloqueia os campos de unidade e descrição
    Permitido = True
    If TBProduto!Tipo <> "S" Then ProcVerifProdUtilizado
    If Permitido = False Then
        With txtDescricaoProduto
            If Chk_compras.Value = 0 Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
        With cmbun
            .Locked = True
            .TabStop = False
        End With
        With Cmb_un_com
            .Locked = True
            .TabStop = False
        End With
    Else
        With txtDescricaoProduto
            .Locked = False
            .TabStop = True
        End With
        With cmbun
            .Locked = False
            .TabStop = True
        End With
        With Cmb_un_com
            'If Chk_vendas.Value = 1 And Chk_compras.Value = 0 Then
                .Locked = False
                .TabStop = True
            'Else
                '.Locked = True
                '.TabStop = False
            'End If
        End With
    End If
    
    ProcCarregaDadosOutros
    ProcCarregaDadosImpostos
    ProcCarregaDadosValoresDesc
    ProcCarregaDadosForn_clientes
    ProcCarregaDadosPC
    ProcCarregaDadosImagem
    ProcEsconderMostrarCC
    
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        Novo_Produto = False
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste registro."), vbExclamation, "CAPRIND v5.0"
        If NomeCampo = "a família" Then GoTo 2 Else GoTo 3
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifProdUtilizado()
On Error GoTo tratar_erro

ProcVerificaRegistroUtilizadoSemMsg "Producao", "desenho = '" & txtdesenhoproduto & "'"
If Permitido = False Then Exit Sub
ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido_lista", "desenho = '" & txtdesenhoproduto & "' and (id_cotacao <> 0 or IDPedido <> 0)"
If Permitido = False Then Exit Sub
'ProcVerificaRegistroUtilizadoSemMsg "vendas_carteira", "desenho = '" & txtdesenhoproduto & "'"
'If Permitido = False Then Exit Sub
ProcVerificaRegistroUtilizadoSemMsg "Estoque_controle", "desenho = '" & txtdesenhoproduto & "'"
If Permitido = False Then Exit Sub
ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "int_Cod_Produto = '" & txtdesenhoproduto & "'"
If Permitido = False Then Exit Sub

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluir()
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
                If USMsgBox("Deseja realmente excluir este(s) registro(s), estrutura(s) e seus códigos de referência?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Projproduto WHERE codproduto = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = Formulario
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Cód. interno: " & TBFI!Desenho
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from PROJPRODUTO WHERE codproduto = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from Projconjunto WHERE codproduto = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from Projconjunto_desc_versao WHERE codproduto = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from item_aplicacoes where codproduto = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) registro(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Registro(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcAtualizalista (1)
    Novo_Produto = False
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
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) documento(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            '==================================
            Modulo = Formulario
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & txtdesenhoproduto
            Documento1 = "Caminho: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
            Conexao.Execute "DELETE from projproduto_documentos where ID = " & .ListItems(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) documento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Documento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimpacampos_doc
    ProcCarregaLista_Doc
    Novo_Produto1 = False
    Frame14.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtcodproduto = ""
Optautomatico.Value = False
Optmanual.Value = False
Chk_vendas.Value = 0
Chk_compras.Value = 0
Chk_PCP.Value = 0
Chk_qualidade.Value = 0
Opt5.Value = False
Opt4.Value = False
opt3.Value = False
Opt2.Value = False
Opt1.Value = False
Opt0.Value = False
Txt_cod_serv = "__.__"
Desenho = ""
With txtdesenhoproduto
    .Text = ""
    .Locked = True
    .TabStop = False
End With
txtreferencia.Text = ""
txtData.Text = Format(Date, "dd/mm/yy")
With txtDescricaoProduto
    .Text = ""
    .Locked = False
    .TabStop = True
End With
With cmbun
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With
txtrevdesproduto.Text = 0
Txt_data_rev = ""
txtleadtime.Text = ""
txtDias_antecipacao = ""
txtespecificacao.Text = ""
cmbfamilia.ListIndex = -1
txtStatus = "Liberado"
txtResponsavel.Text = pubUsuario
txtDataAlteracao = ""
txtResponsavelAlteracao = ""
txtDtValidacao = ""
txtRespValidacao = ""

'Outros
txtIdFabricante = ""
cmbFabricante.Clear
Txt_part_number_fabricante = ""
txtpeso = ""
cmbunkg.ListIndex = -1
txtPesoBruto = "0,000000"
txtPesoLiquido = "0,000000"
Txt_qtde_embalagem = "0,0000"
txtEstMinimo = "0,0000"
txtEstReal = "0,0000"
chkEstoque.Value = 0
Chk_insp_recebimento.Value = 1
chkProcesso.Value = 0
Chk_tem_plano.Value = 0
txtDtCompra = ""
txtDtVenda = ""
Txt_Dt_ordem = ""
chknacional.Value = 0
chkimportacao.Value = 0
chkexportacao.Value = 0
txtFiname = ""
txtCor = ""
txtespessura = ""
txtLargura = ""
txtComprimento = ""
txtDureza = ""
Txt_skip_lote = ""
ProcCarregaComboSetor Cmb_centro, "Setor IS NOT NULL and (DtBloq IS NULL or DtBloq = N'') and (Consolidacao = 'False' or Consolidacao is null)", "", False, True, False, "", True, False
cmbGenero.ListIndex = -1
Txt_GTIN = ""
Txt_cod_serv_NFSe = ""
cmbClassificacao_produto.ListIndex = -1
txtinspecao = ""
txtembalagem = ""
txtGravacao = ""
txt_observacoes = ""
txtPPAP_Rev = ""
txtPPAP_Datarev = "__/__/____"
chkSubmetido.Value = 0
txtQtde_LoteMinimo = ""

'Impostos
Txt_ID_CF = ""
Txt_CF = ""
Txt_ID_CFOP = ""
Txt_CFOP = ""
Txt_natureza_operacao = ""
Txt_ID_CFOP1 = ""
Txt_CFOP1 = ""
Txt_natureza_operacao1 = ""
Chk_servico_executado_cliente.Value = 0
Chk_credito_ICMS.Value = 0
Txt_INSS = ""

'Valores e descontos
txtmarglucro = "0,00"
TxtPCusto = "0,00000"
txtPConsumo = "0,00000"
TxtPRevenda = "0,00000"
chkBloquear_valor = 0

'Clientes e fornecedores
txtIDfornecedor = ""
cmbFornecedor.Clear
Chk_grupo.Value = 0
cmbGrupo.ListIndex = -1
Chk_cliente.Value = 0
txtIDcliente = ""
cmbcliente.Clear
txtLeadTime_forn = ""
txtPcusto_forn = "0,00000"
txtConsumo_forn = "0,00000"
txtRevenda_forn = "0,00000"
Txt_ID_CF_cliente = ""
Txt_CF_cliente = ""

'Conta contábil
Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""
Txt_ID_PC1 = 0
Txt_codigo_PC1 = ""
Txt_descricao_PC1 = ""

'Imagem
txt_Caminho = ""

'Documentos
Proclimpacampos_doc

CodigoLista = 0
ProcCarregaUnidades
ProcCarregaFamilia
    
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
Txt_caminho_doc = ""
Txt_obs_doc = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

ProcLimpaVariaveisPrincipais
'ProcCarregaUnidades
'ProcCarregaFamilia
'cmbFamilia.ListIndex = -1
If txtcodproduto <> "" Then ProcCarregaCamposCombo
If Compras_Produtos = True Then
    If txtdesenhoproduto = "" Then Caption = "Compras - Produtos e serviços"
    Formulario = "Compras/Produtos e serviços"
End If
If Vendas_Produtos = True Then
    If txtdesenhoproduto = "" Then Caption = "Vendas - Produtos e serviços"
    Formulario = "Vendas/Produtos e serviços"
End If
If Engenharia_Produtos = True Then
    If txtdesenhoproduto = "" Then Caption = "Engenharia - Produtos e serviços"
    Formulario = "Engenharia/Produtos e serviços"
End If
Direitos

Formulario_produto = Formulario



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from projproduto where Codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Classe) = False And TBAbrir!Classe <> "" Then cmbfamilia.Text = TBAbrir!Classe
    If IsNull(TBAbrir!Unidade) = False And TBAbrir!Unidade <> "" Then cmbun.Text = TBAbrir!Unidade
    If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com.Text = TBAbrir!Unidade_com
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362E" Then frmproj_produto_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmproj_produto_atualizar
        If .Chk1.Value = 1 Then
            'Atualizar tipo dos códigos interno (manual ou automático)
            Set TBFamilia = CreateObject("adodb.recordset")
            TBFamilia.Open "Select Letra, Familia from projfamilia order by familia", Conexao, adOpenKeyset, adLockOptimistic
            If TBFamilia.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBFamilia.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBFamilia.EOF = False
                    CompLetra = Len(TBFamilia!Letra)
                    'Verifica comercial
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select desenho,codmanual from projproduto where classe = '" & TBFamilia!Familia & "' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5) ", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Do While TBProduto.EOF = False
                            Letra = Right(TBProduto!Desenho, CompLetra)
                            If Letra <> TBFamilia!Letra Then TBProduto!CodManual = True Else TBProduto!CodManual = False
                            TBProduto.Update
                            TBProduto.MoveNext
                        Loop
                    End If
                    
                    'Verifica industrial
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select desenho,codmanual from projproduto where classe = '" & TBFamilia!Familia & "' and (subtipoitem = 2 or subtipoitem = 3)", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Do While TBProduto.EOF = False
                            Letra = Left(TBProduto!Desenho, CompLetra)
                            If Letra <> TBFamilia!Letra Then TBProduto!CodManual = True Else TBProduto!CodManual = False
                            TBProduto.Update
                            TBProduto.MoveNext
                        Loop
                    End If
                    TBProduto.Close
                    Contador = Contador + 1
                    PBLista.Value = Contador
                    TBFamilia.MoveNext
                Loop
            End If
            TBFamilia.Close
        End If
        
        If .Chk2.Value = 1 Then
            'Atualizar tipo, controle do processo e estoque
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from projproduto order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBProduto.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBProduto.EOF = False
                    If TBProduto!Unidade <> "SE" And TBProduto!Unidade <> "SV" And TBProduto!Unidade <> "HS" Then TBProduto!Tipo = "P" Else TBProduto!Tipo = "S"
                    
                    'Verifica se tem processo
                    Set TBProcessos = CreateObject("adodb.recordset")
                    TBProcessos.Open "Select * from processos where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProcessos.EOF = False Then
                        TBProduto!Processo = True
                    Else
                        TBProduto!Processo = False
                    End If
                    TBProcessos.Close
                    
                    'Verifica se controla estoque
                    If TBProduto!Unidade <> "SE" And TBProduto!Unidade <> "SV" And TBProduto!Unidade <> "HS" Then TBProduto!Estoque = True Else TBProduto!Estoque = False
                                    
                    TBProduto.Update
                    Contador = Contador + 1
                    PBLista.Value = Contador
                    TBProduto.MoveNext
                Loop
            End If
            TBProduto.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Corrigir códigos internos duplicados
Inicio:
            Desenho = ""
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from projproduto where Codmanual = 'False' order by classe, desenho", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBAbrir.RecordCount
                PBLista.Value = 1
                Contador = 0
                
                Do While TBAbrir.EOF = False
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where Codmanual = 'False' and Codproduto <> " & TBAbrir!Codproduto & " and Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Set TBFamilia = CreateObject("adodb.recordset")
                        TBFamilia.Open "Select * from projfamilia where familia = '" & TBProduto!Classe & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
                        TBFamilia.Close
                        
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "Select * from projproduto where Codmanual = 'False' and Desenho = '" & TBAbrir!Desenho & "' order by Codproduto", Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            TBItem.MoveLast
NovoProduto:
                            Set TBComponente = CreateObject("adodb.recordset")
                            TBComponente.Open "Select * from projproduto where Codproduto <> " & TBItem!Codproduto & " and Classe = '" & TBAbrir!Classe & "' and codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5) order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
                            If TBComponente.EOF = False Then
                                TBComponente.MoveLast
                                CompLetra = Len(Letra)
                                If Left(TBComponente!Desenho, CompLetra) = Letra Then
                                    GoTo NovoItem
                                    Exit Sub
                                End If
                                
                                CompLetra = CompLetra + 1
                                Numero = Left(TBComponente!Desenho, Len(TBComponente!Desenho) - CompLetra)
                                Numero = Numero + 1
                                Select Case Len(Numero)
                                    Case 5: Desenho = Numero & "-" & Letra
                                    Case 4: Desenho = "0" & Numero & "-" & Letra
                                    Case 3: Desenho = "00" & Numero & "-" & Letra
                                    Case 2: Desenho = "000" & Numero & "-" & Letra
                                    Case 1: Desenho = "0000" & Numero & "-" & Letra
                                End Select
                            Else
NovoItem:
                                Set TBComponente = CreateObject("adodb.recordset")
                                TBComponente.Open "Select * from projproduto where Codproduto <> " & TBItem!Codproduto & " and Classe = '" & TBAbrir!Classe & "' and codmanual = 'False' and (subtipoitem = 2 or subtipoitem = 3) order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
                                If TBComponente.EOF = False Then
                                    TBComponente.MoveLast
                                    CompLetra = Len(Letra)
                                    If Right(TBComponente!Desenho, CompLetra) = Letra Then
                                        GoTo NovoProduto
                                        Exit Sub
                                    End If
                                    
                                    CompLetra = CompLetra + 1
                                    Numero = Right(TBComponente!Desenho, Len(TBComponente!Desenho) - CompLetra)
                                    Numero = Numero + 1
                                    Select Case Len(Numero)
                                        Case 5: Desenho = Letra & "-" & Numero
                                        Case 4: Desenho = Letra & "-0" & Numero
                                        Case 3: Desenho = Letra & "-00" & Numero
                                        Case 2: Desenho = Letra & "-000" & Numero
                                        Case 1: Desenho = Letra & "-0000" & Numero
                                    End Select
                                Else
                                    Desenho = Letra & "-00001"
                                End If
                            End If
                            TBComponente.Close
                            Novo_Produto = True
                            cmbfamilia = TBItem!Classe
                            txtdesenhoproduto = Desenho
                            TBItem!Desenho = txtdesenhoproduto
                            TBItem.Update
                            TBItem.Close
                            
                            Novo_Produto = False
                            cmbfamilia.ListIndex = -1
                            txtdesenhoproduto = ""
                            GoTo Inicio
                        End If
                    End If
                    TBProduto.Close
                    
                    Contador = Contador + 1
                    PBLista.Value = Contador
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
                        
            Conexao.Execute "Update CPL set CPL.Codproduto = P.Codproduto from Compras_pedido_lista CPL INNER JOIN Projproduto P on CPL.Desenho = P.Desenho"
            Conexao.Execute "Update ELA set ELA.IDProduto = P.Codproduto from Estoque_Localarmazenamento ELA INNER JOIN Projproduto P on ELA.codinterno = P.Desenho"
            Conexao.Execute "Update QPPAP set QPPAP.IDProduto = P.Codproduto from QualidadePPAP QPPAP INNER JOIN Projproduto P on QPPAP.codinterno = P.Desenho"
            Conexao.Execute "Update QPPAPPC set QPPAPPC.IDProduto = P.Codproduto from QualidadePPAP_PlanoControle QPPAPPC INNER JOIN Projproduto P on QPPAPPC.codinterno = P.Desenho"
            Conexao.Execute "Update NFP set NFP.Codproduto = P.Codproduto from tbl_Detalhes_Nota NFP INNER JOIN Projproduto P on NFP.int_Cod_Produto = P.Desenho"
        End If
        
        If .Chk4.Value = 1 Then
            'Transferir códigos de referência para interno
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select * from item_aplicacoes where n_referencia is not null order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBItem.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBItem.EOF = False
                    If IsNull(TBItem!Codproduto) = False And TBItem!Codproduto <> "" And TBItem!N_referencia <> "" Then
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select * from projproduto where codproduto = " & TBItem!Codproduto & " and Desenho <> '" & TBItem!N_referencia & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = False Then
                            Conexao.Execute "Update Certificado_qualidade Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Certificado_Quimica Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update CFI Set codigo_produto = '" & TBItem!N_referencia & "' where codigo_produto = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Compras_pedido_lista Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Compras_programa_item Set codigo = '" & TBItem!N_referencia & "' where codigo = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Compras_recebimento Set codigo = '" & TBItem!N_referencia & "' where codigo = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Controle_projetos_historico Set n_cod = '" & TBItem!N_referencia & "' where n_cod = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Cotacao_item Set coditem = '" & TBItem!N_referencia & "' where coditem = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Estoque_Controle Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Estoque_controle_recebimento Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Estoque_Localarmazenamento Set codinterno = '" & TBItem!N_referencia & "' where codinterno = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Estoque_movimentacao Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Ferramentas Set numero = '" & TBItem!N_referencia & "' where numero = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Instrumentos Set numero = '" & TBItem!N_referencia & "' where numero = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Liquido_penetrante Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Medicao Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Medicaodimensao_instrumentos Set Instutilizado = '" & TBItem!N_referencia & "' where Instutilizado = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Plano Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Planodimensao Set Instrumento = '" & TBItem!N_referencia & "' where Instrumento = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Producao Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Producao Set Codigo_produto = '" & TBItem!N_referencia & "' where Codigo_produto = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Producao_NF_Consignada Set Codinterno = '" & TBItem!N_referencia & "' where Codinterno = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Producaomaterial Set Codigo = '" & TBItem!N_referencia & "' where Codigo = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Programas Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Projconjunto Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update QualidadePPAP Set CodInterno = '" & TBItem!N_referencia & "' where CodInterno = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update QualidadePPAP_PlanoControle Set CodInterno = '" & TBItem!N_referencia & "' where CodInterno = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Requisicao_materiais_lista Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update tbl_Detalhes_Nota Set int_Cod_Produto = '" & TBItem!N_referencia & "' where int_Cod_Produto = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update tbl_Detalhes_Nota_pedidos Set Codinterno = '" & TBItem!N_referencia & "' where Codinterno = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update UltraSom Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update UltraSom Set Metal_adicao = '" & TBItem!N_referencia & "' where Metal_adicao = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update UltraSom Set Metal_base = '" & TBItem!N_referencia & "' where Metal_base = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Vendas_analise Set Codinterno = '" & TBItem!N_referencia & "' where Codinterno = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Vendas_analise_setores Set Codinterno = '" & TBItem!N_referencia & "' where Codinterno = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update vendas_carteira Set desenho = '" & TBItem!N_referencia & "' where desenho = '" & TBProduto!Desenho & "'"
                            Conexao.Execute "Update Vendas_programa_item Set codigo = '" & TBItem!N_referencia & "' where codigo = '" & TBProduto!Desenho & "'"
                            
                            TBProduto!Desenho = TBItem!N_referencia
                            TBProduto!CodManual = True
                            TBProduto.Update
                        End If
                        TBProduto.Close
                    End If
                    Conexao.Execute "DELETE from item_aplicacoes where iditem = " & TBItem!IDitem
                    
                    Contador = Contador + 1
                    PBLista.Value = Contador
                    TBItem.MoveNext
                Loop
            End If
            TBItem.Close
        End If
        
        If .Chk5.Value = 1 Then
            'Transferir códigos manuais para automático
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select * from projproduto where Codmanual = 'True' and Vendas = 'True' and SubTipoItem = 1 order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBItem.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBItem.EOF = False
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto & " and n_referencia = '" & TBItem!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = True Then
                        TBProduto.AddNew
                        TBProduto!Codproduto = TBItem!Codproduto
                        TBProduto!N_referencia = TBItem!Desenho
                        TBProduto!Descricao = TBItem!Descricao
                        TBProduto.Update
                    End If
                    TBProduto.Close
                    
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from projfamilia where familia = '" & TBItem!Classe & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
                    TBFamilia.Close
                    
                    Set TBComponente = CreateObject("adodb.recordset")
                    TBComponente.Open "Select * from projproduto where classe = '" & TBItem!Classe & "' and codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5) order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
                    If TBComponente.EOF = False Then
                        TBComponente.MoveLast
                        CompLetra = Len(Letra) + 1
                        Numero = Left(TBComponente!Desenho, Len(TBComponente!Desenho) - CompLetra)
                        Numero = Numero + 1
                        Select Case Len(Numero)
                            Case 5: Desenho = Numero & "-" & Letra
                            Case 4: Desenho = "0" & Numero & "-" & Letra
                            Case 3: Desenho = "00" & Numero & "-" & Letra
                            Case 2: Desenho = "000" & Numero & "-" & Letra
                            Case 1: Desenho = "0000" & Numero & "-" & Letra
                        End Select
                        
VerifCodigo:
                        Set TBFIltro = CreateObject("adodb.recordset")
                        TBFIltro.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFIltro.EOF = False Then
                            Set TBFamilia = CreateObject("adodb.recordset")
                            TBFamilia.Open "Select * from projfamilia where familia= '" & TBItem!Classe & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
                            TBFamilia.Close
                            CompLetra = Len(Letra) + 1
                            Numero = Left(Desenho, Len(Desenho) - CompLetra)
                            Numero = Numero + 1
                            Select Case Len(Numero)
                                Case 5: Desenho = Numero & "-" & Letra
                                Case 4: Desenho = "0" & Numero & "-" & Letra
                                Case 3: Desenho = "00" & Numero & "-" & Letra
                                Case 2: Desenho = "000" & Numero & "-" & Letra
                                Case 1: Desenho = "0000" & Numero & "-" & Letra
                            End Select
                            GoTo VerifCodigo
                        End If
                    Else
                        Desenho = "00001" & "-" & Letra
                    End If
                    TBComponente.Close
                                
                    Conexao.Execute "Update Certificado_qualidade Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Certificado_Quimica Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update CFI Set codigo_produto = '" & Desenho & "' where codigo_produto = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Compras_pedido_lista Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Compras_programa_item Set codigo = '" & Desenho & "' where codigo = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Compras_recebimento Set codigo = '" & Desenho & "' where codigo = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Controle_projetos_historico Set n_cod = '" & Desenho & "' where n_cod = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Cotacao_item Set coditem = '" & Desenho & "' where coditem = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Estoque_Controle Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Estoque_controle_recebimento Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Estoque_Localarmazenamento Set codinterno = '" & Desenho & "' where codinterno = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Estoque_movimentacao Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Ferramentas Set numero = '" & Desenho & "' where numero = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Instrumentos Set numero = '" & Desenho & "' where numero = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Liquido_penetrante Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Medicao Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Medicaodimensao_instrumentos Set Instutilizado = '" & Desenho & "' where Instutilizado = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Plano Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Planodimensao Set Instrumento = '" & Desenho & "' where Instrumento = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Producao Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Producao Set Codigo_produto = '" & Desenho & "' where Codigo_produto = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Producao_NF_Consignada Set Codinterno = '" & Desenho & "' where Codinterno = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Producaomaterial Set Codigo = '" & Desenho & "' where Codigo = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Programas Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Projconjunto Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update QualidadePPAP Set CodInterno = '" & Desenho & "' where CodInterno = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update QualidadePPAP_PlanoControle Set CodInterno = '" & Desenho & "' where CodInterno = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Requisicao_materiais_lista Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    'Conexao.Execute "Update tbl_Detalhes_Nota Set int_Cod_Produto = '" & Desenho & "' where int_Cod_Produto = '" & TBItem!Desenho & "'"
                    'Conexao.Execute "Update tbl_Detalhes_Nota_pedidos Set Codinterno = '" & Desenho & "' where Codinterno = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update UltraSom Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update UltraSom Set Metal_adicao = '" & Desenho & "' where Metal_adicao = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update UltraSom Set Metal_base = '" & Desenho & "' where Metal_base = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Vendas_analise Set Codinterno = '" & Desenho & "' where Codinterno = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Vendas_analise_setores Set Codinterno = '" & Desenho & "' where Codinterno = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update vendas_carteira Set desenho = '" & Desenho & "' where desenho = '" & TBItem!Desenho & "'"
                    Conexao.Execute "Update Vendas_programa_item Set codigo = '" & Desenho & "' where codigo = '" & TBItem!Desenho & "'"
                                
                    TBItem!Desenho = Desenho
                    TBItem!CodManual = False
                    TBItem.Update
                    
                    Contador = Contador + 1
                    PBLista.Value = Contador
                    TBItem.MoveNext
                Loop
            End If
            TBItem.Close
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
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

Private Sub Imgcalendario2_Click()
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
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = True
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

Private Sub ProcCodRef()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtdesenhoproduto.Text = "" Then
    NomeCampo = "o registro"
    Acao = "cadastrar códigos de referência"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Produto = True Then
    USMsgBox ("Salve o registro antes de cadastrar códigos de referência."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmproj_produto_referencia.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSimilar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtdesenhoproduto.Text = "" Then
    NomeCampo = "o registro"
    Acao = "cadastrar produtos similares"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Produto = True Then
    USMsgBox ("Salve o registro antes de cadastrar produtos similares."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmproj_produto_similar.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
  
frmProjproduto_Abrir.Show 1

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
cmbfamilia.Locked = False
cmbfamilia.TabStop = True
txtreferencia.Locked = False
frmprojproduto_novo.Show 1

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
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "documento", "criar novo", True, True) = False Then Exit Sub
Proclimpacampos_doc
Novo_Produto1 = True
Frame14.Enabled = True
Cmd_localizar_doc_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Produto = True Then
    If USMsgBox("O registro ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar
        If Novo_Produto = True Then Exit Sub Else Unload Me
    Else
        If txtcodproduto.Text <> "" Then
            Sair = True
            ProcExcluir
        End If
    End If
End If
If Novo_Produto1 = True Then
    If USMsgBox("O documento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar_doc
        If Novo_Produto1 = True Then Exit Sub Else Unload Me
    End If
End If
Conexao.Execute "DELETE from projproduto where Desenho = 'Null'"
Novo_Produto = False
Novo_Produto1 = False
Unload Me
Formulario_produto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarFornCli()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "este cliente/fornecedor", "salvar", True, True) = False Then Exit Sub
Acao = "salvar"
If txtIDfornecedor <> "" Then
    valor = IIf(txtPcusto_forn = "", 0, txtPcusto_forn)
    If valor <= 0 Then
        NomeCampo = "valor de custo"
        ProcVerificaAcao
        TxtPCusto.SetFocus
        Exit Sub
    End If
    
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select * from Projproduto_fornecedor where codproduto = " & txtcodproduto & " and idfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        '==================================
        Modulo = Formulario
        Evento = "Alterar fornecedor"
        ID_documento = txtIDfornecedor
        Documento = "Cód. interno: " & txtdesenhoproduto
        Documento1 = "Fornecedor: " & cmbFornecedor
        ProcGravaEvento
        '==================================
    Else
        TBFornecedor.AddNew
        '==================================
        Modulo = Formulario
        Evento = "Novo fornecedor"
        ID_documento = txtIDfornecedor
        Documento = "Cód. interno: " & txtdesenhoproduto
        Documento1 = "Fornecedor: " & cmbFornecedor
        ProcGravaEvento
        '==================================
    End If
    ProcSalvarUltimaAlteracao txtcodproduto
    
    TBFornecedor!Codproduto = txtcodproduto
    TBFornecedor!IDFornecedor = txtIDfornecedor
    TBFornecedor!PCusto = txtPcusto_forn
    TBFornecedor!Leadtime = IIf(txtLeadTime_forn = "", Null, txtLeadTime_forn)
    TBFornecedor.Update
    TBFornecedor.Close
End If
If Chk_grupo.Value = 1 And cmbGrupo <> "" Or Chk_cliente.Value = 1 And txtIDcliente <> "" Then
    valor = IIf(txtConsumo_forn = "", 0, txtConsumo_forn)
    If txtConsumo_forn = "" Or valor < 0 Then
        NomeCampo = "valor de consumo"
        ProcVerificaAcao
        txtConsumo_forn.SetFocus
        Exit Sub
    End If
    valor = IIf(txtRevenda_forn = "", 0, txtRevenda_forn)
    If txtRevenda_forn = "" Or valor < 0 Then
        NomeCampo = "valor de revenda"
        ProcVerificaAcao
        txtRevenda_forn.SetFocus
        Exit Sub
    End If
    
    If Chk_grupo.Value = 1 And cmbGrupo <> "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select IDCliente, NomeRazao from Clientes where IDgrupo = " & cmbGrupo.ItemData(cmbGrupo.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                ProcGravarCliente TBAbrir!IDCliente, TBAbrir!NomeRazao
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        ProcCarregaComboCliForn cmbcliente, True
    Else
        ProcGravarCliente txtIDcliente, cmbcliente
    End If
End If
If txtIDfornecedor <> "" Or Chk_grupo.Value = 1 And cmbGrupo <> "" Or Chk_cliente.Value = 1 And txtIDcliente <> "" Then
    USMsgBox ("Registro salvo com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
1:
Else
    If Chk_grupo.Value = 1 And cmbGrupo = "" Then
        USMsgBox ("Informe o grupo antes de salvar."), vbExclamation, "CAPRIND v5.0"
        cmbGrupo.SetFocus
    Else
        USMsgBox ("Informe o fornecedor ou o cliente antes de salvar."), vbExclamation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarCliente(IDCliente As Long, Cliente As String)
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Projproduto_clientes where codproduto = " & txtcodproduto & " and idcliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    '==================================
    Modulo = Formulario
    Evento = "Alterar cliente"
    ID_documento = IDCliente
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = "Cliente: " & Cliente
    ProcGravaEvento
    '==================================
Else
    TBClientes.AddNew
    '==================================
    Modulo = Formulario
    Evento = "Novo cliente"
    ID_documento = IDCliente
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = "Cliente: " & Cliente
    ProcGravaEvento
    '==================================
End If
TBClientes!Codproduto = txtcodproduto
TBClientes!IDCliente = IDCliente
TBClientes!PConsumo = txtConsumo_forn
TBClientes!PRevenda = txtRevenda_forn
TBClientes!ID_CF = IIf(Txt_ID_CF_cliente = "", Null, Txt_ID_CF_cliente)

TBClientes.Update
TBClientes.Close
ProcSalvarUltimaAlteracao txtcodproduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarOutros()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "os dados adicionais", "salvar", True, True) = False Then Exit Sub
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "este registro", True) = False Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Salvar outros"
    ID_documento = txtcodproduto
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = ""
    ProcGravaEvento
    '==================================
    If txtIdFabricante <> "" Then
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from Projproduto_fabricante where codproduto = " & txtcodproduto & " and idfabricante = " & txtIdFabricante, Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = True Then TBFornecedor.AddNew
        TBFornecedor!Codproduto = txtcodproduto
        TBFornecedor!Idfabricante = txtIdFabricante
        TBFornecedor!Part_number = Txt_part_number_fabricante
        TBFornecedor.Update
        TBFornecedor.Close
    End If
    
    If FunLiberaCamposEstrutura = False Then
        If txtpeso <> Format(TBProduto!peso_metro, "###,##0.0000000000") Or cmbunkg <> TBProduto!Un_Kg Then
            valor = IIf(txtpeso = "", 0, txtpeso)
            NovoValor = Replace(valor, ",", ".")
            Conexao.Execute "UPDATE Projconjunto Set Pesometro = " & NovoValor & ", Un_kg = '" & IIf(cmbunkg = "", "N/a", cmbunkg) & "' where Desenho = '" & txtdesenhoproduto & "'"
            If cmbunkg <> "" Then
                Select Case cmbunkg
                    Case "Mt/L": NomeCampoFiltro = "Peso = ROUND((Pesometro / 1000) * Dimensoes, 5)"
                    Case "Pç": NomeCampoFiltro = "Peso = ROUND(Pesometro, 5)"
                    Case "Mt²": NomeCampoFiltro = "Peso = ROUND(((Dimensoes * Pesometro) / 1000) / 1000, 5)"
                    Case "N/a": NomeCampoFiltro = "Peso = 0"
                End Select
                Conexao.Execute "UPDATE Projconjunto Set " & NomeCampoFiltro & " where Desenho = '" & txtdesenhoproduto & "'"
                Conexao.Execute "UPDATE Projconjunto Set Pesototal = ROUND(Peso * Quantidade, 5) where Desenho = '" & txtdesenhoproduto & "'"
            End If
            Select Case cmbun
                Case "KG": NomeCampoFiltro = "Valortotal = ROUND(Valor * Pesototal, 2)"
                Case "MM": NomeCampoFiltro = "Valortotal = ROUND((Valor * Dimensoes) * Quantidade, 2)"
                Case "MT": NomeCampoFiltro = "Valortotal = ROUND(((Valor / 1000) * Dimensoes) * Quantidade, 2)"
            End Select
            If cmbun <> "KG" And cmbun <> "MM" And cmbun <> "MT" Then NomeCampoFiltro = "Valortotal = ROUND(Valor * Quantidade, 2)"
            Conexao.Execute "UPDATE Projconjunto Set " & NomeCampoFiltro & " where Desenho = '" & txtdesenhoproduto & "'"
        End If
    End If
        
    TBProduto!peso_metro = IIf(txtpeso = "", 0, txtpeso)
    TBProduto!Un_Kg = IIf(cmbunkg = "", Null, cmbunkg)

        
    TBProduto!Qtde_embalagem = IIf(Txt_qtde_embalagem = "", Null, Txt_qtde_embalagem)
    
    valor = IIf(txtPesoBruto = "", 0, txtPesoBruto)
    If TBProduto!PBruto <> valor And cmbun = "KG" Then
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select Estoque_real, Estoque_real_PC, Un from Estoque_Controle where Desenho = '" & TBProduto!Desenho & "' and Estoque_real > 0 and Estoque_real_PC = 0 order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            Do While TBEstoque.EOF = False
                TBEstoque!estoque_real_PC = FunCalculaQtdePC(TBProduto!Desenho, TBEstoque!estoque_real, True, TBEstoque!Un)
                TBEstoque.Update
                TBEstoque.MoveNext
            Loop
        End If
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select EM.Entrada, EM.Entrada_PC, EM.Saida, EM.Saida_PC, EC.Un from Estoque_movimentacao EM INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque where EC.Desenho = '" & TBProduto!Desenho & "' and EC.Estoque_real > 0 and (EM.Entrada > 0 and EM.Entrada_PC = 0 or EM.Saida > 0 and EM.Saida_PC = 0) order by EC.idEstoque", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            Do While TBEstoque.EOF = False
                If TBEstoque!Entrada > 0 Then TBEstoque!Entrada_PC = FunCalculaQtdePC(TBProduto!Desenho, TBEstoque!Entrada, True, TBEstoque!Un) Else TBEstoque!Saida_PC = FunCalculaQtdePC(TBProduto!Desenho, TBEstoque!Saida, True, TBEstoque!Un)
                TBEstoque.Update
                TBEstoque.MoveNext
            Loop
        End If
    End If
    
    'Atualiza estoque
    If TBProduto!Estoque = True And chkEstoque.Value = 1 Or TBProduto!Estoque = False And chkEstoque.Value = 0 Then
        'No banco controla estoque na tela não controla
        If TBProduto!Estoque = True And chkEstoque.Value = 1 Then
            Conexao.Execute "Update Estoque_Controle Set Estoque_venda = 0, Estoque_real = 0, valor_unitario = 0, valor_total = 0 where Desenho = '" & txtdesenhoproduto & "'"
        Else
            Entrada = 0
            Saida = 0
            Total = 0
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_Controle where Desenho = '" & TBProduto!Desenho & "' order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_movimentacao where idEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Do While TBAbrir.EOF = False
                            
                            'Verificar se o produto é remessa e marca como não controla estoque
                            Set TBItem = CreateObject("adodb.recordset")
                            TBItem.Open "Select CPL.IDlista from Estoque_controle_recebimento ECR INNER JOIN Compras_pedido_lista CPL ON ECR.IDPedido = CPL.IDPedido and ECR.IdLista = CPL.IdLista and ECR.Desenho = CPL.Desenho where ECR.Id = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento) & " and CPL.remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBItem.EOF = False Then ControlaEstoque = False
                            TBItem.Close
                                                        
                            If TBAbrir!Operacao <> "DEVOLUCAO_ALMOXARIFADO C/ PROB." Then Entrada = Entrada + IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada)
                            Saida = Saida + IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            
                            If TBAbrir!Operacao = "ENTRADA_INVENTÁRIO" Then
                                If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                    'Verif. valor unitário no cadastro do produto
                                    Set TBFI = CreateObject("adodb.recordset")
                                    TBFI.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBFI.EOF = False Then
                                        valor = IIf(IsNull(TBFI!PCusto), 0, TBFI!PCusto)
                                    End If
                                    TBFI.Close
                                End If
                            ElseIf TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Then
                                    'Verif. valor unitário na ordem
                                    Set TBFI = CreateObject("adodb.recordset")
                                    TBFI.Open "Select * from producao where Ordem = " & TBAbrir!LOTE, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBFI.EOF = False Then
                                                                      'ORDEM      QTDE. PREVISTA                          QTDE. OK                                        QT. PROD.(OK+NC)                                                                             CUSTO LOTE                                  CUSTO PEÇA                          CUSTO TERCEIROS                                 CUSTO MATERIAL                                    CUSTO OUTRAS                                  ORDEM CONSIGNADA
                                        valor = FunCalculaValorUnitOrdem(TBFI!Ordem, IIf(IsNull(TBFI!Quant), 0, TBFI!Quant), IIf(IsNull(TBFI!QuantProd), 0, TBFI!QuantProd), IIf(IsNull(TBFI!QuantProd), 0, TBFI!QuantProd) + IIf(IsNull(TBFI!QuantNC), 0, TBFI!QuantNC), IIf(IsNull(TBFI!CTTReal), 0, TBFI!CTTReal), IIf(IsNull(TBFI!CPR), 0, TBFI!CPR), IIf(IsNull(TBFI!CTServico), 0, TBFI!CTServico), IIf(IsNull(TBFI!CTMaterial), 0, TBFI!CTMaterial), IIf(IsNull(TBFI!CTOutras), 0, TBFI!CTOutras), TBFI!Consignacao)
                                        OF = TBFI!Ordem
                                    End If
                                    TBFI.Close
                                ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Then
                                        If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                            'Verif. valor unitário no cadastro do produto
                                            Set TBFI = CreateObject("adodb.recordset")
                                            TBFI.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                            If TBFI.EOF = False Then
                                                valor = IIf(IsNull(TBFI!PCusto), 0, TBFI!PCusto)
                                            End If
                                            TBFI.Close
                                        End If
                                    ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL" Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                            TBFIltro.Open "Select IDlista, ID_empresa from Estoque_controle_recebimento where ID = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento), Conexao, adOpenKeyset, adLockOptimistic
                                            If TBFIltro.EOF = False Then
                                                'Verifica dados da NF
                                                Set TBFI = CreateObject("adodb.recordset")
                                                TBFI.Open "Select NF.ID_empresa, NFP.Int_codigo, NFP.int_Qtd, NFP.Valor_frete, NFP.Valor_seguro, NFP.Valor_acessorias, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, NFP.Total_CSLL_prod, NFP.Total_IRPJ_prod, NFP.dbl_ValorUnitario, NFP.txt_Unid, NFP.Unidade_com from (tbl_Detalhes_Nota NFP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NFP.ID_nota = NF.ID) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo where NFPP.ID_carteira = " & TBFIltro!IDlista & " and NFPP.Codinterno = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                                If TBFI.EOF = False Then
                                                    'Verifica valor do ICMS
                                                    ValorICMS = 0
                                                    Valor1 = 0
                                                    Set TBAliquota = CreateObject("adodb.recordset")
                                                    TBAliquota.Open "Select Valor_ICMS, Valor_ICMS_ST, Valor_ICMS_SN from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & TBFI!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBAliquota.EOF = False Then
                                                        If IsNull(TBAliquota!Valor_ICMS) = False And TBAliquota!Valor_ICMS <> 0 Then
                                                            ValorICMS = TBAliquota!Valor_ICMS
                                                        ElseIf IsNull(TBAliquota!Valor_ICMS_ST) = False And TBAliquota!Valor_ICMS_ST <> 0 Then
                                                                ValorICMS = TBAliquota!Valor_ICMS_ST
                                                            ElseIf IsNull(TBAliquota!Valor_ICMS_SN) = False And TBAliquota!Valor_ICMS_SN <> 0 Then
                                                                    ValorICMS = TBAliquota!Valor_ICMS_SN
                                                        End If
                                                    End If
                                                    If ValorICMS <> 0 Then Valor1 = Format(ValorICMS / TBFI!int_Qtd, "###,##0.0000000000") 'Valor unitário de ICMS
                                                    
                                                    Set TBAliquota = CreateObject("adodb.recordset")
                                                    TBAliquota.Open "Select Simples, Real from Empresa where Codigo = " & TBFIltro!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBAliquota.EOF = False Then
                                                        If TBAliquota!Simples = True Then
                                                            Valor2 = Format(TBFI!Valor_frete / TBFI!int_Qtd, "###,##0.0000000000")
                                                            ValorPagar = Format(TBFI!Valor_seguro / TBFI!int_Qtd, "###,##0.0000000000")
                                                            ValorPago = Format(TBFI!Valor_acessorias / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_PIS_Prod = Format(TBFI!Total_PIS_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_Cofins_Prod = Format(TBFI!Total_Cofins_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_CSLL_Prod = Format(TBFI!Total_CSLL_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            Valor_IRPJ_Prod = Format(TBFI!Total_IRPJ_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                            'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário - Valor desc.) + (Valor ICMS + Frete + Seguro + Desp. acessórias) + Valor PIS + Valor Cofins + Valor CSLL + Valor IRPJ)
                                                            valor = Format((IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario * FunVerificaTabelaConversaoUnidade(TBFI!txt_Unid, TBFI!Unidade_com)) - QuantsolicitadoN2) + (Valor1 + Valor2 + ValorPagar + ValorPago + Valor_PIS_Prod + Valor_Cofins_Prod + Valor_CSLL_Prod + Valor_IRPJ_Prod), "###,##0.0000000000")
                                                        ElseIf TBAliquota!Real = True Then
                                                                Valor_PIS_Prod = Format(TBFI!Total_PIS_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                                Valor_Cofins_Prod = Format(TBFI!Total_Cofins_prod / TBFI!int_Qtd, "###,##0.0000000000")
                                                                'VALOR UNITÁRIO DO ESTOQUE = Valor unitário - (Valor desc. + Valor ICMS + Valor PIS + Valor Cofins)
                                                                valor = Format(IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario * FunVerificaTabelaConversaoUnidade(TBFI!txt_Unid, TBFI!Unidade_com)) - (QuantsolicitadoN2 + Valor1 + Valor_PIS_Prod + Valor_Cofins_Prod), "###,##0.0000000000")
                                                            Else
                                                                'VALOR UNITÁRIO DO ESTOQUE = Valor unitário - (Valor desc. + Valor ICMS)
                                                                valor = Format(IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario * FunVerificaTabelaConversaoUnidade(TBFI!txt_Unid, TBFI!Unidade_com)) - (QuantsolicitadoN2 + Valor1), "###,##0.0000000000")
                                                            End If
                                                    End If
                                                Else
                                                    Set TBPedido = CreateObject("adodb.recordset")
                                                    TBPedido.Open "Select CPL.Quant_Comp, CPL.preco_unitario_desconto, CPL.vlrICMS from Compras_pedido_lista CPL INNER JOIN Compras_comercial CC ON CC.IDpedido = CPL.IDpedido where CPL.IdLista = " & TBFIltro!IDlista & " and CC.Moeda = 'REAL'", Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBPedido.EOF = False Then
                                                        If TBPedido!Quant_Comp <> 0 Then valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto) - (IIf(IsNull(TBPedido!vlrICMS), "0", TBPedido!vlrICMS) / IIf(IsNull(TBPedido!Quant_Comp), "0", TBPedido!Quant_Comp)) Else valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
                                                    End If
                                                    TBPedido.Close
                                                End If
                                                TBFI.Close
                                            End If
                                            TBFIltro.Close
                            End If
                            TBEstoque!valor_unitario = Format(valor, "###,##0.0000000000")
                            TBAbrir!VlrUnit = Format(valor, "###,##0.0000000000")
                            If IsNull(TBAbrir!Entrada) = False And TBAbrir!Entrada <> "0" Then TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada) Else TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            TBAbrir.Update
                            TBAbrir.MoveNext
                        Loop
                    End If
                    
                    Total = Entrada - Saida
                    
                    If TBEstoque!local_armaz = "" Or IsNull(TBEstoque!local_armaz) = True Then TBEstoque!local_armaz = "N/A"
                    TBEstoque!estoque_venda = Total
                    TBEstoque!estoque_real = Total
                    TBEstoque!Valor_total = Format(valor * TBEstoque!estoque_real, "###,##0.00")
                    TBEstoque.Update
                    
                    Entrada = 0
                    Saida = 0
                    Total = 0
                    TBEstoque.MoveNext
                Loop
            End If
            TBEstoque.Close
        End If
    End If
    
    TBProduto!PBruto = IIf(txtPesoBruto = "", Null, txtPesoBruto)
    TBProduto!PLiquido = IIf(txtPesoLiquido = "", Null, txtPesoLiquido)
    TBProduto!Estoque_minimo = IIf(txtEstMinimo = "", 0, txtEstMinimo)
    If chkEstoque.Value = 1 Then TBProduto!Estoque = False Else TBProduto!Estoque = True
    If Chk_insp_recebimento.Value = 1 Then TBProduto!Insp_recebimento = True Else TBProduto!Insp_recebimento = False
    If chknacional.Value = True Then TBProduto!nacional = True Else TBProduto!nacional = False
    If chkimportacao.Value = True Then TBProduto!importacao = True Else TBProduto!importacao = False
    If chkexportacao.Value = 1 Then TBProduto!exportacao = True Else TBProduto!exportacao = False
    TBProduto!FINAME = txtFiname
    TBProduto!Cor = txtCor
    TBProduto!Comprimento = IIf(txtComprimento = "", Null, txtComprimento)
    TBProduto!Largura = IIf(txtLargura = "", Null, txtLargura)
    TBProduto!Espessura = IIf(txtespessura = "", Null, txtespessura)
    TBProduto!Dureza = txtDureza
    TBProduto!Skip_lote = IIf(Txt_skip_lote = "", Null, Txt_skip_lote)
    If Cmb_centro <> "" Then TBProduto!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex) Else TBProduto!ID_CC = Null
'    If cmbGenero <> "" Then TBProduto!ID_Genero = cmbGenero.ItemData(cmbGenero.ListIndex) Else TBProduto!ID_Genero = Null
    TBProduto!GTIN = IIf(Txt_GTIN = "", Null, Txt_GTIN)
    TBProduto!Cod_servico_NFSE = IIf(Txt_cod_serv_NFSe = "", Null, Txt_cod_serv_NFSe)
    If cmbClassificacao_produto <> "" Then TBProduto!ID_Tipo = cmbClassificacao_produto.ItemData(cmbClassificacao_produto.ListIndex) Else TBProduto!ID_Tipo = Null
    TBProduto!Observacoes = txt_observacoes
    TBProduto!Inspecao = txtinspecao
    TBProduto!Embalagem = txtembalagem
    TBProduto!Gravacao = txtGravacao
    TBProduto!PPAP_Rev = txtPPAP_Rev
    TBProduto!PPAP_Datarev = IIf(txtPPAP_Datarev = "__/__/____", Null, txtPPAP_Datarev)
    If chkSubmetido.Value = 1 Then TBProduto!Submetido = True Else TBProduto!Submetido = False
    TBProduto!qtde_LoteMinimo = IIf(txtQtde_LoteMinimo = "", Null, txtQtde_LoteMinimo)
End If

TBProduto.Update
TBProduto.Close
ProcSalvarUltimaAlteracao txtcodproduto
ProcAtualizaConjunto
ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarImpostos()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "os impostos", "salvar", True, True) = False Then Exit Sub

If Chk_servico_executado_cliente.Value = 1 And Txt_INSS = "" Then
    NomeCampo = "a alíquota do INSS"
    ProcVerificaAcao
    Txt_INSS.SetFocus
    Exit Sub
End If

If cmbGenero.Text = "" Then
    USMsgBox "Atenção!!! " & vbCrLf & "É obrigatório indicar o gênero do produto em relação ao bloco K", vbInformation, "CAPRIND v5.0"
    cmbGenero.SetFocus
    Exit Sub
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Salvar impostos"
    ID_documento = txtcodproduto
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = ""
    ProcGravaEvento
        '==================================================
    ' Novos artnel
    '==================================================
        TBProduto!uTrib = IIf(cmbuTrib = "", Null, cmbuTrib)
        TBProduto!vTrib = txtvuTrib.Text
    '==================================
    If Chk_servico_executado_cliente.Value = 1 Then TBProduto!Servico_cliente = True Else TBProduto!Servico_cliente = False
    If Chk_credito_ICMS.Value = 1 Then TBProduto!Credita_ICMS = False Else TBProduto!Credita_ICMS = True
    If Txt_INSS <> "" Then TBProduto!INSS = Txt_INSS
    TBProduto!ID_CF = IIf(Txt_ID_CF = "", 0, Txt_ID_CF)
    TBProduto!ID_CFOP = IIf(Txt_ID_CFOP = "", 0, Txt_ID_CFOP)
    TBProduto!ID_CFOP1 = IIf(Txt_ID_CFOP1 = "", 0, Txt_ID_CFOP1)
    If cmbGenero <> "" Then TBProduto!ID_Genero = cmbGenero.ItemData(cmbGenero.ListIndex) Else TBProduto!ID_Genero = Null
    TBProduto.Update
End If
TBProduto.Close
ProcSalvarUltimaAlteracao txtcodproduto
ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaConjunto()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from ProjConjunto where desenho = '" & txtdesenhoproduto.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Do While TBProduto.EOF = False
        TBProduto!Descricao = txtDescricaoProduto.Text
        If cmbun <> "" Then TBProduto!Unidade = cmbun.Text
        If TxtPCusto <> "" Then TBProduto!valor = TxtPCusto.Text
        TBProduto.Update
        ProcVerificaValor
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from projproduto where codproduto = " & TBProduto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            TBItem!PCusto = TBProduto!ValorTotal
            TBItem.Update
        End If
        TBItem.Close
        TBProduto.MoveNext
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaValor()
On Error GoTo tratar_erro

Set TBMateriaprima = CreateObject("adodb.recordset")
TBMateriaprima.Open "Select * from ProjConjunto where desenho = '" & txtdesenhoproduto.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMateriaprima.EOF = False Then
    If TBMateriaprima!valor <> "" And TBMateriaprima!quantidade <> "" And TBMateriaprima!Dimensoes <> "" Then
        Select Case TBMateriaprima!Unidade
            Case "KG": TBMateriaprima!ValorTotal = Format(TBMateriaprima!valor * TBMateriaprima!PesoTotal, "###,##0.00")
            Case "MM": TBMateriaprima!ValorTotal = Format((TBMateriaprima!valor * TBMateriaprima!Dimensoes) * TBMateriaprima!quantidade, "###,##0.00")
            Case "MT": TBMateriaprima!ValorTotal = Format(((TBMateriaprima!valor / 1000) * TBMateriaprima!Dimensoes) * TBMateriaprima!quantidade, "###,##0.00")
        End Select
        If TBMateriaprima!Unidade <> "KG" And TBMateriaprima!Unidade <> "MM" And TBMateriaprima!Unidade <> "MT" Then TBMateriaprima!ValorTotal = Format(TBMateriaprima!valor * TBMateriaprima!quantidade, "###,##0.00")
    End If
    TBMateriaprima.Update
End If
TBMateriaprima.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarValoresDesc()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "os valores e descontos", "salvar", True, True) = False Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Salvar valores/descontos"
    ID_documento = txtcodproduto
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = ""
    ProcGravaEvento
    '==================================
    If txtmarglucro.Text <> "" Then TBProduto!MLucro = txtmarglucro.Text
    
    If TxtPCusto.Text <> "" Then
        valor = TxtPCusto
        
        If TxtPCusto <> TBProduto!PCusto Then
            NovoValor = Replace(valor, ",", ".")
            Conexao.Execute "UPDATE Projconjunto Set Valor = " & NovoValor & " where Desenho = '" & txtdesenhoproduto & "'"
            
            Select Case cmbun
                Case "KG": NomeCampoFiltro = "Valortotal = ROUND(Valor * Pesototal, 2)"
                Case "MM": NomeCampoFiltro = "Valortotal = ROUND((Valor * Dimensoes) * Quantidade, 2)"
                Case "MT": NomeCampoFiltro = "Valortotal = ROUND(((Valor / 1000) * Dimensoes) * Quantidade, 2)"
            End Select
            If cmbun <> "KG" And cmbun <> "MM" And cmbun <> "MT" Then NomeCampoFiltro = "Valortotal = ROUND(Valor * Quantidade, 2)"
            Conexao.Execute "UPDATE Projconjunto Set " & NomeCampoFiltro & " where Desenho = '" & txtdesenhoproduto & "'"
        End If
        
        If IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto) <> valor Then
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select Projproduto_fornecedor.* FROM Projproduto_fornecedor INNER JOIN Compras_fornecedores ON Projproduto_fornecedor.idfornecedor = Compras_fornecedores.IDCliente where Projproduto_fornecedor.codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                If USMsgBox("Deseja atualizar o preço de custo deste produto/serviço em todos os fornecedores vinculados?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    NovoValor = Replace(valor, ",", ".")
                    Conexao.Execute "Update Projproduto_fornecedor Set PCusto = " & NovoValor & " where codproduto = " & txtcodproduto
                End If
            End If
            TBFornecedor.Close
        End If
        TBProduto!PCusto = TxtPCusto.Text
    End If
    
    If txtPConsumo <> "" Or TxtPRevenda <> "" Then
        valor = txtPConsumo
        Valor1 = TxtPRevenda
        NovoValor = Replace(valor, ",", ".")
        NovoValor1 = Replace(Valor1, ",", ".")
        
        Permitido = False
        If IIf(IsNull(TBProduto!PConsumo), 0, TBProduto!PConsumo) <> valor And IIf(IsNull(TBProduto!PRevenda), 0, TBProduto!PRevenda) <> Valor1 Then
            Permitido = True
            TextoMsg = "o preço de consumo e revenda"
            TextoFiltro = "PConsumo = " & NovoValor & ", PRevenda = " & NovoValor1
        ElseIf IIf(IsNull(TBProduto!PConsumo), 0, TBProduto!PConsumo) <> valor Then
                Permitido = True
                TextoMsg = "o preço de consumo"
                TextoFiltro = "PConsumo = " & NovoValor
            ElseIf IIf(IsNull(TBProduto!PRevenda), 0, TBProduto!PRevenda) <> Valor1 Then
                    Permitido = True
                    TextoMsg = "o preço de revenda"
                    TextoFiltro = "PRevenda = " & NovoValor1
        End If
        If Permitido = True Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select Projproduto_Clientes.* FROM Projproduto_Clientes INNER JOIN Clientes ON Projproduto_Clientes.idCliente = clientes.IDCliente where Projproduto_clientes.codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                If USMsgBox("Deseja atualizar o " & TextoMsg & " deste produto/serviço em todos os clientes vinculados?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    Conexao.Execute "Update Projproduto_Clientes Set " & TextoFiltro & " where codproduto = " & txtcodproduto
                End If
            End If
            TBClientes.Close
        End If
        If txtPConsumo <> "" Then TBProduto!PConsumo = txtPConsumo.Text
        If TxtPRevenda.Text <> "" Then TBProduto!PRevenda = TxtPRevenda.Text
    End If
    
    If chkBloquear_valor = 1 Then TBProduto!Valor_bloqueado = True Else TBProduto!Valor_bloqueado = False
    TBProduto.Update
End If
TBProduto.Close
ProcSalvarUltimaAlteracao txtcodproduto
ProcAtualizaConjunto
ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
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
                    If FunVerificaRegistroValidadoSemMsg("projproduto", "codproduto = " & .ListItems(InitFor), True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    If .ListItems(InitFor).ListSubItems(1) <> "-" Then
                        ProcVerificaRegistroUtilizadoSemMsg "Projproduto", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "' and ID_similar IS NOT NULL"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "projconjunto", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "CadMaquinas_acessorios", "ID_produto = " & .ListItems(InitFor)
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "processos", "codproduto = " & .ListItems(InitFor)
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "Producao", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido_lista", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "vendas_carteira", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "int_Cod_Produto = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "Estoque_controle", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "Estoque_fisico", "codproduto = " & .ListItems(InitFor)
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "Estoque_Localarmazenamento", "IDProduto = " & .ListItems(InitFor)
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                        ProcVerificaRegistroUtilizadoSemMsg "Requisicao_materiais_lista", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                    End If
                ElseIf Cmb_opcao_lista = "Status" Then
                        'Mensagem = "Não é permitido alterar o status desse registro, pois o mesmo está sendo utilizado no módulo"
                        'ProcVerificaRegistroUtilizadoSemMsg "projconjunto", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        'If Permitido = False Then
                            '.ListItems.Item(InitFor).Checked = False
                            'GoTo Proximo
                        'End If
                    ElseIf Cmb_opcao_lista = "Validação estrutura" Then
                            Set TBLISTA = CreateObject("adodb.recordset")
                            TBLISTA.Open "Select Codproduto from ProjConjunto where Codproduto = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                            If TBLISTA.EOF = True Then
                                .ListItems.Item(InitFor).Checked = False
                                GoTo Proximo
                            End If
                            TBLISTA.Close
                        ElseIf Cmb_opcao_lista = "Validação plano de inspeção" Then
                            Set TBLISTA = CreateObject("adodb.recordset")
                            TBLISTA.Open "Select desenho from plano where desenho = '" & .ListItems.Item(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBLISTA.EOF = True Then
                                .ListItems.Item(InitFor).Checked = False
                                GoTo Proximo
                            End If
                            TBLISTA.Close
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
                If FunVerificaRegistroValidado("projproduto", "codproduto = " & .ListItems(InitFor), "mesmo", "registro", "excluir este", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If .ListItems(InitFor).ListSubItems(1) <> "-" Then
                    Mensagem = "Não é permitido excluir este registro, pois o mesmo está sendo utilizado no módulo"
                    ProcVerificaRegistroUtilizado "Projproduto", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "' and ID_similar IS NOT NULL", "Engenharia/Produtos e serviços/Cadastro de produtos similares"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "projconjunto", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Engenharia/Estrutura"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "CadMaquinas_acessorios", "ID_produto = " & .ListItems(InitFor), "PCP/Postos de trabalho"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "processos", "codproduto = " & .ListItems(InitFor), "Engenharia/Processos"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "Producao", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'", "PCP/Gerenciamento de ordem"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "Compras_pedido_lista", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Compras"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "vendas_carteira", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Vendas"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "tbl_Detalhes_Nota", "int_Cod_Produto = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Nota fiscal"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "Estoque_controle", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Estoque/Movimentação"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "Estoque_fisico", "codproduto = " & .ListItems(InitFor), "Estoque/Inventário"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "Estoque_Localarmazenamento", "IDProduto = " & .ListItems(InitFor), "Estoque/Local de armazenamento"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    ProcVerificaRegistroUtilizado "Requisicao_materiais_lista", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Estoque/Requisição de materiais"
                    If Permitido = False Then .ListItems.Item(InitFor).Checked = False
                End If
            ElseIf Cmb_opcao_lista = "Status" Then
                    'Mensagem = "Não é permitido alterar o status desse registro, pois o mesmo está sendo utilizado no módulo"
                    'ProcVerificaRegistroUtilizado "projconjunto", "desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Engenharia/Estrutura"
                    'If Permitido = False Then .ListItems.Item(InitFor).Checked = False
                ElseIf Cmb_opcao_lista = "Validação estrutura" Then
                        Set TBLISTA = CreateObject("adodb.recordset")
                        TBLISTA.Open "Select Codproduto from ProjConjunto where Codproduto = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBLISTA.EOF = True Then
                            USMsgBox ("Não é possivel validar estrutura, pois a mesma não esta cadastrada."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                        End If
                        TBLISTA.Close
                    ElseIf Cmb_opcao_lista = "Validação plano de inspeção" Then
                        Set TBLISTA = CreateObject("adodb.recordset")
                        TBLISTA.Open "Select desenho from plano where desenho = '" & .ListItems.Item(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBLISTA.EOF = True Then
                            USMsgBox ("Não é possivel validar plano de inspeção pois o mesmo não esta cadastrado."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                        End If
                        TBLISTA.Close
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
ProcLimpaCampos
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * From projproduto where codproduto = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDados
    cmbfamilia.Locked = False
    cmbfamilia.TabStop = True
    CodigoLista = Lista.SelectedItem.index
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_cod_serv_NFSe_Change()
On Error GoTo tratar_erro

If Txt_cod_serv_NFSe <> "" Then
    VerifNumero = Txt_cod_serv_NFSe
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_cod_serv_NFSe = ""
        Txt_cod_serv_NFSe.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_INSS_Change()
On Error GoTo tratar_erro

If Txt_INSS.Text <> "" Then
    VerifNumero = Txt_INSS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_INSS.Text = ""
        Txt_INSS.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_embalagem_Change()
On Error GoTo tratar_erro

If Txt_qtde_embalagem.Text <> "" Then
    VerifNumero = Txt_qtde_embalagem.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_embalagem.Text = ""
        Txt_qtde_embalagem.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_embalagem_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_embalagem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_embalagem_LostFocus()
On Error GoTo tratar_erro

Txt_qtde_embalagem.Text = Format(Txt_qtde_embalagem.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_skip_lote_Change()
On Error GoTo tratar_erro

If Txt_skip_lote <> "" Then
    VerifNumero = Txt_skip_lote
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_skip_lote = ""
        Txt_skip_lote.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_Change()
On Error GoTo tratar_erro

If txtComprimento.Text <> "" Then
    VerifNumero = txtComprimento.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComprimento.Text = ""
        txtComprimento.SetFocus
        Exit Sub
    End If
End If
ProcCalculaPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_LostFocus()
On Error GoTo tratar_erro

txtComprimento = Format(txtComprimento, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtConsumo_forn_Change()
On Error GoTo tratar_erro

If txtConsumo_forn.Text <> "" Then
    VerifNumero = txtConsumo_forn.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtConsumo_forn.Text = ""
        txtConsumo_forn.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtConsumo_forn_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtConsumo_forn

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtConsumo_forn_LostFocus()
On Error GoTo tratar_erro

txtConsumo_forn = Format(txtConsumo_forn, "###,##0.0000000000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesenhoproduto_Change()
On Error GoTo tratar_erro

Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from projfamilia where familia = '" & cmbfamilia & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    Letra = TBFamilia!Letra
End If
TBFamilia.Close

If Novo_Produto = True And Copiar_Produto = False And Desenho <> "" Then
VerifCodigo:
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        CompLetra = Len(Letra)
        If Right(Desenho, CompLetra) = Letra Then
            Numero = Left(Desenho, Len(Desenho) - (CompLetra + 1))
            Numero = Numero + 1
            Select Case Len(Numero)
                Case 5: Desenho = Numero & "-" & Letra
                Case 4: Desenho = "0" & Numero & "-" & Letra
                Case 3: Desenho = "00" & Numero & "-" & Letra
                Case 2: Desenho = "000" & Numero & "-" & Letra
                Case 1: Desenho = "0000" & Numero & "-" & Letra
            End Select
        Else
            Numero = Right(Desenho, Len(Desenho) - (CompLetra + 1))
            Numero = Numero + 1
            Select Case Len(Numero)
                Case 5: Desenho = Letra & "-" & Numero
                Case 4: Desenho = Letra & "-0" & Numero
                Case 3: Desenho = Letra & "-00" & Numero
                Case 2: Desenho = Letra & "-000" & Numero
                Case 1: Desenho = Letra & "-0000" & Numero
            End Select
        End If
        GoTo VerifCodigo
    Else
        txtdesenhoproduto = Desenho
    End If
    Copiar_Produto = False
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDias_antecipacao_Change()
On Error GoTo tratar_erro

If txtDias_antecipacao <> "" Then
    VerifNumero = txtDias_antecipacao
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDias_antecipacao = ""
        txtDias_antecipacao.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_Change()
On Error GoTo tratar_erro

If txtespessura.Text <> "" Then
    VerifNumero = txtespessura.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtespessura.Text = ""
        txtespessura.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_LostFocus()
On Error GoTo tratar_erro

txtespessura = Format(txtespessura, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEstMinimo_Change()
On Error GoTo tratar_erro

If txtEstMinimo.Text <> "" Then
    VerifNumero = txtEstMinimo.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtEstMinimo.Text = ""
        txtEstMinimo.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEstMinimo_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtEstMinimo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEstMinimo_LostFocus()
On Error GoTo tratar_erro

txtEstMinimo.Text = Format(txtEstMinimo.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_Change()
On Error GoTo tratar_erro

If txtLargura.Text <> "" Then
    VerifNumero = txtLargura.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLargura.Text = ""
        txtLargura.SetFocus
        Exit Sub
    End If
End If
ProcCalculaPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_LostFocus()
On Error GoTo tratar_erro

txtLargura = Format(txtLargura, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLeadTime_forn_LostFocus()
On Error GoTo tratar_erro

If txtLeadTime_forn <> "" Then
    VerifNumero = txtLeadTime_forn
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLeadTime_forn = ""
        txtLeadTime_forn.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtmarglucro_Change()
On Error GoTo tratar_erro

If txtmarglucro.Text <> "" Then
    VerifNumero = txtmarglucro.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtmarglucro.Text = ""
        txtmarglucro.SetFocus
        Exit Sub
    End If
    If Chk_vendas.Value = 1 And Chk_compras.Value = 1 Then ProcMLucro
Else
    txtPConsumo = "0,00000"
    TxtPRevenda = "0,00000"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMLucro()
On Error GoTo tratar_erro

If txtmarglucro <> "" And TxtPCusto <> "" Then
    MLucro = txtmarglucro
    PCustos = TxtPCusto
    txtPConsumo = Format(PCustos + (PCustos * MLucro) / 100, "###,##0.0000000000")
    TxtPRevenda = Format(PCustos + (PCustos * MLucro) / 100, "###,##,0.0000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtmarglucro_LostFocus()
On Error GoTo tratar_erro

txtmarglucro.Text = Format(txtmarglucro.Text, "###,##0.00")

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

Private Sub txtPConsumo_Change()
On Error GoTo tratar_erro

If txtPConsumo.Text <> "" Then
    VerifNumero = txtPConsumo.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPConsumo.Text = ""
        txtPConsumo.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPConsumo_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtPConsumo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPConsumo_LostFocus()
On Error GoTo tratar_erro

txtPConsumo.Text = Format(txtPConsumo.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtPCusto_Change()
On Error GoTo tratar_erro

If TxtPCusto.Text <> "" Then
    VerifNumero = TxtPCusto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        TxtPCusto.Text = ""
        TxtPCusto.SetFocus
        Exit Sub
    End If
    If Chk_vendas.Value = 1 And Chk_compras.Value = 1 Then ProcMLucro
Else
    txtPConsumo = "0,00000"
    TxtPRevenda = "0,00000"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPcusto_forn_Change()
On Error GoTo tratar_erro

If txtPcusto_forn.Text <> "" Then
    VerifNumero = txtPcusto_forn.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPcusto_forn.Text = ""
        txtPcusto_forn.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPcusto_forn_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtPcusto_forn

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPcusto_forn_LostFocus()
On Error GoTo tratar_erro

txtPcusto_forn.Text = Format(txtPcusto_forn.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtPCusto_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus TxtPCusto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtPCusto_LostFocus()
On Error GoTo tratar_erro

TxtPCusto.Text = Format(TxtPCusto.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpeso_Change()
On Error GoTo tratar_erro

If txtpeso <> "" Then
    VerifNumero = txtpeso
    ProcVerificaNumero
    If VerifNumero = False Then
        txtpeso = ""
        txtpeso.SetFocus
        Exit Sub
    End If
End If
ProcCalculaPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPesoBruto_Change()
On Error GoTo tratar_erro

If txtPesoBruto.Text <> "" Then
    VerifNumero = txtPesoBruto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPesoBruto.Text = ""
        txtPesoBruto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPesoBruto_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPesoBruto_LostFocus()
On Error GoTo tratar_erro

txtPesoBruto.Text = Format(txtPesoBruto.Text, "###,##0.000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPesoLiquido_Change()
On Error GoTo tratar_erro

If txtPesoLiquido.Text <> "" Then
    VerifNumero = txtPesoLiquido.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPesoLiquido.Text = ""
        txtPesoLiquido.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPesoLiquido_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtPesoLiquido

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPesoLiquido_LostFocus()
On Error GoTo tratar_erro

txtPesoLiquido.Text = Format(txtPesoLiquido.Text, "###,##0.000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtPRevenda_Change()
On Error GoTo tratar_erro

If TxtPRevenda.Text <> "" Then
    VerifNumero = TxtPRevenda.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        TxtPRevenda.Text = ""
        TxtPRevenda.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarImagem()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "a imagem", "salvar", True, True) = False Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBProduto!imagem = txt_Caminho
    TBProduto.Update
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Salvar imagem"
    ID_documento = txtcodproduto
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcSalvarUltimaAlteracao txtcodproduto
    
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
1:
End If
TBProduto.Close

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravar_doc()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("projproduto", "codproduto = " & txtcodproduto, "registro", "o documento", "salvar", True, True) = False Then Exit Sub

If Frame14.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_caminho_doc = "" Then
    NomeCampo = "o caminho"
    ProcVerificaAcao
    Cmd_localizar_doc.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from projproduto_documentos where ID = " & txtID_doc, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Codproduto = txtcodproduto
TBGravar!Data = IIf(txtData_doc = "", Date, txtData_doc)
TBGravar!Responsavel = IIf(txtResponsavel_doc = "", pubUsuario, txtResponsavel_doc)
TBGravar!caminho = Txt_caminho_doc
TBGravar!Obs = Trim(Txt_obs_doc)
TBGravar.Update
txtID_doc = TBGravar!ID
TBGravar.Close
ProcCarregaLista_Doc
If Novo_Produto1 = True Then
    USMsgBox ("Novo documento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo documento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar documento"
    If CodigoLista1 <> 0 And Lista_doc.ListItems.Count <> 0 Then
        Lista_doc.SelectedItem = Lista_doc.ListItems(CodigoLista1)
        Lista_doc.SetFocus
    End If
End If
1:
    '==================================
    Modulo = Formulario
    ID_documento = txtID_doc
    Documento = "Cód. interno: " & txtdesenhoproduto
    Documento1 = "Caminho: " & txt_Caminho
    ProcGravaEvento
    '==================================
    Novo_Produto1 = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
frmproj_produto_MenuImpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 19, True
ProcCarregaToolBar2 Me, 15192, 8, False
ProcCarregaToolBar3 Me, 15192, 9, False

ProcLimpaVariaveisPrincipais
Compras_Fornecedores = False

Sql_Produto = ""
ProcCarregaUnidades
ProcCarregaFamilia
procCarregaTipoGenero
If txtcodproduto <> "" Then ProcCarregaCamposCombo
Cmb_opcao_lista = "Validação"

If Compras_Produtos = True Then
    ProcVerifMostrarEsconderTab "Compras/Produtos e serviços/Valores e descontos", 3
    ProcVerifMostrarEsconderTab "Compras/Produtos e serviços/Clientes e fornecedores", 4
    
    Caption = "Compras - Produtos e serviços"
    Formulario = "Compras/Produtos e serviços"
    ProcMostrarValores
End If
If Vendas_Produtos = True Then
    ProcVerifMostrarEsconderTab "Vendas/Produtos e serviços/Valores e descontos", 3
    ProcVerifMostrarEsconderTab "Vendas/Produtos e serviços/Clientes e fornecedores", 4
    
    Caption = "Vendas - Produtos e serviços"
    Formulario = "Vendas/Produtos e serviços"
    ProcMostrarValores
End If
If Engenharia_Produtos = True Then
    With SSTab1
        .TabVisible(3) = False
        .TabVisible(4) = False
        .TabsPerRow = 6
        Numero_Abas = 6
    End With
        
    Caption = "Engenharia - Produtos e serviços"
    Formulario = "Engenharia/Produtos e serviços"
    ProcEsconderValores
End If
Direitos
SSTab1.Tab = 0

Formulario_produto = Formulario
'ActiveResize1.ResizeControls

ProcRemoveObjetosResize Me
'frmproj_produto.Refresh
'frmMDI.SkinFramework1.ApplyOptions = xtpSkinApplyColors


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifMostrarEsconderTab(Formulario As String, NTab As Integer)
On Error GoTo tratar_erro

If NTab = 3 Then Contador = 8
Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = True Then
    SSTab1.TabVisible(NTab) = False
    Contador = Contador - 1
Else
    SSTab1.TabVisible(NTab) = True
End If
TBAcessos.Close
SSTab1.TabsPerRow = Contador
Numero_Abas = Contador

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaUnidades()
On Error GoTo tratar_erro

ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False
ProcCarregaComboUnidade cmbuTrib, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFamilia()
On Error GoTo tratar_erro

Set TBFamilia = CreateObject("adodb.recordset")
If Compras_Produtos = True Then ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
If Vendas_Produtos = True Then ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", True
If Engenharia_Produtos = True Then ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarValores()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
With Lista.ColumnHeaders
    If Formulario = "Vendas/Produtos e serviços" Then
        .Item(12).Width = 900
        .Item(13).Width = 900
        Label1(2).Visible = False
        txtmarglucro.Visible = False
        
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Vendas/Produtos e serviços/Valores e descontos/Visualizar valor de custo'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            .Item(5).Width = 2707
            .Item(11).Width = 0
            Label1(3).Visible = False
            TxtPCusto.Visible = False
            Texto = "V"
        Else
            .Item(5).Width = 1807
            .Item(11).Width = 900
            Texto = "VC"
        End If
        TBAcessos.Close
        cmdcalc_peso.Visible = False
    ElseIf Formulario = "Compras/Produtos e serviços" Then
            .Item(5).Width = 3607
            .Item(11).Width = 900
            .Item(12).Width = 0
            .Item(13).Width = 0
            Label1(4).Visible = False
            txtPConsumo.Visible = False
            Label1(5).Visible = False
            TxtPRevenda.Visible = False
            'chkBloquear_valor.Left = 4560
            Texto = "C"
    End If
End With
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    TBLISTA!Responsavel = pubUsuario
    TBLISTA!Modulo = Formulario
    TBLISTA!Texto = Texto
    TBLISTA.Update
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEsconderValores()
On Error GoTo tratar_erro

With Lista.ColumnHeaders
    .Item(5).Width = 4507
    .Item(11).Width = 0
    .Item(12).Width = 0
    .Item(13).Width = 0
End With
ProcExcluirDadosProducaoRelatoriosTotal
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    TBLISTA!Responsavel = pubUsuario
    TBLISTA!Modulo = Formulario
    TBLISTA!Texto = "E"
    TBLISTA.Update
End If
TBLISTA.Close

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
Set TBLISTA_Produto = CreateObject("adodb.recordset")
If Sql_Produto <> "" Then
    TBLISTA_Produto.Open Sql_Produto, Conexao, adOpenKeyset, adLockReadOnly
Else
    TBLISTA_Produto.Open "Select * from projproduto where desenho = '" & txtdesenhoproduto.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
End If
If TBLISTA_Produto.EOF = False Then ProcExibePagina (Pagina)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Produto.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Produto.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Produto.PageSize
ContadorReg = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Produto.RecordCount - IIf(Pagina > 1, (TBLISTA_Produto.PageSize * (Pagina - 1)), 0), TBLISTA_Produto.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Produto.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Produto!Codproduto
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Produto!Desenho), "", (TBLISTA_Produto!Desenho))
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select N_Referencia from item_aplicacoes where codproduto = " & TBLISTA_Produto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            .Item(.Count).SubItems(2) = IIf(IsNull(TBItem!N_referencia), "", (TBItem!N_referencia))
        Else
            .Item(.Count).SubItems(2) = ""
        End If
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Part_number from Projproduto_fabricante where codproduto = " & TBLISTA_Produto!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            .Item(.Count).SubItems(3) = IIf(IsNull(TBItem!Part_number), "", (TBItem!Part_number))
        Else
            .Item(.Count).SubItems(3) = ""
        End If
        
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Produto!Descricao), "", (TBLISTA_Produto!Descricao))
        
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Revisao, Data from Projproduto_revisao where Codproduto = " & TBLISTA_Produto!Codproduto & " order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            .Item(.Count).SubItems(5) = TBItem!Revisao
            .Item(.Count).SubItems(6) = IIf(IsNull(TBItem!Data), "", Format(TBItem!Data, "dd/mm/yy"))
        Else
            If TBLISTA_Produto!RevDesenho <> "" Then .Item(.Count).SubItems(5) = TBLISTA_Produto!RevDesenho
            .Item(.Count).SubItems(6) = ""
        End If
        
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Produto!Espessura), "", Format(TBLISTA_Produto!Espessura, "###,##0.00")) & "X" & IIf(IsNull(TBLISTA_Produto!Largura), "", Format(TBLISTA_Produto!Largura, "###,##0.00")) & "X" & IIf(IsNull(TBLISTA_Produto!Comprimento), "", Format(TBLISTA_Produto!Comprimento, "###,##0.00"))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Produto!Dureza), "", TBLISTA_Produto!Dureza)
        
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select IDIntClasse from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBLISTA_Produto!ID_CF), 0, TBLISTA_Produto!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            .Item(.Count).SubItems(9) = IIf(IsNull(TBItem!IDIntClasse), "", TBItem!IDIntClasse)
        Else
            .Item(.Count).SubItems(9) = ""
        End If
        TBItem.Close
        
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Produto!PCusto), "", Format(TBLISTA_Produto!PCusto, "###,##0.0000000000"))
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Produto!PConsumo), "", Format(TBLISTA_Produto!PConsumo, "###,##0.0000000000"))
        .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA_Produto!PRevenda), "", Format(TBLISTA_Produto!PRevenda, "###,##0.0000000000"))
        If TBLISTA_Produto!Data <> "" Then .Item(.Count).SubItems(13) = Format(TBLISTA_Produto!Data, "dd/mm/yy") Else .Item(.Count).SubItems(13) = ""
        If Cmb_opcao_lista = "Validação estrutura" Then
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_Produto!DtValidacaoConj) = False, "Sim", "Não")
        ElseIf Cmb_opcao_lista = "Validação plano de inspeção" Then
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_Produto!DtValidacaoPlano) = False, "Sim", "Não")
        Else
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_Produto!DtValidacao) = False, "Sim", "Não")
        End If
    End With
    TBLISTA_Produto.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Produto.RecordCount
If TBLISTA_Produto.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Produto.PageCount
ElseIf TBLISTA_Produto.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Produto.PageCount & " de: " & TBLISTA_Produto.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Produto.AbsolutePage - 1 & " de: " & TBLISTA_Produto.PageCount
End If
If txtreferencia.Text <> "" Then txtreferencia.Locked = True Else txtreferencia.Locked = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista_Doc()
On Error GoTo tratar_erro

Lista_doc.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ID, Caminho from projproduto_documentos where Codproduto = " & txtcodproduto & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_doc.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!caminho), "", TBLISTA!caminho)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtcodproduto = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If

ProcCorrigeForm

Lista.Visible = True
Frame1.Visible = True

If SSTab1.Tab = 0 Then
    USToolBar1.Visible = True
    USToolBar2.Visible = False
    USToolBar3.Visible = False
ElseIf SSTab1.Tab = 7 Then
        USToolBar1.Visible = False
        USToolBar2.Visible = False
        USToolBar3.Visible = True
    Else
        USToolBar1.Visible = False
        USToolBar2.Visible = True
        USToolBar3.Visible = False
End If
If SSTab1.Tab = 2 Then
    Frame5.Visible = True
    Frame13.Visible = True
Else
    Frame5.Visible = False
    Frame13.Visible = False
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Select Case SSTab1.Tab
        Case 0:
            cmbun.Visible = True
            Cmb_un_com.Visible = True
            cmbfamilia.Visible = True
            cmbFornecedor.Visible = False
            cmbGrupo.Visible = False
            cmbcliente.Visible = False
            cmbFabricante.Visible = False
            txtDescricaoProduto.Visible = True
            cmdfamilia.Visible = True
            cmddescricao.Visible = True
            If Lista.Visible = True Then Lista.SetFocus
        Case 1:
            cmbun.Visible = False
            Cmb_un_com.Visible = False
            cmbfamilia.Visible = False
            cmbFabricante.Visible = True
            cmbGrupo.Visible = False
            cmbcliente.Visible = False
            cmbFornecedor.Visible = False
            txtDescricaoProduto.Visible = False
            cmdfamilia.Visible = False
            cmddescricao.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            Lista.SetFocus
        Case 2:
            If Chk_vendas.Value = 1 And Chk_compras.Value = 1 Then
                Frame5.Enabled = True
                Frame13.Enabled = True
            ElseIf Chk_compras.Value = 1 Then
                    Frame5.Enabled = True
                    Frame13.Enabled = False
                Else
                    Frame5.Enabled = False
                    Frame13.Enabled = True
            End If
            cmbun.Visible = False
            Cmb_un_com.Visible = False
            cmbfamilia.Visible = False
            cmbFornecedor.Visible = False
            cmbGrupo.Visible = False
            cmbcliente.Visible = False
            cmbFabricante.Visible = False
            txtDescricaoProduto.Visible = False
            cmdfamilia.Visible = False
            cmddescricao.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            Lista.SetFocus
        Case 3:
            cmbun.Visible = False
            Cmb_un_com.Visible = False
            cmbfamilia.Visible = False
            cmbFornecedor.Visible = False
            cmbGrupo.Visible = False
            cmbcliente.Visible = False
            cmbFabricante.Visible = False
            txtDescricaoProduto.Visible = False
            cmdfamilia.Visible = False
            cmddescricao.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            Lista.SetFocus
        Case 4:
            cmbun.Visible = False
            Cmb_un_com.Visible = False
            cmbfamilia.Visible = False
            cmbFabricante.Visible = False
            cmbFornecedor.Visible = True
            cmbGrupo.Visible = True
            cmbcliente.Visible = True
            txtDescricaoProduto.Visible = False
            cmdfamilia.Visible = False
            cmddescricao.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            Lista.SetFocus
        Case 5:
            ProcArrumaCC
            cmbun.Visible = False
            Cmb_un_com.Visible = False
            cmbfamilia.Visible = False
            cmbFornecedor.Visible = False
            cmbFabricante.Visible = False
            cmbGrupo.Visible = False
            cmbcliente.Visible = False
            txtDescricaoProduto.Visible = False
            cmdfamilia.Visible = False
            cmddescricao.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            Lista.SetFocus
        Case 6:
            cmbun.Visible = False
            Cmb_un_com.Visible = False
            cmbfamilia.Visible = False
            cmbFornecedor.Visible = False
            cmbFabricante.Visible = False
            cmbGrupo.Visible = False
            cmbcliente.Visible = False
            txtDescricaoProduto.Visible = False
            cmdfamilia.Visible = False
            cmddescricao.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            cmdImportar.SetFocus
        Case 7:
            Lista.Visible = False
            Frame1.Visible = False
            cmbun.Visible = False
            Cmb_un_com.Visible = False
            cmbfamilia.Visible = False
            cmbFornecedor.Visible = False
            cmbFabricante.Visible = False
            cmbGrupo.Visible = False
            cmbcliente.Visible = False
            txtDescricaoProduto.Visible = False
            cmdfamilia.Visible = False
            cmddescricao.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            Lista_doc.SetFocus
            ProcCarregaLista_Doc
    End Select
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeForm()
On Error GoTo tratar_erro

With Lista
    Select Case SSTab1.Tab
        Case 0: .Top = Frame2.Top + Frame2.Height
        Case 1: .Top = Frame4.Top + Frame4.Height
        Case 2: .Top = Frame16.Top + Frame16.Height
        Case 3: .Top = Frame9.Top + Frame9.Height
        Case 4: .Top = Frame15.Top + Frame15.Height
        Case 5: ProcArrumaCC
        Case 6: .Top = Frame10.Top + Frame10.Height
    End Select
    .Height = Frame1.Top - .Top
End With


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Produto = True Then
    USMsgBox ("Salve o registro antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcArrumaCC()
On Error GoTo tratar_erro

If Chk_vendas.Value = 1 And Chk_compras.Value = 1 Then
    With Frame19
        .Visible = True
        .Caption = "Compras"
    End With
    With Frame11
        .Visible = True
        .Top = Frame19.Top + Frame19.Height
        .Caption = "Vendas"
    End With
    Lista.Top = Frame11.Top + Frame11.Height
ElseIf Chk_compras.Value = 1 Or Chk_vendas.Value = 1 Then
    Frame19.Caption = ""
    Frame11.Caption = ""
    If Chk_compras.Value = 1 Then
        Frame19.Visible = True
        Frame11.Visible = False
        Lista.Top = Frame19.Top + Frame19.Height
    Else
        Frame19.Visible = False
        With Frame11
            .Visible = True
            .Top = Frame19.Top
        End With
        Lista.Top = Frame11.Top + Frame11.Height
    End If
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEsconderMostrarCC()
On Error GoTo tratar_erro

With SSTab1
    .TabVisible(5) = True
    .TabsPerRow = Numero_Abas
    If Chk_vendas.Value = 0 And Chk_compras.Value = 0 Then
        If SSTab1.Tab = 5 Then SSTab1.Tab = 0
        .TabVisible(5) = False
        .TabsPerRow = Numero_Abas - 1
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosOutros()
On Error GoTo tratar_erro

cmbFabricante.Clear
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select Fabricante_marca.Fabricante FROM Projproduto_fabricante INNER JOIN Fabricante_marca ON Projproduto_fabricante.idfabricante = Fabricante_marca.Id where Projproduto_fabricante.codproduto = " & txtcodproduto & " and Fabricante_marca.Fabricante <> 'Null'", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    Do While TBFornecedor.EOF = False
        cmbFabricante.AddItem TBFornecedor!Fabricante
        TBFornecedor.MoveNext
    Loop
End If
TBFornecedor.Close

txtEstMinimo = IIf(IsNull(TBProduto!Estoque_minimo), "", Format(TBProduto!Estoque_minimo, "###,##0.0000"))
ProcCarregaEstoqueReal
txtpeso = IIf(IsNull(TBProduto!peso_metro), "", Format(TBProduto!peso_metro, "###,##0.0000000000"))
If IsNull(TBProduto!Un_Kg) = False Then cmbunkg.Text = TBProduto!Un_Kg
Txt_qtde_embalagem = IIf(IsNull(TBProduto!Qtde_embalagem), "", Format(TBProduto!Qtde_embalagem, "###,##0.0000"))
If TBProduto!Estoque = True Then chkEstoque.Value = 0 Else chkEstoque.Value = 1
If TBProduto!Insp_recebimento = True Then Chk_insp_recebimento.Value = 1 Else Chk_insp_recebimento.Value = 0
If TBProduto!Processo = True Then chkProcesso.Value = 1 Else chkProcesso.Value = 0
If TBProduto!Plano_inspecao = True Then Chk_tem_plano.Value = 1 Else Chk_tem_plano.Value = 0
If Chk_vendas.Value = 1 Then ProcVerifDtUltimaVenda
If Chk_compras.Value = 1 Then ProcVerifDtUltimaCompra
ProcVerifDtUltimaOrdem
If TBProduto!nacional = True Then chknacional.Value = True Else chknacional.Value = False
If TBProduto!importacao = True Then chkimportacao.Value = True Else chkimportacao.Value = False
If TBProduto!exportacao = True Then chkexportacao.Value = 1 Else chkexportacao.Value = 0
txtFiname = IIf(IsNull(TBProduto!FINAME), "", TBProduto!FINAME)
txtCor = IIf(IsNull(TBProduto!Cor), "", TBProduto!Cor)
txtespessura = IIf(IsNull(TBProduto!Espessura), "", Format(TBProduto!Espessura, "###,##0.00"))
txtLargura = IIf(IsNull(TBProduto!Largura), "", Format(TBProduto!Largura, "###,##0.00"))
txtComprimento = IIf(IsNull(TBProduto!Comprimento), "", Format(TBProduto!Comprimento, "###,##0.00"))
txtDureza = IIf(IsNull(TBProduto!Dureza), "", TBProduto!Dureza)
Txt_skip_lote = IIf(IsNull(TBProduto!Skip_lote), "", TBProduto!Skip_lote)

txtPesoBruto = IIf(IsNull(TBProduto!PBruto), "", Format(TBProduto!PBruto, "###,##0.000000"))
txtPesoLiquido = IIf(IsNull(TBProduto!PLiquido), "", Format(TBProduto!PLiquido, "###,##0.000000"))
'=================================================
' valores tributados Artnel
'=================================================
cmbuTrib.Text = IIf(IsNull(TBProduto!uTrib), TBProduto!Unidade, TBProduto!uTrib)
txtvuTrib.Text = IIf(IsNull(TBProduto!vTrib), "", Format(TBProduto!vTrib, "###,##0.000000"))

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select ID, CODIGO, DtBloq, Setor from Usuarios_setor where ID = " & IIf(IsNull(TBProduto!ID_CC), 0, TBProduto!ID_CC), Conexao, adOpenKeyset, adLockOptimistic
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

If IsNull(TBProduto!Cod_servico) = False And TBProduto!Cod_servico <> "" Then Txt_cod_serv = Left(TBProduto!Cod_servico, 2) & "." & Right(TBProduto!Cod_servico, 2)
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from projproduto_genero where ID = " & IIf(IsNull(TBProduto!ID_Genero), 0, TBProduto!ID_Genero), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then cmbGenero = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Descricao), "", TBFI!Descricao)
TBFI.Close

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from projproduto_Tipo where ID = " & IIf(IsNull(TBProduto!ID_Tipo), 0, TBProduto!ID_Tipo), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then cmbClassificacao_produto = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Descricao), "", TBFI!Descricao)
TBFI.Close

Txt_GTIN = IIf(IsNull(TBProduto!GTIN), "", TBProduto!GTIN)
Txt_cod_serv_NFSe = IIf(IsNull(TBProduto!Cod_servico_NFSE), "", TBProduto!Cod_servico_NFSE)
txtinspecao = IIf(IsNull(TBProduto!Inspecao), "", TBProduto!Inspecao)
txtembalagem = IIf(IsNull(TBProduto!Embalagem), "", TBProduto!Embalagem)
txtGravacao = IIf(IsNull(TBProduto!Gravacao), "", TBProduto!Gravacao)
txt_observacoes = IIf(IsNull(TBProduto!Observacoes), "", TBProduto!Observacoes)
txtQtde_LoteMinimo = IIf(IsNull(TBProduto!qtde_LoteMinimo), "0,0000", Format(TBProduto!qtde_LoteMinimo, "###,##0.0000"))

Set TBCQ = CreateObject("adodb.recordset")
TBCQ.Open "Select * FROM QualidadePPAP where IDProduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBCQ.EOF = False Then
    txtPPAP_Rev = IIf(IsNull(TBCQ!revPPAP), "", TBCQ!revPPAP)
    txtPPAP_Datarev = IIf(IsNull(TBCQ!DataRevisao), "__/__/____", TBCQ!DataRevisao)
    FramePPAP.Enabled = False
Else
    txtPPAP_Rev = IIf(IsNull(TBProduto!PPAP_Rev), "", TBProduto!PPAP_Rev)
    txtPPAP_Datarev = IIf(IsNull(TBProduto!PPAP_Datarev), "__/__/____", Format(TBProduto!PPAP_Datarev, "dd/mm/yyyy"))
    FramePPAP.Enabled = True
End If
TBCQ.Close
If TBProduto!Submetido = True Then chkSubmetido.Value = 1 Else chkSubmetido.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosForn_clientes()
On Error GoTo tratar_erro

With cmbFornecedor
    .Clear
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select Compras_fornecedores.IDCliente, Compras_fornecedores.Nome_Razao FROM Projproduto_fornecedor INNER JOIN Compras_fornecedores ON Projproduto_fornecedor.idfornecedor = Compras_fornecedores.IDCliente where Projproduto_fornecedor.codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        Do While TBFornecedor.EOF = False
            .AddItem TBFornecedor!Nome_Razao
            .ItemData(.NewIndex) = TBFornecedor!IDCliente
            TBFornecedor.MoveNext
        Loop
    End If
    TBFornecedor.Close
End With

Chk_cliente.Value = 1
With cmbcliente
    .Clear
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select Clientes.IDCliente, Clientes.NomeRazao FROM Projproduto_Clientes INNER JOIN Clientes ON Projproduto_Clientes.idCliente = clientes.IDCliente where Projproduto_clientes.codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        Do While TBClientes.EOF = False
            .AddItem TBClientes!NomeRazao
            .ItemData(.NewIndex) = TBClientes!IDCliente
            TBClientes.MoveNext
        Loop
    End If
    TBClientes.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifDtUltimaCompra()
On Error GoTo tratar_erro

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select Compras_pedido.data FROM compras_pedido_lista INNER JOIN Compras_pedido ON compras_pedido_lista.idpedido = Compras_pedido.IDpedido where compras_pedido_lista.desenho = '" & txtdesenhoproduto & "' and Compras_pedido.Data is not null order by compras_pedido_lista.idlista", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    TBCompras.MoveLast
    txtDtCompra = Format(TBCompras!Data, "dd/mm/yy")
End If
TBCompras.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifDtUltimaVenda()
On Error GoTo tratar_erro

Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select datavendas from vendas_carteira where desenho = '" & txtdesenhoproduto & "' and datavendas is not null order by datavendas", Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    TBVendas.MoveLast
    txtDtVenda = Format(TBVendas!Datavendas, "dd/mm/yy")
End If
TBVendas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifDtUltimaOrdem()
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select data from Producao where desenho = '" & txtdesenhoproduto & "' order by data", Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    TBOrdem.MoveLast
    If IsNull(TBOrdem!Data) = False And TBOrdem!Data <> "" Then Txt_Dt_ordem = Format(TBOrdem!Data, "dd/mm/yy")
End If
TBOrdem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosImpostos()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Idclass, IDIntClasse  from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBProduto!ID_CF), 0, TBProduto!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_CF = TBAbrir!Idclass
    Txt_CF = IIf(IsNull(TBAbrir!IDIntClasse), "", TBAbrir!IDIntClasse)
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_CFOP = TBAbrir!IDCountCfop
    Txt_CFOP = IIf(IsNull(TBAbrir!ID_CFOP), "", TBAbrir!ID_CFOP)
    Txt_natureza_operacao = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & IIf(IsNull(TBProduto!ID_CFOP1), 0, TBProduto!ID_CFOP1), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_CFOP1 = TBAbrir!IDCountCfop
    Txt_CFOP1 = IIf(IsNull(TBAbrir!ID_CFOP), "", TBAbrir!ID_CFOP)
    Txt_natureza_operacao1 = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
End If
TBAbrir.Close

txtmarglucro.Text = IIf(IsNull(TBProduto!MLucro), "0,00", Format(TBProduto!MLucro, "###,##0.00"))
TxtPCusto.Text = IIf(IsNull(TBProduto!PCusto), "0,00000", Format(TBProduto!PCusto, "###,##0.0000000000"))
txtPConsumo.Text = IIf(IsNull(TBProduto!PConsumo), "0,00000", Format(TBProduto!PConsumo, "###,##0.0000000000"))
TxtPRevenda.Text = IIf(IsNull(TBProduto!PRevenda), "0,00000", Format(TBProduto!PRevenda, "###,##0.0000000000"))
If TBProduto!Servico_cliente = True Then Chk_servico_executado_cliente.Value = 1 Else Chk_servico_executado_cliente.Value = 0
If TBProduto!Credita_ICMS = True Then Chk_credito_ICMS.Value = 0 Else Chk_credito_ICMS.Value = 1
Txt_INSS = IIf(IsNull(TBProduto!INSS), "", TBProduto!INSS)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosValoresDesc()
On Error GoTo tratar_erro

txtmarglucro.Text = IIf(IsNull(TBProduto!MLucro), "0,00", Format(TBProduto!MLucro, "###,##0.00"))
TxtPCusto.Text = IIf(IsNull(TBProduto!PCusto), "0,00000", Format(TBProduto!PCusto, "###,##0.0000000000"))
If Engenharia_Produtos = False Then
    txtPConsumo.Text = IIf(IsNull(TBProduto!PConsumo), "0,00000", Format(TBProduto!PConsumo, "###,##0.0000000000"))
    TxtPRevenda.Text = IIf(IsNull(TBProduto!PRevenda), "0,00000", Format(TBProduto!PRevenda, "###,##0.0000000000"))
End If
If TBProduto!Valor_bloqueado = True Then chkBloquear_valor = 1 Else chkBloquear_valor = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosImagem()
On Error GoTo tratar_erro

txt_Caminho = IIf(IsNull(TBProduto!imagem), "", TBProduto!imagem)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados_Doc()
On Error GoTo tratar_erro

txtID_doc = TBMaterial!ID
txtData_doc = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
txtResponsavel_doc = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
Txt_caminho_doc = IIf(IsNull(TBMaterial!caminho), "", TBMaterial!caminho)
Txt_obs_doc = IIf(IsNull(TBMaterial!Obs), "", TBMaterial!Obs)
Novo_Produto1 = False
Frame14.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtleadtime_LostFocus()
On Error GoTo tratar_erro

If txtleadtime.Text <> "" Then
    VerifNumero = txtleadtime.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtleadtime.Text = ""
        txtleadtime.SetFocus
        Exit Sub
    End If
End If
    
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

Private Sub TxtPRevenda_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus TxtPRevenda

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtPRevenda_LostFocus()
On Error GoTo tratar_erro

TxtPRevenda.Text = Format(TxtPRevenda.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_LoteMinimo_Change()
On Error GoTo tratar_erro

If txtQtde_LoteMinimo.Text <> "" Then
    VerifNumero = txtQtde_LoteMinimo.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_LoteMinimo.Text = ""
        txtQtde_LoteMinimo.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_LoteMinimo_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde_LoteMinimo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_LoteMinimo_LostFocus()
On Error GoTo tratar_erro

txtQtde_LoteMinimo = Format(txtQtde_LoteMinimo, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaEstoqueReal()
On Error GoTo tratar_erro

txtEstReal = Format(FunVerificaQtdeEstoque(txtdesenhoproduto, 0, "and Estoque_real <> 0"), "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosPC()
On Error GoTo tratar_erro

Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * FROM Projproduto where codproduto = " & txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    Txt_ID_PC = IIf(IsNull(TBFamilia!ID_PC), 0, TBFamilia!ID_PC)
    Txt_ID_PC1 = IIf(IsNull(TBFamilia!ID_PC1), 0, TBFamilia!ID_PC1)
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * FROM tbl_familia where int_codfamilia = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        Txt_codigo_PC = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
        Txt_descricao_PC = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
    End If
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * FROM tbl_familia where int_codfamilia = " & Txt_ID_PC1, Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        Txt_codigo_PC1 = IIf(IsNull(TBFamilia!CODIGO), "", TBFamilia!CODIGO)
        Txt_descricao_PC1 = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
    End If
End If
TBFamilia.Close

If SSTab1.Tab = 5 Then ProcArrumaCC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtRevenda_forn_Change()
On Error GoTo tratar_erro

If txtRevenda_forn.Text <> "" Then
    VerifNumero = txtRevenda_forn.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtRevenda_forn.Text = ""
        txtRevenda_forn.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtRevenda_forn_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtRevenda_forn

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtRevenda_forn_LostFocus()
On Error GoTo tratar_erro

txtRevenda_forn = Format(txtRevenda_forn, "###,##0.0000000000")
    
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
    Case 3: ProcGravar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcStatus
    Case 9: ProcCodRef
    Case 10: ProcSimilar
    Case 11: ProcEstrutura
    Case 12: ProcCopiar
    Case 13: ProcRevisao
    Case 14:
        If Cmb_opcao_lista = "Validação" Then
            If Compras_Produtos = True Then ProcValidarRegistros Lista, "Compras/Produtos e serviços"
            If Vendas_Produtos = True Then ProcValidarRegistros Lista, "Vendas/Produtos e serviços"
            If Engenharia_Produtos = True Then ProcValidarRegistros Lista, "Engenharia/Produtos e serviços"
        ElseIf Cmb_opcao_lista = "Validação estrutura" Then
                If Compras_Produtos = True Then ProcValidarRegistros Lista, "Compras/Produtos e serviços/Validar estrutura"
                If Vendas_Produtos = True Then ProcValidarRegistros Lista, "Vendas/Produtos e serviços/Validar estrutura"
                If Engenharia_Produtos = True Then ProcValidarRegistros Lista, "Engenharia/Produtos e serviços/Validar estrutura"
            Else
                If Compras_Produtos = True Then ProcValidarRegistros Lista, "Compras/Produtos e serviços/Validar plano de inspeção"
                If Vendas_Produtos = True Then ProcValidarRegistros Lista, "Vendas/Produtos e serviços/Validar plano de inspeção"
                If Engenharia_Produtos = True Then ProcValidarRegistros Lista, "Engenharia/Produtos e serviços/Validar plano de inspeção"
        End If
    Case 15: ProcAtualizar
    Case 16: procAtualizaProdutosNuvem
    Case 18: ProcAjuda
    Case 19: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualizaProdutosNuvem()
On Error GoTo tratar_erro

FunAbreBDSite
If ConexaoMySql.State = 1 Then

If USMsgBox("Deseja realmente atualizar todos os cadastros dos produtos na nuvem?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    
    Set TBAbrir = CreateObject("adodb.recordset")
    
StrSql = "Select pp.Desenho,PP.Unidade,PP.Descricao,PP.Prevenda, isnull(QEPP.Estoque_disponivel,0) as Estoque_Disponivel, isnull(QEPP.Qtde_empenhada,0) as Vendido, isnull(Sum(QEPP.Estoque_Disponivel-QEPP.Qtde_empenhada),0) As Saldo from projproduto PP Left join Qtde_estoque_produto QEPP on QEPP.Desenho = pp.Desenho Where PP.Vendas = '1' group by pp.Desenho,PP.Unidade,PP.Descricao,PP.Prevenda,QEPP.Estoque_disponivel,QEPP.Qtde_empenhada"
    TBAbrir.Open StrSql & " order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBMySQL = New ADODB.Recordset
            '=================================================================
            ' Salvar produtos na nuvem
            '=================================================================
            StrSql = "Select * From Vendas_Produtos where Descricao = '" & TBAbrir!Descricao & "' and CNPJ_Empresa = '" & CNPJ_Empresa & "'"
            'Debug.print StrSql
             TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
              If TBMySQL.EOF = True Then
                    TBMySQL.AddNew
              End If
                    TBMySQL.Fields!CODIGO = TBAbrir!Desenho
                    TBMySQL.Fields!Descricao = TBAbrir!Descricao
                    TBMySQL.Fields!Unidade = TBAbrir!Unidade
                    TBMySQL.Fields!vlr_unit = TBAbrir!PRevenda
                    TBMySQL.Fields!Estoque = TBAbrir!Estoque_Disponivel
                    TBMySQL.Fields!Vendido = TBAbrir!Vendido
 '                   TBMySQL.Fields!Disponivel = TBAbrir!Saldo
                    
                    TBMySQL.Fields!CNPJ_Empresa = CNPJ_Empresa
                    
                    TBMySQL.Update
        TBAbrir.MoveNext
        Loop
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Vendas/Clientes"
        Evento = "Atualizar clientes nuvem"
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

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1:
        Select Case SSTab1.Tab
            Case 1: ProcGravarOutros
            Case 2: ProcGravarImpostos
            Case 3: ProcGravarValoresDesc
            Case 4: ProcGravarFornCli
            Case 5: ProcGravarPC
            Case 6: ProcGravarImagem
        End Select
    Case 2: ProcImprimir
    Case 3: ProcAnterior
    Case 4: ProcProximo
    Case 6: ProcAjuda
    Case 7: ProcSair
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
    Case 2: ProcGravar_doc
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

Sub procCarregaTipoGenero()
On Error GoTo tratar_erro

With cmbGenero
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Projproduto_genero order by codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
            .AddItem TBCarregarCombo!CODIGO & " - " & TBCarregarCombo!Descricao
            .ItemData(.NewIndex) = TBCarregarCombo!ID
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

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

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaPesoBruto()
On Error GoTo tratar_erro

valor = IIf(txtpeso = "", 0, txtpeso)
Valor1 = IIf(txtLargura = "", 0, txtLargura)
Valor2 = IIf(txtComprimento = "", 0, txtComprimento)
If cmbunkg = "Mt²" Then
    txtPesoBruto = Format(((valor * Valor2) / 1000) * (Valor1 / 1000), "###,##0.000000")
ElseIf cmbunkg = "Mt/L" Then
        txtPesoBruto = Format(((valor * Valor2) / 1000), "###,##0.000000")
    Else
        txtPesoBruto = Format(valor, "###,##0.000000")
End If
txtPesoLiquido = txtPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarUltimaAlteracao(IDProjproduto As Long)
On Error GoTo tratar_erro

Conexao.Execute "Update projproduto Set DtAlteracao = '" & Now & "', RespAlteracao = '" & pubUsuario & "' where codproduto = " & IDProjproduto
txtDataAlteracao = Now
txtResponsavelAlteracao = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
