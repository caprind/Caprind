VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_Aprovar_Pedido 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Pedido - Aprovar"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   23
      Top             =   5760
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
         TabIndex        =   12
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtPagIr 
         Height          =   315
         Left            =   9540
         TabIndex        =   13
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   17
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Aprovar_Pedido.frx":0000
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
         TabIndex        =   16
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Aprovar_Pedido.frx":37A7
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
         TabIndex        =   14
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
         TabIndex        =   15
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Aprovar_Pedido.frx":72B9
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
         TabIndex        =   18
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Aprovar_Pedido.frx":B3AE
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
         Left            =   4410
         TabIndex        =   30
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label32 
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
         TabIndex        =   27
         Top             =   240
         Width           =   645
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   240
         Width           =   1275
      End
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
      Height          =   1545
      Left            =   55
      TabIndex        =   20
      Top             =   990
      Width           =   15195
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   10290
         TabIndex        =   31
         Top             =   240
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   4
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   3
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   5
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.CheckBox Chk_nao_aprovados 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não aprovados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   13470
         TabIndex        =   10
         Top             =   1125
         Width           =   1605
      End
      Begin VB.CheckBox Chk_aprovados 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aprovados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12120
         TabIndex        =   9
         Top             =   1125
         Width           =   1245
      End
      Begin VB.CheckBox Chk_aprovar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "À aprovar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10860
         TabIndex        =   8
         Top             =   1125
         Width           =   1185
      End
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frmCompras_Aprovar_Pedido.frx":EC51
         Left            =   180
         List            =   "frmCompras_Aprovar_Pedido.frx":EC53
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   5715
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         ItemData        =   "frmCompras_Aprovar_Pedido.frx":EC55
         Left            =   5910
         List            =   "frmCompras_Aprovar_Pedido.frx":EC77
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4305
      End
      Begin VB.TextBox txtTexto 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1065
         Width           =   10575
      End
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmCompras_Aprovar_Pedido.frx":ECE9
         Left            =   180
         List            =   "frmCompras_Aprovar_Pedido.frx":ECEB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1065
         Width           =   10575
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   2670
         TabIndex        =   29
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   7642
         TabIndex        =   22
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   4732
         TabIndex        =   21
         Top             =   840
         Width           =   1470
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   28
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   7
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
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
      ButtonWidth1    =   36
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Aprovação"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Aprovar/Não aprovar (F3)"
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
      ButtonWidth2    =   60
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Cancelar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Cancelar operação (F4)"
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
      ButtonLeft3     =   102
      ButtonTop3      =   2
      ButtonWidth3    =   50
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   154
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
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
      ButtonLeft5     =   158
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
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
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   224
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9660
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_Aprovar_Pedido.frx":ECED
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3180
      Left            =   60
      TabIndex        =   11
      Top             =   2550
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   5609
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Pedido"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Vlr. total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cond. de pagamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Observações"
         Object.Width           =   6536
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "IDempresa"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_itens 
      Height          =   3360
      Left            =   60
      TabIndex        =   19
      Top             =   6390
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   5927
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5389
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Vlr. unit."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. unit. desc."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Vlr. total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Prazo entr."
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Observações"
         Object.Width           =   2646
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   26
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "frmCompras_Aprovar_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql_Compras_AprovarPedido As String 'OK
Dim TBLISTA_Compras_AprovarPedido As ADODB.Recordset 'OK

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Lista_itens.ListItems.Clear
If Sql_Compras_AprovarPedido = "" Then Exit Sub
Set TBLISTA_Compras_AprovarPedido = CreateObject("adodb.recordset")
TBLISTA_Compras_AprovarPedido.Open Sql_Compras_AprovarPedido, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Compras_AprovarPedido.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear
TBLISTA_Compras_AprovarPedido.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Compras_AprovarPedido.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Compras_AprovarPedido.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Compras_AprovarPedido.RecordCount - IIf(Pagina > 1, (TBLISTA_Compras_AprovarPedido.PageSize * (Pagina - 1)), 0), TBLISTA_Compras_AprovarPedido.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Compras_AprovarPedido.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Compras_AprovarPedido!IDpedido
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!Pedido), "", TBLISTA_Compras_AprovarPedido!Pedido)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!Data), "", Format(TBLISTA_Compras_AprovarPedido!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!Responsavel), "", TBLISTA_Compras_AprovarPedido!Responsavel)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!Fornecedor), "", TBLISTA_Compras_AprovarPedido!Fornecedor)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!dbl_valor_total), "0,00", Format(TBLISTA_Compras_AprovarPedido!dbl_valor_total, "###,##0.00"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!condicoes), "", TBLISTA_Compras_AprovarPedido!condicoes)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!Observacoes), "", TBLISTA_Compras_AprovarPedido!Observacoes)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Compras_AprovarPedido!ID_empresa), 0, TBLISTA_Compras_AprovarPedido!ID_empresa)
    End With
    TBLISTA_Compras_AprovarPedido.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Compras_AprovarPedido.RecordCount
If TBLISTA_Compras_AprovarPedido.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Compras_AprovarPedido.PageCount
ElseIf TBLISTA_Compras_AprovarPedido.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Compras_AprovarPedido.PageCount & " de: " & TBLISTA_Compras_AprovarPedido.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Compras_AprovarPedido.AbsolutePage - 1 & " de: " & TBLISTA_Compras_AprovarPedido.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_aprovados_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_aprovar_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_nao_aprovados_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear
txtTexto.Visible = True
cmbTexto.Visible = False
If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    ProcCarregaComboFamilia cmbTexto, "Familia is not null", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

StatusProdServ = ""
If Chk_aprovar.Value = 1 Then StatusProdServ = "CPL.Status_Item = 'AGUARDANDO APROVAÇÃO'"
If Chk_aprovados.Value = 1 Then
   If StatusProdServ = "" Then StatusProdServ = "CPL.Status_Item = 'APROVADO' or CPL.Status_Item = 'N_RECEBIDO'" Else StatusProdServ = StatusProdServ & " or CPL.Status_Item = 'APROVADO' or CPL.Status_Item = 'N_RECEBIDO'"
End If
If Chk_nao_aprovados.Value = 1 Then
   If StatusProdServ = "" Then StatusProdServ = "CPL.Status_Item = 'CANCELADO'" Else StatusProdServ = StatusProdServ & " or CPL.Status_Item = 'CANCELADO'"
End If
If StatusProdServ = "" Then StatusProdServ = "CPL.Status_Item IS NOT NULL"

CamposFiltro = "CP.IDpedido, CP.Pedido, CP.Data, CP.Responsavel, CP.Fornecedor, CP.dbl_valor_total, CP.ID_empresa, CDC.condicoes, CDC.observacoes"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.Idpedido = CPL.Idpedido) INNER JOIN Compras_comercial CDC ON CDC.IDpedido = CP.IDpedido) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = CPL.codproduto"
TextoFiltroPadrao = "CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and CP.Status_pedido <> 'ENCERRADO' and CP.Status_pedido <> 'PARCIAL' and CPL.Idpedido <> 0 and CP.DtValidacao IS NOT NULL and (" & StatusProdServ & ") group by " & CamposFiltro & " order by CP.Idpedido desc"
If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
    If cmbfiltrarpor = "Família" Then
        Sql_Compras_AprovarPedido = INNERJOINTEXTO & " where CPL.Familia = '" & cmbTexto & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor
            Case "Pedido": TextoFiltro = "CP.pedido"
            Case "Fornecedor": TextoFiltro = "CP.fornecedor"
            Case "Código interno": TextoFiltro = "CPL.desenho"
            Case "Descrição": TextoFiltro = "CPL.descricao"
            Case "Descrição comercial": TextoFiltro = "CPL.descricao_comercial"
            Case "Detalhe": TextoFiltro = "CPL.Detalheitem"
            Case "Ordem": TextoFiltro = "CPL.Ordem"
            Case "OS": TextoFiltro = "CPL.OS"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
        End Select
        Sql_Compras_AprovarPedido = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    Sql_Compras_AprovarPedido = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Compras_AprovarPedido.AbsolutePage <> 2 Then
    If TBLISTA_Compras_AprovarPedido.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Compras_AprovarPedido.PageCount - 1)
    Else
        TBLISTA_Compras_AprovarPedido.AbsolutePage = TBLISTA_Compras_AprovarPedido.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Compras_AprovarPedido.AbsolutePage)
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
    TBLISTA_Compras_AprovarPedido.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Compras_AprovarPedido.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Compras_AprovarPedido.AbsolutePage = 1
ProcExibePagina (TBLISTA_Compras_AprovarPedido.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Compras_AprovarPedido.AbsolutePage <> -3 Then
    If TBLISTA_Compras_AprovarPedido.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Compras_AprovarPedido.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Compras_AprovarPedido.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Compras_AprovarPedido.AbsolutePage = TBLISTA_Compras_AprovarPedido.PageCount
ProcExibePagina (TBLISTA_Compras_AprovarPedido.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcAprovacao
    Case vbKeyF4: ProcCancelar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
Formulario = "Compras/Pedido/Aprovar"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Pedido"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Pedido/Aprovar"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmestoque_item_imprimir.Show 1

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
procCarregalista_Itens

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_itens_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_itens
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_itens, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

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

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_itens.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcAprovacao
    Case 3: ProcCancelar
    'Case 5: ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCarregalista_Itens()
On Error GoTo tratar_erro

Lista_itens.ListItems.Clear
Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * FROM compras_pedido_lista where IDpedido = " & Lista.SelectedItem & " order by desenho, idlista desc", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBCompras_Lista.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBCompras_Lista.EOF = False
        With Lista_itens.ListItems.Add(, , TBCompras_Lista!IDlista)
            .SubItems(1) = IIf(IsNull(TBCompras_Lista!Desenho), "", TBCompras_Lista!Desenho)
            .SubItems(2) = IIf(IsNull(TBCompras_Lista!Descricao), "", TBCompras_Lista!Descricao)
            .SubItems(3) = IIf(IsNull(TBCompras_Lista!Un), "", TBCompras_Lista!Un)
            .SubItems(4) = IIf(IsNull(TBCompras_Lista!Familia), "", TBCompras_Lista!Familia)
            .SubItems(5) = IIf(IsNull(TBCompras_Lista!Quant_Comp), 0, Format(TBCompras_Lista!Quant_Comp, "###,##0.0000"))
            .SubItems(6) = IIf(IsNull(TBCompras_Lista!preco_unitario), "", Format(TBCompras_Lista!preco_unitario, "###,##0.0000000000"))
            .SubItems(7) = IIf(IsNull(TBCompras_Lista!preco_unitario_desconto), 0, Format(TBCompras_Lista!preco_unitario_desconto, "###,##0.0000000000"))
            .SubItems(8) = IIf(IsNull(TBCompras_Lista!preco_total), 0, Format(TBCompras_Lista!preco_total, "###,##0.00"))
            .SubItems(9) = IIf(IsNull(TBCompras_Lista!Prazo), "", Format(TBCompras_Lista!Prazo, "dd/mm/yy"))
            If TBCompras_Lista!Status_Item = "AGUARDANDO APROVAÇÃO" Or TBCompras_Lista!Status_Item = "APROVADO" Or TBCompras_Lista!Status_Item = "RECEBIDO" Or TBCompras_Lista!Status_Item = "CANCELADO" Then
                Status_Item = TBCompras_Lista!Status_Item
            ElseIf TBCompras_Lista!Status_Item = "N_RECEBIDO" Then
                    Status_Item = "COMPRADO"
                Else
                    Status_Item = "RECEBIDO PARCIAL"
            End If
            .SubItems(10) = Status_Item
            .SubItems(11) = IIf(IsNull(TBCompras_Lista!Obs), "", TBCompras_Lista!Obs)
        End With
        TBCompras_Lista.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBCompras_Lista.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAprovacao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente aprovar/não aprovar esse(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    
    'Verifica se o usuario pode aprovar o pedido de acordo com o limite cadastrado
    valor = 0
    With Lista_itens
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then valor = valor + .ListItems(InitFor).ListSubItems(8)
        Next InitFor
    End With
    If valor > 0 Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select A.Valor_Limite from usuarios U INNER JOIN acessos A ON A.IDUsuario = U.IDUsuario where U.usuario = '" & pubUsuario & "' and A.Acesso = 'Compras/Pedido/Aprovar' and A.Validacao = 'True' and Valor_Limite IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtd = TBAbrir!Valor_Limite
            If valor > Qtd Then
                USMsgBox ("Atenção usuário " & pubUsuario & ", você não tem autorização para aprovar esse(s) produto(s)/serviço(s) pois ultrapassou o valor permitido."), vbExclamation, "CAPRIND v5.0"
                TBAbrir.Close
                Exit Sub
            End If
        End If
        TBAbrir.Close
    End If
    
    With Lista_itens
        For InitFor = 1 To .ListItems.Count
            CamposTexto = ""
            Familiatext = ""
            If .ListItems.Item(InitFor).Checked = True Then
                'Aprovados
                If FunVerifStatusAprovadoPC(Lista.SelectedItem.ListSubItems(8)) = True Then TextoStatus = "APROVADO" Else TextoStatus = "N_RECEBIDO"
                Familiatext = ""
            Else
                'Não aprovados
                TextoStatus = "CANCELADO"
                If .ListItems(InitFor).SubItems(10) <> "CANCELADO" Then
1:
                    Familiatext = InputBox("Favor informar o motivo da não aprovação do produto/serviço " & .ListItems(InitFor).SubItems(1) & " - " & .ListItems(InitFor).SubItems(2) & ".")
                    If StrPtr(Familiatext) = 0 Then
                        USMsgBox ("Não é permitido cancelar essa operação."), vbExclamation, "CAPRIND v5.0"
                        GoTo 1
                    Else
                        If Familiatext = "" Then GoTo 1
                    End If
                End If
            End If
            If Familiatext <> "" Then
                CamposTexto = ", Resp_cancelado = '" & pubUsuario & "', Data_cancelado = '" & Date & "', Motivo_cancelado = '" & Familiatext & "'"
            ElseIf TextoStatus <> "CANCELADO" Then
                    CamposTexto = ", Resp_cancelado = 'NULL', Data_cancelado = NULL, Motivo_cancelado = 'NULL'"
            End If
            Conexao.Execute "UPDATE Compras_pedido_lista Set Status_Item = '" & TextoStatus & "'" & CamposTexto & " where IDlista = " & .ListItems(InitFor)
        Next InitFor
    End With
    Call FunAtualizaStatusPC(Lista.SelectedItem)
    ProcGravarTotaisPC (Lista.SelectedItem)
    
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from Compras_pedido where IDpedido = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = False Then
        If TBCompras!Status_pedido = "APROVADO" Or TBCompras!Status_pedido = "ABERTO" Then
            FunAlterarProdSimiliarOrdemPC TBCompras!ID_empresa, TBCompras!IDpedido
            ProcCriarRMOrdemPC TBCompras!IDpedido, TBCompras!ID_empresa
        
            TBCompras!Resp_aprovado = pubUsuario
            TBCompras!Data_aprovado = Format(Date, "dd/mm/yy")
        Else
            ProcExcluirRMOrdemPC TBCompras!IDpedido, TBCompras!ID_empresa
            
            TBCompras!Resp_aprovado = Null
            TBCompras!Data_aprovado = Null
        End If
        TBCompras.Update
    End If
    TBCompras.Close
    USMsgBox ("Operação de aprovação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Compras/Pedido/Aprovar"
    Evento = "Aprovar/não aprovar pedido de compra"
    With Lista
        ID_documento = .SelectedItem
        Documento = "Nº pedido: " & .SelectedItem.ListSubItems(1)
    End With
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Lista.ListItems.Clear
    Lista_itens.ListItems.Clear
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With Lista
    If .ListItems.Count = 0 Then Exit Sub
    If USMsgBox("Deseja realmente cancelar a operação de aprovação deste pedido de compra?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select ID_empresa, IDpedido, Pedido, Resp_aprovado, Data_aprovado, Status_pedido from Compras_pedido where IDpedido = " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        TBCompras!Resp_aprovado = Null
        TBCompras!Data_aprovado = Null
        TBCompras!Status_pedido = "AGUARDANDO APROVAÇÃO"
        TBCompras.Update
        
        Conexao.Execute "UPDATE Compras_pedido_lista Set Status_Item = 'AGUARDANDO APROVAÇÃO', Resp_cancelado = 'NULL', Data_cancelado = NULL, Motivo_cancelado = 'NULL' where IDpedido = " & TBCompras!IDpedido
        ProcExcluirRMOrdemPC TBCompras!IDpedido, TBCompras!ID_empresa
        
        ProcGravarTotaisPC (Lista.SelectedItem)
        
        '==================================
        Modulo = "Compras/Pedido/Aprovar"
        Evento = "Cancelar operação de aprovação do pedido de compra"
        ID_documento = TBCompras!IDpedido
        Documento = "Nº pedido: " & TBCompras!Pedido
        Documento1 = ""
        ProcGravaEvento
        '==================================
        USMsgBox ("Operação de aprovação cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
        Lista.ListItems.Clear
        Lista_itens.ListItems.Clear
        ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


