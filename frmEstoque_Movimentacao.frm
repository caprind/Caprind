VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmestoque_Movimentacao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Movimentação"
   ClientHeight    =   10125
   ClientLeft      =   3870
   ClientTop       =   2325
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10125
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   1035
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   1826
      ButtonCount     =   6
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
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   51
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Exportar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Exportar relatório"
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
      ButtonLeft3     =   93
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
      ButtonLeft4     =   145
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   58
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
      ButtonLeft5     =   149
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
      ButtonLeft6     =   187
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13020
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmEstoque_Movimentacao.frx":0000
         Count           =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7785
      Left            =   0
      TabIndex        =   16
      Top             =   2250
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   13732
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Relação de itens (Saldos)"
      TabPicture(0)   =   "frmEstoque_Movimentacao.frx":3125
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "PBLista"
      Tab(0).Control(1)=   "Frame9"
      Tab(0).Control(2)=   "GridProdutos"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Relação de RE´s (Lotes)"
      TabPicture(1)   =   "frmEstoque_Movimentacao.frx":3141
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GridRE"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Movimentações do lote"
      TabPicture(2)   =   "frmEstoque_Movimentacao.frx":315D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GridMv"
      Tab(2).ControlCount=   1
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   -74970
         TabIndex        =   29
         Top             =   6960
         Width           =   15255
         _ExtentX        =   26908
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
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   17
         Top             =   7140
         Width           =   15265
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
            Left            =   3120
            TabIndex        =   19
            Text            =   "22"
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
            Left            =   9690
            TabIndex        =   18
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11910
            TabIndex        =   20
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Movimentacao.frx":3179
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
            Left            =   11370
            TabIndex        =   21
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Movimentacao.frx":691D
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
            Left            =   10260
            TabIndex        =   22
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
            Left            =   10830
            TabIndex        =   23
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Movimentacao.frx":A426
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
            Left            =   12450
            TabIndex        =   24
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Movimentacao.frx":E515
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
            Left            =   3750
            TabIndex        =   28
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2430
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13200
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
      End
      Begin FlexCell.Grid GridMv 
         Height          =   7395
         Left            =   -74970
         TabIndex        =   30
         Top             =   330
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   13044
         Appearance      =   0
         BackColorBkg    =   16777215
         Cols            =   12
         DefaultFontSize =   8.25
         DisplayFocusRect=   0   'False
         GridColor       =   12632256
         ReadOnly        =   -1  'True
         Rows            =   30
      End
      Begin FlexCell.Grid GridRE 
         Height          =   7395
         Left            =   30
         TabIndex        =   31
         Top             =   330
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   13044
         Appearance      =   0
         BackColorBkg    =   16777215
         Cols            =   10
         DefaultFontSize =   8.25
         DisplayFocusRect=   0   'False
         GridColor       =   12632256
         Rows            =   26
         DateFormat      =   2
      End
      Begin FlexCell.Grid GridProdutos 
         Height          =   6615
         Left            =   -74970
         TabIndex        =   32
         Top             =   330
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   11668
         Appearance      =   0
         BackColorBkg    =   16777215
         Cols            =   5
         DefaultFontSize =   8.25
         DisplayFocusRect=   0   'False
         GridColor       =   14737632
         Rows            =   26
         DateFormat      =   2
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
      FormHeightDT    =   10590
      FormWidthDT     =   15480
      FormScaleHeightDT=   10125
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para filtrar    "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1245
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1020
      Width           =   15345
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   825
         Left            =   14070
         TabIndex        =   13
         ToolTipText     =   "Filtrar itens no estoque"
         Top             =   300
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1455
         DibPicture      =   "frmEstoque_Movimentacao.frx":11DA1
         Caption         =   "Filtrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         Theme           =   4
         ToolTipIcon     =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin VB.CheckBox chkBloqueados 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Itens bloqueados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4470
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmEstoque_Movimentacao.frx":153F1
         Left            =   180
         List            =   "frmEstoque_Movimentacao.frx":153F3
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Empresa."
         Top             =   630
         Width           =   4155
      End
      Begin VB.OptionButton optIgual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Igual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13170
         TabIndex        =   8
         Top             =   360
         Width           =   705
      End
      Begin VB.OptionButton Optmeio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Meio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11730
         TabIndex        =   7
         Top             =   360
         Width           =   645
      End
      Begin VB.OptionButton Optinicio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Início"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton Optfim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fim"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12450
         TabIndex        =   5
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtTexto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   10860
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   630
         Width           =   3105
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmEstoque_Movimentacao.frx":153F5
         Left            =   7710
         List            =   "frmEstoque_Movimentacao.frx":1540E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   630
         Width           =   2745
      End
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEstoque_Movimentacao.frx":15451
         Left            =   10860
         List            =   "frmEstoque_Movimentacao.frx":15453
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   630
         Width           =   3105
      End
      Begin MSComCtl2.DTPicker Ate 
         Height          =   315
         Left            =   6300
         TabIndex        =   11
         ToolTipText     =   "Data fim do período do filtro"
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   198508545
         CurrentDate     =   43822
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8550
         TabIndex        =   15
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1755
         TabIndex        =   14
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6690
         TabIndex        =   12
         Top             =   420
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmestoque_Movimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Documento_Ordem As String 'OK
Dim Sql_Estoque_Movimentacao As String 'OK
Dim TBLISTA_Estoque_Movimentacao As ADODB.Recordset 'OK
Dim Status_movimentacao As String 'OK

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

txtlocalização = ""
Txt_cod_ref = ""
Txt_n_serie = ""
Txt_qtde_estoqueRE = "0,0000"
Txt_qtde_estoque_PCRE = "0,0000"
Txt_qtde_empenhoRE = "0,0000"
Txt_qtde_est_dispRE = "0,0000"
Txt_qtde_est_disp_PCRE = "0,0000"
Txt_qtde_est_tercRE = "0,0000"
Txt_valor_total_estRE = "0,00"
Txt_valor_unitRE = "0,0000"

Txt_qtde_estoque = "0,0000"
Txt_qtde_estoque_PC = "0,0000"
Txt_qtde_empenho = "0,0000"
Txt_qtde_est_disp = "0,0000"
Txt_qtde_est_disp_PC = "0,0000"
Txt_qtde_est_terc = "0,0000"
Txt_valor_total_est = "0,00"
Txt_custo_medio_unit = "0,00000"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

'lblRegistros.Caption = "Nº de registros: 0"
'lblPaginas.Caption = "Página: 0 de: 0"
GridProdutos.rows = 1

If Sql_Estoque_Movimentacao = "" Then Exit Sub
Set TBLISTA_Estoque_Movimentacao = CreateObject("adodb.recordset")
'Debug.print Sql_Estoque_Movimentacao
TBLISTA_Estoque_Movimentacao.Open Sql_Estoque_Movimentacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Estoque_Movimentacao.EOF = False Then ProcExibePaginaGrid (Pagina) Else CodigoLista = 0
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
TBLISTA_Estoque_Movimentacao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Estoque_Movimentacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_Movimentacao.PageSize
ContadorReg = 1

'PBLista.Min = 0
'PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_Movimentacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_Movimentacao.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_Movimentacao.PageSize)
'PBLista.Value = 1
Contador = 0
Do While TBLISTA_Estoque_Movimentacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , IIf(IsNull(TBLISTA_Estoque_Movimentacao!IDEstoque), 0, TBLISTA_Estoque_Movimentacao!IDEstoque)
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Data), "", Format(TBLISTA_Estoque_Movimentacao!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!LOTE), "", TBLISTA_Estoque_Movimentacao!LOTE)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Desenho), "", TBLISTA_Estoque_Movimentacao!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Ref), "", TBLISTA_Estoque_Movimentacao!Ref)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Descricao), "", TBLISTA_Estoque_Movimentacao!Descricao)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Unidade), "", TBLISTA_Estoque_Movimentacao!Unidade)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Classe), "", TBLISTA_Estoque_Movimentacao!Classe)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!local_armaz), "", TBLISTA_Estoque_Movimentacao!local_armaz)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Corrida), "", TBLISTA_Estoque_Movimentacao!Corrida)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Certificado), "", TBLISTA_Estoque_Movimentacao!Certificado)
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Numero_serie), "", TBLISTA_Estoque_Movimentacao!Numero_serie)
        
        Cliente = ""
        If IsNull(TBLISTA_Estoque_Movimentacao!Fornecedor) = False And TBLISTA_Estoque_Movimentacao!Fornecedor <> "" Then
            Cliente = TBLISTA_Estoque_Movimentacao!Fornecedor
        ElseIf IsNull(TBLISTA_Estoque_Movimentacao!Cliente) = False And TBLISTA_Estoque_Movimentacao!Cliente <> "" Then
                Cliente = TBLISTA_Estoque_Movimentacao!Cliente
        End If
        .Item(.Count).SubItems(12) = Cliente
        
        .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!estoque_real), "", Format(TBLISTA_Estoque_Movimentacao!estoque_real, "###,##0.0000"))
        .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!estoque_real_PC), "", Format(TBLISTA_Estoque_Movimentacao!estoque_real_PC, "###,##0.0000"))
        .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!valor_unitario), "", Format(TBLISTA_Estoque_Movimentacao!valor_unitario, "###,##0.0000000000"))
        .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Valor_total), "", Format(TBLISTA_Estoque_Movimentacao!Valor_total, "###,##0.00"))
        .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Liberado), "", TBLISTA_Estoque_Movimentacao!Liberado)
    End With
    TBLISTA_Estoque_Movimentacao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    'PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_Movimentacao.RecordCount
If TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_Movimentacao.PageCount
ElseIf TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.PageCount & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaGridOLD(Pagina)
On Error GoTo tratar_erro

TBLISTA_Estoque_Movimentacao.PageSize = IIf(txtNreg = "", 22, txtNreg)
TBLISTA_Estoque_Movimentacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_Movimentacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_Movimentacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_Movimentacao.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_Movimentacao.PageSize)
PBLista.Value = 1

Contador = 0
TotalSaldo = 0
TotalEntrada = 0
TotalSaida = 0
TotalEmpenho = 0

GridRE.rows = 1

Do While TBLISTA_Estoque_Movimentacao.EOF = False And (ContadorReg <= TamanhoPagina)

                 EstoqueEntrada = TBLISTA_Estoque_Movimentacao!ttEntrada
                 EstoqueSaida = TBLISTA_Estoque_Movimentacao!ttsaida
                 EstoqueEmpenho = TBLISTA_Estoque_Movimentacao!Qtde_empenhada
                 EstoqueSaldo = EstoqueEntrada - EstoqueSaida - EstoqueEmpenho


    GridRE.AddItem Contador + 1 & vbTab & _
                 TBLISTA_Estoque_Movimentacao!Data & vbTab & _
                 TBLISTA_Estoque_Movimentacao!IDEstoque & vbTab & _
                 TBLISTA_Estoque_Movimentacao!LOTE & vbTab & _
                 TBLISTA_Estoque_Movimentacao!status & vbTab & _
                 TBLISTA_Estoque_Movimentacao!Desenho & vbTab & _
                 TBLISTA_Estoque_Movimentacao!Descricao & vbTab & _
                 TBLISTA_Estoque_Movimentacao!Unidade & vbTab & _
                 Format(TBLISTA_Estoque_Movimentacao!ttEntrada, "0.00") & vbTab & _
                 Format(TBLISTA_Estoque_Movimentacao!ttsaida, "0.00") & vbTab & _
                 Format(TBLISTA_Estoque_Movimentacao!Qtde_empenhada, "0.00") & vbTab & _
                 Format(EstoqueSaldo, "0.00") 'Saldo calculado
                 'Format(TBLISTA_Estoque_Movimentacao!Estoque_disponivel, "0.00") '(Saldo do banco)
                 TotalEntrada = TotalEntrada + EstoqueEntrada
                 TotalSaida = TotalSaida + EstoqueSaida
                 TotalEmpenho = TotalEmpenho + EstoqueEmpenho
                 'Saldo = Saldo + EstoqueSaldo
    TBLISTA_Estoque_Movimentacao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
    'frmestoque_Movimentacao.Refresh
Loop
TotalSaldo = TotalEntrada - TotalSaida - TotalEmpenho
GridRE.AddItem Contador & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Saldo:" & vbTab & Format(TotalSaldo, "0.00")

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_Movimentacao.RecordCount
If TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_Movimentacao.PageCount
ElseIf TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.PageCount & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
End If

GridRE.AutoRedraw = True
GridRE.Refresh


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaGrid(Pagina)
On Error GoTo tratar_erro


TBLISTA_Estoque_Movimentacao.PageSize = IIf(txtNreg = "", 22, txtNreg)
TBLISTA_Estoque_Movimentacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_Movimentacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_Movimentacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_Movimentacao.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_Movimentacao.PageSize)
PBLista.Value = 1

Contador = 0
TotalSaldo = 0
TotalEntrada = 0
TotalSaida = 0
TotalEmpenho = 0

GridProdutos.rows = 1

Do While TBLISTA_Estoque_Movimentacao.EOF = False And (ContadorReg <= TamanhoPagina)
 EstoqueSaldo = TBLISTA_Estoque_Movimentacao!Saldo 'EstoqueEntrada - EstoqueSaida
 GridProdutos.AddItem Contador + 1 & vbTab & _
 TBLISTA_Estoque_Movimentacao!Desenho & vbTab & _
 TBLISTA_Estoque_Movimentacao!Descricao & vbTab & _
 TBLISTA_Estoque_Movimentacao!Unidade & vbTab & _
 Format(EstoqueSaldo, "0.00")
 TBLISTA_Estoque_Movimentacao.MoveNext
 ContadorReg = ContadorReg + 1
 Contador = Contador + 1
 PBLista.Value = Contador
Loop

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_Movimentacao.RecordCount
If TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_Movimentacao.PageCount
ElseIf TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.PageCount & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
End If

GridProdutos.AutoRedraw = True
GridProdutos.Refresh


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

If IDempresa = 0 Then
Exit Sub
End If

If cmbfiltrarpor.Text = "" Then
USMsgBox "Tem que informar uma das opções para pesquisa", vbCritical, "CAPRIND v5.0"
cmbfiltrarpor.SetFocus
Exit Sub
End If

SSTab1.Tab = 0
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkperiodo_Click()
On Error GoTo tratar_erro

If chkPeriodo.Value = 1 Then
De.Enabled = True
Ate.Enabled = True
Else
De.Enabled = False
Ate.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Change()
On Error GoTo tratar_erro

IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

txtTexto.Visible = True
cmbTexto.Visible = False
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Local de armazenamento" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Operação" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    cmbTexto.Clear
    
    Select Case cmbfiltrarpor
        Case "Local de armazenamento": ProcCarregaComboLA cmbTexto, True, True
        Case "Família": ProcCarregaComboFamilia cmbTexto, "Familia is not null", False
        Case "Grupo": ProcCarregaComboGrupoFamilia cmbTexto, "Grupo is not null", False
        Case "Operação": ProcCarregaComboOperacao cmbTexto, "Operacao is not null", False
    End Select
ElseIf cmbfiltrarpor = "RE" And txtTexto <> "" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlterar_valor()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista.ListItems.Count = 0 Then Exit Sub
If USMsgBox("Deseja realmente atualizar o valor unitário deste RE?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem1:
     Valorunitario = InputBox("Favor informar o valor unitário.")
     If Valorunitario = "" Then Exit Sub
     If IsNumeric(Valorunitario) = False Then
         USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
         GoTo Mensagem1
     End If
     If Valorunitario = 0 Then
         USMsgBox ("Não é permitido colocar zero no valor unitário."), vbExclamation, "CAPRIND v5.0"
         GoTo Mensagem1
     End If
     valor = Valorunitario
     NovoValor = Replace(valor, ",", ".")
     Conexao.Execute "UPDATE Estoque_Controle Set valor_unitario = " & NovoValor & " where IDEstoque = " & Lista.SelectedItem
     Conexao.Execute "UPDATE Estoque_Controle Set Valor_total = ROUND(valor_unitario * Estoque_real, 2) where IDEstoque = " & Lista.SelectedItem
     Conexao.Execute "UPDATE Estoque_movimentacao Set VlrUnit = " & NovoValor & " where IDEstoque = " & Lista.SelectedItem
     Conexao.Execute "UPDATE Estoque_movimentacao Set VlrTotal = ROUND(VlrUnit * Entrada, 2) where IDEstoque = " & Lista.SelectedItem & " and Entrada <> 0"
     Conexao.Execute "UPDATE Estoque_movimentacao Set VlrTotal = ROUND(VlrUnit * Saida, 2) where IDEstoque = " & Lista.SelectedItem & " and Saida <> 0"
     Conexao.Execute "Update CC set CC.Valor = EM.VlrTotal from CC_realizado CC INNER JOIN Estoque_movimentacao EM on CC.ID_estoque = EM.Idoperacao where ID_estoque = " & Lista.SelectedItem
    
     Set TBFIltro = CreateObject("adodb.recordset")
     TBFIltro.Open "Select Documento from Estoque_movimentacao where IDEstoque = " & Lista.SelectedItem & " and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL') group by Documento", Conexao, adOpenKeyset, adLockOptimistic
     If TBFIltro.EOF = False Then
         Do While TBFIltro.EOF = False
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IDlista from Compras_pedido_lista where Ordem = " & TBFIltro!Documento & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                ProcAtualizaCTMaterialOrdem Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBFIltro!Documento
            End If
            TBAbrir.Close
             TBFIltro.MoveNext
         Loop
     End If
     TBFIltro.Close
     USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
     '==================================
     Modulo = "Estoque/Movimentação"
     Evento = "Alterar valor unitário do inventário"
     ID_documento = txtlocalização
     Documento = "Cód. interno: " & Lista.SelectedItem.ListSubItems(3)
     Documento1 = ""
     ProcGravaEvento
     '==================================
     ProcFiltrar
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcFiltrar2()
On Error GoTo tratar_erro

Acao = "filtrar"

If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If

Conexao.Execute ("update Estoque_movimentacao Set Grupo = PF.Grupo from Estoque_movimentacao EM inner Join Projfamilia PF On EM.Familia = PF.Familia")

TextoFiltroQtde = ""

'Itens bloqueados
If chkBloqueados.Value = 1 Then
StatusFiltro = " EM.bloqueado = 'True' AND EM.id_empresa = '" & IDempresa & "'"
FormulaRelatorioCampo = " And {Estoque_Controle_Saldo_RE.saldo} > 0 And {Estoque_Controle_Saldo_RE.bloqueado} = True AND {Estoque_Controle_Saldo_RE.ID_empresa} = " & IDempresa & " and {Empresa.codigo} = " & IDempresa & ""
Else
StatusFiltro = " EM.bloqueado = 'False' AND EM.id_empresa = '" & IDempresa & "'"
FormulaRelatorioCampo = " And {Estoque_Controle_Saldo_RE.saldo} > 0 And {Estoque_Controle_Saldo_RE.bloqueado} = False AND {Estoque_Controle_Saldo_RE.ID_empresa} = " & IDempresa & " and {Empresa.codigo} = " & IDempresa & ""
End If

'Executa o link entre tabelas
INNERJOINTEXTO = "SELECT EM.Grupo, Familia, EM.Codigo, EM.Descricao,EM.Un, SUM(EM.Saldo) AS saldo, SUM(Valor_total) AS valor_Total, EM.Estoque from Estoque_Controle_Saldo_RE EM"

'Filtro padrão
TextoFiltroPadrao = TextoFiltroQtde & TextoFiltroTerc & TextoFiltroEmp & StatusFiltro & " group by " & "EM.Grupo, EM.Familia, EM.Codigo, EM.Descricao, EM.Un,EM.Estoque HAVING (SUM(EM.Saldo) > 0)"
'TextoFiltroPadrao = TextoFiltroQtde & TextoFiltroTerc & TextoFiltroEmp & StatusFiltro & " group by " & "EM.Desenho,PP.Classe, EM.Descricao, PP.Unidade" & " order by PP.Classe, Desenho"

'Filtro por período
TextoFiltroPadrao = "EM.Data <= '" & Ate & "' AND " & TextoFiltroPadrao

'=================================================================================================
' Filtro para reltório
'=================================================================================================
strDataFim = "Date(" & Format(Ate.Value, "yyyy,mm,dd") & ")"
FormulaRelatorio = "{Estoque_Controle_Saldo_RE.Data} <= " & strDataFim & ""
'=================================================================================================


'Campo a ser filtrado
If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
    Select Case cmbfiltrarpor
        
        Case "Família":
        TextoFiltro = "EM.Familia"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Familia}='" & cmbTexto & "'"
        
        Case "Grupo":
        TextoFiltro = "EM.Grupo"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Grupo}='" & cmbTexto & "'"
        
        Case "Código interno":
        TextoFiltro = "EM.Codigo"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.codigo}='" & txtTexto & "'"
        
        
        Case "Descrição":
        TextoFiltro = "EM.descricao"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.descricao} like '" & txtTexto & "*'"
        
        Case "Lote":
        TextoFiltro = "EM.Lote"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Lote}='" & txtTexto & "'"
        
       
        Case "Documento":
        TextoFiltro = "EM.Documento"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Documento}='" & txtTexto & "'"
        
        Case "RE":
        TextoFiltro = "EM.IDestoque"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.IDestoque}=" & txtTexto & ""
        
        Case "Operação":
        TextoFiltro = "EM.Operacao"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Status}='" & cmbTexto & "'"
           
    End Select
    
    If cmbfiltrarpor = "Local de armazenamento" Or cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Operação" Then
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
    Else
        If cmbfiltrarpor = "Documento" Then
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & txtTexto & "'" & " and " & TextoFiltroPadrao
        Else
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
        End If
    End If
   

Else
    Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If


FormulaRelatorio = FormulaRelatorio & FormulaRelatorioCampo & OpcaoFiltro

'Debug.print FormulaRelatorio
'Debug.print Sql_Estoque_Movimentacao

ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Public Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"

If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If

'Conexao.Execute ("update Estoque_movimentacao Set Bloqueado = EC.Bloqueado from Estoque_movimentacao EM inner Join Estoque_controle EC On EM.idestoque = EC.idEstoque")
'Conexao.Execute ("update Estoque_movimentacao Set ID_empresa = EC.id_empresa from Estoque_movimentacao EM inner Join Estoque_controle EC On EM.idestoque = EC.idEstoque")
Conexao.Execute ("update Estoque_movimentacao Set Familia = PP.Classe, Unidade = PP.Unidade from Estoque_movimentacao EM inner Join ProjProduto PP On EM.Desenho = PP.Desenho")
Conexao.Execute ("update Estoque_movimentacao Set Grupo = PF.Grupo from Estoque_movimentacao EM inner Join Projfamilia PF On EM.Familia = PF.Familia")
Conexao.Execute ("update Estoque_movimentacao Set Bloqueado = 'False' from Estoque_movimentacao Where Bloqueado Is null")

TextoFiltroQtde = ""
StatusFiltro = ""
'Itens bloqueados
If chkBloqueados.Value = 1 Then
StatusFiltro = " EM.bloqueado = 'True' AND EM.id_empresa = '" & IDempresa & "'"
FormulaRelatorioCampo = " And {Estoque_Controle_Saldo_RE.saldo} > 0 And {Estoque_Controle_Saldo_RE.bloqueado} = True AND {Estoque_Controle_Saldo_RE.ID_empresa} = " & IDempresa & " and {Empresa.codigo} = " & IDempresa & ""
Else
StatusFiltro = " EM.bloqueado = 'False' AND EM.id_empresa = '" & IDempresa & "'"
FormulaRelatorioCampo = " And {Estoque_Controle_Saldo_RE.saldo} > 0 And {Estoque_Controle_Saldo_RE.bloqueado} = False AND {Estoque_Controle_Saldo_RE.ID_empresa} = " & IDempresa & " and {Empresa.codigo} = " & IDempresa & ""
End If

'Executa o link entre tabelas
'INNERJOINTEXTO = "SELECT EM.Grupo, Familia, EM.Codigo, EM.Descricao,EM.Un, SUM(EM.Saldo) AS saldo, SUM(Valor_total) AS valor_Total, EM.Estoque from Estoque_Controle_Saldo_RE EM"
INNERJOINTEXTO = "SELECT EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Unidade, SUM(EM.Entrada)-SUM(EM.Saida) AS saldo from Estoque_Movimentacao EM "
'Filtro padrão

TextoFiltroPadrao = TextoFiltroQtde & TextoFiltroTerc & TextoFiltroEmp & StatusFiltro & " group by " & "EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao, EM.Unidade"
'TextoFiltroPadrao = TextoFiltroQtde & TextoFiltroTerc & TextoFiltroEmp & StatusFiltro & " group by " & "EM.Desenho,PP.Classe, EM.Descricao, PP.Unidade" & " order by PP.Classe, Desenho"
'Debug.print TextoFiltroPadrao

'Filtro por período
TextoFiltroPadrao = "EM.Data <= '" & Ate & "' AND " & TextoFiltroPadrao
'TextoFiltroPadrao = "EM.Data <= '" & Ate & TextoFiltroPadrao

'Debug.print TextoFiltroPadrao

'=================================================================================================
' Filtro para reltório
'=================================================================================================
strDataFim = "Date(" & Format(Ate.Value, "yyyy,mm,dd") & ")"
FormulaRelatorio = "{Estoque_Controle_Saldo_RE.Data} <= " & strDataFim & ""
'=================================================================================================


'Campo a ser filtrado
If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
    Select Case cmbfiltrarpor
        
        Case "Família":
        TextoFiltro = "EM.Familia"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Familia}='" & cmbTexto & "'"
        
        Case "Grupo":
        TextoFiltro = "EM.Grupo"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Grupo}='" & cmbTexto & "'"
        
        Case "Código interno":
        TextoFiltro = "EM.Desenho"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.codigo}='" & txtTexto & "'"
        
        
        Case "Descrição":
        TextoFiltro = "EM.descricao"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.descricao} like '" & txtTexto & "*'"
        
        Case "Lote":
        TextoFiltro = "EM.Lote"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Lote}='" & txtTexto & "'"
        
       
        Case "Documento":
        TextoFiltro = "EM.Documento"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Documento}='" & txtTexto & "'"
        
        Case "RE":
        TextoFiltro = "EM.IDestoque"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.IDestoque}=" & txtTexto & ""
        
        Case "Operação":
        TextoFiltro = "EM.Operacao"
        FormulaRelatorioCampo = FormulaRelatorioCampo & " And {Estoque_Controle_Saldo_RE.Status}='" & cmbTexto & "'"
           
    End Select
    
    If cmbfiltrarpor = "Local de armazenamento" Or cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Operação" Then
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
    Else
        If cmbfiltrarpor = "Documento" Then
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & txtTexto & "'" & " and " & TextoFiltroPadrao
        Else
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
        End If
    End If
   

Else
    Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If


FormulaRelatorio = FormulaRelatorio & FormulaRelatorioCampo & OpcaoFiltro

'Debug.print FormulaRelatorio
'Debug.print Sql_Estoque_Movimentacao

ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltarNovo()
On Error GoTo tratar_erro

Acao = "filtrar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If

'Filtrar com saldo maior que zero
If chkEstoquePositivo.Value = 1 Then
 TextoFiltroQtde = " and EP.estoque_real > 0"
Else
 TextoFiltroQtde = ""
End If

'Filtrar estoque em poder de terceiros
If chkTerceiros.Value = 1 Then
TextoFiltroTerc = "and EP.destino = 'Terceiros' and EP.Terceiros = 'True'"
Else
TextoFiltroTerc = ""
End If

'Filtrar produtos bloqueados
If chkBloqueados.Value = 1 Then
StatusFiltro = ""
Else
StatusFiltro = " and EP.bloqueado = 'False'"
End If


CamposFiltro = "idestoque, Data,Familia, LOTE, Desenho, Ref, Descricao, Corrida, Certificado, Numero_serie, Fornecedor, Cliente, Unidade, estoque_real, valor_unitario, Valor_Total"
INNERJOINTEXTO = "Select " & CamposFiltro & " from estoque_movimentacao"
TextoFiltroPadrao = "(ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " or ID_empresa is null) " & TextoFiltroQtde & TextoFiltroTerc & TextoFiltroEmp & StatusFiltro & " group by " & CamposFiltro & " order by desenho, idestoque"

If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
    Select Case cmbfiltrarpor
        Case "Família": TextoFiltro = "Familia"
        Case "Código interno": TextoFiltro = "desenho"
        Case "Código de referência": TextoFiltro = "Ref"
        Case "Corrida": TextoFiltro = "EP.Corrida"
        Case "Certificado": TextoFiltro = "EP.Certificado"
        Case "Descrição": TextoFiltro = "EP.descricao"
        Case "Lote": TextoFiltro = "EP.Lote"
        Case "Documento": TextoFiltro = "EM.Documento"
        Case "Número de série": TextoFiltro = "EP.Numero_serie"
        Case "Part number": TextoFiltro = "PFAB.Part_number"
        Case "RE": TextoFiltro = "IDestoque"
    End Select
    If cmbfiltrarpor = "Local de armazenamento" Or cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
    Else
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
'ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_sucata()
On Error GoTo tratar_erro
Dim ID_Antigo As Integer 'OK

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If RE = 0 Then Exit Sub

Qtde = 0
Qtd = 0
IDlista = 0
Set TBExecucao = CreateObject("adodb.recordset")
TBExecucao.Open "select * from Estoque_Controle where IdEstoque = " & RE & " and idLote_sucata <> 0 and idLote_sucata IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBExecucao.EOF = False Then
    If TBExecucao!status = "ENTRADA_SUCATA" Then
        MsgTexto = "sucata"
        MsgTexto1 = "Sucata"
    Else
        MsgTexto = "retalho"
        MsgTexto1 = "Retalho"
    End If
    If USMsgBox("Deseja realmente excluir este RE de " & MsgTexto & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        
        'Verifica se existe movimentação de saida e bloqueia exclusão
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select Saida from estoque_movimentacao where idestoque = " & TBExecucao!IDEstoque & " and Saida <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = False Then
            USMsgBox ("Não é permitido excluir este RE, pois o mesmo já foi movimentado."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
                
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "select * from estoque_controle where IdEstoque = " & TBExecucao!idLote_sucata & " and idestoque <> " & TBExecucao!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from estoque_controle where idEstoque = " & TBExecucao!idLote_sucata, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBExecucao!status = "ENTRADA_SUCATA" Then
                    If TBAbrir!Desenho <> TBExecucao!Desenho_sucata Then
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "select * from estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
                        TBGravar.AddNew
                        Set TBAcessos = CreateObject("adodb.recordset")
                        TBAcessos.Open "Select * from estoque_movimentacao where idestoque = " & TBAbrir!IDEstoque & " and desenho = '" & TBAbrir!Desenho & "' and Operacao = 'ENTRADA_SUCATA'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAcessos.EOF = False Then
                            Do While TBAcessos.EOF = False
                                Qtde = Qtde + TBAcessos!Entrada
                                TBAcessos!IDEstoque = TBGravar!IDEstoque
                                TBAcessos.Update
                                TBAcessos.MoveNext
                            Loop
                        End If
                        TBAcessos.Close
                        
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "select * from projproduto where desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            TBGravar!Classe = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
                            TBGravar!valor_unitario = IIf(IsNull(TBItem!PCusto), "", TBItem!PCusto)
                            ValorTotal = IIf(IsNull(TBItem!PCusto), "", TBItem!PCusto)
                            TBGravar!Valor_total = Format(Qtde * ValorTotal, "###,##0.00")
                        End If
                        TBItem.Close
                        TBGravar!estoque_real = Format(Qtde, "###,##0.0000")
                        TBGravar!Qtde = TBGravar!estoque_real
                        TBGravar!estoque_venda = TBGravar!estoque_real
                        TBGravar!Desenho = TBAbrir!Desenho
                        TBGravar!LOTE = TBAbrir!LOTE
                        TBGravar!Descricao = TBAbrir!Descricao
                        TBGravar!Un = TBAbrir!Un
                        TBGravar!idLote_sucata = TBAbrir!IDEstoque
                        TBGravar!Desenho_sucata = TBAbrir!Desenho_sucata
                        TBGravar!Data = Format(Date, "dd/mm/yy")
                        TBGravar!Responsavel = pubUsuario
                        TBGravar!Fornecedor = IIf(IsNull(TBAbrir!Fornecedor), "", TBAbrir!Fornecedor)
                        TBGravar!Certificado = IIf(IsNull(TBAbrir!Certificado), "", TBAbrir!Certificado)
                        TBGravar!local_armaz = IIf(IsNull(TBAbrir!local_armaz), "", TBAbrir!local_armaz)
                        TBGravar!status = IIf(IsNull(TBAbrir!status), "", TBAbrir!status)
                        TBGravar!Corrida = IIf(IsNull(TBAbrir!Corrida), "", TBAbrir!Corrida)
                        If TBAbrir!Consignacao = True Then TBGravar!Consignacao = True
                        TBGravar.Update
                        TBGravar.Close
                        TBAbrir!Desenho = TBAbrir!Desenho_sucata
                        TBAbrir!idLote_sucata = Null
                        TBAbrir!Desenho_sucata = Null
                    End If
                    
'======================================================================
' Devolve a saida para o saldo da RE
'======================================================================
' Se for unidade pç para kg, busca quanto kg x peça
'======================================================================
EstoqueEntrada = 0

If TBExecucao!Un = "KG" Then

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "select * from projproduto where desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Pesoestoque = TBItem!PLiquido
End If
TBItem.Close

EstoqueEntrada = TBExecucao!estoque_real / Pesoestoque
Else
EstoqueEntrada = TBExecucao!estoque_real
End If


                    Qtd = TBAbrir!estoque_real + EstoqueEntrada
                    TBAbrir!estoque_real = Format(Qtd - Qtde, "###,##0.0000")
                Else
                    'Verifica qtde. de saída para o retalho
                    Set TBAcessos = CreateObject("adodb.recordset")
                    TBAcessos.Open "Select Saida from estoque_movimentacao where idestoque = " & TBAbrir!IDEstoque & " and Operacao = 'SAIDA_RETALHO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAcessos.EOF = False Then
                        TBAbrir!estoque_real = TBAbrir!estoque_real + TBAcessos!Saida
                    End If
                    TBAcessos.Close
                End If
                'TBAbrir!Qtde = TBAbrir!estoque_real
                TBAbrir!estoque_venda = TBAbrir!estoque_real
                TBAbrir!Valor_total = Format(TBAbrir!valor_unitario * TBAbrir!estoque_real, "###,##0.00")
                TBAbrir.Update
                
                IDlista = TBAbrir!IDEstoque
            End If
            TBAbrir.Close
            
'            Qtde = 0
'            Set TBAbrir = CreateObject("adodb.recordset")
'            TBAbrir.Open "Select * from estoque_movimentacao where idestoque = " & IDlista & " order by Data desc, IDoperacao desc, Conexao, adOpenKeyset, adLockOptimistic"
'            If TBAbrir.EOF = False Then
'                If TBAbrir!Operacao = "ENTRADA_SUCATA" Then
'                    Qtde = TBAbrir!Entrada
'                    Set TBGravar = CreateObject("adodb.recordset")
'                    TBGravar.Open "select * from estoque_movimentacao where idestoque = " & IDlista & " and operacao = 'ENTRADA_SUCATA'", Conexao, adOpenKeyset, adLockOptimistic
'                    If TBGravar.EOF = False Then
'                        TBGravar!Entrada = Format(TBGravar!Entrada + Qtde, "###,##0.00")
'                        TBGravar.Update
'                    End If
'                    TBGravar.Close
'                End If
'            End If
'            TBAbrir.Close
            
            If TBExecucao!status = "ENTRADA_SUCATA" Then TextoFiltro = "(Operacao = 'ENTRADA_SUCATA' or Operacao = 'SAIDA_SUCATA')" Else TextoFiltro = "(Operacao = 'ENTRADA_RETALHO' or Operacao = 'SAIDA_RETALHO')"
            Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & RE & " and Data = '" & TBExecucao!Data & "' and " & TextoFiltro
            Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & IDlista & " and Data = '" & TBExecucao!Data & "' and " & TextoFiltro
        Else
            Set TBAcessos = CreateObject("adodb.recordset")
            TBAcessos.Open "Select * from estoque_controle where IdEstoque = " & TBExecucao!idLote_sucata, Conexao, adOpenKeyset, adLockOptimistic
            If TBAcessos.EOF = False Then
                Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & RE & " and desenho = '" & Lista.SelectedItem.SubItems(3) & "' and lote = '" & Lista.SelectedItem.SubItems(2) & "' and (Operacao = 'ENTRADA_SUCATA' or Operacao = 'SAIDA_SUCATA')"
                Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & TBExecucao!idLote_sucata & " and documento = '" & Lista.SelectedItem.SubItems(3) & "' and lote = '" & Lista.SelectedItem.SubItems(2) & "' and (Operacao = 'ENTRADA_SUCATA' or Operacao = 'SAIDA_SUCATA')"
                
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "select * from projproduto where desenho = '" & TBExecucao!Desenho_sucata & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    TBExecucao!Desenho = TBExecucao!Desenho_sucata
                    TBExecucao!Descricao = TBItem!Descricao
                    TBExecucao!Un = TBItem!Unidade
                    TBExecucao!Classe = TBItem!Classe
                End If
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "select Compras_pedido_lista.* from compras_pedido_lista inner join compras_pedido on compras_pedido_lista.idpedido = compras_pedido.idpedido where compras_pedido.pedido = '" & TBExecucao!LOTE & "' and compras_pedido_lista.desenho = '" & TBExecucao!Desenho_sucata & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBPedido.EOF = False Then
                    ValorTotal = IIf(IsNull(TBPedido!preco_unitario), "0", TBPedido!preco_unitario)
                End If
                TBPedido.Close
                quantestoque = TBExecucao!estoque_real
                TBExecucao!Valor_total = Format(quantestoque * ValorTotal, "###,##0.00")
                TBExecucao!Desenho_sucata = ""
                TBExecucao!idLote_sucata = 0
            Else
                TBExecucao!Desenho = TBExecucao!Desenho_sucata
                Conexao.Execute "Update estoque_controle Set idLote_sucata = " & TBExecucao!IDEstoque & " where idLote_sucata = " & TBExecucao!idLote_sucata & " and idestoque <> " & RE
                TBExecucao!idLote_sucata = Null
            End If
            TBExecucao.Update
            TBAcessos.Close
        End If
        TBFIltro.Close
        
        TBExecucao.Delete
        
        USMsgBox (MsgTexto1 & " excluído com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Movimentação"
        Evento = "Excluir " & MsgTexto
        ID_documento = RE
        Documento = "Cód. interno: " & Desenho
        Documento1 = ""
        ProcGravaEvento
        '==================================
        ProcFiltrar
    End If
Else
    USMsgBox ("Favor selecionar um RE de sucata/retalho antes de excluir."), vbExclamation, "CAPRIND v5.0"
End If
TBExecucao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro
Dim Entrada As Boolean 'OK

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If

If IDlista = 0 Then
    USMsgBox ("Informe a movimentação antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If USMsgBox("Deseja realmente excluir esta movimentação?", vbYesNo, "CAPRIND v5.0") = vbNo Then
    Exit Sub
End If
'===========================================================================================
' Se a movimentação for saida por requisição
'===========================================================================================
If Status_movimentacao = "SAIDA_REQUISICAO_PARCIAL" Or Operacao = "SAIDA_REQUISICAO" Then
'Verif. número do documento
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Estoque_movimentacao where idestoque = " & RE & " and (operacao = 'SAIDA_REQUISICAO_PARCIAL' or operacao = 'SAIDA_REQUISICAO')", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Documento_Ordem = IIf(IsNull(TBAbrir!Documento), "", TBAbrir!Documento)
End If
TBAbrir.Close
End If
'============================================================================================
' "ENTRADA_ORDEM" "ENTRADA_ORDEM_PARCIAL" "ENTRADA_INVENTÁRIO" "ENTRADA_DEVOLUÇÃO"
'============================================================================================
           
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from estoque_movimentacao where idoperacao = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
IDEstoque = TBAbrir!IDEstoque

valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)

'Define valor na variável
'============================================================================================
' "ENTRADA_ORDEM" "ENTRADA_ORDEM_PARCIAL" "ENTRADA_INVENTÁRIO" "ENTRADA_DEVOLUÇÃO"
'============================================================================================
If TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Or TBAbrir!Operacao = "ENTRADA_INVENTÁRIO" Or TBAbrir!Operacao = "ENTRADA_DEVOLUÇÃO" Then
QuantSolicitado = IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada)
    
'Exclui o empenho no produto em estoque para o pedido
Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBAbrir!IDEstoque
    
 'Atualiza quantidade de entrada no empenho da ordem
'============================================================================================
' "ENTRADA_ORDEM" "ENTRADA_ORDEM_PARCIAL"
'============================================================================================
   If TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Then
       qtdeliberada = QuantSolicitado
       Set TBFI = CreateObject("adodb.recordset")
       TBFI.Open "Select PP.Qtde_entrada from (producao_pedidos PP INNER JOIN vendas_carteira VC ON PP.IDcarteira = VC.Codigo) INNER JOIN Producao P ON P.Ordem = PP.Ordem where P.Ordem = " & TBAbrir!LOTE & " and P.Desenho = '" & TBAbrir!Desenho & "' and ISNULL(Qtde_entrada , 0) > 0 order by VC.Prazofinal desc", Conexao, adOpenKeyset, adLockOptimistic
       If TBFI.EOF = False Then
           Do While TBFI.EOF = False
               If qtdeliberada >= 0 Then
                   If qtdeliberada >= TBFI!Qtde_entrada Then
                       TBFI!Qtde_entrada = 0
                       qtdeliberada = 0
                   Else
                       qtdeliberada = TBFI!Qtde_entrada - qtdeliberada
                       TBFI!Qtde_entrada = qtdeliberada
                   End If
                   TBFI.Update
               End If
               TBFI.MoveNext
           Loop
       End If
       TBFI.Close
   End If
'==============================================================================================
'Else
'    EstoqueSaida = IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
End If
'==============================================================================================
' Se for saida de nota fiscal
'==============================================================================================
If IsNumeric(Documento) = True And (Status_movimentacao = "SAIDA_NOTA" Or Status_movimentacao = "SAIDA_NOTA_PARCIAL") Then
    'Atualiza qtde. expedida
    Qtd = EstoqueSaida
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select PP.Qtdeexpedida, PP.Dataexpedicao, NFPP.* from vendas_carteira PP INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.Codigo and NFPP.Codinterno = PP.Desenho where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " order by PP.PrazoFinal", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        Do While TBGravar.EOF = False
            If Qtd >= TBGravar!qtdeexpedida Then qt = TBGravar!qtdeexpedida Else qt = Qtd
            TBGravar!qtdeexpedida = TBGravar!qtdeexpedida - qt
            Qtd = Qtd - qt
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Data from Estoque_movimentacao where Idoperacao <> " & IDlista & " and ID_prod_NF = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF) & " and (Operacao = 'SAIDA_NOTA' or Operacao = 'SAIDA_NOTA_PARCIAL') order by Data desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                TBGravar!dataexpedicao = TBFI!Data
            Else
                TBGravar!dataexpedicao = Null
            End If
            TBFI.Close
            TBGravar.Update
            
            'Desvincula pedido da ordem para estoque
            If IsNumeric(TBAbrir!LOTE) = True Then
                Set TBAliquota = CreateObject("adodb.recordset")
                TBAliquota.Open "Select * from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.EOF = False Then
                    TBAliquota!Qtde_empenho = TBAliquota!Qtde_empenho - qt
                    TBAliquota!Qtde_entrada = TBAliquota!Qtde_empenho
                    TBAliquota.Update
                    
                    If TBAliquota!Qtde_empenho <= 0 Then Conexao.Execute "DELETE from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'"
                End If
                TBAliquota.Close
            End If
            
            Do While qt > 0
                'Atualiza qtde. de saída no empenho
                Set TBAliquota = CreateObject("adodb.recordset")
                TBAliquota.Open "Select EE.Qtde_saida from Estoque_Controle_Empenho_Vendas EE INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON EE.ID_carteira = NFPP.ID_carteira where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " and EE.ID_estoque = " & TBAbrir!IDEstoque & " and EE.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.EOF = False Then
                    If TBAliquota!Qtde_saida >= qt Then
                        TBAliquota!Qtde_saida = TBAliquota!Qtde_saida - qt
                        qt = 0
                    Else
                        qt = qt - TBAliquota!Qtde_saida
                        TBAliquota!Qtde_saida = 0
                    End If
                    TBAliquota.Update
                Else
                    GoTo Prosseguir
                End If
                TBAliquota.Close
            Loop
Prosseguir:
            If Qtd <= 0 Then GoTo Prosseguir1
            TBGravar.MoveNext
        Loop
    End If
End If

Prosseguir1:
'====================================================================================
' Se for saida por requisição
'====================================================================================
                If Status_movimentacao = "SAIDA_REQUISICAO" Or Status_movimentacao = "SAIDA_REQUISICAO_PARCIAL" Then
                    Set TBMateriaprima = CreateObject("adodb.recordset")
                    TBMateriaprima.Open "Select * from Requisicao_materiais where requisicao = '" & TBAbrir!Documento & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBMateriaprima.EOF = False Then
                        
                        NovoValor = Replace(EstoqueSaida, ",", ".")
                        
                        Set TBMaterial = CreateObject("adodb.recordset")
                        TBMaterial.Open "Select RML.*, P.ID_PC from Requisicao_materiais_lista RML INNER JOIN Projproduto P ON P.Desenho = RML.Desenho where RML.idrequisicao = " & TBMateriaprima!ID & " and RML.desenho = '" & Desenho & "' and RML.quant_saida >= " & NovoValor & " and (RML.Status = 'RETIRADO' or RML.Status = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                        If TBMaterial.EOF = False Then
                            Qtde = IIf(IsNull(TBMaterial!quant_saida), 0, TBMaterial!quant_saida)
                            Qtde = Qtde - EstoqueSaida
                            TBMaterial!quant_saida = Format(Qtde, "###,##0.0000")
                            If Qtde > 0 Then TBMaterial!status = "PARCIAL" Else TBMaterial!status = "REQUISIT."
                            TBMaterial.Update
                            
                            If IsNull(TBMaterial!Ordem) = False And TBMaterial!Ordem <> 0 And IsNull(TBMaterial!ID_PC) = False And TBMaterial!ID_PC <> 0 Then
                                Set TBOrdem = CreateObject("adodb.recordset")
                                TBOrdem.Open "Select * from Producao_outras_despesas where Ordem = " & TBMaterial!Ordem & " and ID_PC = " & TBMaterial!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                                If TBOrdem.EOF = False Then
                                    If TBOrdem!valor - (Format(valor * EstoqueSaida, "###,##0.00")) <= 0 Then
                                        TBOrdem.Delete
                                    Else
                                        TBOrdem!valor = TBOrdem!valor - (Format(valor * EstoqueSaida, "###,##0.00"))
                                        TBOrdem.Update
                                    End If
                                End If
                                Valor1 = 0
                                Set TBOrdem = CreateObject("adodb.recordset")
                                TBOrdem.Open "Select Sum(Valor) as Valor1 from Producao_outras_despesas where Ordem = " & TBMaterial!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                                If TBOrdem.EOF = False Then
                                    Valor1 = IIf(IsNull(TBOrdem!Valor1), 0, TBOrdem!Valor1)
                                End If
                                TBOrdem.Close
                                NovoValor = Replace(Valor1, ",", ".")
                                Conexao.Execute "Update Producao Set CTOutras = " & NovoValor & " where Ordem = " & TBMaterial!Ordem
                                
                            End If
                        End If
                        TBMaterial.Close
                        
                        ProcAtualizaStatus_RM
                    End If
                End If
'=================================================================================
 ' Se for entrada de local de armazenamento ou saida de local de armazenamento
 '================================================================================
                If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Or Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                    If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Then
                        'Achando a movimentação de saida
                        Set TBNivel14 = CreateObject("adodb.recordset")
                        TBNivel14.Open "Select IdEstoque, idoperacao from Estoque_movimentacao where idoperacao = " & TBAbrir!IdTrocaLocal, Conexao, adOpenKeyset, adLockOptimistic
                        If TBNivel14.EOF = False Then
                            'achando o RE de saida
                            Set TBMateriaprima = CreateObject("adodb.recordset")
                            TBMateriaprima.Open "Select estoque_real_PC, estoque_real, IDestoque from Estoque_controle where IdEstoque = " & TBNivel14!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                            If TBMateriaprima.EOF = False Then
                                procVoltarEmpenhoLocal TBAbrir!IDEstoque, TBMateriaprima!IDEstoque
                                TBMateriaprima!estoque_real = TBMateriaprima!estoque_real + TBAbrir!Entrada
                                TBMateriaprima!estoque_real_PC = TBMateriaprima!estoque_real_PC + TBAbrir!Entrada_PC
                                TBMateriaprima.Update
                                TBNivel14.Delete 'Exclui movimentação de saida
                            End If
                            TBMateriaprima.Close
                        End If
                        TBNivel14.Close
                    Else
                        'Achando a movimentação de entrada
                        Set TBNivel14 = CreateObject("adodb.recordset")
                        TBNivel14.Open "Select IdEstoque, idoperacao from Estoque_movimentacao where IdTrocaLocal = " & TBAbrir!IDoperacao, Conexao, adOpenKeyset, adLockOptimistic
                        If TBNivel14.EOF = False Then
                            procVoltarEmpenhoLocal TBNivel14!IDEstoque, TBAbrir!IDEstoque
                            Conexao.Execute "DELETE FROM Estoque_controle WHERE IdEstoque = " & TBNivel14!IDEstoque 'Exclui estoque controle entrada
                            TBNivel14.Delete 'Exclui movimentação entrada
                        End If
                        TBNivel14.Close
                    End If
                End If
            End If
            TBAbrir.Close
'Apaga a movimentação do estoque
Conexao.Execute "DELETE from estoque_movimentacao where idoperacao = " & IDlista
'================================================================================
            'Corrige retirada na tabela producaomaterial
            quantidade = 0
            QuantidadePC = 0
'================================================================================
'Verifica se é consignado
'================================================================================
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from Estoque_controle where IDEstoque = " & RE & " and consignacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                TextoFiltro = "IDEstoque = " & Lista.SelectedItem
            Else
                TextoFiltro = "Desenho = '" & Desenho & "'"
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Saida) as Quantidade, Sum(ISNULL(Saida_PC, 0)) as QuantidadePC from estoque_movimentacao where " & TextoFiltro & " and documento = '" & Documento & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                quantidade = IIf(IsNull(TBAbrir!quantidade), 0, TBAbrir!quantidade)
                QuantidadePC = IIf(IsNull(TBAbrir!QuantidadePC), 0, TBAbrir!QuantidadePC)
            End If
            TBAbrir.Close
            
            If IsNumeric(Documento) = True Then
                Set TBproducao = CreateObject("adodb.recordset")
                TBproducao.Open "Select * from producaomaterial where codigo = '" & Desenho & "' and Ordem = " & Int(Documento), Conexao, adOpenKeyset, adLockOptimistic
                If TBproducao.EOF = False Then
                    If quantidade = 0 Then
                        TBproducao!Saida = "NÃO"
                    ElseIf quantidade >= TBproducao!Requisitado Or QuantidadePC >= TBproducao!Total_pc Then
                        TBproducao!Saida = "SIM"
                    Else
                        TBproducao!Saida = "PARCIAL"
                    End If
                    
                    TBproducao!Valor_saida_estoque = Format(IIf(IsNull(TBproducao!Valor_saida_estoque), 0, TBproducao!Valor_saida_estoque) - (valor * EstoqueSaida), "###,##0.00")
                    TBproducao.Update
                    
                    'Atualiza qtde. de saída do empenho da ordem
                    QuantEmpenho = 0
                    QuantEmpenhoPC = 0
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select Sum(Saida) as QuantEmpenho, Sum(ISNULL(Saida_PC, 0)) as QuantEmpenhoPC from estoque_movimentacao where IDestoque = " & Lista.SelectedItem & " and oe = '" & TBproducao!Ordem & "' and desenho = '" & TBproducao!CODIGO & "' and documento = '" & TBproducao!Ordem & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, Format(TBAbrir!QuantEmpenho, "###,##0.0000"))
                        QuantEmpenhoPC = IIf(IsNull(TBAbrir!QuantEmpenhoPC), 0, Format(TBAbrir!QuantEmpenhoPC, "###,##0.0000"))
                    End If
                    NovoValor = Replace(QuantEmpenho, ",", ".")
                    NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
                    Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida = " & NovoValor & " where IDestoque = " & Lista.SelectedItem & " and Ordem = " & TBproducao!Ordem & " and Codinterno = '" & TBproducao!CODIGO & "'"
                    Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida_PC = " & NovoValor1 & " where IDestoque = " & Lista.SelectedItem & " and Ordem = " & TBproducao!Ordem & " and Codinterno = '" & TBproducao!CODIGO & "' and Quantidade_PC IS NOT NULL and Quantidade_PC > 0"
                End If
                TBproducao.Close
            End If
            
            Permitido1 = True
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from projproduto where Desenho = '" & Desenho & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = True Then
                Permitido1 = False
            End If
            TBProduto.Close
            
            'Centro de custo
            Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & IDlista
            
            'Corrige estoque real
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_controle where idestoque = " & RE, Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                
                '==================================
                Modulo = "Estoque/Movimentação"
                Evento = "Excluir movimentação"
                ID_documento = IDlista
                Documento = "Cód. interno: " & TBEstoque!Desenho & " - Nº lote: " & TBEstoque!LOTE & " - Nº corrida: " & TBEstoque!Corrida & " - Nº certificado: " & TBEstoque!Certificado & " - Local armaz.: " & TBEstoque!local_armaz
                Documento1 = "Operação: " & Status_movimentacao & " - Data: " & Data & " - Entrada: " & EstoqueEntrada & " - Saída: " & EstoqueSaida & " - Documento: " & Documento
                ProcGravaEvento
                '==================================
                
                If Permitido1 = True Then
                    Qtde = Valor1
                    TBEstoque!estoque_real = Format((TBEstoque!estoque_real + EstoqueSaida) - EstoqueEntrada, "###,##0.0000")
                    TBEstoque!Qtde = Format((TBEstoque!Qtde + EstoqueSaida) - EstoqueEntrada, "###,##0.0000")
                    TBEstoque!estoque_real_PC = Format((TBEstoque!estoque_real_PC + EstoqueSaida) - EstoqueEntrada, "###,##0.0000")
                Else
                    TBEstoque!estoque_real = 0
                    TBEstoque!estoque_real_PC = 0
                    Qtde = 0
                End If
                                        
                'Atualiza valor do material no estoque
                'Estoque_controle
                TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * Qtde, "###,##0.00")
                
                TBEstoque.Update
                 Set TBMaterial = CreateObject("adodb.recordset")
                 TBMaterial.Open "Select * from Estoque_movimentacao where IDEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                 If TBMaterial.EOF = True Then TBEstoque.Delete
                 TBMaterial.Close
            End If
            TBEstoque.Close
            
            If Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Then
                Conexao.Execute "Update estoque_controle Set Status = 'ENTRADA_ORDEM_PARCIAL' where lote = '" & LOTE & "' and status = 'ENTRADA_ORDEM' and ID_empresa = " & IDempresa
                Conexao.Execute "Update estoque_movimentacao Set Operacao = 'ENTRADA_ORDEM_PARCIAL' where lote = '" & LOTE & "' and Operacao = 'ENTRADA_ORDEM'"
            End If
            
            If IsNumeric(Documento) = True And (Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Or Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL") Then
                'Atualiza qtde. produzida
                Qtde = 0
                qtdeliberada = 0
                Qtd = 0
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from producao where Ordem = " & Documento & " and Controlado_estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Then Qtd = EstoqueEntrada
                    If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then Qtd = EstoqueSaida
                    Set TBExecucao = CreateObject("adodb.recordset")
                    TBExecucao.Open "select * from estoque_controle where Lote = '" & Lista.SelectedItem.SubItems(3) & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao.EOF = False Then
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select Sum(Entrada) as Qtde, Sum(Saida) as qtdeliberada from estoque_movimentacao where idestoque = " & TBExecucao!IDEstoque & " and documento = '" & Documento & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBCorretiva.EOF = False Then
                            Qtde = IIf(IsNull(TBCorretiva!Qtde), 0, TBCorretiva!Qtde) + IIf(IsNull(TBCorretiva!qtdeliberada), 0, TBCorretiva!qtdeliberada)
                        End If
                        TBCorretiva.Close
                    End If
                    TBExecucao.Close
                    If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Then
                        If Qtde < Qtd Then ProcOrdem
                    End If
                    If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
                        If Qtde = 0 Then ProcOrdem
                    End If
                End If
                TBCiclo.Close
                
                'Custo material
                If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then ProcAtualizaCTMaterialOrdem Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Int(Documento)
            End If
        'End If
'     Next InitFor
'End With

USMsgBox ("Movimentação excluída com sucesso."), vbInformation, "CAPRIND v5.0"
'ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
'If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
'    Lista.SelectedItem = Lista.ListItems(CodigoLista)
'    ProcCarregaDadosLote
'    Lista.SetFocus
'Else
'2:
'    ProcLimpaCamposTotais
'End If
ProcFiltrar

Exit Sub
tratar_erro:
   ' If Err.Number = "35600" Then GoTo 2
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLocalArmazenamento()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "alterar"
If txtlocalização = "" Then
    NomeCampo = "o lote na lista"
    ProcVerificaAcao
    Exit Sub
End If
If Lista.SelectedItem.SubItems(8) = "" Then
    NomeCampo = "o local de armazenamento"
    ProcVerificaAcao
    Exit Sub
End If
frmEstoque_item_localarmaz.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSucata()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If RE = 0 Then
    USMsgBox ("Infome o RE antes de gerar sucata/retrabalho."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select IDestoque from estoque_controle where idestoque = " & RE & " and idlote_sucata <> 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido gerar sucata/retrabalho, pois o mesmo já é uma sucata."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close
frmEstoque_sucata.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If RE = 0 Then
    USMsgBox ("Informe o RE antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmEstoque_item_bloq.Show 1
ProcFiltrar
'ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEstruturadoitem()
On Error GoTo tratar_erro

If Desenho = "" Then
    USMsgBox ("Informe o código interno antes de abrir a estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    With frmproj_conjunto
        .Show
        .ProcLimpaCampos
        .Txt_cod_produto = TBItem!Codproduto
        .txtdesenhoproduto.Text = TBItem!Desenho
        .txtDescricaoProduto.Text = TBItem!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then .txtRef.Text = TBAbrir("n_referencia")
        TBAbrir.Close
        .ProcAtualizalista (1)
    End With
Else
    USMsgBox ("Não foi encontrado nenhum registro para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Cmd_salvar_ref_serie_Click()
On Error GoTo tratar_erro

Acao = "salvar"
If txtlocalização = "" Then
    NomeCampo = "o RE"
    ProcVerificaAcao
    txtlocalização.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente alterar o código de referência e número de série deste RE?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    'Verifica se o código de referencia está vinculado a outro produto
    'If Txt_cod_ref <> "" Then If FunVerifiCodRefUtilizado(Lista.SelectedItem.ListSubItems(3), Txt_cod_ref) = True Then Exit Sub
    
    Conexao.Execute "Update estoque_controle Set Ref = '" & Txt_cod_ref & "', Numero_serie = '" & Txt_n_serie & "' where IDestoque = " & txtlocalização
    Conexao.Execute "Update Estoque_fisico Set Cod_ref = '" & Txt_cod_ref & "', Numero_serie = '" & Txt_n_serie & "' where IDestoque = " & txtlocalização
    Conexao.Execute "Update EF set EF.Cod_ref = '" & Txt_cod_ref & "', EF.Numero_serie = '" & Txt_n_serie & "' from Estoque_fisico EF INNER JOIN Estoque_movimentacao EM ON EM.ID_inventario = EF.ID where EM.IDestoque = " & txtlocalização & " and EM.ID_inventario IS NOT NULL and EM.ID_inventario <> 0"
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Estoque/Movimentação"
    Evento = "Alterar código de referência e número de série"
    ID_documento = txtlocalização
    Documento = "Cód. interno: " & Lista.SelectedItem.SubItems(3)
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Lista_Movimentacao.ListItems.Clear
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
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
If TBLISTA_Estoque_Movimentacao.AbsolutePage <> 2 Then
    If TBLISTA_Estoque_Movimentacao.AbsolutePage = -3 Then
        ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.PageCount - 1)
    Else
        TBLISTA_Estoque_Movimentacao.AbsolutePage = TBLISTA_Estoque_Movimentacao.AbsolutePage - 2
        ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)
    End If
Else
    ProcExibePaginaGrid (1)
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
    TBLISTA_Estoque_Movimentacao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Estoque_Movimentacao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Movimentacao.AbsolutePage = 1
ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_Movimentacao.AbsolutePage <> -3 Then
    If TBLISTA_Estoque_Movimentacao.AbsolutePage = 1 Then
        ProcExibePaginaGrid (2)
    Else
        ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)
    End If
Else
    ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Movimentacao.AbsolutePage = TBLISTA_Estoque_Movimentacao.PageCount
ProcExibePaginaGrid (TBLISTA_Estoque_Movimentacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcEstruturadoitem
    Case vbKeyF8: ProcSucata
    Case vbKeyF9: ProcExcluir_sucata
    Case vbKeyF10: ProcLocalArmazenamento
    Case vbKeyF11: ProcAlterar_valor
    Case vbKeyF12: ProcCC
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCC()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista_Movimentacao.ListItems.Count = 0 Then
    USMsgBox ("Informe a movimentação antes de visualizar o(s) centro(s) de custo."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Status_movimentacao = Lista_Movimentacao.SelectedItem.ListSubItems(2)
If Status_movimentacao = "SAIDA_REQUISICAO" Or Status_movimentacao = "SAIDA_REQUISICAO_PARCIAL" Or Status_movimentacao = "ENTRADA_INVENTÁRIO" Or Status_movimentacao = "SAIDA_INVENTÁRIO" Or Status_movimentacao = "ENTRADA_DEVOLUÇÃO" Then
    Estoque_recebimento = False
Else
    Estoque_recebimento = True
End If
frmEstoque_item_lista_CC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
Formulario = "Estoque/Movimentação"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False

'ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "Produtos/Serviços", "T", True

If Permitido = False Then
cmbfiltrarpor = "Código interno"
End If

ProcRemoveObjetosResize Me
ProcAjustaGridProduto
ProcAjustaGridRE
ProcAjustaGridMV

'De.Value = "31/12/" & Year(Date) - 10
Ate.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridMV()
On Error GoTo tratar_erro

    GridMv.AllowUserPaste = cellTextOnly
    GridMv.AllowUserResizing = False
    GridMv.ExtendLastCol = True
    GridMv.BoldFixedCell = False
    GridMv.DisplayDateTimeMask = True
    GridMv.DisplayFocusRect = False
    GridMv.SelectionMode = cellSelectionByRow

    GridMv.DrawMode = cellOwnerDraw
    
    GridMv.Appearance = Flat
    GridMv.ScrollBarStyle = Flat
    GridMv.FixedRowColStyle = Flat
    GridMv.Cell(0, 1).Text = "Item"
    GridMv.Cell(0, 2).Text = "Operação"
    GridMv.Cell(0, 3).Text = "Data"
    GridMv.Cell(0, 4).Text = "Entrada"
    GridMv.Cell(0, 5).Text = "Saida"
    GridMv.Cell(0, 6).Text = "Documento"
    GridMv.Cell(0, 7).Text = "Responsável"
    GridMv.Cell(0, 8).Text = "Requisitante"
    GridMv.Cell(0, 9).Text = "Destino"
  '  GridMv.Cell(0, 10).Text = "PC\PI"
  '  GridMv.Cell(0, 11).Text = "Cliente\Fornecedor"
   ' GridMv.Cell(0, 12).Text = "Observações"
        
    GridMv.Column(1).CellType = cellTextBox
    GridMv.Column(1).Alignment = cellCenterCenter
    
    GridMv.Column(2).CellType = cellTextBox
    GridMv.Column(2).Alignment = cellCenterCenter
    
    GridMv.Column(3).CellType = cellDate
    GridMv.Column(3).Alignment = cellCenterCenter
    GridMv.Column(3).FormatString = "DD/MM/YYYY"
    
    GridMv.Column(4).CellType = cellTextBox
    GridMv.Column(4).Alignment = cellRightCenter
    
    GridMv.Column(5).CellType = cellTextBox
    GridMv.Column(5).Alignment = cellRightCenter 'cellHyperLink
    
    GridMv.Column(6).CellType = cellTextBox 'cellButton
    GridMv.Column(6).Alignment = cellCenterCenter 'cellHyperLink
    
    GridMv.Column(7).CellType = cellTextBox 'cellHyperLink
    GridMv.Column(7).Alignment = cellCenterCenter 'cellHyperLink
    
    GridMv.Column(8).CellType = cellTextBox 'cellHyperLink
    GridMv.Column(8).Alignment = cellCenterCenter 'cellHyperLink
    
    GridMv.Column(9).CellType = cellTextBox 'cellHyperLink
    GridMv.Column(9).Alignment = cellCenterCenter 'cellHyperLink
    
    GridMv.Column(10).CellType = cellTextBox 'cellHyperLink
    GridMv.Column(10).Alignment = cellCenterCenter 'cellHyperLink
   
 
    GridMv.Column(0).Width = 10
    GridMv.Column(1).Width = 30
    GridMv.Column(2).Width = 350
    GridMv.Column(3).Width = 90
    GridMv.Column(4).Width = 50
    GridMv.Column(5).Width = 50
    GridMv.Column(6).Width = 100
    GridMv.Column(7).Width = 220
    GridMv.Column(8).Width = 220
    GridMv.Column(9).Width = 0
    GridMv.Column(10).Width = 0
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridRE()
On Error GoTo tratar_erro

    GridRE.AllowUserPaste = cellTextOnly
    GridRE.AllowUserResizing = False
    GridRE.ExtendLastCol = True
    GridRE.BoldFixedCell = False
    GridRE.DisplayDateTimeMask = True
    GridRE.DisplayFocusRect = False
    GridRE.SelectionMode = cellSelectionByRow

    GridRE.DrawMode = cellOwnerDraw
    GridRE.Cols = 13
    
    GridRE.Appearance = Flat
    GridRE.ScrollBarStyle = Flat
    GridRE.FixedRowColStyle = Flat
    GridRE.Cell(0, 1).Text = "Item"
    GridRE.Cell(0, 2).Text = "Data"
    GridRE.Cell(0, 3).Text = "RE"
    GridRE.Cell(0, 4).Text = "Lote"
    GridRE.Cell(0, 5).Text = "Status"
    GridRE.Cell(0, 6).Text = "Codigo"
    GridRE.Cell(0, 7).Text = "Descrição"
    GridRE.Cell(0, 8).Text = "UN"
    GridRE.Cell(0, 9).Text = "Entrada"
    GridRE.Cell(0, 10).Text = "Saida"
    GridRE.Cell(0, 11).Text = "Empenho"
    GridRE.Cell(0, 12).Text = "Saldo"
        
    GridRE.Column(1).CellType = cellTextBox 'Item
    GridRE.Column(1).Alignment = cellCenterCenter
    
    GridRE.Column(2).CellType = cellDate 'Data
    GridRE.Column(2).Alignment = cellCenterCenter
    GridRE.Column(2).FormatString = "DD/MM/YYYY"
    
    GridRE.Column(3).CellType = cellTextBox 'RE
    GridRE.Column(3).Alignment = cellCenterCenter
    
    GridRE.Column(4).CellType = cellTextBox 'Lote
    GridRE.Column(4).Alignment = cellCenterCenter
    
    GridRE.Column(5).CellType = cellTextBox 'Status
    GridRE.Column(5).Alignment = cellCenterCenter 'cellHyperLink
    
    GridRE.Column(6).CellType = cellTextBox 'Codigo item'cellButton
    GridRE.Column(6).Alignment = cellCenterCenter 'cellHyperLink
    
    GridRE.Column(7).CellType = cellTextBox 'Descricao item 'cellHyperLink
    GridRE.Column(7).Alignment = cellLeftCenter 'cellHyperLink
    
    GridRE.Column(8).CellType = cellTextBox ' Unidade'cellHyperLink
    GridRE.Column(8).Alignment = cellCenterCenter 'cellHyperLink
    
    GridRE.Column(9).CellType = cellTextBox 'Entrada'cellHyperLink
    GridRE.Column(9).Alignment = cellRightCenter 'cellHyperLink
    
    GridRE.Column(10).CellType = cellTextBox 'Saida'cellHyperLink
    GridRE.Column(10).Alignment = cellRightCenter 'cellHyperLink
    
    GridRE.Column(11).CellType = cellTextBox 'Empenho'cellHyperLink
    GridRE.Column(11).Alignment = cellRightCenter 'cellHyperLink
    
    GridRE.Column(12).CellType = cellTextBox 'Saldo'cellHyperLink
    GridRE.Column(12).Alignment = cellRightCenter 'cellHyperLink
 
    GridRE.Column(0).Width = 10
    GridRE.Column(1).Width = 30 ' Item
    GridRE.Column(2).Width = 70 ' Data
    GridRE.Column(3).Width = 40 ' RE
    GridRE.Column(4).Width = 70 'Lote
    GridRE.Column(5).Width = 150 ' Status
    GridRE.Column(6).Width = 80 'Codigo
    GridRE.Column(7).Width = 350 'Descricao
    GridRE.Column(8).Width = 30 ' Unidade
    GridRE.Column(9).Width = 50 ' Entrada
    GridRE.Column(10).Width = 50 ' Saida
    GridRE.Column(11).Width = 50 ' Empenho
    GridRE.Column(12).Width = 50 ' Saldo
    
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridProduto()
On Error GoTo tratar_erro

    GridProdutos.AllowUserPaste = cellTextOnly
    GridProdutos.AllowUserResizing = False
    GridProdutos.ExtendLastCol = True
    GridProdutos.BoldFixedCell = False
    GridProdutos.DisplayDateTimeMask = True
    GridProdutos.DisplayFocusRect = False
    GridProdutos.SelectionMode = cellSelectionByRow

    GridProdutos.DrawMode = cellOwnerDraw
    GridProdutos.Cols = 6 '
    
    GridProdutos.Appearance = Flat
    GridProdutos.ScrollBarStyle = Flat
    GridProdutos.FixedRowColStyle = Flat
    GridProdutos.Cell(0, 1).Text = "Item"
    GridProdutos.Cell(0, 2).Text = "Codigo"
    GridProdutos.Cell(0, 3).Text = "Descrição"
    GridProdutos.Cell(0, 4).Text = "UN"
    GridProdutos.Cell(0, 5).Text = "Saldo"
        
    GridProdutos.Column(1).CellType = cellTextBox 'Item
    GridProdutos.Column(1).Alignment = cellCenterCenter
        
    GridProdutos.Column(2).CellType = cellTextBox 'Codigo
    GridProdutos.Column(2).Alignment = cellCenterCenter
    
    GridProdutos.Column(3).CellType = cellTextBox 'Descricao
    GridProdutos.Column(3).Alignment = cellLeftCenter
    
    GridProdutos.Column(4).CellType = cellTextBox 'un
    GridProdutos.Column(4).Alignment = cellCenterCenter 'cellHyperLink
        
    GridProdutos.Column(5).CellType = cellTextBox ' Empenho
    GridProdutos.Column(5).Alignment = cellRightCenter 'cellHyperLink
        
 
    GridProdutos.Column(0).Width = 10
    GridProdutos.Column(1).Width = 50 ' Item
    GridProdutos.Column(2).Width = 140 'Codigo
    GridProdutos.Column(3).Width = 400 'Descricao
    GridProdutos.Column(4).Width = 30 ' Unidade
    GridProdutos.Column(5).Width = 80 ' Saldo
    
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Estoque/Movimentação"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

 With GridProdutos.PageSetup
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Tahoma"
        .HeaderFont.size = 20
        .PrintCellBorders = True
        .PrintTitleColumns = True
        .PrintFixedColumn = True
        .PrintFixedRow = True
        .PrintGridlines = True
        .Header = "Estoque Movimentação Resumido"
        
        .PaperSize = cellPaperA4
        .LeftMargin = 1
        .TopMargin = 3
        .RightMargin = 1
        .BottomMargin = 1
        .HeaderMargin = 1
        .FooterMargin = 1
  
        '.Header = "FlexCell Studio"
        '.HeaderAlignment = cellLeft
        '.HeaderFont.Name = "Courier New"
        '.HeaderFont.Size = 16
        '.HeaderFont.Bold = True
        
        .Footer = "Pag &P de &N"
        .FooterAlignment = cellRight
        .FooterFont.Name = "Tahoma"
        .FooterFont.size = 8
    End With

GridProdutos.PrintPreview 100


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362E" Then frmestoque_item_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro
Dim EntradaPC As Double, Saida As Double, SaidaPC As Double, Total As Double, TotalPC As Double 'OK

Entrada = 0
EntradaPC = 0
Saida = 0
SaidaPC = 0
Total = 0
valor = 0
Valor1 = 0
Valor2 = 0

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmestoque_item_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza estoque controle de movimentação sem estoque controle
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_movimentacao order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_controle where IDestoque = " & TBEstoque!IDEstoque & " or Lote = '" & TBEstoque!LOTE & "' and Desenho = '" & TBEstoque!Desenho & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        TBAbrir.AddNew
                        TBAbrir!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                        TBAbrir!LOTE = TBEstoque!LOTE
                        TBAbrir!Desenho = TBEstoque!Desenho
                        TBAbrir!Descricao = TBEstoque!Descricao
                        TBAbrir!estoque_venda = 0
                        TBAbrir!estoque_real = 0
                        TBAbrir!estoque_real_PC = 0
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from projproduto where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            TBAbrir!Un = TBFI!Unidade
                            TBAbrir!Classe = TBFI!Classe
                        End If
                        TBFI.Close
                        
                        TBAbrir!Data = Date
                        TBAbrir!Responsavel = pubUsuario
                        
                        Set TBPedido = CreateObject("adodb.recordset")
                        TBPedido.Open "Select * from compras_pedido where Pedido = '" & TBEstoque!LOTE & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBPedido.EOF = False Then
                            
                            TBAbrir!status = "ENTRADA_NOTA_FISCAL"
                            
                            Set TBCompras = CreateObject("adodb.recordset")
                            TBCompras.Open "Select * from Estoque_controle_recebimento where ID = " & TBEstoque!IDEstoque_recebimento, Conexao, adOpenKeyset, adLockOptimistic
                            If TBCompras.EOF = False Then
                                TBAbrir!Certificado = TBCompras!Certificado
                                TBAbrir!Corrida = TBCompras!Corrida
                                TBAbrir!local_armaz = TBCompras!local_armaz
                            Else
                                GoTo 1:
                            End If
                            TBCompras.Close
                        Else
                            If IsNumeric(TBEstoque!LOTE) = True Then
                                Set TBproducao = CreateObject("adodb.recordset")
                                TBproducao.Open "Select * from Producao where Ordem = " & TBEstoque!LOTE, Conexao, adOpenKeyset, adLockOptimistic
                                If TBproducao.EOF = False Then
                                    TBAbrir!status = "ENTRADA_ORDEM"
                                Else
                                    TBAbrir!status = "ENTRADA_INVENTÁRIO"
                                End If
                                TBproducao.Close
                            Else
                                TBAbrir!status = "ENTRADA_INVENTÁRIO"
                            End If
1:
                            TBAbrir!Certificado = 0
                            TBAbrir!Corrida = 0
                            
                            Set TBFIltro = CreateObject("adodb.recordset")
                            TBFIltro.Open "Select * from estoque_controle where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFIltro.EOF = False Then
                                TBAbrir!local_armaz = TBFIltro!local_armaz
                            Else
                                TBAbrir!local_armaz = "N/C"
                            End If
                            TBFIltro.Close
                        End If
                        TBPedido.Close
                        
                        TBAbrir.Update
                        TBEstoque!IDEstoque = TBAbrir!IDEstoque
                        TBEstoque.Update
                        
                    'Else
                        'If TBEstoque!IdEstoque <> TBAbrir!IdEstoque Then
                            'Conexao.Execute "Update Estoque_movimentacao Set idestoque = " & TBAbrir!IdEstoque & " where IDestoque = " & TBEstoque!IdEstoque
                        'End If
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_Controle order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_Controle Where Desenho = '" & TBEstoque!Desenho & "' and Lote = '" & TBEstoque!LOTE & "' and Corrida = '" & TBEstoque!Corrida & "' and Certificado = '" & TBEstoque!Certificado & "' and local_armaz = '" & TBEstoque!local_armaz & "' and Idestoque <> " & TBEstoque!IDEstoque & " and ID_empresa = " & TBEstoque!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Do While TBAbrir.EOF = False
                            TBEstoque!estoque_venda = Format(TBEstoque!estoque_venda + TBAbrir!estoque_venda, "###,##0.0000000000")
                            TBEstoque!estoque_real = Format(TBEstoque!estoque_real + TBAbrir!estoque_real, "###,##0.0000000000")
                            TBEstoque!estoque_real_PC = Format(TBEstoque!estoque_real_PC + TBAbrir!estoque_real_PC, "###,##0.0000000000")
                            TBEstoque!Qtde = Format(TBEstoque!Qtde + TBAbrir!Qtde, "###,##0.0000000000")
                            TBEstoque.Update
                            
                            Conexao.Execute "Update Estoque_movimentacao Set IdEstoque = " & TBEstoque!IDEstoque & " where Idestoque = " & TBAbrir!IDEstoque
                            Conexao.Execute "DELETE from Estoque_Controle where Idestoque = " & TBAbrir!IDEstoque
                            TBAbrir.MoveNext
                        Loop
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            
            'Deleta movimentação sem estoque_controle
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_movimentacao order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_controle where Idestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        Conexao.Execute "DELETE from Estoque_movimentacao where Idestoque = " & TBEstoque!IDEstoque
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        End If
            
        If .Chk2.Value = 1 Then
            'Custo material na ordens
            Conexao.Execute "Update producao Set CPR = 0, CTTReal = 0, CTMaterial = 0, CTServico = 0"
            
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from Producao order by Tipo desc, Ordem", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                TBCiclo.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCiclo.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCiclo.MoveFirst
                Do While TBCiclo.EOF = False
                    valor = 0
                    Valor1 = 0
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select IDlista from Compras_pedido_lista where Ordem = " & TBCiclo!Ordem & " and Tipo = 'P' and (Status_Item = 'N_RECEBIDO' or Status_Item = 'RECEBIDO' or Status_Item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select CODIGO from producaomaterial where Ordem = " & TBCiclo!Ordem & " order by codigo", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            Do While TBFI.EOF = False
                                Set TBEstoque = CreateObject("adodb.recordset")
                                TBEstoque.Open "Select Sum(EM.VlrTotal) as Valor from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.Documento = '" & TBCiclo!Ordem & "' and EM.Desenho = '" & TBFI!CODIGO & "' and EC.Consignacao = 'False' and (EM.Operacao = 'SAIDA_ORDEM' or EM.Operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                                If TBEstoque.EOF = False Then
                                    valor = IIf(IsNull(TBEstoque!valor), 0, TBEstoque!valor)
                                End If
                                TBFI!Valor_saida_estoque = valor
                                Valor1 = Valor1 + valor
                                TBFI.Update
                                
                                TBEstoque.Close
                                TBFI.MoveNext
                            Loop
                        End If
                        TBFI.Close
                        TBCiclo!CTMaterial = Format(Valor1, "###,##0.00")
                        TBCiclo.Update
                    End If
                    TBAbrir.Close
                    TBCiclo.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        End If
        
        If .Chk3.Value = 1 Then
            'Atualiza valor unitário do iventário que esta zerado
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_movimentacao where VlrUnit = 0 and (Operacao = 'ENTRADA_INVENTÁRIO' or Operacao = 'SAIDA_INVENTÁRIO') order by idEstoque, Data", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_movimentacao where idEstoque = " & TBEstoque!IDEstoque & " and Idoperacao <> " & TBEstoque!IDoperacao & " and VlrUnit <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        TBEstoque!VlrUnit = TBAbrir!VlrUnit
                        If TBEstoque!Operacao = "ENTRADA_INVENTÁRIO" Then
                            TBEstoque!vlrTotal = TBAbrir!VlrUnit * TBEstoque!Entrada
                        Else
                            TBEstoque!vlrTotal = TBAbrir!VlrUnit * TBEstoque!Saida
                        End If
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBEstoque.Close
            
            'Valor unitário de entrada e local de armazenamento na lista de materias da ordem
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_Controle order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        If TBProduto!Estoque = True Then ControlaEstoque = True Else ControlaEstoque = False
                        TBEstoque!Un = TBProduto!Unidade
                    End If
                    TBProduto.Close
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_movimentacao where idEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Do While TBAbrir.EOF = False
                            
                            'Verificar se o produto é remessa e marca como não controla estoque
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select CPL.IDlista from Estoque_controle_recebimento ECR INNER JOIN Compras_pedido_lista CPL ON ECR.IDPedido = CPL.IDPedido and ECR.IdLista = CPL.IdLista and ECR.Desenho = CPL.Desenho where ECR.Id = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento) & " and CPL.remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then ControlaEstoque = False
                            TBProduto.Close
                                                        
                            If TBAbrir!Operacao <> "DEVOLUCAO_ALMOXARIFADO C/ PROB." Then
                                Entrada = Entrada + IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada)
                                EntradaPC = EntradaPC + IIf(IsNull(TBAbrir!Entrada_PC), 0, TBAbrir!Entrada_PC)
                            End If
                                
                            Saida = Saida + IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            SaidaPC = SaidaPC + IIf(IsNull(TBAbrir!Saida_PC), 0, TBAbrir!Saida_PC)
                            
                            If TBAbrir!Operacao = "ENTRADA_INVENTÁRIO" Then
                                If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                    'Verif. valor unitário no cadastro do produto
                                    Set TBProduto = CreateObject("adodb.recordset")
                                    TBProduto.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBProduto.EOF = False Then
                                        valor = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
                                    End If
                                    TBProduto.Close
                                Else
                                    valor = TBAbrir!VlrUnit
                                End If
                            ElseIf TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Then
                                    'Verif. valor unitário na ordem
                                    Set TBProduto = CreateObject("adodb.recordset")
                                    TBProduto.Open "Select Ordem, Quant, QuantProd, QuantNC, CTTReal, CPR, CTServico, CTMaterial, CTOutras, consignacao from producao where Ordem = " & TBAbrir!LOTE, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBProduto.EOF = False Then
                                                                      'ORDEM           QTDE. PREVISTA                                    QTDE. OK                                                  QT. PROD.(OK+NC)                                                                                                 CUSTO LOTE                                            CUSTO PEÇA                                    CUSTO TERCEIROS                                           CUSTO MATERIAL                                              CUSTO OUTRAS                                            ORDEM CONSIGNADA
                                        valor = FunCalculaValorUnitOrdem(TBProduto!Ordem, IIf(IsNull(TBProduto!Quant), 0, TBProduto!Quant), IIf(IsNull(TBProduto!QuantProd), 0, TBProduto!QuantProd), IIf(IsNull(TBProduto!QuantProd), 0, TBProduto!QuantProd) + IIf(IsNull(TBProduto!QuantNC), 0, TBProduto!QuantNC), IIf(IsNull(TBProduto!CTTReal), 0, TBProduto!CTTReal), IIf(IsNull(TBProduto!CPR), 0, TBProduto!CPR), IIf(IsNull(TBProduto!CTServico), 0, TBProduto!CTServico), IIf(IsNull(TBProduto!CTMaterial), 0, TBProduto!CTMaterial), IIf(IsNull(TBProduto!CTOutras), 0, TBProduto!CTOutras), TBProduto!Consignacao)
                                        OF = TBProduto!Ordem
                                    End If
                                    TBProduto.Close
                                ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Then
                                        If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                            'Verif. valor unitário no cadastro do produto
                                            Set TBProduto = CreateObject("adodb.recordset")
                                            TBProduto.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                            If TBProduto.EOF = False Then
                                                valor = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
                                            End If
                                            TBProduto.Close
                                        End If
                                    ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL" Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                            TBFIltro.Open "Select IDlista, ID_empresa from Estoque_controle_recebimento where ID = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento), Conexao, adOpenKeyset, adLockOptimistic
                                            If TBFIltro.EOF = False Then
                                                
                                                'Verifica dados da NF
                                                Set TBFI = CreateObject("adodb.recordset")
                                                TBFI.Open "Select NF.ID_empresa, NF.Estado, NFP.Int_codigo, NFP.txt_Unid, NFP.Unidade_com, NFP.int_Qtd, NFP.Valor_frete, NFP.Valor_seguro, NFP.Valor_acessorias, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, NFP.Total_CSLL_prod, NFP.Total_IRPJ_prod, NFP.dbl_ValorUnitario, NFP.txt_Unid, NFP.Unidade_com from (tbl_Detalhes_Nota NFP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NFP.ID_nota = NF.ID) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo where NFPP.ID_carteira = " & TBFIltro!IDlista & " and NFPP.Codinterno = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                                If TBFI.EOF = False Then
                                                    qt = 1
                                                    If TBFI!txt_Unid <> TBFI!Unidade_com And TBFI!Qtde_estoque > 0 Then qt = TBFI!int_Qtd / TBAbrir!Entrada
                                                    
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
                                                    If ValorICMS <> 0 Then Valor1 = Format(qt * (ValorICMS / TBFI!int_Qtd), "0.0000000000") 'Valor unitário de ICMS
                                                    
                                                    EstoqueSaidaN2 = Format(qt * (IIf(IsNull(TBFI!Valor_desconto), 0, TBFI!Valor_desconto) / TBFI!int_Qtd), "0.0000000000") 'Valor unitário de desconto
                                                    Valor2 = Format(qt * (TBFI!Valor_frete / TBFI!int_Qtd), "0.0000000000")
                                                    ValorPagar = Format(qt * (TBFI!Valor_seguro / TBFI!int_Qtd), "0.0000000000")
                                                    ValorPago = Format(qt * (TBFI!Valor_acessorias / TBFI!int_Qtd), "0.0000000000")
                                                    
                                                    Set TBAliquota = CreateObject("adodb.recordset")
                                                    TBAliquota.Open "Select Simples, Real from Empresa where Codigo = " & TBFIltro!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBAliquota.EOF = False Then
                                                        If TBAliquota!Simples = True Then
                                                            If TBFI!Estado = "EX" Then
                                                                'Quando é nota de importação os valores de PIS e Cofins já estão inclusos nas despesas acessorias
                                                                Valor_PIS_Prod = 0
                                                                Valor_Cofins_Prod = 0
                                                            Else
                                                                Valor_PIS_Prod = Format(qt * (TBFI!Total_PIS_prod / TBFI!int_Qtd), "0.0000000000")
                                                                Valor_Cofins_Prod = Format(qt * (TBFI!Total_Cofins_prod / TBFI!int_Qtd), "0.0000000000")
                                                            End If
                                                            Valor_CSLL_Prod = Format(qt * (TBFI!Total_CSLL_prod / TBFI!int_Qtd), "0.0000000000")
                                                            Valor_IRPJ_Prod = Format(qt * (TBFI!Total_IRPJ_prod / TBFI!int_Qtd), "0.0000000000")
                                                            'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário - Valor desc.) + (Valor ICMS + Valor do frete + Valor Seguro + Valor despesas + Valor PIS + Valor Cofins + Valor CSLL + Valor IRPJ)
                                                            Qtd = Format(qt * (IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario) - EstoqueSaidaN2) + (Valor1 + Valor2 + ValorPagar + ValorPago + Valor_PIS_Prod + Valor_Cofins_Prod + Valor_CSLL_Prod + Valor_IRPJ_Prod), "0.0000000000")
                                                        ElseIf TBAbrir!Real = True Then
                                                                Valor_PIS_Prod = Format(qt * (TBFI!Total_PIS_prod / TBFI!int_Qtd), "0.0000000000")
                                                                Valor_Cofins_Prod = Format(qt * (TBFI!Total_Cofins_prod / TBFI!int_Qtd), "0.0000000000")
                                                                'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário + Valor do frete + Valor Seguro + Valor despesas) - (Valor desc. + Valor ICMS + Valor PIS + Valor Cofins)
                                                                Qtd = (qt * IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario)) + Valor2 + ValorPagar + ValorPago
                                                                Qtd = Format(Qtd - (EstoqueSaidaN2 + Valor1 + Valor_PIS_Prod + Valor_Cofins_Prod), "0.0000000000")
                                                            Else
                                                                'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário + Valor do frete + Valor Seguro + Valor despesas) - (Valor desc. + Valor ICMS)
                                                                Qtd = (qt * IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario)) + Valor2 + ValorPagar + ValorPago
                                                                Qtd = Format(Qtd - (EstoqueSaidaN2 + Valor1), "0.0000000000")
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
                            
                            TBAbrir!Familia = TBEstoque!Classe
                            TBEstoque!valor_unitario = Format(valor, "###,##0.0000000000")
                            TBAbrir!VlrUnit = Format(valor, "###,##0.0000000000")
                            If IsNull(TBAbrir!Entrada) = False And TBAbrir!Entrada <> "0" Then TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada) Else TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            TBAbrir.Update
                            TBAbrir.MoveNext
                        Loop
                    End If
                    
                    Total = Format(Entrada - Saida, "###,##0.0000000000")
                    TotalPC = Format(EntradaPC - SaidaPC, "###,##0.0000000000")
                    
                    If TBEstoque!local_armaz = "" Or IsNull(TBEstoque!local_armaz) = True Then TBEstoque!local_armaz = "N/A"
                    If ControlaEstoque = True Then
                        TBEstoque!estoque_venda = Total
                        TBEstoque!estoque_real = Total
                        TBEstoque!estoque_real_PC = TotalPC
                        TBEstoque!Valor_total = Format(valor * TBEstoque!estoque_real, "###,##0.00")
                    Else
                        TBEstoque!estoque_venda = 0
                        TBEstoque!estoque_real = 0
                        TBEstoque!estoque_real_PC = 0
                        TBEstoque!Valor_total = 0
                    End If
                    TBEstoque.Update
                    
                    Entrada = 0
                    EntradaPC = 0
                    Saida = 0
                    SaidaPC = 0
                    Total = 0
                    TotalPC = 0
                    valor = 0
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBEstoque.Close
        End If
        
        If .Chk4.Value = 1 Then
            'Verifica pedidos de compra com centro de custo e produto que controla estoque
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select CP.ID_empresa, CPL.Desenho, CPL.Descricao, CPL.Familia, CPL.UN, CPL.Unidade_com, P.peso_metro, P.un_kg, EC.IDestoque, EC.Estoque_real from (((Compras_pedido_lista CPL INNER JOIN Compras_pedido_lista_custo CPLC ON CPLC.IDLista = CPL.IdLista) INNER JOIN projproduto P ON P.Desenho = CPL.Desenho) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDpedido) INNER JOIN Estoque_Controle EC ON EC.Lote = CP.Pedido and EC.Desenho = CPL.Desenho where P.Estoque = 'True' and EC.Estoque_real > 0 and CPL.Tipo = 'P' and CPLC.ID_CC IS NOT NULL and CPLC.ID_CC <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBPedido.RecordCount
                PBLista.Value = 1
                Contador = 0
                
                'Cria requisição
                Set TBCompras = CreateObject("adodb.recordset")
                TBCompras.Open "Select * from Requisicao_materiais", Conexao, adOpenKeyset, adLockOptimistic
                TBCompras.AddNew
                TBCompras!ID_empresa = TBPedido!ID_empresa
                TBCompras!Responsavel = "PROCAM"
                TBCompras!Data = Date
                TBCompras!status = "RETIRADA"
                TBCompras!DtValidacao = Now
                TBCompras!RespValidacao = "PROCAM"
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Requisicao_materiais where Year(data) = '" & Year(Date) & "' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Numero = Left(TBAbrir!requisicao, Len(TBAbrir!requisicao) - 3)
                    Numero = Right(Numero, 5) + 1
                Else
                    Numero = 1
                End If
                TBAbrir.Close
                a = "RM-" & FunTamanhoTextoZeroEsq(Numero, 5) & "/" & Right(Year(Date), 2)
                TBCompras!requisicao = a
                TBCompras!Obs = "ACERTO DO ESTOQUE *** PRODUTO QUE CONTROLA ESTOQUE E TEM CENTRO DE CUSTO NO PEDIDO DE COMPRA ***"
                TBCompras.Update
                ID_RM = TBCompras!ID
                
                Do While TBPedido.EOF = False
                    'Salva o produto na RM
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from Requisicao_materiais_lista", Conexao, adOpenKeyset, adLockOptimistic
                    TBCompras.AddNew
                    TBCompras!IDRequisicao = ID_RM
                    TBCompras!Data = Date
                    TBCompras!Responsavel = "PROCAM"
                    TBCompras!status = "RETIRADO"
                    TBCompras!Desenho = TBPedido!Desenho
                    TBCompras!Quant = TBPedido!estoque_real
                    TBCompras!quant_saida = TBPedido!estoque_real
                    TBCompras!Familia = TBPedido!Familia
                    TBCompras!Descricao = TBPedido!Descricao
                    TBCompras!Un = TBPedido!Un
                    TBCompras!Unidade_com = TBPedido!Unidade_com
                    TBCompras!ID_CC = Null
                    TBCompras!Data_autorizacao = Null
                    TBCompras!Autorizado = ""
                    TBCompras!Obs = Null
                    TBCompras.Update
                    TBCompras.Close
                
                    'Retira o produto do estoque
                    Set TBEstoque = CreateObject("adodb.recordset")
                    TBEstoque.Open "Select * from estoque_controle where IDestoque = " & TBPedido!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBEstoque.EOF = False Then
                        qtdeliberada = 0
                        qtdeliberadaPC = 0
                        qtdeliberar = 0
                        qtdeliberarPC = 0
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select Sum(Entrada) as qtdeliberada, Sum(ISNULL(Entrada_PC, 0)) as qtdeliberadaPC, Sum(Saida) as qtdeliberar, Sum(ISNULL(Saida_PC, 0)) as qtdeliberarPC from Estoque_movimentacao where IDestoque = " & TBPedido!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            qtdeliberada = IIf(IsNull(TBFI!qtdeliberada), 0, TBFI!qtdeliberada)
                            'qtdeliberadaPC = IIf(IsNull(TBFI!qtdeliberadaPC), 0, TBFI!qtdeliberadaPC)
                            qtdeliberar = IIf(IsNull(TBFI!qtdeliberar), 0, TBFI!qtdeliberar)
                            'qtdeliberarPC = IIf(IsNull(TBFI!qtdeliberarPC), 0, TBFI!qtdeliberarPC)
                            QtdeEstoque = Format(qtdeliberada - (qtdeliberar + TBPedido!estoque_real), "###,##0.0000")
                            'QtdeEstoquePC = Format(qtdeliberadaPC - (qtdeliberarPC + QtdeSaidaPC), "###,##0.0000")
                        End If
                        TBFI.Close
                        
                        TBEstoque!peso_unit = TBPedido!peso_metro
                        'TBEstoque!Pedido = IIf(txtPedidoCompra = "", Null, txtPedidoCompra)
                       
                        TBEstoque!estoque_real = QtdeEstoque
                        'TBEstoque!estoque_real_PC = QtdeEstoquePC
                        TBEstoque!estoque_venda = QtdeEstoque
                        TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * QtdeEstoque, "###,##0.00")
                               
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
                        TBProduto.AddNew
                        TBProduto!Operacao = "SAIDA_REQUISICAO"
                        TBProduto!Documento = a
                        TBProduto!LOTE = TBEstoque!LOTE
                        TBProduto!Desenho = TBEstoque!Desenho
                        TBProduto!Data = Date
                        TBProduto!Descricao = TBEstoque!Descricao
                        TBProduto!Familia = TBEstoque!Classe
                        TBProduto!Requisitante = "PROCAM"
                        TBProduto!Responsavel = "PROCAM"
                        TBProduto!IDEstoque = TBEstoque!IDEstoque
                        TBProduto!OE = a
                        TBProduto!Destino = "Interno"
                        TBProduto!Terceiros = False
                        
                        TBProduto!Saida = TBPedido!estoque_real
                        'TBProduto!Saida_PC = IIf(txtquantretirado_PC = "", 0, txtquantretirado_PC)
                        TBProduto!estoque_venda = QtdeEstoque
                    
                        'Atualiza valor do material no estoque
                        TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.0000000000"))
                        TBProduto!vlrTotal = Format(TBPedido!estoque_real * TBProduto!VlrUnit, "###,##0.00")
                    
                        TBEstoque.Update
                        TBProduto.Update
                        TBProduto.Close
                    End If
                    TBEstoque.Close
                    TBPedido.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBPedido.Close
        End If
        
        If .Chk5.Value = 1 Then
            'Corrige empenho no RE
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select Ordem, Ordemempenho, Qtde_empenho from Producao_pedidos where Ordemempenho IS NOT NULL and Ordemempenho <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBPedido.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBPedido.EOF = False
                    valor = TBPedido!Qtde_empenho
                    
                    Set TBEstoque = CreateObject("adodb.recordset")
                    TBEstoque.Open "Select IDestoque, Desenho, Data, Responsavel from Estoque_Controle where Lote = '" & TBPedido!Ordem & "' and LEFT(status, 13) = 'ENTRADA_ORDEM' and Estoque_real > 0", Conexao, adOpenKeyset, adLockOptimistic
                    Do While TBEstoque.EOF = False
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ISNULL(Sum(Entrada), 0) as Valor1 from Estoque_movimentacao where IDestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Valor1 = TBAbrir!Valor1
                        End If
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ISNULL(Sum(Quantidade), 0) as Valor3 from Producao_NF_Consignada where Ordem = " & TBPedido!OrdemEmpenho & " and IDestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Valor3 = TBAbrir!Valor3
                            valor = valor - Valor3
                        End If
                        TBAbrir.Close
                        If valor > 0 And Valor1 - Valor3 > 0 Then
                            Set TBGravar = CreateObject("adodb.recordset")
                            TBGravar.Open "Select * from Producao_NF_Consignada", Conexao, adOpenKeyset, adLockOptimistic
                            TBGravar.AddNew
                            TBGravar!Ordem = TBPedido!OrdemEmpenho
                            TBGravar!Codinterno = TBEstoque!Desenho
                            If valor <= Valor1 Then TBGravar!quantidade = valor Else TBGravar!quantidade = Valor1
                            TBGravar!IDEstoque = TBEstoque!IDEstoque
                            TBGravar!Data = TBEstoque!Data
                            TBGravar!Responsavel = TBEstoque!Responsavel
                            TBGravar!Qtde_saida = 0
                            TBGravar!Quantidade_PC = TBGravar!quantidade
                            TBGravar!Qtde_saida_PC = 0
                            TBGravar.Update
                            TBGravar.Close
                        End If
                        TBEstoque.MoveNext
                    Loop
                    TBEstoque.Close
                    TBPedido.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBPedido.Close
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Movimentação"
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
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Procentrada()
On Error GoTo tratar_erro
  
frmestoque_entrada.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcRetirada()
On Error GoTo tratar_erro
  
frmestoque_Retirar.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub GridMv_CellChange(ByVal Row As Long, ByVal Col As Long)
On Error GoTo tratar_erro

'vRow = GridMv.ActiveCell.Row
'IDlista = Int(GridMv.Cell(vRow, 9).Text)
'Status_movimentacao = GridMv.Cell(vRow, 2).Text
'EstoqueEntrada = GridMv.Cell(vRow, 4).Text
'EstoqueSaida = GridMv.Cell(vRow, 5).Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub GridMv_Click()
On Error GoTo tratar_erro

vRow = GridMv.ActiveCell.Row

If vRow > 0 Then
IDlista = Int(GridMv.Cell(vRow, 9).Text)
Status_movimentacao = GridMv.Cell(vRow, 2).Text
EstoqueEntrada = GridMv.Cell(vRow, 4).Text
EstoqueSaida = GridMv.Cell(vRow, 5).Text
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
ProcCarregaDadosLote

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaDadosLote()
On Error GoTo tratar_erro

TTE = 0
quantestoque = 0
quantestoquelote = 0
IDempresa = 0

Lista_Movimentacao.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM estoque_produtos where idestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    NomeCampo = "o local de armazenamento"
    If IsNull(TBAbrir!local_armaz) = False And TBAbrir!local_armaz <> "" Then txtLocal_armaz = TBAbrir!local_armaz
    NomeCampo = "a família"
    If IsNull(TBAbrir!Classe) = False And TBAbrir!Classe <> "" Then cmbfamilia = TBAbrir!Classe
1:
    txtlocalização.Text = Lista.SelectedItem
    
    'RE
    Txt_cod_ref = IIf(IsNull(TBAbrir!Ref), "", TBAbrir!Ref)
    Txt_n_serie = IIf(IsNull(TBAbrir!Numero_serie), "", TBAbrir!Numero_serie)
    Txt_qtde_estoqueRE = Format(TBAbrir!estoque_real, "###,##0.0000")
    Txt_qtde_estoque_PCRE = IIf(IsNull(TBAbrir!estoque_real_PC), "0,0000", Format(TBAbrir!estoque_real_PC, "###,##0.0000"))
    Txt_qtde_empenhoRE = Format(TBAbrir!Qtde_empenhada, "###,##0.0000")
    Txt_qtde_est_dispRE = Format(TBAbrir!Estoque_Disponivel, "###,##0.0000")
    Txt_qtde_est_disp_PCRE = Format(TBAbrir!Estoque_disponivel_PC, "###,##0.0000")
    
    Qtde = 0
    Qtd = 0
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Saida) as qtde from estoque_movimentacao EM INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = EM.ID_prod_NF where EM.idestoque = " & Lista.SelectedItem & " and EM.destino = 'Terceiros' and NFP.Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtde = IIf(IsNull(TBFI!Qtde), 0, TBFI!Qtde)
    End If
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Entrada) as qtd from ((estoque_movimentacao EM INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = EM.Documento and NF.txt_Razao_Nome = EC.Fornecedor) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID and NFP.int_Cod_Produto = EM.Desenho where EM.idestoque = " & Lista.SelectedItem & " and EM.pedidocompra IS NOT NULL and NFP.Retorno = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtd = IIf(IsNull(TBFI!Qtd), 0, TBFI!Qtd)
    End If
    Txt_qtde_est_terc = Format(Qtde - Qtd, "###,##0.0000")
    
    Txt_valor_total_estRE = Format(TBAbrir!Valor_total, "###,##0.00")
    Txt_valor_unitRE = Format(TBAbrir!valor_unitario, "###,##0.0000000000")
            
    'Código interno
    Txt_qtde_estoque = Format(FunVerificaQtdeEstoque(Lista.SelectedItem.ListSubItems(3), Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
    Txt_qtde_estoque_PC = Format(qt, "###,##0.0000")
    Txt_qtde_empenho = Format(QuantEmpenho, "###,##0.0000")
    Txt_qtde_est_disp = Format(quantestoque, "###,##0.0000")
    Txt_qtde_est_disp_PC = Format(quantnovo, "###,##0.0000")
        
    Qtde = 0
    Qtd = 0
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Saida) as qtde from estoque_movimentacao EM INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = EM.ID_prod_NF where EM.Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and EM.destino = 'Terceiros' and NFP.Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtde = IIf(IsNull(TBFI!Qtde), 0, TBFI!Qtde)
    End If
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Entrada) as qtd from ((estoque_movimentacao EM INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = EM.Documento and NF.txt_Razao_Nome = EC.Fornecedor) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID and NFP.int_Cod_Produto = EM.Desenho where EM.Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and EM.pedidocompra IS NOT NULL and NFP.Retorno = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtd = IIf(IsNull(TBFI!Qtd), 0, TBFI!Qtd)
    End If
    Txt_qtde_est_terc = Format(Qtde - Qtd, "###,##0.0000")
    
    Txt_valor_total_est = Format(Valor_total, "###,##0.00")
    Txt_custo_medio_unit = Format(CTMedioEst, "###,##0.0000000000")
    
    ProcCarregaListaMovimentacao
    CodigoLista = Lista.SelectedItem.index
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaListaMovimentacao()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub

Lista_Movimentacao.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from estoque_movimentacao where idestoque = " & Lista.SelectedItem & " order by Data desc, Idoperacao desc", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBAbrir.EOF = False
        With Lista_Movimentacao.ListItems
            .Add , , TBAbrir!IDoperacao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Operacao), "", TBAbrir!Operacao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Entrada), "0,0000", Format(TBAbrir!Entrada, "###,##0.0000"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Entrada_PC), "0,0000", Format(TBAbrir!Entrada_PC, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Saida), "0,0000", Format(TBAbrir!Saida, "###,##0.0000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!Saida_PC), "", Format(TBAbrir!Saida_PC, "###,##0.0000"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!Documento), "", TBAbrir!Documento)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBAbrir!Requisitante), "", TBAbrir!Requisitante)
            If TBAbrir!Destino = "Terceiros" Then .Item(.Count).SubItems(11) = "Terceiros (Remessa forn.)" Else .Item(.Count).SubItems(11) = "Interno/Cliente"
            
            If TBAbrir!Entrada > 0 Or TBAbrir!Destino = "Terceiros" Then
                .Item(.Count).SubItems(12) = IIf(IsNull(TBAbrir!Pedidocompra), "", TBAbrir!Pedidocompra)
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "Select Fornecedor from Compras_pedido where idpedido = " & IIf(IsNull(TBAbrir!IDpedido), 0, TBAbrir!IDpedido), Conexao, adOpenKeyset, adLockOptimistic
                If TBPedido.EOF = False Then .Item(.Count).SubItems(13) = IIf(IsNull(TBPedido!Fornecedor), "", TBPedido!Fornecedor)
                TBPedido.Close
            Else
                If TBAbrir!Operacao = "SAIDA_NOTA" Or TBAbrir!Operacao = "SAIDA_NOTA_PARCIAL" Then
                    Set TBControleNF = CreateObject("adodb.recordset")
                    TBControleNF.Open "Select TDNF.txt_Razao_Nome from tbl_Detalhes_Nota TDN INNER JOIN tbl_Dados_Nota_Fiscal TDNF ON TDN.ID_Nota = TDNF.ID where TDN.Int_codigo = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF), Conexao, adOpenKeyset, adLockOptimistic
                    If TBControleNF.EOF = False Then
                        Set TBPedido = CreateObject("adodb.recordset")
                        TBPedido.Open "Select VP.Ncotacao from (tbl_Detalhes_Nota_pedidos TDNP INNER JOIN vendas_carteira VC ON TDNP.ID_carteira = VC.Codigo) INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.cotacao where TDNP.ID_prod_NF = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF), Conexao, adOpenKeyset, adLockOptimistic
                        If TBPedido.EOF = False Then .Item(.Count).SubItems(12) = IIf(IsNull(TBPedido!Ncotacao), "", TBPedido!Ncotacao)
                        TBPedido.Close
                        
                        .Item(.Count).SubItems(13) = IIf(IsNull(TBControleNF!txt_Razao_Nome), "", TBControleNF!txt_Razao_Nome)
                    End If
                    TBControleNF.Close
                End If
            End If
            .Item(.Count).SubItems(14) = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
        End With
        TBAbrir.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_movimentacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_Movimentacao
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Status_movimentacao = .ListItems(InitFor).SubItems(2)
                If Status_movimentacao = "SAIDA_ALMOXARIFADO" Or Status_movimentacao = "ENTRADA_ALMOXARIFADO" Or Status_movimentacao = "DEVOLUCAO_ALMOXARIFADO C/ PROB." Or Status_movimentacao = "SAIDA_SUCATA" Or Status_movimentacao = "ENTRADA_SUCATA" Or Status_movimentacao = "SAIDA_RETALHO" Or Status_movimentacao = "ENTRADA_RETALHO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL" Then
                    GoTo Proximo
                ElseIf Status_movimentacao = "ENTRADA_INVENTÁRIO" Or Status_movimentacao = "SAIDA_INVENTÁRIO" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from estoque_movimentacao where idoperacao = " & .ListItems(InitFor) & " and id_inventario <> 0 and ID_inventario IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        TBAbrir.Close
                        GoTo Proximo
                    End If
                    TBAbrir.Close
                Else
                    'Verifica se a entrada esta vinculada a ordem
                    If Left(Status_movimentacao, 7) = "ENTRADA" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select PNC.Ordem from estoque_movimentacao EM INNER JOIN Producao_NF_Consignada PNC ON EM.IDestoque = PNC.Idestoque where EM.idoperacao = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            'Verifica qtde. de entrada do RE
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select ROUND(SUM(ISNULL(Entrada, 0)), 3) as Valor from estoque_movimentacao where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                            End If
                            'Verifica qtde. empenhada do RE
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select ROUND(SUM(ISNULL(Quantidade, 0)), 3) as Valor1 from Producao_NF_Consignada where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
                            End If
                            Permitido = True
                            If valor - Valor1 <= 0 Then
                                Permitido = False
                            ElseIf (valor - .ListItems(InitFor).SubItems(4)) - Valor1 < 0 Then
                                Permitido = False
                            End If
                            If Permitido = False Then
                                TBAbrir.Close
                                GoTo Proximo
                            End If
                        End If
                        TBAbrir.Close
                    End If
                    
                    'Verifica se o resultado da ordem esta validado
                    If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Or Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select P.Ordem from estoque_movimentacao EM INNER JOIN Producao P ON EM.Documento = P.Ordem where EM.idoperacao = " & .ListItems(InitFor) & " and P.RespValidacao_Custo IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            TBAbrir.Close
                            GoTo Proximo
                        End If
                        TBAbrir.Close
                    End If
                    
                    'Verifica se a RE tem movimentações mais recentes
                    If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Or Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                        If Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                            Set TBAfericao = CreateObject("adodb.recordset")
                            TBAfericao.Open "SELECT IdEstoque, idoperacao FROM estoque_movimentacao WHERE IdTrocaLocal = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                            If TBAfericao.EOF = False Then
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & TBAfericao!IDEstoque & " AND Idoperacao <> " & TBAfericao!IDoperacao & " AND idoperacao > " & TBAfericao!IDoperacao, Conexao, adOpenKeyset, adLockReadOnly
                                If TBAbrir.EOF = False Then
                                    TBAbrir.Close
                                    TBAfericao.Close
                                    GoTo Proximo
                                End If
                                TBAbrir.Close
                            End If
                            TBAfericao.Close
                        Else
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & Lista.SelectedItem & " AND Idoperacao <> " & .ListItems(InitFor) & " AND  idoperacao > " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                            If TBAbrir.EOF = False Then
                                TBAbrir.Close
                                GoTo Proximo
                            End If
                            TBAbrir.Close
                        End If
                    End If
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_Movimentacao, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizaStatus_RM()
On Error GoTo tratar_erro

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Requisicao_materiais where requisicao = '" & TBAbrir!Documento & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Set TBMateriaprima = CreateObject("adodb.recordset")
    TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBMateriaprima.EOF = False Then
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'REQUISIT.'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = True Then
            TBproducao!status = "ABERTA"
        Else
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'RETIRADO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = True Then
                TBproducao!status = "RETIRADA"
            Else
                TBproducao!status = "PARCIAL"
            End If
        End If
        TBproducao.Update
    End If
    TBMateriaprima.Close
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcOrdem()
On Error GoTo tratar_erro

If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and SubTipoItem <> 0 and SubTipoItem <> 4", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        If TBCiclo!QuantProd - Qtd <= 0 Then TBCiclo!QuantProd = 0 Else TBCiclo!QuantProd = TBCiclo!QuantProd - Qtd
    End If
    TBProduto.Close
    If TBCiclo!QuantProd <> 0 Then
        TBCiclo!CPR = IIf(IsNull(TBCiclo!CTTReal), 0, TBCiclo!CTTReal) / TBCiclo!QuantProd
    Else
        TBCiclo!CPR = 0
        TBCiclo!Controlado_estoque = False
    End If
    If TBCiclo!QuantProd < TBCiclo!Quant Then
        TBCiclo!DataEntrega = Null
        TBCiclo!Concluida = False
        TBCiclo!pronta = "NÃO"
        If TBCiclo!status <> "Entregue" Then TBCiclo!status = "Aberta"
        
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from producao where Ordem = " & TBCiclo!Ordem & " and Ap_backup = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            NomeTabelaAp = "ProducaoFases_Backup"
        Else
            NomeTabelaAp = "ProducaoFases"
        End If
        
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from ordemservico where Ordem = " & TBCiclo!Ordem & " and pronto = 'SIM'", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            Do While TBproducao.EOF = False
                TBproducao!Pronto = "NÃO"
                TBproducao!DataConclusao = Null
                TBproducao!status = Null
                TBproducao.Update
                'Filtra todos os eventos desta OS na tabela producaofases para marcar como fase pronta
                Conexao.Execute "Update " & NomeTabelaAp & " Set pronto = 'NÃO' where idfase = " & TBproducao!IDProducao
                TBproducao.MoveNext
            Loop
        End If
        TBproducao.Close
    End If
    TBCiclo.Update
End If
'==================================
Modulo = "Estoque/Movimentação/Entrada"
Evento = "Alterar OF p/ não concluída"
ID_documento = TBCiclo!NOF
Documento = "Ordem: " & TBCiclo!Ordem
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_movimentacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Status_movimentacao = .ListItems(InitFor).SubItems(2)
            If Status_movimentacao = "SAIDA_ALMOXARIFADO" Or Status_movimentacao = "ENTRADA_ALMOXARIFADO" Or Status_movimentacao = "DEVOLUCAO_ALMOXARIFADO C/ PROB." Or Status_movimentacao = "SAIDA_SUCATA" Or Status_movimentacao = "ENTRADA_SUCATA" Or Status_movimentacao = "SAIDA_RETALHO" Or Status_movimentacao = "ENTRADA_RETALHO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL" Then
                If Status_movimentacao = "SAIDA_SUCATA" Or Status_movimentacao = "SAIDA_RETALHO" Then
                    USMsgBox ("Não é permitido excluir esta movimentação."), vbExclamation, "CAPRIND v5.0"
                ElseIf Status_movimentacao = "ENTRADA_SUCATA" Or Status_movimentacao = "ENTRADA_RETALHO" Then
                        USMsgBox ("Só é permitido excluir o lote desta movimentação, utilizando o botão (Excluir sucata/ret.)."), vbExclamation, "CAPRIND v5.0"
                    Else
                        USMsgBox ("Não é permitido excluir este tipo de movimentação neste módulo."), vbExclamation, "CAPRIND v5.0"
                End If
                .ListItems.Item(InitFor).Checked = False
            ElseIf Status_movimentacao = "ENTRADA_INVENTÁRIO" Or Status_movimentacao = "SAIDA_INVENTÁRIO" Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select idoperacao from estoque_movimentacao where idoperacao = " & .ListItems(InitFor) & " and id_inventario <> 0 and ID_inventario IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    USMsgBox ("Não é permitido excluir este tipo de movimentação neste módulo."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                TBAbrir.Close
            Else
                'Verifica se a entrada esta vinculada a ordem
                If Left(Status_movimentacao, 7) = "ENTRADA" And Status_movimentacao <> "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select PNC.Ordem from estoque_movimentacao EM INNER JOIN Producao_NF_Consignada PNC ON EM.IDestoque = PNC.Idestoque where EM.idoperacao = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        'Verifica qtde. de entrada do RE
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ROUND(SUM(ISNULL(Entrada, 0)), 3) as Valor from estoque_movimentacao where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                        End If
                        'Verifica qtde. empenhada do RE
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ROUND(SUM(ISNULL(Quantidade, 0)), 3) as Valor1 from Producao_NF_Consignada where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
                        End If
                        Permitido = True
                        If valor - Valor1 <= 0 Then
                            Permitido = False
                        ElseIf (valor - .ListItems(InitFor).SubItems(4)) - Valor1 < 0 Then
                            Permitido = False
                        End If
                        If Permitido = False Then
                            OPTexto = ""
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select Ordem from Producao_NF_Consignada where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                Do While TBAbrir.EOF = False
                                    If OPTexto = "" Then OPTexto = TBAbrir!Ordem Else OPTexto = OPTexto & " | " & TBAbrir!Ordem
                                    TBAbrir.MoveNext
                                Loop
                            End If
                            USMsgBox ("Não é permitido excluir esta movimentação, pois a mesma está sendo vinculada a(s) ordem(ns): " & vbCrLf & OPTexto), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    End If
                    TBAbrir.Close
                End If
                
                'Verifica se o resultado da ordem esta validado
                If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Or Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select P.Ordem from estoque_movimentacao EM INNER JOIN Producao P ON EM.Documento = P.Ordem where EM.idoperacao = " & .ListItems(InitFor) & " and P.RespValidacao_Custo IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido excluir esta movimentação, pois o resultado da ordem " & TBAbrir!Ordem & " já foi validado."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        TBAbrir.Close
                        Exit Sub
                    End If
                    TBAbrir.Close
                End If
                
                'Verifica se a RE tem movimentações mais recentes
                If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Or Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                    If Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                        Set TBAfericao = CreateObject("adodb.recordset")
                        TBAfericao.Open "SELECT IdEstoque, idoperacao FROM estoque_movimentacao WHERE IdTrocaLocal = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                        If TBAfericao.EOF = False Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & TBAfericao!IDEstoque & " AND Idoperacao <> " & TBAfericao!IDoperacao & " AND idoperacao > " & TBAfericao!IDoperacao, Conexao, adOpenKeyset, adLockReadOnly
                            If TBAbrir.EOF = False Then
                                USMsgBox ("Não é permitido excluir esta movimentação, pois exitem movimentações mais recentes no RE " & TBAbrir!IDEstoque & "."), vbExclamation, "CAPRIND v5.0"
                                .ListItems.Item(InitFor).Checked = False
                                TBAbrir.Close
                                TBAfericao.Close
                                Exit Sub
                            End If
                            TBAbrir.Close
                        End If
                        TBAfericao.Close
                    Else
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & Lista.SelectedItem & " AND Idoperacao <> " & .ListItems(InitFor) & " AND  idoperacao > " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                        If TBAbrir.EOF = False Then
                            USMsgBox ("Não é permitido excluir esta movimentação, pois exitem movimentações mais recentes no RE."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            TBAbrir.Close
                            Exit Sub
                        End If
                        TBAbrir.Close
                    End If
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub GridMv_DblClick()
On Error GoTo tratar_erro
'========================================
' Carrega os dados do lote (RE)
'========================================
vRow = GridRE.ActiveCell.Row
'========================================
' Passa os dados da RE
'========================================
With frmEstoque_CorrigirLote
.DataRE = GridRE.Cell(vRow, 2).Text
.txtRE = GridRE.Cell(vRow, 3).Text
.txtLote = GridRE.Cell(vRow, 4).Text
.txtStatus = GridRE.Cell(vRow, 5).Text
.txtCodigo = GridRE.Cell(vRow, 6).Text
.txtdescricao = GridRE.Cell(vRow, 7).Text
.txtUN = GridRE.Cell(vRow, 8).Text
.txtSaldoRE = GridRE.Cell(vRow, 12).Text
End With
'========================================
' Carrega os dados da movimentação
'========================================
vRow = GridMv.ActiveCell.Row
'========================================
' Passa dados da movimentação
'========================================
With frmEstoque_CorrigirLote
.txtOperacao = GridMv.Cell(vRow, 2).Text
.Data = GridMv.Cell(vRow, 3).Text
.txtEntrada = GridMv.Cell(vRow, 4).Text
.txtSaida = GridMv.Cell(vRow, 5).Text
.txtDocumento = GridMv.Cell(vRow, 6).Text
.txtId = Int(GridMv.Cell(vRow, 9).Text)
.Show 1
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridProdutos_CellChange(ByVal Row As Long, ByVal Col As Long)
On Error GoTo tratar_erro

ProcCarregaGridRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridProdutos_Click()
On Error GoTo tratar_erro

ProcCarregaGridRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridProdutos_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
On Error GoTo tratar_erro

'ProcCarregaGridRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridRE_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
On Error GoTo tratar_erro

ProcCarregaGridMV

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaGridMV()
On Error GoTo tratar_erro

vRow = GridRE.ActiveCell.Row
GridMv.rows = 1

If GridRE.Cell(vRow, 3).Text <> "" And vRow > 0 Then
RE = Int(GridRE.Cell(vRow, 3).Text)
'Operacao
Desenho = GridRE.Cell(vRow, 6).Text
EstoqueEmpenho = GridRE.Cell(vRow, 11).Text
Documento = GridRE.Cell(vRow, 4).Text
'If chkperiodo.Value = 1 Then
StrSql = "select * from Estoque_movimentacao where idEstoque ='" & RE & "' and Data <= '" & Ate.Value & "' ORDER BY IDOperacao"
'Else
'StrSql = "select * from Estoque_movimentacao where idEstoque ='" & RE & "' ORDER BY IDOperacao"
'End If

 Set TBEstoque = CreateObject("adodb.recordset")
 TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
 Contador = 1
 If TBEstoque.EOF = False Then
 Do While TBEstoque.EOF = False
    GridMv.AddItem Contador & vbTab & _
                 TBEstoque!Operacao & vbTab & _
                 TBEstoque!Data & vbTab & _
                 Format(TBEstoque!Entrada, "0.00") & vbTab & _
                 Format(TBEstoque!Saida, "0.00") & vbTab & _
                 TBEstoque!Documento & vbTab & _
                 TBEstoque!Responsavel & vbTab & _
                 TBEstoque!Requisitante & vbTab & _
                 TBEstoque!IDoperacao
Contador = Contador + 1
TBEstoque.MoveNext
Loop
End If
TBEstoque.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Public Sub ProcCarregaGridRE()
On Error GoTo tratar_erro
Contador = 0
vRow = GridProdutos.ActiveCell.Row
GridRE.rows = 1

If GridProdutos.Cell(vRow, 2).Text <> "" And vRow > 0 Then
Desenho = GridProdutos.Cell(vRow, 2).Text

StrSql = "Select EP.Data, EP.IdEstoque,EP.Qtde_empenhada, EP.status, EP.Desenho ,EP.Descricao, EP.Unidade, EP.Lote, SUM(EM.Entrada) as Entrada, SUM(EM.Saida) as Saida, SUM(EM.entrada-EM.Saida) as Saldo from Estoque_produtos EP inner join Estoque_movimentacao EM on EP.IdEstoque = EM.IdEstoque Where EM.Data <= '" & Ate.Value & "'  AND EP.Desenho = '" & Desenho & "' group by EP.data, EP.IdEstoque, EP.Desenho,EP.Qtde_empenhada, EP.status,EP.Classe,EP.Descricao, EP.Unidade, EP.Lote order by EP.idestoque"
'Debug.print StrSql

Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
Contador = 1
 
'If chkEstoquePositivo.Value = 1 Then
 
 If TBEstoque.EOF = False Then
 Do While TBEstoque.EOF = False
 If TBEstoque!Saldo > 0 Then
    GridRE.AddItem Contador & vbTab & _
                 TBEstoque!Data & vbTab & _
                 TBEstoque!IDEstoque & vbTab & _
                 TBEstoque!LOTE & vbTab & _
                 TBEstoque!status & vbTab & _
                 TBEstoque!Desenho & vbTab & _
                 TBEstoque!Descricao & vbTab & _
                 TBEstoque!Unidade & vbTab & _
                 Format(TBEstoque!Entrada, "0.00") & vbTab & _
                 Format(TBEstoque!Saida, "0.00") & vbTab & _
                 Format(TBEstoque!Qtde_empenhada, "0.00") & vbTab & _
                 Format(TBEstoque!Saldo, "0.00")
Contador = Contador + 1
End If
TBEstoque.MoveNext
Loop
End If
TBEstoque.Close
'Else
' If TBEstoque.EOF = False Then
' Do While TBEstoque.EOF = False
'    GridRE.AddItem contador & vbTab & _
'                 TBEstoque!Data & vbTab & _
'                 TBEstoque!IDestoque & vbTab & _
'                 TBEstoque!LOTE & vbTab & _
'                 TBEstoque!status & vbTab & _
'                 TBEstoque!Desenho & vbTab & _
'                 TBEstoque!Descricao & vbTab & _
'                 TBEstoque!Unidade & vbTab & _
'                 Format(TBEstoque!Entrada, "0.00") & vbTab & _
'                 Format(TBEstoque!Saida, "0.00") & vbTab & _
'                 Format(TBEstoque!Qtde_empenhada, "0.00") & vbTab & _
'                 Format(TBEstoque!Saldo, "0.00")
'contador = contador + 1
'TBEstoque.MoveNext
'Loop
'End If
'TBEstoque.Close
'End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro


If cmbfiltrarpor = "RE" And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 3: frmEstoque_Movimentacao_Exportar.Show 1
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procVoltarEmpenhoLocal(idEstoqueEntrada As Long, idEstoqueSaida As Long)
On Error GoTo tratar_erro

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT ID_estoque, IdAntigoLocal, Qtde_empenhada FROM Estoque_Controle_Empenho_Vendas where id_estoque = " & idEstoqueEntrada, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Do While TBCFOP.EOF = False
        If IsNull(TBCFOP!IdAntigoLocal) = True Then
            TBCFOP!ID_estoque = idEstoqueSaida
            TBCFOP.Update
        Else
            Set TBCorretiva = CreateObject("adodb.recordset")
            TBCorretiva.Open "SELECT Qtde_empenhada FROM Estoque_Controle_Empenho_Vendas where id = " & TBCFOP!IdAntigoLocal, Conexao, adOpenKeyset, adLockOptimistic
            If TBCorretiva.EOF = False Then
                TBCorretiva!Qtde_empenhada = TBCFOP!Qtde_empenhada + TBCorretiva!Qtde_empenhada
                TBCorretiva.Update
            End If
            TBCorretiva.Close
            TBCFOP.Delete
        End If
        TBCFOP.MoveNext
    Loop
End If
TBCFOP.Close

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT IDestoque, IdAntigoLocal, Quantidade, Quantidade_PC FROM Producao_NF_Consignada where idestoque = " & idEstoqueEntrada, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Do While TBCFOP.EOF = False
        If IsNull(TBCFOP!IdAntigoLocal) = True Then
            TBCFOP!IDEstoque = idEstoqueSaida
            TBCFOP.Update
        Else
            Set TBCorretiva = CreateObject("adodb.recordset")
            TBCorretiva.Open "SELECT * FROM Producao_NF_Consignada where id = " & TBCFOP!IdAntigoLocal, Conexao, adOpenKeyset, adLockOptimistic
            If TBCorretiva.EOF = False Then
                TBCorretiva!quantidade = TBCFOP!quantidade + TBCorretiva!quantidade
                TBCorretiva!Quantidade_PC = TBCFOP!Quantidade_PC + TBCorretiva!Quantidade_PC
                TBCorretiva.Update
            End If
            TBCorretiva.Close
            TBCFOP.Delete
        End If
        TBCFOP.MoveNext
    Loop
End If
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
