VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_ListaProduto 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Pedido - Localizar"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   14385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompras_ListaProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
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
      FormHeightDT    =   9525
      FormWidthDT     =   14505
      FormScaleHeightDT=   9060
      FormScaleWidthDT=   14385
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   51
      Top             =   8760
      Width           =   14325
      _ExtentX        =   25268
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
      Height          =   9120
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   16087
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Necessidade"
      TabPicture(0)   =   "frmCompras_ListaProduto.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListaNecessidade"
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(2)=   "Frame1(23)"
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(4)=   "Frame1(4)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Solicitação"
      TabPicture(1)   =   "frmCompras_ListaProduto.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(5)"
      Tab(1).Control(1)=   "Frame1(0)"
      Tab(1).Control(2)=   "USToolBar3"
      Tab(1).Control(3)=   "Lista_solicitados"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Cadastrados"
      TabPicture(2)   =   "frmCompras_ListaProduto.frx":0D02
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "USToolBar2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FlexGrid"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame1(6)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   6
         Left            =   65
         TabIndex        =   74
         Top             =   8130
         Width           =   14325
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
            Index           =   2
            Left            =   3930
            TabIndex        =   42
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
            Index           =   2
            Left            =   8220
            TabIndex        =   43
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   2
            Left            =   10440
            TabIndex        =   47
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":0D1E
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
            Index           =   2
            Left            =   9900
            TabIndex        =   46
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":44C5
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
            Index           =   2
            Left            =   8790
            TabIndex        =   44
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
            Index           =   2
            Left            =   9360
            TabIndex        =   45
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":7FD1
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
            Index           =   2
            Left            =   10980
            TabIndex        =   48
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":C0C3
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
         Begin VB.Label Label8 
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
            Index           =   2
            Left            =   3240
            TabIndex        =   77
            Top             =   240
            Width           =   2760
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
            Index           =   2
            Left            =   180
            TabIndex        =   76
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
            Index           =   2
            Left            =   12090
            TabIndex        =   75
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   5
         Left            =   -74935
         TabIndex        =   70
         Top             =   8130
         Width           =   14325
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
            Index           =   1
            Left            =   8220
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
            Index           =   1
            Left            =   3930
            TabIndex        =   26
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   1
            Left            =   10440
            TabIndex        =   31
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":F951
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
            Index           =   1
            Left            =   9900
            TabIndex        =   30
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":130F8
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
            Index           =   1
            Left            =   8790
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
            Index           =   1
            Left            =   9360
            TabIndex        =   29
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":16C05
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
            Index           =   1
            Left            =   10980
            TabIndex        =   32
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":1ACF6
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
            Index           =   1
            Left            =   12090
            TabIndex        =   73
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
            Index           =   1
            Left            =   180
            TabIndex        =   72
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label8 
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
            Index           =   1
            Left            =   3240
            TabIndex        =   71
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   4
         Left            =   -74935
         TabIndex        =   66
         Top             =   8130
         Width           =   14325
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
            Index           =   0
            Left            =   3930
            TabIndex        =   10
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
            Index           =   0
            Left            =   8220
            TabIndex        =   11
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   0
            Left            =   10440
            TabIndex        =   15
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":1E584
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
            Index           =   0
            Left            =   9900
            TabIndex        =   14
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":21D2B
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
            Index           =   0
            Left            =   8790
            TabIndex        =   12
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
            Index           =   0
            Left            =   9360
            TabIndex        =   13
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":25838
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
            Index           =   0
            Left            =   10980
            TabIndex        =   16
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_ListaProduto.frx":2992A
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
         Begin VB.Label Label8 
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
            Index           =   0
            Left            =   3240
            TabIndex        =   69
            Top             =   240
            Width           =   2760
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
            Index           =   0
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
            Index           =   0
            Left            =   12090
            TabIndex        =   67
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   495
         Left            =   -74935
         TabIndex        =   65
         Top             =   1320
         Width           =   14325
         Begin VB.OptionButton Opt_PCP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Por PCP"
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
            Left            =   180
            TabIndex        =   0
            Top             =   210
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton Opt_vendas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Por vendas"
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
            Left            =   1320
            TabIndex        =   1
            Top             =   210
            Width           =   1245
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Height          =   555
         Left            =   65
         TabIndex        =   63
         Top             =   1320
         Width           =   14325
         Begin VB.CheckBox Chk_prazo_todos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Utilizar prazo para todos:"
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
            Left            =   4020
            TabIndex        =   81
            Top             =   210
            Width           =   2445
         End
         Begin VB.CheckBox Chk_fornecedor 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Filtrar produtos/serviços do fornecedor"
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
            Height          =   255
            Left            =   180
            TabIndex        =   33
            Top             =   180
            Width           =   4185
         End
         Begin MSComCtl2.DTPicker Cmb_prazo 
            Height          =   315
            Left            =   6540
            TabIndex        =   82
            ToolTipText     =   "Prazo final."
            Top             =   180
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   179896323
            CurrentDate     =   39057
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   2
         Left            =   65
         TabIndex        =   59
         Top             =   1890
         Width           =   14325
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2760
            TabIndex        =   80
            Top             =   210
            Width           =   4785
            Begin VB.OptionButton optIgual_cad 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               Height          =   255
               Left            =   3930
               TabIndex        =   38
               Top             =   180
               Width           =   705
            End
            Begin VB.OptionButton optMeio_cad 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               Height          =   255
               Left            =   1470
               TabIndex        =   36
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optInicio_cad 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início frase"
               Height          =   255
               Left            =   180
               TabIndex        =   35
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton optFim_cad 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               Height          =   255
               Left            =   2760
               TabIndex        =   37
               Top             =   180
               Width           =   1155
            End
         End
         Begin VB.ComboBox Cmb_ordenar 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmCompras_ListaProduto.frx":2D1B8
            Left            =   11850
            List            =   "frmCompras_ListaProduto.frx":2D1C2
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   41
            ToolTipText     =   "Ordenar por."
            Top             =   390
            Width           =   2265
         End
         Begin VB.ComboBox cmbfiltrarpor_cad 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmCompras_ListaProduto.frx":2D1E1
            Left            =   180
            List            =   "frmCompras_ListaProduto.frx":2D203
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2505
         End
         Begin VB.TextBox txtTexto_cad 
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
            Left            =   7620
            TabIndex        =   39
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   4215
         End
         Begin VB.ComboBox cmbTexto_cad 
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
            ItemData        =   "frmCompras_ListaProduto.frx":2D290
            Left            =   7620
            List            =   "frmCompras_ListaProduto.frx":2D292
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            ToolTipText     =   "Familia."
            Top             =   390
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenar por"
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
            Left            =   12472
            TabIndex        =   64
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   8985
            TabIndex        =   61
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label1 
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
            Index           =   3
            Left            =   1012
            TabIndex        =   60
            Top             =   180
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   0
         Left            =   -74935
         TabIndex        =   56
         Top             =   1320
         Width           =   14325
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2910
            TabIndex        =   79
            Top             =   210
            Width           =   4785
            Begin VB.OptionButton optFim_sol 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               Height          =   255
               Left            =   2760
               TabIndex        =   20
               Top             =   180
               Width           =   1155
            End
            Begin VB.OptionButton Optinicio_Sol 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início frase"
               Height          =   255
               Left            =   180
               TabIndex        =   18
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton optMeio_Sol 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               Height          =   255
               Left            =   1470
               TabIndex        =   19
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optIgual_Sol 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               Height          =   255
               Left            =   3930
               TabIndex        =   21
               Top             =   180
               Width           =   705
            End
         End
         Begin VB.ComboBox cmbfiltrarpor_sol 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmCompras_ListaProduto.frx":2D294
            Left            =   180
            List            =   "frmCompras_ListaProduto.frx":2D2B0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2655
         End
         Begin VB.TextBox txtTexto_sol 
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
            Left            =   7770
            TabIndex        =   22
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6345
         End
         Begin MSComCtl2.DTPicker Txtprazo_sol 
            Height          =   315
            Left            =   7770
            TabIndex        =   24
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Visible         =   0   'False
            Width           =   6345
            _ExtentX        =   11192
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
            Format          =   143982593
            CurrentDate     =   39057
         End
         Begin VB.ComboBox cmbTexto_sol 
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
            ItemData        =   "frmCompras_ListaProduto.frx":2D31F
            Left            =   7770
            List            =   "frmCompras_ListaProduto.frx":2D321
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            ToolTipText     =   "Familia."
            Top             =   390
            Visible         =   0   'False
            Width           =   6345
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   10200
            TabIndex        =   58
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   1087
            TabIndex        =   57
            Top             =   180
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   23
         Left            =   -74935
         TabIndex        =   53
         Top             =   1830
         Width           =   14325
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2820
            TabIndex        =   78
            Top             =   210
            Width           =   4785
            Begin VB.OptionButton optIgual_necess 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               Height          =   255
               Left            =   3930
               TabIndex        =   6
               Top             =   180
               Width           =   705
            End
            Begin VB.OptionButton Optmeio_necess 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               Height          =   255
               Left            =   1470
               TabIndex        =   4
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton Optinicio_necess 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início frase"
               Height          =   255
               Left            =   180
               TabIndex        =   3
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Optfim_necess 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               Height          =   255
               Left            =   2760
               TabIndex        =   5
               Top             =   180
               Width           =   1155
            End
         End
         Begin VB.ComboBox Cmb_filtrar 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmCompras_ListaProduto.frx":2D323
            Left            =   11850
            List            =   "frmCompras_ListaProduto.frx":2D32D
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Tipo de necessidade."
            Top             =   390
            Width           =   2295
         End
         Begin VB.ComboBox cmbfiltrarpor_necess 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmCompras_ListaProduto.frx":2D357
            Left            =   180
            List            =   "frmCompras_ListaProduto.frx":2D36A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2565
         End
         Begin VB.TextBox txtTexto_necess 
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
            Left            =   7680
            TabIndex        =   7
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   4155
         End
         Begin VB.ComboBox cmbTexto_necess 
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
            ItemData        =   "frmCompras_ListaProduto.frx":2D3B5
            Left            =   7680
            List            =   "frmCompras_ListaProduto.frx":2D3B7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Familia."
            Top             =   390
            Visible         =   0   'False
            Width           =   4155
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de necessidade"
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
            Left            =   12142
            TabIndex        =   62
            Top             =   180
            Width           =   1710
         End
         Begin VB.Label Label1 
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
            Index           =   24
            Left            =   1042
            TabIndex        =   55
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   9015
            TabIndex        =   54
            Top             =   180
            Width           =   1485
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74940
         TabIndex        =   50
         Top             =   330
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   1720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   42
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Adicionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Adicionar selecionados (F3)"
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
         ButtonLeft2     =   46
         ButtonTop2      =   2
         ButtonWidth2    =   61
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonAlignment3=   2
         ButtonType3     =   1
         ButtonStyle3    =   -1
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   -1
         ButtonLeft3     =   109
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
         ButtonCaption4  =   "Ajuda"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Ajuda (F1)"
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
         ButtonLeft4     =   113
         ButtonTop4      =   2
         ButtonWidth4    =   41
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Sair"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Sair (Esc)"
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
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   188
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   5970
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_ListaProduto.frx":2D3B9
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74935
         TabIndex        =   52
         Top             =   330
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   1720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   42
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Adicionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Adicionar selecionados (F3)"
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
         ButtonLeft2     =   46
         ButtonTop2      =   2
         ButtonWidth2    =   61
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonAlignment3=   2
         ButtonType3     =   1
         ButtonStyle3    =   -1
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   -1
         ButtonLeft3     =   109
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
         ButtonCaption4  =   "Ajuda"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Ajuda (F1)"
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
         ButtonLeft4     =   113
         ButtonTop4      =   2
         ButtonWidth4    =   41
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Sair"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Sair (Esc)"
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
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   188
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   5970
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_ListaProduto.frx":3013C
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_solicitados 
         Height          =   5925
         Left            =   -74935
         TabIndex        =   25
         Top             =   2190
         Width           =   14325
         _ExtentX        =   25268
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nº solicitação"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   5353
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Quant. est."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Quant. com."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Detalhe"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "D"
            Text            =   "Prazo entr."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "Obs."
            Object.Width           =   0
         EndProperty
      End
      Begin DrawSuite2022.USFlexGrid FlexGrid 
         Height          =   5325
         Left            =   60
         TabIndex        =   83
         Top             =   2760
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   9393
         BackColorEvenRows=   14737632
         BackColorSelected1=   16643298
         BackColorSelected2=   16643298
         FocusRectColor  =   15181413
         GridColor       =   16247519
         HeaderGradientColor2=   12632256
         ProgressBarColor2=   2277891
         ForeColorSelected=   0
         AllowColumnResizing=   -1  'True
         CaptionHeight   =   28
         ColumnHeaderSmall=   -1  'True
         ColumnSort      =   -1  'True
         Editable        =   -1  'True
         FocusRowHighlightKeepTextForeColor=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFormatString=   $"frmCompras_ListaProduto.frx":32EBF
         MinRowHeight    =   14
         TotalLineShow   =   0   'False
      End
      Begin MSComctlLib.ListView ListaNecessidade 
         Height          =   5400
         Left            =   -74940
         TabIndex        =   84
         Top             =   2700
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   9525
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
         NumItems        =   6
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
            Text            =   "Descrição"
            Object.Width           =   16730
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Necessidade"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Necess. PÇ"
            Object.Width           =   2117
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   65
         TabIndex        =   85
         Top             =   330
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   1720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   42
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Adicionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Adicionar selecionados (F3)"
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
         ButtonLeft2     =   46
         ButtonTop2      =   2
         ButtonWidth2    =   61
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonAlignment3=   2
         ButtonType3     =   1
         ButtonStyle3    =   -1
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   -1
         ButtonLeft3     =   109
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
         ButtonCaption4  =   "Ajuda"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Ajuda (F1)"
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
         ButtonLeft4     =   113
         ButtonTop4      =   2
         ButtonWidth4    =   41
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Sair"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Sair (Esc)"
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
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   188
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   5970
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_ListaProduto.frx":32FF0
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frmCompras_ListaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSqlLoc_ProdServComprasNec As String
Dim StrSqlLoc_ProdServComprasSol As String
Dim StrSqlLoc_ProdServComprasCad As String

Private Sub Chk_fornecedor_Click()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_prazo_todos_Click()
On Error GoTo tratar_erro

FlexGrid.Clear
If Chk_prazo_todos.Value = 1 Then
    With Cmb_prazo
        .Enabled = True
        If Vendas_Proposta = False Then .SetFocus
    End With
Else
    With Cmb_prazo
        .Value = Date
        .Enabled = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_filtrar_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ordenar_Click()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_prazo_Change()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear
If cmbfiltrarpor_necess = "Família" Then
    txtTexto_necess.Visible = False
    cmbTexto_necess.Visible = True
    ProcCarregaComboFamilia cmbTexto_necess, "Compras = 'True' and familia <> 'Null'", True
Else
    txtTexto_necess.Visible = True
    cmbTexto_necess.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Solicitacao()
On Error GoTo tratar_erro

lblRegistros(1).Caption = "Nº de reg.: 0"
lblPaginas(1).Caption = "Página: 0 de: 0"
Lista_solicitados.ListItems.Clear
If StrSqlLoc_ProdServComprasSol = "" Then Exit Sub
Set TBLocalizar_produto_padrao1 = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao1.Open StrSqlLoc_ProdServComprasSol, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao1.EOF = False Then ProcExibePagina_Solicitacao (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Solicitacao(Pagina)
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear
TBLocalizar_produto_padrao1.PageSize = IIf(txtNreg(1) = "", 30, txtNreg(1))
TBLocalizar_produto_padrao1.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao1.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao1.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao1.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao1.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao1.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_solicitados.ListItems
        .Add , , TBLocalizar_produto_padrao1!IDlista
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao1!Status_Item), "", TBLocalizar_produto_padrao1!Status_Item)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao1!Requisicaotexto), "", TBLocalizar_produto_padrao1!Requisicaotexto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao1!Desenho), "", TBLocalizar_produto_padrao1!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao1!Descricao), "", TBLocalizar_produto_padrao1!Descricao)
        
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao1!Un), "", TBLocalizar_produto_padrao1!Un)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao1!Unidade_com), "", TBLocalizar_produto_padrao1!Unidade_com)
        If TBLocalizar_produto_padrao1!Un <> TBLocalizar_produto_padrao1!Unidade_com Then valor = FunConversaoFinalUn(TBLocalizar_produto_padrao1!Un, TBLocalizar_produto_padrao1!Unidade_com, TBLocalizar_produto_padrao1!quant_req, TBLocalizar_produto_padrao1!Desenho, True) Else valor = TBLocalizar_produto_padrao1!quant_req
        .Item(.Count).SubItems(7) = FunFormataCasasDecimais(4, valor)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_produto_padrao1!quant_req), "", FunFormataCasasDecimais(4, TBLocalizar_produto_padrao1!quant_req))
        
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao1!detalheitem), "", TBLocalizar_produto_padrao1!detalheitem)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLocalizar_produto_padrao1!prazoreq), "", Format(TBLocalizar_produto_padrao1!prazoreq, "dd/mm/yy"))
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLocalizar_produto_padrao1!Obs), "", TBLocalizar_produto_padrao1!Obs)
    End With
    ContadorReg = ContadorReg + 1
    TBLocalizar_produto_padrao1.MoveNext
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(1).Caption = "Nº de reg.: " & TBLocalizar_produto_padrao1.RecordCount
If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
   lblPaginas(1).Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao1.PageCount
ElseIf TBLocalizar_produto_padrao1.AbsolutePage = adPosEOF Then
        lblPaginas(1).Caption = "Pág.: " & TBLocalizar_produto_padrao1.PageCount & " de: " & TBLocalizar_produto_padrao1.PageCount
    Else
        lblPaginas(1).Caption = "Pág.: " & TBLocalizar_produto_padrao1.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao1.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAdicionar_Solicitacao()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
If Sit_Nota = 1 Then
    MsgTexto = "produto"
    MsgTexto1 = "Produto"
Else
    MsgTexto = "serviço"
    MsgTexto1 = "Serviço"
End If
With Lista_solicitados
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente adicionar este(s) " & MsgTexto & "(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    If USMsgBox("Algum " & MsgTexto & " selecionado será adicionado com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
                Else
                    Exit Sub
                End If
            End If
            Permitido = True
            IDlista = .ListItems.Item(InitFor)
            Desenho = .ListItems(InitFor).SubItems(3)
            If Permitido1 = True Then
                Compras_Pedido = True
                Vendas_PI = False
                Compras_Cotacao = False
                Faturamento = False
                Sit_Data = 2
                Permitido2 = True
                Sit_REG = 2
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then Exit Sub
            Else
                valor = .ListItems(InitFor).SubItems(8)
                frmCompras_Pedido.ProcAlterar_Solicitacao False
            End If
        End If
    Next InitFor
End With

If Permitido = False Then
    USMsgBox ("Informe o(s) " & MsgTexto & "(s) antes de adicionar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox (MsgTexto1 & " adicionado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    If Sit_Nota = 2 Then frmCompras_Pedido.ProcAtualizalistaServ Else frmCompras_Pedido.ProcAtualizalista
End If
ProcCarregalista_Solicitacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Solicitacao()
On Error GoTo tratar_erro

If Sit_Nota = 2 Then Tipo = "S" Else Tipo = "P"
CamposFiltro = "CR.ID_requisicao, CR.Requisicaotexto, CPL.IdLista, CPL.Status_Item, CPL.desenho, CPL.descricao, CPL.Un, CPL.Unidade_com, CPL.quant_req, CPL.detalheitem, CPL.prazoreq, CPL.obs"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (Compras_requisicao CR INNER JOIN Compras_pedido_lista CPL ON CR.ID_Requisicao = CPL.ID_Requisicao) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = CPL.codproduto"
TextoFiltroPadrao = "CPL.status_item = 'REQUISIT.' and CR.Status = 'LIBERADA' and CPL.Tipo = '" & Tipo & "' group by " & CamposFiltro & " order by CR.ID_requisicao"

If txtTexto_sol.Visible = True And txtTexto_sol <> "" Or cmbTexto_sol.Visible = True And cmbTexto_sol <> "" Or Txtprazo_sol.Visible = True Then
    If cmbfiltrarpor_sol = "Família" Then
        StrSqlLoc_ProdServComprasSol = INNERJOINTEXTO & " where CPL.Familia = '" & cmbTexto_sol & "' and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
            StrSqlLoc_ProdServComprasSol = INNERJOINTEXTO & " where CPL.Prazoreq = '" & Format(Txtprazo_sol.Value, "Short Date") & "' and " & TextoFiltroPadrao
        Else
            Select Case cmbfiltrarpor_sol
                Case "Solicitação": TextoFiltro = "CR.Requisicaotexto"
                Case "Código interno": TextoFiltro = "CPL.desenho"
                Case "Descrição": TextoFiltro = "CPL.descricao"
                Case "Descrição comercial": TextoFiltro = "CPL.descricao_comercial"
                Case "Detalhe": TextoFiltro = "CPL.Detalheitem"
                Case "Part number": TextoFiltro = "PFAB.Part_number"
            End Select
            StrSqlLoc_ProdServComprasSol = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_sol, optMeio_sol, optFim_sol, optIgual_sol, txtTexto_sol) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLoc_ProdServComprasSol = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Solicitacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear
Txtprazo_sol.Value = Date
If cmbfiltrarpor_sol = "Família" Then
    txtTexto_sol.Visible = False
    cmbTexto_sol.Visible = True
    Txtprazo_sol.Visible = False
    ProcCarregaComboFamilia cmbTexto_sol, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", True
ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
        txtTexto_sol.Visible = False
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = True
    Else
        txtTexto_sol.Visible = True
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Cadastrados()
On Error GoTo tratar_erro

If Sit_Nota = 1 Then
    varTipoProduto = "P.Tipo = 'P' and (P.Compras = 'True' or P.Producao = 'True')"
Else
    varTipoProduto = "P.Tipo = 'S' and P.Compras = 'True'"
End If
If Cmb_ordenar = "Código interno" Then Ordenar = "P.desenho" Else Ordenar = "P.Descricao"
CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.Unidade_com, P.Classe"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto) LEFT JOIN Projproduto_fornecedor PF ON PF.Codproduto = P.Codproduto) LEFT JOIN Compras_pedido_lista CPL ON CPL.Desenho = P.Desenho) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
TextoFiltroForn = ""
If Chk_fornecedor.Value = 1 Then
    IDPlano = frmCompras_Pedido.txtIDfornecedor
    TextoFiltroForn = "and PF.Idfornecedor = " & IDPlano
End If
TextoFiltroPadrao = varTipoProduto & " and P.bloqueado = 'False' and P.DtValidacao IS NOT NULL " & TextoFiltroForn & " group by " & CamposFiltro & " order by " & Ordenar

If txtTexto_cad.Visible = True And txtTexto_cad <> "" Or cmbTexto_cad.Visible = True And cmbTexto_cad <> "" Then
    If cmbfiltrarpor_cad = "Fornecedor" Then
        StrSqlLoc_ProdServComprasCad = INNERJOINTEXTO & " where PF.IDfornecedor = " & cmbTexto_cad.ItemData(cmbTexto_cad.ListIndex) & " and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor_cad = "Família" Then
            StrSqlLoc_ProdServComprasCad = INNERJOINTEXTO & " where P.classe = '" & cmbTexto_cad & "' and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor_cad = "Comprimento" Or cmbfiltrarpor_cad = "Largura" Or cmbfiltrarpor_cad = "Espessura" Then
                Select Case cmbfiltrarpor_cad
                    Case "Comprimento": TextoFiltro = "P.Comprimento"
                    Case "Largura": TextoFiltro = "P.Largura"
                    Case "Espessura": TextoFiltro = "P.Espessura"
                End Select
                valor = txtTexto_cad
                NovoValor = Replace(valor, ",", ".")
                StrSqlLoc_ProdServComprasCad = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
            Else
                Select Case cmbfiltrarpor_cad
                    Case "Código interno": TextoFiltro = "P.desenho"
                    Case "Descrição": TextoFiltro = "P.descricao"
                    Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                    Case "Código de referência": TextoFiltro = "IA.N_referencia"
                    Case "Part number": TextoFiltro = "PFAB.Part_number"
                End Select
                StrSqlLoc_ProdServComprasCad = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_cad, optMeio_cad, optFim_cad, optIgual_cad, txtTexto_cad) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLoc_ProdServComprasCad = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Cadastrados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Cadastrados()
On Error GoTo tratar_erro

lblRegistros(2).Caption = "Nº de reg.: 0"
lblPaginas(2).Caption = "Página: 0 de: 0"
FlexGrid.Clear
If StrSqlLoc_ProdServComprasCad = "" Then Exit Sub
Set TBLocalizar_produto_padrao2 = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao2.Open StrSqlLoc_ProdServComprasCad, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao2.EOF = False Then ProcExibePagina_Cadastrados (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_cad_Click()
On Error GoTo tratar_erro

FlexGrid.Clear
With cmbTexto_cad
    If cmbfiltrarpor_cad = "Família" Or cmbfiltrarpor_cad = "Fornecedor" Then
        txtTexto_cad.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor_cad = "Família" Then
            If Sit_Nota = 1 Then ProcCarregaComboFamilia cmbTexto_cad, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", True Else ProcCarregaComboFamilia cmbTexto_cad, "familia <> 'Null' and Compras = 'True'", True
        ElseIf cmbfiltrarpor_cad = "Fornecedor" Then
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select IDCliente, Nome_Razao from Compras_fornecedores where Nome_Razao <> 'Null' order by Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    Do While TBFornecedor.EOF = False
                        .AddItem Trim(TBFornecedor!Nome_Razao)
                        .ItemData(.NewIndex) = TBFornecedor!IDCliente
                        TBFornecedor.MoveNext
                    Loop
                    .Text = Trim(frmCompras_Pedido.txtFornecedor)
                End If
                TBFornecedor.Close
        End If
    Else
        txtTexto_cad.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_necess_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_cad_Click()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
            If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
                ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.PageCount - 1)
            Else
                TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
                ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
            End If
        Else
            ProcExibePagina_Necessidade (1)
        End If
    Case 1:
        If TBLocalizar_produto_padrao1.AbsolutePage <> 2 Then
            If TBLocalizar_produto_padrao1.AbsolutePage = -3 Then
                ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao1.PageCount - 1)
            Else
                TBLocalizar_produto_padrao1.AbsolutePage = TBLocalizar_produto_padrao1.AbsolutePage - 2
                ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao1.AbsolutePage)
            End If
        Else
            ProcExibePagina_Solicitacao (1)
        End If
    Case 2:
        If TBLocalizar_produto_padrao2.AbsolutePage <> 2 Then
            If TBLocalizar_produto_padrao2.AbsolutePage = -3 Then
                ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.PageCount - 1)
            Else
                TBLocalizar_produto_padrao2.AbsolutePage = TBLocalizar_produto_padrao2.AbsolutePage - 2
                ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
            End If
        Else
            ProcExibePagina_Cadastrados (1)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4))
If Quant <= 1 Or txtPagIr(index) > Quant Then Exit Sub
If txtPagIr(index).Text >= 1 And txtPagIr(index).Text <= Quant Then
    Select Case SSTab1.Tab
        Case 0:
            TBLocalizar_produto_padrao.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
        Case 1:
            TBLocalizar_produto_padrao1.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao1.AbsolutePage)
        Case 2:
            TBLocalizar_produto_padrao2.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        TBLocalizar_produto_padrao.AbsolutePage = 1
        ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
    Case 1:
        TBLocalizar_produto_padrao1.AbsolutePage = 1
        ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao1.AbsolutePage)
    Case 2:
        TBLocalizar_produto_padrao2.AbsolutePage = 1
        ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
            If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
                ProcExibePagina_Necessidade (2)
            Else
                ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
            End If
        Else
            ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.PageCount)
        End If
    Case 1:
        If TBLocalizar_produto_padrao1.AbsolutePage <> -3 Then
            If TBLocalizar_produto_padrao1.AbsolutePage = 1 Then
                ProcExibePagina_Solicitacao (2)
            Else
                ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao1.AbsolutePage)
            End If
        Else
            ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao1.PageCount)
        End If
    Case 2:
        If TBLocalizar_produto_padrao2.AbsolutePage <> -3 Then
            If TBLocalizar_produto_padrao2.AbsolutePage = 1 Then
                ProcExibePagina_Cadastrados (2)
            Else
                ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
            End If
        Else
            ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.PageCount)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
        ProcExibePagina_Necessidade (TBLocalizar_produto_padrao.AbsolutePage)
    Case 1:
        TBLocalizar_produto_padrao1.AbsolutePage = TBLocalizar_produto_padrao1.PageCount
        ProcExibePagina_Solicitacao (TBLocalizar_produto_padrao1.AbsolutePage)
    Case 2:
        TBLocalizar_produto_padrao2.AbsolutePage = TBLocalizar_produto_padrao2.PageCount
        ProcExibePagina_Cadastrados (TBLocalizar_produto_padrao2.AbsolutePage)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long, vNewValue As String, Cancel As Boolean)
On Error GoTo tratar_erro

With FlexGrid
    If Col = 7 Then vNewValue = Format(vNewValue, "DD/MM/YY")
    If Col = 8 Then
        If vNewValue <> "" Then
            valor = vNewValue
            vNewValue = FunFormataCasasDecimais(4, valor)
            
            If .CellText(Row, 4) <> .CellText(Row, 5) Then valor = FunConversaoFinalUn(.CellText(Row, 4), .CellText(Row, 5), valor, .CellText(Row, 2), True)
            .CellText(Row, 9) = FunFormataCasasDecimais(4, valor)
        Else
            .CellText(Row, 9) = ""
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGrid_CellClick(ByVal Row As Long, ByVal Col As Long, Shift As Integer)
On Error GoTo tratar_erro

Permitido = True
With FlexGrid
    If Col = 1 And .CellChecked(Row, 1) = True Then
        'Verifica se é obrigatório ter cotação valida
        If FunVerifCotacaoValida(frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), .CellText(Row, 2), IIf(Sit_Nota = 1, True, False), True, "selecionar", frmCompras_Pedido.txtIDfornecedor) = False Then
            .CellChecked(Row, 1) = False
            Exit Sub
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGrid_ColumnClick(ByVal Col As Long)
On Error GoTo tratar_erro

'If Col = 1 And Permitido = False Then
'    With FlexGrid
'        For InitFor = 0 To (.Rows + 1)
'            If .CellChecked(InitFor, 1) = True Then
'                .CellChecked(InitFor, 1) = False
'            Else
'                If FlexGrid.CellText(InitFor, 7) = "" Then GoTo Proximo
'                If FlexGrid.CellText(InitFor, 8) = "" Then GoTo Proximo
'
'            .CellChecked(InitFor, 1) = True
'Proximo:
'            End If
'        Next InitFor
'    End With
'End If
'Permitido = False

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
            Case vbKeyF2: ProcFiltrar_Necessidade
            Case vbKeyF3: ProcAdicionar_Necessidade
            Case vbKeyEscape: Unload Me
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar_Solicitacao
            Case vbKeyF3: ProcAdicionar_Solicitacao
            Case vbKeyEscape: Unload Me
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrar_Cadastrados
            Case vbKeyF3: procAdicionar_Cadastrados
            Case vbKeyEscape: Unload Me
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 14325, 6, True
ProcCarregaToolBar3 Me, 14325, 6, True
ProcCarregaToolBar2 Me, 14325, 6, True
Cmb_prazo.Value = Date
With SSTab1
    If Sit_Nota = 1 Then
        Caption = Caption & " produtos"
        .Tab = 0
        
        ProcFiltroPadrao cmbfiltrarpor_necess, Optmeio_necess, Optfim_necess, optIgual_necess, frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), "Produtos/Serviços", "C", True
        If Permitido = False Then cmbfiltrarpor_necess = "Código interno"
        Cmb_filtrar = "Com necessidade"
    Else
        Caption = Caption & " serviços"
        .TabVisible(0) = False
        .TabsPerRow = 2
        .Tab = 1
    End If
End With

ProcFiltroPadrao cmbfiltrarpor_sol, optMeio_sol, optFim_sol, optIgual_sol, frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), "Produtos/Serviços", "C", True
ProcFiltroPadrao cmbfiltrarpor_cad, optMeio_cad, optFim_cad, optIgual_cad, frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), "Produtos/Serviços", "C", True
If Permitido = False Then
    cmbfiltrarpor_necess = "Código interno"
    cmbfiltrarpor_sol = "Código interno"
    cmbfiltrarpor_cad = "Código interno"
End If

Cmb_ordenar = "Código interno"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_solicitados
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select IDlista from Compras_pedido_lista where idlista = " & .ListItems.Item(InitFor) & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then Permitido = True Else Permitido = False
                TBLISTA.Close
                
                'Verifica se é obrigatório ter cotação valida
                If FunVerifCotacaoValida(frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), .ListItems.Item(InitFor).ListSubItems(3), Permitido, False, "", frmCompras_Pedido.txtIDfornecedor) = False Then GoTo Proximo
                
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_solicitados, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitados_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_solicitados
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select IDlista from Compras_pedido_lista where idlista = " & .ListItems.Item(InitFor) & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then Permitido = True Else Permitido = False
            TBLISTA.Close
            
            'Verifica se é obrigatório ter cotação valida
            If FunVerifCotacaoValida(frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), .ListItems.Item(InitFor).ListSubItems(3), Permitido, True, "selecionar", frmCompras_Pedido.txtIDfornecedor) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_PCP_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_vendas_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptFim_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_cad_Click()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_cad_Click()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_cad_Click()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_necess_Click()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_sol_Click()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_cad_Click()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0: If cmbfiltrarpor_necess.Visible = True Then cmbfiltrarpor_necess.SetFocus
    Case 1: If cmbfiltrarpor_sol.Visible = True Then cmbfiltrarpor_sol.SetFocus
    Case 2: cmbfiltrarpor_cad.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtprazo_sol_Change()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change(index As Integer)
On Error GoTo tratar_erro

If txtNreg(index) <> "" Then
    VerifNumero = txtNreg(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg(index) = ""
        txtNreg(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) <> "" Then
    VerifNumero = txtPagIr(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr(index) = ""
        txtPagIr(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_necess_Change()
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAdicionar_Necessidade()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With ListaNecessidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente adicionar este(s) produto(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    If USMsgBox("Algum produto selecionado será adicionado com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
                Else
                    Exit Sub
                End If
            End If
            Permitido = True
            IDlista = .ListItems.Item(InitFor)
            Desenho = .ListItems(InitFor).SubItems(1)
            If Permitido1 = True Then
                Qtde = .ListItems(InitFor).SubItems(4)
                
                Compras_Pedido = True
                Vendas_PI = False
                Compras_Cotacao = False
                Faturamento = False
                Sit_Data = 1
                Permitido2 = True
                Sit_REG = 2
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then Exit Sub
            Else
                valor = .ListItems(InitFor).SubItems(4)
                frmCompras_Pedido.ProcNovo_Necessidade Opt_vendas
            End If
        End If
    Next InitFor
End With

If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de adicionar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) adicionado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    frmCompras_Pedido.ProcAtualizalista
End If
ProcCarregalista_Necessidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_sol_Change()
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_cad_Change()
On Error GoTo tratar_erro

FlexGrid.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar_Necessidade
    Case 2: ProcAdicionar_Necessidade
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar_Cadastrados
    Case 2: procAdicionar_Cadastrados
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar_Solicitacao
    Case 2: ProcAdicionar_Solicitacao
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Necessidade()
On Error GoTo tratar_erro

If Opt_PCP.Value = True Then NomeTabela = "Estoque_necessidade_resumido" Else NomeTabela = "Estoque_necessidade_resumido_PIEST"
CamposFiltro = "ENR.Codproduto, ENR.Desenho, ENR.Descricao, ENR.Unidade, ENR.Unidade_Com, ENR.Necessidade, ENR.Necessidade_estoque, ENR.Necessidade_PC"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (" & NomeTabela & " ENR LEFT JOIN item_aplicacoes IA ON ENR.codproduto = IA.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = ENR.codproduto"
If Cmb_filtrar = "Com necessidade" Then TextoFiltroEstoque = " and ENR.Necessidade > 0" Else TextoFiltroEstoque = " and ENR.Necessidade_estoque > 0"
TextoFiltroPadrao = "ENR.Compras = 'True' and ENR.ID_empresa = " & frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex) & TextoFiltroEstoque & " group by " & CamposFiltro & " order by ENR.desenho"

If txtTexto_necess.Visible = True And txtTexto_necess <> "" Or cmbTexto_necess.Visible = True And cmbTexto_necess <> "" Then
    If cmbfiltrarpor_necess = "Família" Then
        StrSqlLoc_ProdServComprasNec = INNERJOINTEXTO & " where ENR.classe = '" & cmbTexto_necess & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor_necess
            Case "Código interno": TextoFiltro = "ENR.Desenho"
            Case "Código de referência": TextoFiltro = "IA.n_referencia"
            Case "Descrição": TextoFiltro = "ENR.Descricao"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
        End Select
        StrSqlLoc_ProdServComprasNec = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_necess, Optmeio_necess, Optfim_necess, optInicio_necess, txtTexto_necess) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLoc_ProdServComprasNec = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Necessidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Necessidade()
On Error GoTo tratar_erro

If StrSqlLoc_ProdServComprasNec = "" Then Exit Sub
lblRegistros(0).Caption = "Nº de registros: 0"
lblPaginas(0).Caption = "Página: 0 de: 0"
ListaNecessidade.ListItems.Clear
If StrSqlLoc_ProdServComprasNec = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLoc_ProdServComprasNec, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina_Necessidade (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExibePagina_Cadastrados(Pagina)
On Error GoTo tratar_erro
Dim L As Long

FlexGrid.Clear
TBLocalizar_produto_padrao2.PageSize = IIf(txtNreg(2) = "", 30, txtNreg(2))

TBLocalizar_produto_padrao2.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao2.PageSize
ContadorReg = 1

Contador = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao2.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao2.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao2.PageSize)
PBLista.Value = 1
Contador2 = 0

Do While TBLocalizar_produto_padrao2.EOF = False And (ContadorReg <= TamanhoPagina)
    With FlexGrid
        L = .AddItem(TBLocalizar_produto_padrao2!Codproduto)
        .RowData(L) = Contador
        .CellText(L, 2) = IIf(IsNull(TBLocalizar_produto_padrao2!Desenho), "", TBLocalizar_produto_padrao2!Desenho)
        .CellText(L, 3) = IIf(IsNull(TBLocalizar_produto_padrao2!Descricao), "", TBLocalizar_produto_padrao2!Descricao)
        .CellText(L, 4) = IIf(IsNull(TBLocalizar_produto_padrao2!Unidade), "", TBLocalizar_produto_padrao2!Unidade)
        .CellText(L, 5) = IIf(IsNull(TBLocalizar_produto_padrao2!Unidade_com), "", TBLocalizar_produto_padrao2!Unidade_com)
        .CellText(L, 6) = IIf(IsNull(TBLocalizar_produto_padrao2!Classe), "", TBLocalizar_produto_padrao2!Classe)

        If Chk_prazo_todos.Value = 1 Then .CellText(L, 7) = Format(Cmb_prazo, "dd/mm/yy")
    End With
    Contador = Contador + 1
    TBLocalizar_produto_padrao2.MoveNext
    ContadorReg = ContadorReg + 1
    Contador2 = Contador2 + 1
    PBLista.Value = Contador2
Loop
lblRegistros(2).Caption = "Nº de registros: " & TBLocalizar_produto_padrao2.RecordCount
If TBLocalizar_produto_padrao2.AbsolutePage = adPosBOF Then
   lblPaginas(2).Caption = "Página: 1 de: " & TBLocalizar_produto_padrao2.PageCount
ElseIf TBLocalizar_produto_padrao2.AbsolutePage = adPosEOF Then
        lblPaginas(2).Caption = "Página: " & TBLocalizar_produto_padrao2.PageCount & " de: " & TBLocalizar_produto_padrao2.PageCount
    Else
        lblPaginas(2).Caption = "Página: " & TBLocalizar_produto_padrao2.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao2.PageCount
End If
FlexGrid.Redraw = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNecessidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaNecessidade
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                'Verifica se é obrigatório ter cotação valida
                If FunVerifCotacaoValida(frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), .ListItems.Item(InitFor).ListSubItems(1), False, False, "", frmCompras_Pedido.txtIDfornecedor) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaNecessidade, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNecessidade_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaNecessidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            'Verifica se é obrigatório ter cotação valida
            If FunVerifCotacaoValida(frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), .ListItems.Item(InitFor).ListSubItems(1), False, True, "selecionar", frmCompras_Pedido.txtIDfornecedor) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Necessidade(Pagina)
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg(0) = "", 30, txtNreg(0))
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaNecessidade.ListItems.Add(, , TBLocalizar_produto_padrao!Codproduto)
        .SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
        .SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
        .SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
        If TBLocalizar_produto_padrao!Unidade <> TBLocalizar_produto_padrao!Unidade_com Then
            If Cmb_filtrar = "Com necessidade" Then qt = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000") Else qt = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000")
            If FunVerifUNConversao(TBLocalizar_produto_padrao!Unidade, TBLocalizar_produto_padrao!Unidade_com) = True Then
                Qtde = FunConverteUN(TBLocalizar_produto_padrao!Unidade_com, TBLocalizar_produto_padrao!Unidade, qt, TBLocalizar_produto_padrao!Desenho)
                .SubItems(4) = Format(Qtde, "###,##0.0000")
            Else
                If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000")
            End If
        Else
            If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000")
        End If
        .SubItems(5) = TBLocalizar_produto_padrao!Necessidade_PC
        If Cmb_filtrar = "Com necess. estoque" Then NReal = Format(TBLocalizar_produto_padrao!Necessidade_estoque, "###,##0.0000") Else NReal = Format(TBLocalizar_produto_padrao!Necessidade, "###,##0.0000")
        If NReal > 0 Then
            .ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbRed
            .ListSubItems(2).ForeColor = vbRed
            .ListSubItems(3).ForeColor = vbRed
            .ListSubItems(4).ForeColor = vbRed
            .ListSubItems(5).ForeColor = vbRed
        End If
    End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(0).Caption = "Nº de registros: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas(0).Caption = "Página: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas(0).Caption = "Página: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas(0).Caption = "Página: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAdicionar_Cadastrados()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido2 = False
If Sit_Nota = 1 Then
    MsgTexto = "produto"
    MsgTexto1 = "Produto"
Else
    MsgTexto = "serviço"
    MsgTexto1 = "Serviço"
End If

Contador = 0
Permitido = False
With FlexGrid
    For InitFor = 1 To (.rows)
        If .CellChecked(Contador, 1) = True Then
            If .CellText(Contador, 7) = "" Then Permitido = True
            If .CellText(Contador, 8) = "" Then Permitido = True
        End If
        Contador = Contador + 1
    Next InitFor
End With
If Permitido = True Then
    USMsgBox ("Informe o prazo e quantidade para o(s) " & IIf(Sit_Nota = 1, "produto(s)", "serviço(s)") & " selecionado(s) e clique na lista para concluir a alteração antes de adicionar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Contador = 0
With FlexGrid
    For InitFor = 1 To (.rows)
        If .CellChecked(Contador, 1) = True Then
            If Permitido2 = False Then
                If USMsgBox("Deseja realmente adicionar este(s) " & MsgTexto & "(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            
            Permitido2 = True
            procCriarItem .CellText(Contador, 0), .CellText(Contador, 8), .CellText(Contador, 7)
        End If
        Contador = Contador + 1
    Next InitFor
End With

If Permitido2 = False Then
    USMsgBox ("Informe o(s) " & MsgTexto & "(s) antes de adicionar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    USMsgBox (MsgTexto1 & " adicionado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With frmCompras_Pedido
        If Sit_Nota = 1 Then
            .ProcAtualizalista
            .Novo_PC1 = False
        Else
            .ProcAtualizalistaServ
            .Novo_PC2 = False
        End If
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCriarItem(Codprod_adicionar As Long, Quantidade_adicionar As Double, Prazo_adicionar As String)
On Error GoTo tratar_erro

With frmCompras_Pedido
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    TBGravar!IDpedido = .txtIDPedido
    
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where Codproduto = " & Codprod_adicionar, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        TBGravar!Codproduto = IIf(IsNull(TBItem!Codproduto), "", TBItem!Codproduto)
        TBGravar!Desenho = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
        TBGravar!Descricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
        TBGravar!Descricao_comercial = IIf(IsNull(TBItem!descricaotecnica), "", TBItem!descricaotecnica)
        TBGravar!Un = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
        TBGravar!Unidade_com = IIf(IsNull(TBItem!Unidade_com), "", TBItem!Unidade_com)
        TBGravar!Familia = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select PCusto from Projproduto_fornecedor where Codproduto = " & TBItem!Codproduto & " and idfornecedor = " & .txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TBGravar!preco_unitario = IIf(IsNull(TBFI!PCusto), "", Format(TBFI!PCusto, "###,##0.0000000000"))
        Else
            TBGravar!preco_unitario = IIf(IsNull(TBItem!PCusto), 0, TBItem!PCusto)
        End If
        TBFI.Close
        TBGravar!preco_unitario_desconto = TBGravar!preco_unitario
        
        If TBItem!Tipo = "S" Then
            If IsNull(TBGravar!ID_CFOP) = False Then TBGravar!ID_CFOP = TBItem!ID_CFOP
            
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Simples, presumido, Real, Simples1 from Empresa where Codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Regime = 0
                If TBFIltro!Simples = True Then Regime = 1
                If TBFIltro!Presumido = True Then Regime = 2
                If TBFIltro!Real = True Then Regime = 3
                If TBFIltro!Simples1 = True Then Regime = 4
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ISS from Impostos where Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBGravar!ISSQN = IIf(IsNull(TBFI!ISS), 0, TBFI!ISS)
                End If
                TBFI.Close
            End If
            
            TBGravar!VlrISSQN = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!ISSQN), 0, TBGravar!ISSQN)) / 100, "###,##0.00")
            TBGravar!Tipo = "S"
        Else
            TBGravar!Quant_Comp_PC = FunCalculaQtdePC(TBItem!Desenho, Quantidade_adicionar, True, TBGravar!Unidade_com)
            If IsNull(TBItem!ID_CF) = False Then TBGravar!ID_CF = TBItem!ID_CF
            
            If IsNull(TBItem!ID_CFOP) = False Then TBGravar!ID_CFOP = TBItem!ID_CFOP
            If IsNull(TBItem!ID_CF) = False Then
                ProcValorImposto .txtPedido, IIf(TBItem!ID_CF = "", 0, TBItem!ID_CF), IIf(.txtIDfornecedor = "", 0, .txtIDfornecedor), .txtFornecedor, .txtuf, .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), True, IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP), 0
                ProcControleImposto IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP), IIf(.txtIDfornecedor = "", 0, .txtIDfornecedor)
                If TemIPI = "SIM" Then TBGravar!IPI = IntIPI Else TBGravar!IPI = 0
                If TemICMS = "SIM" Then TBGravar!ICMS = IntICMS Else TBGravar!ICMS = 0
            End If
            
            TBGravar!VlrIPI = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!IPI), 0, TBGravar!IPI)) / 100, "###,##0.00")
            TBGravar!vlrICMS = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!ICMS), 0, TBGravar!ICMS)) / 100, "###,##0.00")
            TBGravar!Tipo = "P"
        End If
    End If
    
    TBGravar!Prazo = Prazo_adicionar
    TBGravar!Quant_Comp = Quantidade_adicionar
    TBGravar!Status_Item = "AGUARDANDO APROVAÇÃO"
    TBGravar!ValorDesconto = 0
    TBGravar!Desconto = 0
    
    TBGravar!preco_total = Format((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) + IIf(IsNull(TBGravar!VlrIPI), 0, TBGravar!VlrIPI), "###,##0.00")

    TBItem.Close
    TBGravar.Update
    
    '==================================
    Modulo = "Compras/Pedido"
    Evento = "Novo produto"
    ID_documento = TBGravar!IDlista
    Documento = "Nº pedido: " & .txtPedido
    Documento1 = "Cód. interno: " & TBGravar!Desenho
    ProcGravaEvento
    '==================================
    
    TBGravar.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
