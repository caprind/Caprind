VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_programacao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Programação"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
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
      Left            =   60
      TabIndex        =   93
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
      ItemData        =   "frmVendas_programacao.frx":0000
      Left            =   240
      List            =   "frmVendas_programacao.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1680
      Width           =   6645
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17701
      _Version        =   393216
      Tab             =   1
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
      TabCaption(0)   =   "Dados principais"
      TabPicture(0)   =   "frmVendas_programacao.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Framelista"
      Tab(0).Control(1)=   "txtID"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "USImageList1"
      Tab(0).Control(4)=   "lista"
      Tab(0).Control(5)=   "USToolBar1"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Produtos"
      TabPicture(1)   =   "frmVendas_programacao.frx":0020
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lista_item"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "USImageList2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtID_item"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Programação"
      TabPicture(2)   =   "frmVendas_programacao.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "txtID_prog"
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(3)=   "Frame10"
      Tab(2).Control(4)=   "Frame8"
      Tab(2).Control(5)=   "USImageList3"
      Tab(2).Control(6)=   "USToolBar3"
      Tab(2).ControlCount=   7
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   75
         TabIndex        =   109
         Top             =   2160
         Width           =   15200
         Begin VB.TextBox Txt_texto 
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
            Left            =   5010
            MaxLength       =   255
            TabIndex        =   26
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   8895
         End
         Begin VB.ComboBox Cmb_filtrarpor 
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
            ItemData        =   "frmVendas_programacao.frx":0058
            Left            =   180
            List            =   "frmVendas_programacao.frx":0068
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   4815
         End
         Begin VB.ComboBox Cmb_texto 
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
            ItemData        =   "frmVendas_programacao.frx":00A5
            Left            =   5010
            List            =   "frmVendas_programacao.frx":00B5
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   8895
         End
         Begin DrawSuite2022.USButton Cmd_filtrar 
            Height          =   330
            Left            =   13980
            TabIndex        =   27
            Top             =   390
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   582
            Caption         =   "Filtrar (F2)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   16777215
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
         End
         Begin VB.Label Label14 
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
            Left            =   8962
            TabIndex        =   111
            Top             =   180
            Width           =   1470
         End
         Begin VB.Label Label11 
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
            Left            =   2167
            TabIndex        =   110
            Top             =   180
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   104
         Top             =   9090
         Width           =   15195
         Begin VB.TextBox txtNreg1 
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
         Begin VB.TextBox txtPagIr1 
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
         Begin DrawSuite2022.USButton cmdPagProx1 
            Height          =   315
            Left            =   11760
            TabIndex        =   35
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":00E5
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
         Begin DrawSuite2022.USButton cmdPagAnt1 
            Height          =   315
            Left            =   11220
            TabIndex        =   34
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":3889
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
         Begin DrawSuite2022.USButton cmdPagIr1 
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
         Begin DrawSuite2022.USButton cmdPagPrim1 
            Height          =   315
            Left            =   10680
            TabIndex        =   33
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":7392
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
         Begin DrawSuite2022.USButton cmdPagUlt1 
            Height          =   315
            Left            =   12300
            TabIndex        =   36
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":B481
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
         Begin VB.Label Label34 
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
            TabIndex        =   115
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label12 
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
            TabIndex        =   107
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros1 
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
            TabIndex        =   106
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas1 
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
            TabIndex        =   105
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   98
         Top             =   9120
         Width           =   11565
         Begin VB.TextBox Txt_qtde_total_estoque 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade total estoque."
            Top             =   390
            Width           =   1725
         End
         Begin VB.TextBox Txt_saldo 
            Alignment       =   1  'Right Justify
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
            Left            =   7170
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Saldo da revisão."
            Top             =   390
            Width           =   1695
         End
         Begin VB.TextBox Txt_qtde_total_vendida 
            Alignment       =   1  'Right Justify
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
            Left            =   3674
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade total vendida da revisão."
            Top             =   390
            Width           =   1725
         End
         Begin VB.TextBox Txt_qtde_total_prev 
            Alignment       =   1  'Right Justify
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
            Left            =   1927
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade total prevista da revisão."
            Top             =   390
            Width           =   1735
         End
         Begin VB.TextBox Txt_qtde_total_faturada_rev 
            Alignment       =   1  'Right Justify
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
            Left            =   5421
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade total faturada da revisão."
            Top             =   390
            Width           =   1745
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. total est."
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
            Left            =   412
            TabIndex        =   103
            Top             =   180
            Width           =   1260
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Left            =   7785
            TabIndex        =   102
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. total vend."
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
            Left            =   3831
            TabIndex        =   101
            Top             =   180
            Width           =   1410
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. total prev."
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
            Left            =   2104
            TabIndex        =   100
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. total fatur."
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
            Left            =   5588
            TabIndex        =   99
            Top             =   180
            Width           =   1410
         End
      End
      Begin VB.TextBox txtID_prog 
         Height          =   285
         Left            =   -72150
         TabIndex        =   96
         Text            =   "0"
         Top             =   4830
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1455
         Left            =   -74925
         TabIndex        =   82
         Top             =   1335
         Width           =   15195
         Begin VB.CheckBox Chk_utiliza_mat_consignado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Utiliza material consignado?"
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
            TabIndex        =   39
            Top             =   450
            Width           =   2415
         End
         Begin VB.TextBox Txt_n_item 
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
            Left            =   7860
            MaxLength       =   6
            TabIndex        =   46
            ToolTipText     =   "Número do produto/item no pedido do cliente."
            Top             =   990
            Width           =   735
         End
         Begin VB.TextBox Txt_pedido_cliente 
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
            Left            =   5790
            MaxLength       =   60
            TabIndex        =   45
            ToolTipText     =   "Pedido do cliente."
            Top             =   990
            Width           =   2055
         End
         Begin VB.TextBox Txt_un 
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
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Unidade de estoque."
            Top             =   990
            Width           =   675
         End
         Begin VB.TextBox Txt_un_com 
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
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Unidade comercial."
            Top             =   990
            Width           =   675
         End
         Begin VB.TextBox txtCodigo2 
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
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Código do item."
            Top             =   390
            Width           =   1665
         End
         Begin VB.CheckBox optFirme 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Venda confirmada"
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
            Height          =   255
            Left            =   13380
            TabIndex        =   49
            Top             =   1020
            Width           =   1605
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
            Left            =   3090
            TabIndex        =   42
            ToolTipText     =   "Quantidade."
            Top             =   990
            Width           =   1305
         End
         Begin VB.TextBox txtDescricao2 
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
            Left            =   1860
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do item."
            Top             =   390
            Width           =   10815
         End
         Begin VB.TextBox txtStatus_prog 
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
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Status da programação."
            Top             =   990
            Width           =   3195
         End
         Begin MSComCtl2.DTPicker txtData_inicio 
            Height          =   315
            Left            =   180
            TabIndex        =   40
            ToolTipText     =   "Inicio do prazo."
            Top             =   990
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   197525505
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtData_fim 
            Height          =   315
            Left            =   1770
            TabIndex        =   41
            ToolTipText     =   "Final do prazo."
            Top             =   990
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   197525505
            CurrentDate     =   39057
         End
         Begin MSMask.MaskEdBox Txt_data_negociada 
            Height          =   315
            Left            =   8610
            TabIndex        =   47
            ToolTipText     =   "Data negociada."
            Top             =   990
            Width           =   1095
            _ExtentX        =   1931
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
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N. item"
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
            Left            =   7972
            TabIndex        =   113
            Top             =   780
            Width           =   510
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. negociada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8655
            TabIndex        =   112
            Top             =   780
            Width           =   1005
         End
         Begin VB.Image imgCalendario 
            Height          =   360
            Left            =   9705
            Picture         =   "frmVendas_programacao.frx":ED0D
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   960
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pedido do cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6210
            TabIndex        =   97
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. est."
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
            Left            =   4455
            TabIndex        =   95
            Top             =   780
            Width           =   585
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. com."
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
            Left            =   5115
            TabIndex        =   94
            Top             =   780
            Width           =   645
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "á"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1590
            TabIndex        =   89
            Top             =   1050
            Width           =   90
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3322
            TabIndex        =   88
            Top             =   780
            Width           =   840
         End
         Begin VB.Label Label13 
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
            Height          =   180
            Left            =   6922
            TabIndex        =   87
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   502
            TabIndex        =   86
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio do prazo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   322
            TabIndex        =   85
            Top             =   780
            Width           =   1050
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final do prazo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1920
            TabIndex        =   84
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
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
            Left            =   11400
            TabIndex        =   83
            Top             =   780
            Width           =   555
         End
      End
      Begin VB.Frame Framelista 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1485
         Left            =   -74925
         TabIndex        =   73
         Top             =   1305
         Width           =   15195
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Cliente."
            Top             =   1020
            Width           =   11535
         End
         Begin VB.TextBox txtPrograma 
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
            Left            =   11740
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Numero do programa."
            Top             =   390
            Width           =   1465
         End
         Begin VB.CommandButton cmdpesquisar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmVendas_programacao.frx":F190
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Localizar cliente."
            Top             =   1020
            Width           =   315
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1215
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8065
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3660
         End
         Begin VB.TextBox txtStatus 
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
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   1020
            Width           =   2115
         End
         Begin VB.TextBox txtID_cli 
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
            Left            =   2310
            TabIndex        =   7
            ToolTipText     =   "Código do cliente."
            Top             =   1020
            Width           =   795
         End
         Begin VB.TextBox txtRev 
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
            Left            =   13225
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Número de revisão da programação."
            Top             =   390
            Width           =   525
         End
         Begin VB.TextBox txtData_rev 
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
            Left            =   13770
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Data da revisão."
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   8640
            TabIndex        =   81
            Top             =   810
            Width           =   495
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
            Left            =   7275
            TabIndex        =   80
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº programa"
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
            Left            =   11932
            TabIndex        =   79
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label Label25 
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
            Left            =   9438
            TabIndex        =   78
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
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
            Left            =   960
            TabIndex        =   77
            Top             =   810
            Width           =   555
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   13300
            TabIndex        =   76
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data rev."
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
            Left            =   13987
            TabIndex        =   75
            Top             =   180
            Width           =   780
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
            Left            =   3210
            TabIndex        =   74
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6315
         Left            =   -74925
         TabIndex        =   67
         Top             =   2790
         Width           =   11565
         Begin DrawSuite2022.USProgressBar PBLista2 
            Height          =   255
            Left            =   180
            TabIndex        =   68
            Top             =   5910
            Width           =   11205
            _ExtentX        =   19764
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
         Begin MSComctlLib.ListView lista_prog 
            Height          =   5715
            Left            =   180
            TabIndex        =   51
            Top             =   180
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   10081
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Inicio prazo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Final prazo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Confirm."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   4260
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Pedido"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Ped. cliente"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "N. item"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "D"
               Text            =   "Dt. negoc."
               Object.Width           =   1587
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Notas fiscais"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7215
         Left            =   -63345
         TabIndex        =   65
         Top             =   2790
         Width           =   3615
         Begin VB.TextBox Txt_qtde_total_faturada 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade total faturada."
            Top             =   6750
            Width           =   1845
         End
         Begin MSComctlLib.ListView lista_nota 
            Height          =   5850
            Left            =   180
            TabIndex        =   50
            Top             =   315
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   10319
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Text            =   "Id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "Nota fiscal"
               Object.Width           =   1826
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1587
            EndProperty
         End
         Begin DrawSuite2022.USProgressBar PBlista3 
            Height          =   255
            Left            =   180
            TabIndex        =   66
            Top             =   6180
            Width           =   3225
            _ExtentX        =   5689
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. total faturada"
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
            Left            =   1635
            TabIndex        =   108
            Top             =   6540
            Width           =   1680
         End
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -72750
         TabIndex        =   64
         Text            =   "0"
         Top             =   3780
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtID_item 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2010
         TabIndex        =   63
         Text            =   "0"
         Top             =   4440
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   59
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
            TabIndex        =   12
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
            TabIndex        =   11
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   16
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":F292
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
            TabIndex        =   15
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":12A36
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
            TabIndex        =   13
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
            TabIndex        =   14
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":1653F
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
            TabIndex        =   17
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_programacao.frx":1A62E
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
            TabIndex        =   114
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
            TabIndex        =   60
            Top             =   240
            Width           =   645
         End
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   -62130
         Top             =   510
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_programacao.frx":1DEBA
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   11850
         Top             =   510
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_programacao.frx":2329E
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   -63480
         Top             =   420
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_programacao.frx":28CC7
         Count           =   1
      End
      Begin MSComctlLib.ListView lista_item 
         Height          =   6050
         Left            =   75
         TabIndex        =   29
         Top             =   3025
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10663
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
         NumItems        =   10
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
            Text            =   "Descrição"
            Object.Width           =   6518
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde. total prev."
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Qtde. total vend."
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde. total fat."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Saldo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Valid."
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView lista 
         Height          =   6270
         Left            =   -74925
         TabIndex        =   10
         Top             =   2805
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11060
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
         NumItems        =   8
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
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Nº programa"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Data rev."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   8828
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   3528
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   90
         Top             =   330
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
         ButtonCaption8  =   "Revisar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Revisar (F7)"
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
         ButtonWidth8    =   44
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Emitir PI"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Emitir PI das programação(ões) no período (F8)"
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
         ButtonLeft9     =   352
         ButtonTop9      =   2
         ButtonWidth9    =   47
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
         ButtonLeft10    =   401
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   405
         ButtonTop11     =   2
         ButtonWidth11   =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   443
         ButtonTop12     =   2
         ButtonWidth12   =   26
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
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
         ButtonState13   =   5
         ButtonLeft13    =   471
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   75
         TabIndex        =   91
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   38
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   77
         ButtonTop3      =   2
         ButtonWidth3    =   39
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   47
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   220
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Emitir PI"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Emitir PI das programação(ões) do produto/item no período (F8)"
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
         ButtonLeft7     =   268
         ButtonTop7      =   2
         ButtonWidth7    =   47
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
         ButtonLeft8     =   317
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   321
         ButtonTop9      =   2
         ButtonWidth9    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   359
         ButtonTop10     =   2
         ButtonWidth10   =   26
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
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
         ButtonState11   =   5
         ButtonLeft11    =   387
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   1005
         Left            =   -74925
         TabIndex        =   92
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1773
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   38
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   77
         ButtonTop3      =   2
         ButtonWidth3    =   39
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Produto anterior."
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
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   47
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo produto."
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
         ButtonLeft6     =   220
         ButtonTop6      =   2
         ButtonWidth6    =   46
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
         ButtonLeft7     =   268
         ButtonTop7      =   4
         ButtonWidth7    =   2
         ButtonHeight7   =   56
         ButtonCaption8  =   "Ajuda"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Ajuda (F1)"
         ButtonKey8      =   "10"
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
         ButtonLeft8     =   272
         ButtonTop8      =   2
         ButtonWidth8    =   36
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Sair"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Sair (Esc)"
         ButtonKey9      =   "11"
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
         ButtonLeft9     =   310
         ButtonTop9      =   2
         ButtonWidth9    =   26
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonKey10     =   "12"
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
         ButtonLeft10    =   338
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         ButtonUseMaskColor10=   0   'False
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   855
         Left            =   75
         TabIndex        =   69
         Top             =   1305
         Width           =   15195
         Begin VB.ComboBox cmbReferencia 
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
            Left            =   2700
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            ToolTipText     =   "Código de referência."
            Top             =   390
            Width           =   1890
         End
         Begin VB.CheckBox Chk_validado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Validado"
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
            Left            =   13950
            TabIndex        =   24
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox txtCodigo 
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
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Código do item."
            Top             =   390
            Width           =   1755
         End
         Begin VB.TextBox txtDescricao 
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
            Left            =   4610
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do item."
            Top             =   390
            Width           =   6405
         End
         Begin VB.CommandButton cmdLocalizar_item 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2280
            Picture         =   "frmVendas_programacao.frx":2FD88
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Localizar produtos (F2)"
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtStatus_item 
            Alignment       =   2  'Center
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
            Left            =   11040
            TabIndex        =   23
            Top             =   390
            Width           =   2805
         End
         Begin VB.CommandButton cmdfiltrar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1950
            Picture         =   "frmVendas_programacao.frx":2FE8A
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Filtrar por código interno."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
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
            Left            =   2970
            TabIndex        =   116
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label16 
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
            Left            =   442
            TabIndex        =   72
            Top             =   180
            Width           =   1230
         End
         Begin VB.Label Label18 
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
            Left            =   7467
            TabIndex        =   71
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
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
            Left            =   12165
            TabIndex        =   70
            Top             =   180
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmVendas_programacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Programacao_Vendas              As Boolean 'OK
Dim Novo_Programacao_Vendas1             As Boolean 'OK
Dim Novo_Programacao_Vendas2             As Boolean 'OK
Public Sql_Programacao_Vendas_Localizar  As String 'OK
Public Sql_Programacao_Vendas_Localizar_produto As String 'OK
Dim TBLISTA_Vendas_programacao   As ADODB.Recordset 'OK
Dim TBLISTA_Vendas_programacao1  As ADODB.Recordset 'OK
Dim Emitir_PI As Boolean 'OK
Public TabelaSN_Prog As Integer 'OK

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Vendas_programa order by Programatexto", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID = " & txtId)
    TBCompras.MovePrevious
    If TBCompras.BOF = False Then
        txtId = TBCompras!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCampos_Item
        ProcLimpaCampos_Prog
        ProcCarregaDados
        ProcFiltrarListaProdutos
        ProcCarregalista_Prog
    Else
        USMsgBox ("Fim dos cadastros de programação de venda."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Vendas1 = False
Novo_Programacao_Vendas2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAnterior_prog()
On Error GoTo tratar_erro

If txtID_item = "" Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Vendas_programa_item where ID = " & txtId & " and codigo <> 'Null' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID_item = " & txtID_item)
    TBCompras.MovePrevious
    If TBCompras.BOF = False Then
        ProcLimpaCampos_Prog
        txtID_item = TBCompras!Id_Item
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_programa_item where ID_Item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
        txtCodigo = TBAbrir!CODIGO
        txtCodigo2 = TBAbrir!CODIGO
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select descricao from projproduto where desenho = '" & TBAbrir!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            txtdescricao = TBProduto!Descricao
            txtDescricao2 = TBProduto!Descricao
        End If
        TBProduto.Close
        ProcCarregalista_Prog
        ProcCarregaListaNF
    Else
        USMsgBox ("Fim dos cadastros de produtos."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Vendas2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExcluir_prog()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_prog
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) programação(ões) do produto?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Vendas_programacao where ID_prog = " & .ListItems(InitFor)
            ProcVerifExcluirPedido "VP.ID_programa = " & txtId & " and VC.ID_programacao = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Programação"
            Evento = "Excluir programação do produto"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev & " - Cód. interno: " & txtCodigo & " - Descrição: " & txtdescricao
            Documento1 = "Data: " & .ListItems(InitFor).SubItems(1) & "-" & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) programação(ões) do produto antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Programação(ões) do produto excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Prog
    ProcCarregalista_Prog
    Frame7.Enabled = False
    optFirme.Enabled = True
    ProcAltera_Status txtID_item, txtId
    Novo_Programacao_Vendas2 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procExcluir_item()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_item
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/item(ns) ?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Vendas_programa_item where id_item = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Vendas_programacao where id_item = " & .ListItems(InitFor)
            ProcVerifExcluirPedido "VP.ID_programa = " & txtId & " and VC.Desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
            '==================================
            Modulo = "Vendas/Programação"
            Evento = "Excluir produto"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1) & " - Descrição: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produtos(s)/item(ns) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produtos(s)/item(ns) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Item
    ProcFiltrarListaProdutos
    Frame4.Enabled = False
    ProcAltera_Status txtID_item, txtId
    ProcLimparTudo
    Novo_Programacao_Vendas1 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_filtrarpor_Click()
On Error GoTo tratar_erro

lista_item.ListItems.Clear
If Cmb_filtrarpor = "Status" Then
    Cmb_texto.Visible = True
    Txt_texto.Visible = False
Else
    Cmb_texto.Visible = False
    Txt_texto.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_texto_Click()
On Error GoTo tratar_erro

lista_item.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_filtrar_Click()
On Error GoTo tratar_erro

ProcFiltrarListaProdutos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

ProcPuxaDadosProduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcPuxaDadosProduto()
On Error GoTo tratar_erro

If txtCodigo <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto, desenho, descricao from projproduto where desenho = '" & txtCodigo & "' and Vendas = 'True' and Tipo = 'P' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtCodigo = TBProduto!Desenho
        txtdescricao = Trim(TBProduto!Descricao)
        
        ProcCarregaComboCodRef cmbReferencia, "P.codproduto = " & TBProduto!Codproduto, txtID_cli, "C", True, True
    End If
    TBProduto.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdLocalizar_item_Click()
On Error GoTo tratar_erro

Proposta_Servicos = False
PI_Servicos = False
Vendas_Proposta = False
Vendas_PI = False
Vendas_Programacao = True
frmVendas_ListaProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procNovo_item()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido criar um novo produto, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos_Item
Novo_Programacao_Vendas1 = True
Frame4.Enabled = True
txtCodigo.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcNovo_prog()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido criar uma nova programação, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifProdBloq(True) = True Then Exit Sub
ProcLimpaCampos_Prog
Novo_Programacao_Vendas2 = True
Frame7.Enabled = True
optFirme.Enabled = True
txtData_inicio.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_programacao.AbsolutePage <> 2 Then
    If TBLISTA_Vendas_programacao.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Vendas_programacao.PageCount - 1)
    Else
        TBLISTA_Vendas_programacao.AbsolutePage = TBLISTA_Vendas_programacao.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Vendas_programacao.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Vendas_programacao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Vendas_programacao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_programacao.AbsolutePage = 1
ProcExibePagina (TBLISTA_Vendas_programacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_programacao.AbsolutePage <> -3 Then
    If TBLISTA_Vendas_programacao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Vendas_programacao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Vendas_programacao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_programacao.AbsolutePage = TBLISTA_Vendas_programacao.PageCount
ProcExibePagina (TBLISTA_Vendas_programacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagAnt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_programacao1.AbsolutePage <> 2 Then
    If TBLISTA_Vendas_programacao1.AbsolutePage = -3 Then
        ProcExibePagina1 (TBLISTA_Vendas_programacao1.PageCount - 1)
    Else
        TBLISTA_Vendas_programacao1.AbsolutePage = TBLISTA_Vendas_programacao1.AbsolutePage - 2
        ProcExibePagina1 (TBLISTA_Vendas_programacao1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagIr1_Click()
On Error GoTo tratar_erro

If txtPagIr1 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas1.Caption, 4))
If Quant <= 1 Or txtPagIr1 > Quant Then Exit Sub
If txtPagIr1.Text >= 1 And txtPagIr1.Text <= Quant Then
    TBLISTA_Vendas_programacao1.AbsolutePage = txtPagIr1.Text
    ProcExibePagina1 (TBLISTA_Vendas_programacao1.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagPrim1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_programacao1.AbsolutePage = 1
ProcExibePagina1 (TBLISTA_Vendas_programacao1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagProx1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_programacao1.AbsolutePage <> -3 Then
    If TBLISTA_Vendas_programacao1.AbsolutePage = 1 Then
        ProcExibePagina1 (2)
    Else
        ProcExibePagina1 (TBLISTA_Vendas_programacao1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (TBLISTA_Vendas_programacao1.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagUlt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_programacao1.AbsolutePage = TBLISTA_Vendas_programacao1.PageCount
ProcExibePagina1 (TBLISTA_Vendas_programacao1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdpesquisar_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False
frmVendas_LocalizarCliente.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Vendas_programa order by Programatexto", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID = " & txtId)
    TBCompras.MoveNext
    If TBCompras.EOF = False Then
        txtId = TBCompras!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCampos_Item
        ProcLimpaCampos_Prog
        ProcCarregaDados
        ProcFiltrarListaProdutos
        ProcCarregalista_Prog
    Else
        USMsgBox ("Fim dos cadastros de programação de venda."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Vendas1 = False
Novo_Programacao_Vendas2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcProximo_prog()
On Error GoTo tratar_erro

If txtID_item = "" Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Vendas_programa_item where ID = " & txtId & " and codigo <> 'Null' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID_item = " & txtID_item)
    TBCompras.MoveNext
    If TBCompras.EOF = False Then
        ProcLimpaCampos_Prog
        txtID_item = TBCompras!Id_Item
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_programa_item where ID_Item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
        txtCodigo = TBAbrir!CODIGO
        txtCodigo2 = TBAbrir!CODIGO
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select descricao from projproduto where desenho = '" & TBAbrir!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            txtdescricao = TBProduto!Descricao
            txtDescricao2 = TBProduto!Descricao
        End If
        TBProduto.Close
        ProcCarregalista_Prog
        ProcCarregaListaNF
    Else
        USMsgBox ("Fim dos cadastros de produtos."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Vendas2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcRevisar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "criar revisão"
If txtPrograma = "" Then
    NomeCampo = "o programa"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Programacao_Vendas = True Then
    USMsgBox ("Salve o programa antes de criar revisão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente criar uma revisão deste programa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    '==================================
    Modulo = "Vendas/Programação"
    ID_documento = txtId
    Evento = "Revisar"
    Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * from Vendas_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    txtRev = IIf(IsNull(TBCotacao!Rev), 0, TBCotacao!Rev) + 1
    txtData_rev = Format(Date, "dd/mm/yy")
    IDlista = TBCotacao!programa
    IDAntigo = txtId
    TBCotacao.AddNew
    TBCotacao!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBCotacao!programa = IDlista
    TBCotacao!programatexto = txtPrograma
    TBCotacao!Data = Date
    TBCotacao!Responsavel = pubUsuario
    TBCotacao!Rev = txtRev
    TBCotacao!status = "ABERTO"
    TBCotacao!ID_Cliente = txtID_cli
    TBCotacao!via = "0"
    TBCotacao.Update
    txtId = TBCotacao!ID
    TBCotacao.Close
    
    Conexao.Execute "UPDATE Vendas_programa Set Status = 'REVISADA', Data_rev = '" & Format(Date, "Short Date") & "' where ID = " & IDAntigo
    Conexao.Execute "UPDATE Vendas_proposta Set ID_programa = " & txtId & " where ID_programa = " & IDAntigo
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Vendas_programa_item where ID = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Do While TBProduto.EOF = False
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Vendas_programa_item", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!CODIGO = TBProduto!CODIGO
            TBGravar!ID = txtId
            TBGravar!Status_Item = "ABERTO"
            TBGravar.Update
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Vendas_programacao where ID_item = " & TBProduto!Id_Item & " and Quantidade > qtdefaturada", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Set TBPrograma = CreateObject("adodb.recordset")
                    TBPrograma.Open "Select * from Vendas_programacao", Conexao, adOpenKeyset, adLockOptimistic
                    TBPrograma.AddNew
                    TBPrograma!ID_Antigo = TBAbrir!ID_prog
                    TBPrograma!ID = txtId
                    TBPrograma!Id_Item = TBGravar!Id_Item
                    TBPrograma!Un = TBAbrir!Un
                    TBPrograma!Unidade_com = TBAbrir!Unidade_com
                    TBPrograma!Data_inicio = TBAbrir!Data_inicio
                    TBPrograma!Data_fim = TBAbrir!Data_fim
                    TBPrograma!quantidade = TBAbrir!quantidade - TBAbrir!QtdeFaturada
                    TBPrograma!Firme = TBAbrir!Firme
                    TBPrograma!Status_prog = "ABERTO"
                    TBPrograma!Ordenar = TBAbrir!Ordenar
                    TBPrograma!PCCliente = TBAbrir!PCCliente
                    TBPrograma!N_item = TBAbrir!N_item
                    TBPrograma.Update
                    
                    Conexao.Execute "UPDATE vendas_carteira Set ID_programacao = " & TBPrograma!ID_prog & " where ID_programacao = " & TBAbrir!ID_prog
                    
                    TBPrograma.Close
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            
            TBGravar.Close
            TBProduto.MoveNext
        Loop
    End If
    TBProduto.Close
    
    USMsgBox ("Revisão do programa criada com sucesso."), vbInformation, "CAPRIND v5.0"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Vendas_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then ProcCarregaDados
    ProcFiltrarListaProdutos
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEmitirPI()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "firmar a(s) programação(ões)"
If txtPrograma = "" Then
    NomeCampo = "o programa"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Programacao_Vendas = True Then
    USMsgBox ("Salve o programa antes de emitir o(s) pedido(s) da(s) programação(ões)."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente emitir o(s) pedido(s) da(s) programação(ões) deste programa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem:
    DataTexto = InputBox("Favor informar o prazo final da(s) programação(ões).")
    If DataTexto = "" Then Exit Sub
    If IsDate(DataTexto) = False Then
        USMsgBox ("Esta data não é válida."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem
    End If
    DataFim = DataTexto
    
    Familiatext = ""
    Permitido = True
    Set TBPrograma = CreateObject("adodb.recordset")
    TBPrograma.Open "Select VPI.codigo, VPI.referencia, VP.*, P.Descricao from (Vendas_programa_item VPI INNER JOIN Vendas_programacao VP ON VPI.Id_item = VP.Id_item) INNER JOIN projproduto P ON P.Desenho = VPI.Codigo where VP.ID = " & txtId & " and VP.Data_fim <= '" & DataFim & "' and VP.Firme = 'False' order by VPI.ID_item, VP.ID_prog", Conexao, adOpenKeyset, adLockOptimistic
    If TBPrograma.EOF = False Then
        TBPrograma.MoveLast
        PBLista.Min = 0
        PBLista.Max = TBPrograma.RecordCount
        PBLista.Value = 1
        Contador = 0
        TBPrograma.MoveFirst
        
        'Verifica se tem alguma programação sem pedido do cliente e número do item
        Do While TBPrograma.EOF = False
            If IsNull(TBPrograma!PCCliente) = True Or TBPrograma!PCCliente = "" Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Código interno: " & TBPrograma!CODIGO & " - Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000") Else Familiatext = "Código interno: " & TBPrograma!CODIGO & " - Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000")
                Permitido = False
            End If
            If IsNull(TBPrograma!N_item) = True Or TBPrograma!N_item = "" Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Código interno: " & TBPrograma!CODIGO & " - Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000") Else Familiatext = "Código interno: " & TBPrograma!CODIGO & " - Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000")
                Permitido = False
            End If
            TBPrograma.MoveNext
        Loop
        If Permitido = False Then
            USMsgBox ("Informe o pedido do cliente e número do item do(s) produto(s) antes de emitir o PI: " & vbCrLf & " " & Familiatext), vbInformation, "CAPRIND v5.0"
            Exit Sub
        End If
        
        Emitir_PI = True
        TBPrograma.MoveFirst
        Do While TBPrograma.EOF = False
            TBPrograma!Firme = True
            If TBPrograma!Status_prog = "PREVISÃO FUTURA" Then
                TBPrograma!Status_prog = "ABERTO"
                TBPrograma!Ordenar = 2
            End If
            TBPrograma.Update
            ProcAlteraStatusProgramacao TBPrograma!ID_prog, TBPrograma!quantidade
            ProcAltera_Status TBPrograma!Id_Item, txtId
            
            ProcGerarPedido TBPrograma!ID_prog, TBPrograma!Data_fim, TBPrograma!PCCliente, TBPrograma!N_item, TBPrograma!CODIGO, TBPrograma!Descricao, TBPrograma!quantidade, TBPrograma!Un, TBPrograma!Unidade_com, TBPrograma!Utiliza_mat_cons, IIf(IsNull(TBPrograma!Referencia), "", TBPrograma!Referencia)
            
            Contador = Contador + 1
            PBLista.Value = Contador
            TBPrograma.MoveNext
        Loop
        ProcGravarTotaisPedido
        FunAtualizaStatusPropPI (IDpedido)
        USMsgBox ("Pedido(s) emitido(s) com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Vendas/Programação"
        ID_documento = txtId
        Evento = "Emitir PI"
        Documento = "Nº programa: " & txtPrograma & " - Nº revisão: " & txtRev
        Documento1 = "Data: " & Format(DataFim, "dd/mm/yy")
        ProcGravaEvento
        '==================================
    Else
        USMsgBox ("Não foi encontrada nenhuma programação com o prazo final menor ou igual a " & Format(DataFim, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
    End If
    TBPrograma.Close
End If
Emitir_PI = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcEmitirPI_item()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "firmar a(s) programação(ões)"
If txtdescricao = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Programacao_Vendas1 = True Then
    USMsgBox ("Salve o produto antes de emitir o(s) pedido(s) da(s) programação(ões)."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente emitir o(s) pedido(s) da(s) programação(ões) deste produto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem:
    DataTexto = InputBox("Favor informar o prazo final da(s) programação(ões).")
    If DataTexto = "" Then Exit Sub
    If IsDate(DataTexto) = False Then
        USMsgBox ("Esta data não é válida."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem
    End If
    DataFim = DataTexto
    
    Familiatext = ""
    Set TBPrograma = CreateObject("adodb.recordset")
    TBPrograma.Open "Select VPI.codigo, VPI.Referencia, VP.*, P.Descricao from (Vendas_programa_item VPI INNER JOIN Vendas_programacao VP ON VPI.Id_item = VP.Id_item) INNER JOIN projproduto P ON P.Desenho = VPI.Codigo where VP.ID = " & txtId & " and VPI.ID_item = " & txtID_item & " and VP.Data_fim <= '" & DataFim & "' and VP.Firme = 'False' order by VPI.ID_item, VP.ID_prog", Conexao, adOpenKeyset, adLockOptimistic
    If TBPrograma.EOF = False Then
        'Verifica se tem alguma programação sem pedido do cliente e número do item
        Do While TBPrograma.EOF = False
            If IsNull(TBPrograma!PCCliente) = True Or TBPrograma!PCCliente = "" Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000") Else Familiatext = "Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000")
                Permitido = False
            End If
            If IsNull(TBPrograma!N_item) = True Or TBPrograma!N_item = "" Then
                If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & " Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000") Else Familiatext = "Início do prazo: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & " - Final do prazo: " & Format(TBPrograma!Data_fim, "dd/mm/yy") & " - Qtde. " & Format(TBPrograma!quantidade, "###,##0.0000")
                Permitido = False
            End If
            TBPrograma.MoveNext
        Loop
        If Permitido = False Then
            USMsgBox ("Informe o pedido do cliente e o número do item da(s) programação(ões) antes de emitir o PI: " & vbCrLf & " " & Familiatext), vbInformation, "CAPRIND v5.0"
            Exit Sub
        End If
        
        TBPrograma.MoveFirst
        Do While TBPrograma.EOF = False
            TBPrograma!Firme = True
            If TBPrograma!Status_prog = "PREVISÃO FUTURA" Then
                TBPrograma!Status_prog = "ABERTO"
                TBPrograma!Ordenar = 2
            End If
            TBPrograma.Update
            ProcAlteraStatusProgramacao TBPrograma!ID_prog, TBPrograma!quantidade
            ProcAltera_Status TBPrograma!Id_Item, txtId
            
            ProcGerarPedido TBPrograma!ID_prog, TBPrograma!Data_fim, TBPrograma!PCCliente, TBPrograma!N_item, TBPrograma!CODIGO, TBPrograma!Descricao, TBPrograma!quantidade, TBPrograma!Un, TBPrograma!Unidade_com, TBPrograma!Utiliza_mat_cons, IIf(IsNull(TBPrograma!Referencia), "", TBPrograma!Referencia)
            
            '==================================
            Modulo = "Vendas/Programação"
            ID_documento = txtID_item
            Evento = "Emitir PI do produto"
            Documento = "Nº programa: " & txtPrograma & " - Nº revisão: " & txtRev & " - Cód. interno: " & TBPrograma!CODIGO & " - Descrição: " & TBPrograma!Descricao
            Documento1 = "Data: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & "-" & Format(TBPrograma!Data_fim, "dd/mm/yy")
            ProcGravaEvento
            '==================================
            
            TBPrograma.MoveNext
        Loop
        ProcGravarTotaisPedido
        FunAtualizaStatusPropPI (IDpedido)
        USMsgBox ("Pedido(s) do produto emitido(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Não foi encontrada nenhuma programação com o prazo final menor ou igual a " & Format(DataFim, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
    End If
    TBPrograma.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procSalvar_item()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido alterar este produto, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If txtCodigo = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodigo.SetFocus
    Exit Sub
End If
If txtdescricao = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
'Verifica se o produto já não está na lista
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_programa_item where ID = " & txtId.Text & " and codigo = '" & txtCodigo & "' and ID_item <> " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Já existe o produto " & txtCodigo & " nesta programação."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_programa_item where id_item = " & txtID_item & " and id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If txtCodigo <> TBGravar!CODIGO Then
        If (txtstatus_item = "PARCIAL" Or txtstatus_item = "FATURADO") Then
            USMsgBox ("Não é permitido alterar este produto, pois o mesmo está " & IIf(txtstatus_item = "FATURADO", "faturado", "faturado parcial") & "."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_programacao where id_item = " & TBGravar!Id_Item, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            USMsgBox ("Não é permitido alterar este produto, pois o mesmo já possui programação."), vbExclamation, "CAPRIND v5.0"
            TBAbrir.Close
            Exit Sub
        End If
        TBAbrir.Close
    End If
End If
TBGravar!CODIGO = txtCodigo
TBGravar!ID = txtId
TBGravar!Status_Item = txtstatus_item
TBGravar!Referencia = cmbReferencia
If Chk_validado.Value = 1 Then TBGravar!Validado = True Else TBGravar!Validado = False
TBGravar.Update
txtID_item = TBGravar!Id_Item
TBGravar.Close
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_programa_item where id = " & txtId & " and status_item = 'FATURADO' or id = " & txtId & " and status_item = 'PARCIAL'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from Vendas_programa where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        TBItem!status = "PARCIAL"
        TBItem.Update
        txtStatus = "PARCIAL"
    End If
    TBItem.Close
End If
TBAbrir.Close
Lista.ListItems.Clear
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
ProcFiltrarListaProdutos
If Novo_Programacao_Vendas1 = True Then
    USMsgBox ("Novo produto cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto"
    If CodigoLista1 <> 0 And lista_item.ListItems.Count <> 0 Then
        lista_item.SelectedItem = lista_item.ListItems(CodigoLista1)
        lista_item.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Programação"
ID_documento = txtID_item
Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
Documento1 = "Cód. interno: " & txtCodigo & " - Descrição: " & txtdescricao
ProcGravaEvento
'==================================
Novo_Programacao_Vendas1 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvar_prog()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame7.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido alterar esta programação, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
valor = IIf(txtQTD = "", 0, txtQTD)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQTD.SetFocus
    Exit Sub
End If
With txtData_fim
    If FunVerificaDataFinal(txtData_inicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
If FunVerifProdBloq(False) = True Then Exit Sub

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_programacao where ID_prog = " & txtID_prog, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!Status_prog = "PREVISÃO FUTURA"
Else
    If TBGravar!Firme = True Then
        If optFirme.Value = 0 Then
            If FunVerifDadosPedido("VC.ID_programacao", txtID_prog, True, "cancelar a venda desta programação") = False Then Exit Sub
        End If
        
        'Verifica se a quantidade da programação foi alterada
        qt = txtQTD
        If TBGravar!quantidade <> qt Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from vendas_carteira where ID_programacao = " & txtID_prog, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                TBAbrir!quantidade = qt
                If qt > TBGravar!quantidade Then TBAbrir!OE = False
                If TBAbrir!QtdeFaturada > 0 Then If TBAbrir!QtdeFaturada >= qt Then TBAbrir!Liberacao = "FATURADO" Else TBAbrir!Liberacao = "FATURADO PARCIAL"
                TBAbrir.Update
                
                FunAtualizaStatusPropPI TBAbrir!Cotacao
            End If
            TBAbrir.Close
        End If
    End If
End If
If optFirme.Value = 1 Then
    If Txt_pedido_cliente = "" Then
        NomeCampo = "o pedido do cliente"
        ProcVerificaAcao
        Txt_pedido_cliente.SetFocus
        Exit Sub
    End If
    If Txt_n_item = "" Then
        NomeCampo = "o número do item"
        ProcVerificaAcao
        Txt_n_item.SetFocus
        Exit Sub
    End If
    pc = Txt_pedido_cliente
End If
If Txt_data_negociada <> "__/__/____" Then
    If IsDate(Txt_data_negociada) = False Then
        NomeCampo = "a data negociada"
        ProcVerificaAcao
        Txt_data_negociada.SetFocus
        Exit Sub
    End If
End If
ProcEnviadados_Prog
TBGravar.Update
txtID_prog = TBGravar!ID_prog

If TBGravar!Firme = True Then
    ProcGerarPedido txtID_prog, txtData_fim, pc, Txt_n_item, txtCodigo, txtdescricao, txtQTD, Txt_un, Txt_un_com, Chk_utiliza_mat_consignado, cmbReferencia
    ProcGravarTotaisPedido
    FunAtualizaStatusPropPI (IDpedido)
Else
    If txtID_prog <> 0 Then ProcVerifExcluirPedido "VC.Id_programacao = " & txtID_prog
End If

TBGravar.Close
ProcCarregalista_Prog
If Novo_Programacao_Vendas2 = True Then
    USMsgBox ("Nova programação do produto cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova programação produto"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar programação produto"
    If CodigoLista2 <> 0 And lista_prog.ListItems.Count <> 0 Then
        lista_prog.SelectedItem = lista_prog.ListItems(CodigoLista2)
        lista_prog.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Programação"
ID_documento = txtID_prog
Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev & " - Cód. interno: " & txtCodigo & " - Descrição " & txtdescricao
Documento1 = "Data: " & txtData_inicio.Value & "-" & txtData_fim.Value
ProcGravaEvento
'==================================
Novo_Programacao_Vendas2 = False

ProcAlteraStatusProgramacao txtID_prog, txtQTD
ProcAltera_Status txtID_item, txtId
optFirme.Enabled = True
Lista.ListItems.Clear
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
ProcFiltrarListaProdutos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunVerifProdBloq(Novo As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifProdBloq = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Codproduto from projproduto where Desenho = '" & txtCodigo & "' and Bloqueado = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido " & IIf(Novo = True, "criar nova", "salvar esta") & " programação, pois o produto " & txtCodigo & " está bloqueado."), vbExclamation, "CAPRIND v5.0"
    FunVerifProdBloq = True
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Private Sub ProcAlteraStatusProgramacao(ID_programacao As Long, Qtde As Double)
On Error GoTo tratar_erro

'Altera status da programação do item
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_programacao where id_prog = " & ID_programacao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ValorTotal = Qtde - IIf(IsNull(TBAbrir!QtdeFaturada), 0, TBAbrir!QtdeFaturada)
    If ValorTotal = 0 Then
        TBAbrir!Status_prog = "FATURADO"
        TBAbrir!Ordenar = 4
    ElseIf ValorTotal = Qtde Then
            If TBAbrir!Firme = True Then
                TBAbrir!Status_prog = "ABERTO"
                TBAbrir!Ordenar = 2
            Else
                TBAbrir!Status_prog = "PREVISÃO FUTURA"
                TBAbrir!Ordenar = 3
            End If
        Else
            TBAbrir!Status_prog = "PARCIAL"
            TBAbrir!Ordenar = 1
    End If
    TBAbrir.Update
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcRevisar
            Case vbKeyF8: ProcEmitirPI
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_item
            Case vbKeyF2: ProcFiltrarListaProdutos
            Case vbKeyF3: procSalvar_item
            Case vbKeyF4: procExcluir_item
            Case vbKeyF5: ProcImprimir
            Case vbKeyF8: ProcEmitirPI_item
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_prog
            Case vbKeyF3: ProcSalvar_prog
            Case vbKeyF4: If Cmb_opcao_lista2 = "Excluir" Then ProcExcluir_prog
            Case vbKeyF5: ProcImprimir
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 13, True
ProcCarregaToolBar2 Me, 15195, 11, True
ProcCarregaToolBar3 Me, 15195, 10, True
Formulario = "Vendas/Programação"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_filtrarpor = "Código interno"
Cmb_opcao_lista2 = "Excluir"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Programação"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmVendas_programacao_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
                If USMsgBox("Deseja realmente excluir este(s) programa(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBProgramas = CreateObject("adodb.recordset")
            TBProgramas.Open "Select * from Vendas_programacao where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
            If TBProgramas.EOF = False Then
                Do While TBProgramas.EOF = False
                    If TBProgramas!ID_Antigo <> 0 Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select VP.ID from Vendas_programa VP INNER JOIN Vendas_programacao VPR ON VP.ID = VPR.ID where VPR.ID_prog = " & TBProgramas!ID_Antigo, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Conexao.Execute "UPDATE VP Set VP.ID_programa = " & TBAbrir!ID & " from Vendas_proposta VP INNER JOIN vendas_carteira VC ON VP.Cotacao = VC.Cotacao where VC.ID_programacao = " & TBProgramas!ID_prog
                        End If
                        TBAbrir.Close
                        
                        Conexao.Execute "UPDATE Vendas_carteira Set ID_programacao = " & TBProgramas!ID_Antigo & " where ID_programacao = " & TBProgramas!ID_prog
                    Else
                        Conexao.Execute "DELETE from vendas_comercial from vendas_comercial INNER JOIN Vendas_proposta ON vendas_comercial.Cotacao = Vendas_proposta.Cotacao Where Vendas_proposta.ID_programa = " & .ListItems(InitFor)
                        Conexao.Execute "DELETE from Vendas_carteira from Vendas_carteira INNER JOIN Vendas_proposta ON Vendas_carteira.Cotacao = Vendas_proposta.Cotacao Where Vendas_proposta.ID_programa = " & .ListItems(InitFor)
                        Conexao.Execute "DELETE from Vendas_proposta where ID_programa = " & .ListItems(InitFor)
                    End If
                    TBProgramas.MoveNext
                Loop
            End If
            TBProgramas.Close
            
            Conexao.Execute "DELETE from Vendas_programa where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Vendas_programa_item where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Vendas_programacao where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Programação"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº programação: " & .ListItems(InitFor).SubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            
            Set TBPrograma = CreateObject("adodb.recordset")
            TBPrograma.Open "Select * from Vendas_programa where Programatexto = '" & .ListItems(InitFor).SubItems(3) & "' and Rev = " & .ListItems(InitFor).SubItems(4) - 1, Conexao, adOpenKeyset, adLockOptimistic
            If TBPrograma.EOF = False Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Vendas_programa_item where id = " & TBPrograma!ID & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = True Then
                    TBPrograma!status = "PREVISÃO FUTURA"
                Else
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Vendas_programa_item where id = " & TBPrograma!ID & " and Status_Item <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        TBPrograma!status = "ABERTO"
                    Else
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Vendas_programa_item where id = " & TBPrograma!ID & " and Status_Item <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = True Then
                            TBPrograma!status = "FATURADO"
                        Else
                            TBPrograma!status = "PARCIAL"
                        End If
                    End If
                End If
                TBAbrir.Close
                TBPrograma!Data_rev = Null
                TBPrograma.Update
            End If
            TBPrograma.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) programa(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Programa(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Lista.ListItems.Clear
    ProcCarregaLista (1)
    Framelista.Enabled = False
    ProcLimparTudo
    Novo_Programacao_Vendas = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtPrograma = "" Then
    USMsgBox ("Informe o programa antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Dataini = 0
DataFim = 0
'Verifica a datainicio e final das programações q estão e não estão como faturadas
Set TBProgramas = CreateObject("adodb.recordset")
If txtStatus <> "FATURADO" Then TextoFiltro = "Status_prog <> 'FATURADO'" Else TextoFiltro = "Status_prog = 'FATURADO'"
TBProgramas.Open "Select * from Vendas_programacao where ID = " & txtId & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProgramas.EOF = False Then
    Dataini = TBProgramas!Data_inicio
    TBProgramas.MoveLast
    DataFim = TBProgramas!Data_fim
End If
TBProgramas.Close
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_programa where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!Dtinicio_rel = Dataini
    TBGravar!Dtfinal_rel = DataFim
    'TBGravar!via = IIf(IsNull(TBGravar!via), 0, TBGravar!via) + 1
    TBGravar.Update
End If
TBGravar.Close

NomeRel = "Vendas_programacao.rpt"
ProcImprimirRel "{Vendas_programa.id} = " & txtId, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Data_Prog = ""
ProcLimpaCampos
Novo_Programacao_Vendas = True
Framelista.Enabled = True
txtData = Format(Date, "dd/mm/yy")
txtStatus = "ABERTO"
cmdpesquisar_Click
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame4.Enabled = False
Frame7.Enabled = False
ProcLimpaCampos_Item
ProcLimpaCampos_Prog
lista_item.ListItems.Clear
lista_prog.ListItems.Clear
lista_nota.ListItems.Clear
Novo_Programacao_Vendas1 = False
Novo_Programacao_Vendas2 = False
Sql_Programacao_Vendas_Localizar_produto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Programacao_Vendas = True Then
    If USMsgBox("O programa ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Programacao_Vendas = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Programacao_Vendas1 = True Then
    If USMsgBox("O produto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_item
        If Novo_Programacao_Vendas1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Programacao_Vendas2 = True Then
    If USMsgBox("A programação do produto ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_prog
        If Novo_Programacao_Vendas2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Programacao_Vendas = False
Novo_Programacao_Vendas1 = False
Novo_Programacao_Vendas2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtPrograma = ""
txtCliente = ""
txtID_cli = ""
txtResponsavel = pubUsuario
txtStatus = "ABERTO"
txtData.Text = Format(Date, "dd/mm/yy")
txtData_rev.Text = ""
txtRev = "0"
CodigoLista = 0
Caption = "Administrativo - Vendas - Programação"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Framelista.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtStatus = "FATURADO" Or txtStatus = "PARCIAL" Or txtStatus = "REVISADA" Then
    Select Case txtStatus
        Case "FATURADO": Mensagem = "faturado"
        Case "PARCIAL": Mensagem = "faturado parcial"
        Case "REVISADA": Mensagem = "revisado"
    End Select
    USMsgBox ("Não é permitido alterar este programa, pois o mesmo está " & Mensagem & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "Salvar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtID_cli = "" Or txtCliente = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    Exit Sub
End If
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * FROM Clientes WHERE idcliente = " & txtID_cli & " and Left(Tipo, 1) = 'J' And idTipoEmpresa = 1", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    If FunVerifRegimeTribCliForn(Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, True) = False Then Exit Sub
End If
TBClientes.Close
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Vendas_programa where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select programa from Vendas_programa where Year(Data) = '" & Year(Date) & "' order by id", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = True Then
        Cont = 1
    Else
        TBCompras.MoveLast
        Cont = TBCompras!programa + 1
    End If
    Data_Prog = Format(Date, "mm/yyyy")
    ProcGeraNumero
    TBCompras.Close

    TBGravar.AddNew
    TBGravar!programa = Cont
    TBGravar!programatexto = a
    TBGravar!Data = Date
    TBGravar!Responsavel = txtResponsavel
    TBGravar!status = "ABERTO"
End If
TBGravar!ID_Cliente = txtID_cli
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar.Update
txtId = TBGravar!ID
txtPrograma.Text = a
TBGravar.Close

Lista.ListItems.Clear
If Novo_Programacao_Vendas = True Then
    USMsgBox ("Novo programa cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Programacao_Vendas_Localizar = "Select Vendas_programa.*, Clientes.NomeRazao FROM Vendas_programa INNER JOIN Clientes ON Vendas_programa.ID_cliente = Clientes.IDCliente where Vendas_programa.ID = " & txtId
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
Modulo = "Vendas/Programação"
ID_documento = txtId
Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Programacao_Vendas = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub imgCalendario_Click()
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
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = True
Outros_solicitacaoPCP = False
Estoque_recebimento = False
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Vendas_programa where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If TBAbrir!status = "FATURADO" Or TBAbrir!status = "PARCIAL" Or TBAbrir!status = "REVISADA" Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                End If
                
                If FunVerifDadosPedido("VPP.ID", .ListItems(InitFor), False, "") = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
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
End Sub

Private Sub lista_item_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_item
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtStatus = "REVISADA" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If .ListItems(InitFor).SubItems(4) = "PARCIAL" Or .ListItems(InitFor).SubItems(4) = "FATURADO" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                
                If FunVerifDadosPedido("VPP.ID_item", .ListItems(InitFor), False, "") = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_item, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub lista_item_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_item
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If txtStatus = "REVISADA" Then
                USMsgBox ("Não é permitido excluir este produto, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If .ListItems(InitFor).SubItems(4) = "PARCIAL" Or .ListItems(InitFor).SubItems(4) = "FATURADO" Then
                USMsgBox ("Não é permitido excluir este produto, pois o mesmo está " & IIf(.ListItems(InitFor).SubItems(4) = "PARCIAL", "faturado parcial", "faturado") & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            If FunVerifDadosPedido("VPP.ID_item", .ListItems(InitFor), True, "excluir este produto") = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunVerifDadosPedido(NomeCampo As String, IDreg As Long, MostrarMsg As Boolean, Mensagem As String) As Boolean
On Error GoTo tratar_erro

FunVerifDadosPedido = True

'Verifica se já existe empenho no estoque ou na produção
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select VC.Codigo from (((Vendas_programa_item VPP INNER JOIN Vendas_programacao VP ON VPP.ID_item = VP.Id_item) INNER JOIN vendas_carteira VC ON VP.ID_prog = VC.ID_programacao) INNER JOIN Vendas_proposta VPR ON VPR.Cotacao = VC.Cotacao) INNER JOIN Empresa E ON E.Codigo = VPR.ID_empresa where " & NomeCampo & " = " & IDreg & " and E.Ativar_empenho_autom = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select ID from Estoque_Controle_Empenho_Vendas where ID_carteira = " & TBAliquota!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        If MostrarMsg = True Then USMsgBox ("Não é permitido " & Mensagem & ", pois existe(m) empenho(s) no estoque para o(s) pedido(s) vinculado(s)."), vbExclamation, "CAPRIND v5.0"
        FunVerifDadosPedido = False
        TBAliquota.Close
        TBProduto.Close
        Exit Function
    End If
    TBProduto.Close
End If
TBAliquota.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select VC.Codigo from ((Vendas_programa_item VPP INNER JOIN Vendas_programacao VP ON VPP.ID_item = VP.Id_item) INNER JOIN vendas_carteira VC ON VP.ID_prog = VC.ID_programacao) INNER JOIN Producao_pedidos PP ON VC.Codigo = PP.IDcarteira where " & NomeCampo & " = " & IDreg, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Mensagem & ", pois existe(m) empenho(s) na produção para o(s) pedido(s) vinculado(s)."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    FunVerifDadosPedido = False
    Exit Function
End If

'Verifica se foi gerado ordem de faturamento
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select VC.Codigo from (Vendas_programa_item VPP INNER JOIN Vendas_programacao VP ON VPP.ID_item = VP.Id_item) INNER JOIN vendas_carteira VC ON VP.ID_prog = VC.ID_programacao where " & NomeCampo & " = " & IDreg & " and (VC.Liberacao = 'FATURAR' or VC.Liberacao = 'FATURAR PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Mensagem & ", pois existe(m) ordem(ns) de faturamento aberta para o(s) pedido(s) vinculado(s)."), vbExclamation, "CAPRIND v5.0"
    FunVerifDadosPedido = False
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub lista_item_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_item.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Vendas_programa_item where ID_item = " & lista_item.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimpaCampos_Item
    Frame4.Enabled = True
    procPuxadadosItem
    CodigoLista1 = lista_item.SelectedItem.index
End If
TBProduto.Close
Novo_Programacao_Vendas1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Vendas_programa where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!status = "FATURADO" Or TBAbrir!status = "PARCIAL" Or TBAbrir!status = "REVISADA" Then
                    Select Case TBAbrir!status
                        Case "FATURADO": Mensagem = "faturado"
                        Case "PARCIAL": Mensagem = "faturado parcial"
                        Case "REVISADA": Mensagem = "revisado"
                    End Select
                    USMsgBox ("Não é permitido excluir este programa, pois o mesmo está " & Mensagem & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            
            If FunVerifDadosPedido("VPP.ID", .ListItems(InitFor), True, "excluir este programa") = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_programa where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close
Framelista.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBAbrir!ID_empresa) = False And TBAbrir!ID_empresa <> "" Then
    ProcPuxaDadosComboEmpresa Cmb_empresa, TBAbrir!ID_empresa
    Empresarel = Cmb_empresa
End If
Cont = TBAbrir!programa
Caption = "Administrativo - Vendas - Programação - (Programação : " & IIf(IsNull(TBAbrir!programatexto), "", TBAbrir!programatexto) & " - Rev. : " & IIf(IsNull(TBAbrir!Rev), "", TBAbrir!Rev) & ")"
txtId = TBAbrir!ID
txtPrograma = TBAbrir!programatexto
txtID_cli = TBAbrir!ID_Cliente
txtResponsavel = TBAbrir!Responsavel
txtStatus = TBAbrir!status
txtData = Format(TBAbrir!Data, "dd/mm/yy")
txtRev = IIf(IsNull(TBAbrir!Rev), "0", TBAbrir!Rev)
txtData_rev = IIf(IsNull(TBAbrir!Data_rev), "", Format(TBAbrir!Data_rev, "dd/mm/yy"))
Novo_Programacao_Vendas = False
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub lista_nota_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lista_nota, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub lista_prog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_prog
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtStatus = "REVISADA" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                
                If .ListItems(InitFor).SubItems(5) <> "ABERTO" And .ListItems(InitFor).SubItems(5) <> "PREVISÃO FUTURA" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                
                If FunVerifDadosPedido("VC.ID_programacao", .ListItems(InitFor), False, "") = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_prog, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub lista_prog_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_prog
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If txtStatus = "REVISADA" Then
                USMsgBox ("Não é permitido excluir esta programação, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            If .ListItems(InitFor).SubItems(5) <> "ABERTO" And .ListItems(InitFor).SubItems(5) <> "PREVISÃO FUTURA" Then
                USMsgBox ("Não é permitido excluir esta programação do produto, pois a mesma está " & IIf(.ListItems(InitFor).SubItems(5) = "FATURADO", "faturada", "faturada parcial") & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            If FunVerifDadosPedido("VC.ID_programacao", .ListItems(InitFor), True, "excluir esta programação") = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub lista_prog_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_prog.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_programacao where ID_prog = " & lista_prog.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_Prog
    
    txtID_prog = TBAbrir!ID_prog
    txtData_inicio.Value = TBAbrir!Data_inicio
    txtData_fim.Value = TBAbrir!Data_fim
    txtQTD = Format(TBAbrir!quantidade, "###,##0.0000")
    Txt_pedido_cliente = IIf(IsNull(TBAbrir!PCCliente), "", TBAbrir!PCCliente)
    Txt_n_item = IIf(IsNull(TBAbrir!N_item), "", TBAbrir!N_item)
    Txt_data_negociada = IIf(IsNull(TBAbrir!Data_negociada), "__/__/____", Format(TBAbrir!Data_negociada, "dd/mm/yyyy"))
    txtStatus_prog.Text = TBAbrir!Status_prog
    If TBAbrir!Firme = True Then optFirme.Value = 1 Else optFirme.Value = 0
    If TBAbrir!Utiliza_mat_cons = True Then Chk_utiliza_mat_consignado.Value = 1 Else Chk_utiliza_mat_consignado.Value = 0
    Txt_un = IIf(IsNull(TBAbrir!Un), "", TBAbrir!Un)
    Txt_un_com = IIf(IsNull(TBAbrir!Unidade_com), "", TBAbrir!Unidade_com)
    
    Frame7.Enabled = True
    CodigoLista2 = lista_prog.SelectedItem.index
    Novo_Programacao_Vendas2 = False
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        Cmb_empresa.Visible = True
        PBLista.Visible = True
        If Cmb_empresa.Visible = True Then Cmb_empresa.SetFocus
    Case 1:
        Cmb_empresa.Visible = False
        PBLista.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        lista_item.SetFocus
        'ProcFiltrarListaProdutos
    Case 2:
        Cmb_empresa.Visible = False
        PBLista.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        If txtCodigo = "" Then
            SSTab1.Tab = 1
            USMsgBox ("Informe o produto antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        If Novo_Programacao_Vendas1 = True Then
            SSTab1.Tab = 1
            USMsgBox ("Salve o produto antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Permitido = False
            Exit Sub
        End If
        lista_prog.SetFocus
        txtCodigo2 = txtCodigo
        txtDescricao2 = txtdescricao
        ProcVerificaUnidade
        ProcCarregalista_Prog
        ProcCarregaListaNF
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Programacao_Vendas = True Then
    SSTab1.Tab = 0
    USMsgBox ("Salve o programa antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    Exit Sub
End If
If txtCliente = "" Then
    SSTab1.Tab = 0
    USMsgBox ("Informe o cliente antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    cmdpesquisar_Click
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerificaUnidade()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Unidade, Unidade_com from projproduto where desenho = '" & txtCodigo2.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Txt_un.Text = TBProduto!Unidade
    Txt_un_com.Text = TBProduto!Unidade_com
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

If Sql_Programacao_Vendas_Localizar = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Set TBLISTA_Vendas_programacao = CreateObject("adodb.recordset")
TBLISTA_Vendas_programacao.Open Sql_Programacao_Vendas_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Vendas_programacao.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Vendas_programacao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Vendas_programacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendas_programacao.PageSize
ContadorReg = 1

If Emitir_PI = False Then
    PBLista.Min = 0
    PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendas_programacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendas_programacao.PageSize * (Pagina - 1)), 0), TBLISTA_Vendas_programacao.PageSize)
    PBLista.Value = 1
    Contador = 0
End If
Do While TBLISTA_Vendas_programacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Vendas_programacao!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendas_programacao!Data), "", Format(TBLISTA_Vendas_programacao!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Vendas_programacao!Responsavel), "", TBLISTA_Vendas_programacao!Responsavel)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Vendas_programacao!programatexto), "", TBLISTA_Vendas_programacao!programatexto)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendas_programacao!Rev), 0, TBLISTA_Vendas_programacao!Rev)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Vendas_programacao!Data_rev), "", Format(TBLISTA_Vendas_programacao!Data_rev, "dd/mm/yy"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Vendas_programacao!NomeRazao), "", TBLISTA_Vendas_programacao!NomeRazao)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Vendas_programacao!status), "", TBLISTA_Vendas_programacao!status)
    End With
    TBLISTA_Vendas_programacao.MoveNext
    ContadorReg = ContadorReg + 1
    If Emitir_PI = False Then
        Contador = Contador + 1
        PBLista.Value = Contador
    End If
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Vendas_programacao.RecordCount
If TBLISTA_Vendas_programacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Vendas_programacao.PageCount
ElseIf TBLISTA_Vendas_programacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Vendas_programacao.PageCount & " de: " & TBLISTA_Vendas_programacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Vendas_programacao.AbsolutePage - 1 & " de: " & TBLISTA_Vendas_programacao.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcLimpaCampos_Item()
On Error GoTo tratar_erro

txtID_item.Text = 0
txtCodigo.Text = ""
txtstatus_item = "ABERTO"
Chk_validado.Value = 0
txtdescricao.Text = ""
cmbReferencia.Clear
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaLista_Item(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros1.Caption = "Nº de registros: 0"
lblPaginas1.Caption = "Página: 0 de: 0"
lista_item.ListItems.Clear
Set TBLISTA_Vendas_programacao1 = CreateObject("adodb.recordset")
TBLISTA_Vendas_programacao1.Open Sql_Programacao_Vendas_Localizar_produto, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Vendas_programacao1.EOF = False Then ProcExibePagina1 (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExibePagina1(Pagina)
On Error GoTo tratar_erro

lista_item.ListItems.Clear
TBLISTA_Vendas_programacao1.PageSize = IIf(txtNreg1 = "", 30, txtNreg1)
TBLISTA_Vendas_programacao1.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendas_programacao1.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendas_programacao1.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendas_programacao1.PageSize * (Pagina - 1)), 0), TBLISTA_Vendas_programacao1.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Vendas_programacao1.EOF = False And (ContadorReg <= TamanhoPagina)
    With lista_item.ListItems
        .Add , , TBLISTA_Vendas_programacao1!Id_Item
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendas_programacao1!CODIGO), "", TBLISTA_Vendas_programacao1!CODIGO)
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select n_referencia from item_aplicacoes where Codproduto = " & TBLISTA_Vendas_programacao1!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
        End If
        TBAbrir.Close
        
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Vendas_programacao1!Descricao), "", TBLISTA_Vendas_programacao1!Descricao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendas_programacao1!Status_Item), "", TBLISTA_Vendas_programacao1!Status_Item)
        .Item(.Count).SubItems(5) = Format(FunVerificaQtdeTotalPrev(TBLISTA_Vendas_programacao1!Id_Item), "###,##0.0000")
        
        valor = FunVerificaQtdeTotalVend(TBLISTA_Vendas_programacao1!Id_Item)
        .Item(.Count).SubItems(6) = Format(valor, "###,##0.0000")
        
        Valor1 = FunVerificaQtdeTotalFaturada(TBLISTA_Vendas_programacao1!Id_Item)
        .Item(.Count).SubItems(7) = Format(Valor1, "###,##0.0000")
        
        .Item(.Count).SubItems(8) = Format(valor - Valor1, "###,##0.0000")
        If TBLISTA_Vendas_programacao1!Validado = True Then .Item(.Count).SubItems(9) = "Sim" Else .Item(.Count).SubItems(9) = "Não"
    End With
    TBLISTA_Vendas_programacao1.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros1.Caption = "Nº de registros: " & TBLISTA_Vendas_programacao1.RecordCount
If TBLISTA_Vendas_programacao1.AbsolutePage = adPosBOF Then
   lblPaginas1.Caption = "Página: 1 de: " & TBLISTA_Vendas_programacao1.PageCount
ElseIf TBLISTA_Vendas_programacao1.AbsolutePage = adPosEOF Then
        lblPaginas1.Caption = "Página: " & TBLISTA_Vendas_programacao1.PageCount & " de: " & TBLISTA_Vendas_programacao1.PageCount
    Else
        lblPaginas1.Caption = "Página: " & TBLISTA_Vendas_programacao1.AbsolutePage - 1 & " de: " & TBLISTA_Vendas_programacao1.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcGeraNumero()
On Error GoTo tratar_erro

a = Cont
Select Case Len(a)
    Case 1: a = "000" & Cont & "-" & Data_Prog
    Case 2: a = "00" & Cont & "-" & Data_Prog
    Case 3: a = "0" & Cont & "-" & Data_Prog
    Case 4: a = Cont & "-" & Data_Prog
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcLimpaCampos_Prog()
On Error GoTo tratar_erro

txtID_prog = 0
txtData_fim.Value = Date
txtData_inicio.Value = Date
txtStatus_prog = "ABERTO"
txtQTD = ""
ProcVerificaUnidade
Txt_pedido_cliente = ""
Txt_n_item = ""
Txt_data_negociada = "__/__/____"
optFirme.Value = 0
Chk_utiliza_mat_consignado.Value = False
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrarListaProdutos()
On Error GoTo tratar_erro

CamposFiltro = "VPI.Id_Item, VPI.CODIGO, P.Descricao, VPI.Status_Item, VPI.Validado, P.Codproduto"
If Txt_texto <> "" Or Cmb_texto <> "" Then
    If Txt_texto.Visible = True Then
        If Cmb_filtrarpor = "Código de referência" Then
            INNERJOINTEXTO = "(Vendas_programa_item VPI INNER JOIN Projproduto P ON VPI.Codigo = P.Desenho) INNER JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto"
            TextoFiltro = "IA.n_referencia"
        Else
            INNERJOINTEXTO = "Vendas_programa_item VPI INNER JOIN Projproduto P ON VPI.Codigo = P.Desenho"
            If Cmb_filtrarpor = "Código interno" Then TextoFiltro = "P.Desenho" Else TextoFiltro = "P.descricao"
        End If
        Sql_Programacao_Vendas_Localizar_produto = "Select " & CamposFiltro & " from " & INNERJOINTEXTO & " where VPI.ID = " & txtId & " and " & TextoFiltro & " Like '" & Txt_texto & "%' group by " & CamposFiltro & " order by VPI.Codigo"
    Else
        Sql_Programacao_Vendas_Localizar_produto = "Select " & CamposFiltro & " from Vendas_programa_item VPI INNER JOIN Projproduto P ON VPI.Codigo = P.Desenho where VPI.ID = " & txtId & " and Status_item = '" & Cmb_texto & "' order by VPI.Codigo"
    End If
Else
    Sql_Programacao_Vendas_Localizar_produto = "Select " & CamposFiltro & " from Vendas_programa_item VPI INNER JOIN Projproduto P ON VPI.Codigo = P.Desenho where VPI.ID = " & txtId & " order by VPI.Codigo"
End If
ProcCarregaLista_Item (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_texto_Change()
On Error GoTo tratar_erro

lista_item.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

txtdescricao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtID_cli_Change()
On Error GoTo tratar_erro

If txtID_cli <> "" Then
    VerifNumero = txtID_cli
    ProcVerificaNumero
    If VerifNumero = False Then
        txtID_cli = ""
        txtID_cli.SetFocus
        Exit Sub
    End If
End If
ProcCarregaCliente

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtID_cli_Click()
On Error GoTo tratar_erro

If txtID_cli = "0" Then txtID_cli = ""

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

Private Sub txtNreg1_Change()
On Error GoTo tratar_erro

If txtNreg1 <> "" Then
    VerifNumero = txtNreg1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg1 = ""
        txtNreg1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtPagIr1_Change()
On Error GoTo tratar_erro

If txtPagIr1 <> "" Then
    VerifNumero = txtPagIr1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr1 = ""
        txtPagIr1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtd_Change()
On Error GoTo tratar_erro

If txtQTD.Text <> "" Then
    VerifNumero = txtQTD.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQTD.Text = ""
        txtQTD.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtd_LostFocus()
On Error GoTo tratar_erro

txtQTD.Text = Format(txtQTD.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcEnviadados_Prog()
On Error GoTo tratar_erro

TBGravar!ID = txtId
TBGravar!Id_Item = txtID_item
TBGravar!Un = Txt_un
TBGravar!Unidade_com = Txt_un_com
TBGravar!Data_inicio = Format(txtData_inicio.Value, "dd/mm/yyyy")
TBGravar!Data_fim = Format(txtData_fim.Value, "dd/mm/yyyy")
TBGravar!quantidade = Format(txtQTD, "###,##0.0000")
TBGravar!PCCliente = Txt_pedido_cliente
TBGravar!N_item = Txt_n_item
TBGravar!Data_negociada = IIf(Txt_data_negociada = "__/__/____", Null, Txt_data_negociada)
If optFirme.Value = 1 Then
    TBGravar!Firme = True
    If TBGravar!Status_prog = "PREVISÃO FUTURA" Then
        TBGravar!Status_prog = "ABERTO"
        TBGravar!Ordenar = 2
    End If
Else
    TBGravar!Firme = False
    TBGravar!Status_prog = "PREVISÃO FUTURA"
    TBGravar!Ordenar = 3
    TBGravar!QtdeFaturada = 0
End If
If Chk_utiliza_mat_consignado.Value = 1 Then TBGravar!Utiliza_mat_cons = True Else TBGravar!Utiliza_mat_cons = False
If TBGravar!QtdeFaturada > 0 Then
    If TBGravar!quantidade > TBGravar!QtdeFaturada Then TBGravar!Status_prog = "PARCIAL" Else TBGravar!Status_prog = "FATURADO"
End If
txtStatus_prog = TBGravar!Status_prog
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcGerarPedido(ID_programacao As Long, Prazo_final As Date, Pedido_cliente As String, N_item As String, Codinterno As String, Descricao As String, Qtde As Double, Un As String, Un_com As String, Utiliza_mat_consignado As Boolean, Referencia As String)
On Error GoTo tratar_erro

IDCFOP = ""
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from Vendas_proposta where ID_programa = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = True Then
    
    Regime = FunVerifRegimeEmpresa(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
    If Regime = 1 Then
        'Verifica se existe mais de uma tabela do simples cadastrada
        TabelaSN_Prog = 0
        Contador = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                TabelaSN_Prog = TBAbrir!Tabela
                Contador = Contador + 1
                TBAbrir.MoveNext
            Loop
            If Contador > 1 Then
                USMsgBox ("Favor informar a tabela do simples nacional utilizada para esse pedido."), vbInformation, "CAPRIND v5.0"
                Vendas_Proposta = False
                Vendas_PI = False
                Vendas_Programacao = True
                frmVendas_proposta_tabelaSN.Show 1
            End If
        Else
            USMsgBox ("Não é permitido gerar pedido, pois não existe nenhuma tabela do simples nacional ativa."), vbExclamation, "CAPRIND v5.0"
            TBAbrir.Close
            Exit Sub
        End If
        TBAbrir.Close
    End If
            
    TBPedido.AddNew
    TBPedido!Regime = Regime
    TBPedido!TabelaSN = TabelaSN_Prog
    TBPedido!Data = Date
    TBPedido!Responsavel = pubUsuario
    TBPedido!status = "VENDIDA"
    
    'Gerar numero do pedido
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_proposta where Year(data) = '" & Year(Date) & "' order by Ordenarproposta", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir.MoveLast
        Cotacao = Left(TBAbrir!Ncotacao, Len(TBAbrir!Ncotacao) - 3) + 1
    Else
        Cotacao = 1
    End If
    Ano = Right(Year(Date), 2)
    Select Case Len(Cotacao)
        Case 1: NumeroCotacao = "000" & Cotacao & "/" & Ano
        Case 2: NumeroCotacao = "00" & Cotacao & "/" & Ano
        Case 3: NumeroCotacao = "0" & Cotacao & "/" & Ano
        Case 4: NumeroCotacao = Cotacao & "/" & Ano
        Case 5: NumeroCotacao = Cotacao & "/" & Ano
    End Select
    TBPedido!Ncotacao = NumeroCotacao
                
    TBPedido!ID_programa = txtId
    TBPedido!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBPedido!Revisao = 0
    TBPedido!IDCliente = txtID_cli
    
    'Gravar dados do cliente
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from Clientes where idcliente = " & txtID_cli, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        TBPedido!Cliente = TBClientes!NomeRazao
        TBPedido!Fax = IIf(IsNull(TBClientes!Fax), Null, TBClientes!Fax)
        TBPedido!Email = IIf(IsNull(TBClientes!Email), Null, TBClientes!Email)
        TBPedido!Tipo_endereco = IIf(IsNull(TBClientes!Tipo_endereco), Null, TBClientes!Tipo_endereco)
        TBPedido!Endereco = IIf(IsNull(TBClientes!Endereco), Null, TBClientes!Endereco)
        TBPedido!Numero = IIf(IsNull(TBClientes!Numero), Null, TBClientes!Numero)
        TBPedido!complemento = IIf(IsNull(TBClientes!complemento), Null, TBClientes!complemento)
        TBPedido!Tipo_bairro = IIf(IsNull(TBClientes!Tipo_bairro), Null, TBClientes!Tipo_bairro)
        TBPedido!Bairro = IIf(IsNull(TBClientes!Bairro), Null, TBClientes!Bairro)
        TBPedido!Cidade = IIf(IsNull(TBClientes!Cidade), Null, TBClientes!Cidade)
        TBPedido!telefone = IIf(IsNull(TBClientes!Tel01), Null, TBClientes!Tel01)
        TBPedido!UF = IIf(IsNull(TBClientes!UF), Null, TBClientes!UF)
        TBPedido!Tipo_cliente = IIf(IsNull(TBClientes!Tipo), Null, TBClientes!Tipo)
    End If
    
    'Grava CFOP se o mesmo existir no cadastro do cliente
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from Clientes_DadosComerciais where idcliente = " & txtID_cli & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        If IsNull(TBClientes!IDCFOP) = False And TBClientes!IDCFOP <> "" Then IDCFOP = TBClientes!IDCFOP
    End If
    TBPedido!Tipo = "PE"
    TBPedido!Datavendas = Date
    
    TBPedido.Update
    Conexao.Execute "Update Vendas_proposta set ordenarproposta = " & TBPedido!Cotacao & " where cotacao = " & TBPedido!Cotacao
End If
IDpedido = TBPedido!Cotacao

'Gravar produtos
Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from Vendas_carteira where cotacao = " & IDpedido & " and ID_programacao = " & ID_programacao, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = True Then
    TBCompras_Lista.AddNew
    TBCompras_Lista!ID_programacao = ID_programacao
    TBCompras_Lista!Liberacao = "VENDIDA"
    TBCompras_Lista!Datavendas = Date
    TBCompras_Lista!Tem_ordem = False
    Novo = True
Else
    Novo = False
End If
TBCompras_Lista!Cotacao = IDpedido
TBCompras_Lista!PrazoFinal = Prazo_final
TBCompras_Lista!PCCliente = Pedido_cliente
TBCompras_Lista!N_item = N_item
TBCompras_Lista!Desenho = Codinterno
TBCompras_Lista!N_referencia = Referencia
TBCompras_Lista!descricao_tecnica = Descricao
TBCompras_Lista!quantidade = Qtde
TBCompras_Lista!Utiliza_mat_cons = Utiliza_mat_consignado
TBCompras_Lista!Qtde_produzir = TBCompras_Lista!quantidade / FunVerificaTabelaConversaoUnidade(Un, Un_com)
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBCompras_Lista!Rev_codinterno = IIf(TBProduto!RevDesenho = "", 0, TBProduto!RevDesenho)
    TBCompras_Lista!Descricao = TBProduto!descricaotecnica
    
    'Código de referencia
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.Codproduto = " & TBProduto!Codproduto & " and IA.ID_cliente_forn = " & txtID_cli & " and IA.Tipo = 'C' and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = True Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.Codproduto = " & TBProduto!Codproduto & " and (IA.ID_cliente_forn = 0 or IA.ID_cliente_forn IS NULL) and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBFIltro.EOF = False Then
        TBCompras_Lista!N_referencia = TBFIltro!N_referencia
    End If
    TBFIltro.Close
    
    If Novo = True Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from Projproduto_clientes where Codproduto = " & TBProduto!Codproduto & " and idcliente = " & txtID_cli, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            If TBPedido!Tipo_cliente <> "JR" And TBPedido!Tipo_cliente <> "FR" Then
                TBCompras_Lista!preco_unitario = IIf(IsNull(TBFIltro!PConsumo), "", Format(TBFIltro!PConsumo / FunVerificaTabelaConversaoUnidade(Un, Un_com), "###,##0.0000000000"))
            Else
                TBCompras_Lista!preco_unitario = IIf(IsNull(TBFIltro!PRevenda), "", Format(TBFIltro!PRevenda / FunVerificaTabelaConversaoUnidade(Un, Un_com), "###,##0.0000000000"))
            End If
            If IsNull(TBFIltro!ID_CF) = False Then
                TBCompras_Lista!ID_CF = TBFIltro!ID_CF
            Else
                If IsNull(TBProduto!ID_CF) = False Then TBCompras_Lista!ID_CF = TBProduto!ID_CF
            End If
        Else
            If TBPedido!Tipo_cliente <> "JR" And TBPedido!Tipo_cliente <> "FR" Then
                TBCompras_Lista!preco_unitario = IIf(IsNull(TBProduto!PConsumo), "", Format(TBProduto!PConsumo / FunVerificaTabelaConversaoUnidade(Un, Un_com), "###,##0.0000000000"))
            Else
                TBCompras_Lista!preco_unitario = IIf(IsNull(TBProduto!PRevenda), "", Format(TBProduto!PRevenda / FunVerificaTabelaConversaoUnidade(Un, Un_com), "###,##0.0000000000"))
            End If
            If IsNull(TBProduto!ID_CF) = False Then TBCompras_Lista!ID_CF = TBProduto!ID_CF
        End If
        TBFIltro.Close
        
        TBCompras_Lista!preco_unitario_desconto = Format(TBCompras_Lista!preco_unitario, "0.0000000000")
        
        If IsNull(TBCompras_Lista!ID_CF) = False And IsNull(TBPedido!UF) = False And TBPedido!UF <> "" Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Carregar_CFOP_ST = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                ProcVerifCFOPST TBCompras_Lista!ID_CF, TBPedido!UF
                If Valido = True Then
                    IDCFOP = IDAntigo
                    TBCompras_Lista!txt_CST = Letra
                End If
            End If
        End If
        If IDCFOP = "" Then
            TBCompras_Lista!ID_CFOP = IIf(IsNull(TBProduto!ID_CFOP1), Null, TBProduto!ID_CFOP1)
        Else
            TBCompras_Lista!ID_CFOP = IDCFOP
        End If
    End If
        
    ProcAgregarProdutoCli TBProduto!Codproduto, txtID_cli, TBPedido!Tipo_cliente, Un, Un_com, TBCompras_Lista!preco_unitario
    
    TBCompras_Lista!Unidade = TBProduto!Unidade
    TBCompras_Lista!Unidade_com = TBProduto!Unidade_com
    TBCompras_Lista!Familia = TBProduto!Classe
    TBCompras_Lista!Comprimento = IIf(TBProduto!Comprimento = "", Null, TBProduto!Comprimento)
    TBCompras_Lista!Largura = IIf(TBProduto!Largura = "", Null, TBProduto!Largura)
    TBCompras_Lista!Espessura = IIf(TBProduto!Espessura = "", Null, TBProduto!Espessura)
    TBCompras_Lista!Dureza = TBProduto!Dureza
    TBCompras_Lista!Embalagem = TBProduto!Embalagem
    TBCompras_Lista!Inspecao = TBProduto!Inspecao
End If

TBCompras_Lista!preco_lote = Format(TBCompras_Lista!preco_unitario_desconto * TBCompras_Lista!quantidade, "###,##0.00")
TBCompras_Lista!Tipo = "P"
TBCompras_Lista!retorno = False

If IsNull(TBCompras_Lista!ID_CF) = False And IsNull(TBCompras_Lista!ID_CFOP) = False And TBCompras_Lista!ID_CFOP <> "" Then ProcAtualizavalores TBCompras_Lista!Cotacao, TBCompras_Lista!ID_CF, TBPedido!IDCliente, TBPedido!Cliente, TBPedido!UF, TBPedido!ID_empresa, TBCompras_Lista!ID_CFOP, IIf(IsNull(TBPedido!Regime), 0, TBPedido!Regime)

Regime = 0
If IsNull(TBCompras_Lista!ID_CFOP) = False And TBCompras_Lista!ID_CFOP <> "" Then
    ProcVerifImpostosEmpresa TBPedido!ID_empresa, False, "", False, 0, False, IIf(IsNull(TBPedido!TabelaSN), 0, TBPedido!TabelaSN), 0
    
    TBCompras_Lista!PIS_Prod = PIS_Prod
    If PIS_Prod <> 0 Then TBCompras_Lista!Total_PIS_prod = Format((TBCompras_Lista!preco_lote * PIS_Prod) / 100, "###,##0.00") Else TBCompras_Lista!Total_PIS_prod = 0
    TBCompras_Lista!Cofins_Prod = Cofins_Prod
    If Cofins_Prod <> 0 Then TBCompras_Lista!Total_Cofins_prod = Format((TBCompras_Lista!preco_lote * Cofins_Prod) / 100, "###,##0.00") Else TBCompras_Lista!Total_Cofins_prod = 0
    TBCompras_Lista!CSLL_Prod = CSLL_Prod
    If CSLL_Prod <> 0 Then TBCompras_Lista!Total_CSLL_prod = Format((TBCompras_Lista!preco_lote * CSLL_Prod) / 100, "###,##0.00") Else TBCompras_Lista!Total_CSLL_prod = 0
    TBCompras_Lista!IRPJ_Prod = IRPJ_Prod
    If IRPJ_Prod <> 0 Then TBCompras_Lista!Total_IRPJ_prod = Format((TBCompras_Lista!preco_lote * IRPJ_Prod) / 100, "###,##0.00") Else TBCompras_Lista!Total_IRPJ_prod = 0
    TBCompras_Lista!DAS = DAS
    If DAS <> 0 Then TBCompras_Lista!Total_DAS = Format((TBCompras_Lista!preco_lote * DAS) / 100, "###,##0.00") Else TBCompras_Lista!Total_DAS = 0
    
    If IsNull(TBCompras_Lista!ID_CF) = False Then
        ProcValorImposto TBPedido!Ncotacao, IIf(IsNull(TBCompras_Lista!ID_CF), 0, TBCompras_Lista!ID_CF), TBPedido!IDCliente, TBPedido!Cliente, TBPedido!UF, TBPedido!ID_empresa, False, IIf(IsNull(TBCompras_Lista!ID_CFOP), 0, TBCompras_Lista!ID_CFOP), IIf(IsNull(TBPedido!Regime), 0, TBPedido!Regime)
        ProcControleImposto IIf(IsNull(TBCompras_Lista!ID_CFOP), 0, TBCompras_Lista!ID_CFOP), TBPedido!IDCliente
        If TemIPI = "SIM" Then
            TBCompras_Lista!int_IPI = IntIPI
            TBCompras_Lista!dbl_ValorIPI = Format((TBCompras_Lista!preco_lote * IntIPI) / 100, "###,##0.00")
        Else
            TBCompras_Lista!int_IPI = 0
            TBCompras_Lista!dbl_ValorIPI = 0
        End If
        If TemICMS = "SIM" Then
            TBCompras_Lista!IntICMS = IntICMS
            TBCompras_Lista!dbl_Valor_ICMS = Format((TBCompras_Lista!preco_lote * IntICMS) / 100, "###,##0.00")
            If IntICMS > 0 Then
                ProcCalculaBC TBPedido!ID_empresa, TBCompras_Lista!ID_CFOP, 0, TBCompras_Lista!preco_lote, TBCompras_Lista!dbl_ValorIPI, SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBCompras_Lista!txt_CST), "", TBCompras_Lista!txt_CST), "P", 0, ""
                TBCompras_Lista!BC_ICMS = BC
            End If
        Else
            TBCompras_Lista!IntICMS = 0
            TBCompras_Lista!dbl_Valor_ICMS = 0
        End If
        
        If IsNull(TBCompras_Lista!txt_CST) = False And TBCompras_Lista!txt_CST <> "" Then
            ProcSubstituicaoTributaria TBPedido!UF, TBCompras_Lista!txt_CST, TBCompras_Lista!ID_CF, IIf(txtID_cli = "", 0, txtID_cli), txtCliente, TBCompras_Lista!preco_unitario_desconto, TBCompras_Lista!quantidade, BC, BCST, 0, 0, 0, False, False, 0
            TBCompras_Lista!Valor_ICMS_ST = ICMSCST
            If ICMSCST <> 0 Then TBCompras_Lista!BC_ICMS_ST = BCICMSCST
        End If
    End If
End If
TBCompras_Lista.Update
IDlista = TBCompras_Lista!CODIGO

Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from Vendas_carteira where Codigo = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = False Then
    If IsNull(TBCompras_Lista!ID_CF) = False Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from tbl_classificacaofiscal where Idclass = " & TBCompras_Lista!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            'Verifica se a CF tem retenção de PIS/Cofins, destaca PIS/Cofins e grava no produto
            If TBFI!Retem_PIS_Cofins = True Then
                TBCompras_Lista!Valor_Retencao_PIS = Format((TBCompras_Lista!preco_lote * IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)) / 100, "###,##0.00")
                TBCompras_Lista!Valor_Retencao_Cofins = Format((TBCompras_Lista!preco_lote * IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)) / 100, "###,##0.00")
            End If
            
            If Regime <> 0 And Regime <> 1 Then
                PIS_Prod = IIf(IsNull(TBFI!PIS_destaca), 0, TBFI!PIS_destaca)
                Cofins_Prod = IIf(IsNull(TBFI!Cofins_destaca), 0, TBFI!Cofins_destaca)
                If PIS_Prod <> 0 Then
                    TBCompras_Lista!PIS_Prod = PIS_Prod
                    TBCompras_Lista!Total_PIS_prod = Format((TBCompras_Lista!preco_lote * PIS_Prod) / 100, "###,##0.00")
                End If
                If Cofins_Prod <> 0 Then
                    TBCompras_Lista!Cofins_Prod = Cofins_Prod
                    TBCompras_Lista!Total_Cofins_prod = Format((TBCompras_Lista!preco_lote * Cofins_Prod) / 100, "###,##0.00")
                End If
            End If
        End If
        TBFI.Close
    End If
    TBCompras_Lista.Update
    
    ProcExcluirEmpenhos Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBCompras_Lista!CODIGO, True
    QuantSolicitado = TBCompras_Lista!Qtde_produzir
    ProcEmpenharProdEstoque Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBCompras_Lista!CODIGO, TBCompras_Lista!Desenho, True, False, TBCompras_Lista!Qtde_produzir
    If QuantSolicitado > 0 Then ProcEmpenharProdProduzindo Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBCompras_Lista!CODIGO, TBCompras_Lista!Desenho, TBCompras_Lista!PrazoFinal, True
End If

'Grava dados comerciais se o mesmo existir no cadastro do cliente
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Clientes_DadosComerciais where idcliente = " & txtID_cli & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select * from vendas_comercial where cotacao = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBPedido.EOF = True Then TBPedido.AddNew
    TBPedido!Cotacao = IDpedido
    TBPedido!calculos = TBClientes!calculos
    TBPedido!impostos = TBClientes!impostos
    TBPedido!condicoes = TBClientes!condicoes
    TBPedido!garantia = TBClientes!garantia
    TBPedido!reajuste = TBClientes!reajuste
    TBPedido!transporte = TBClientes!transporte
    TBPedido!validade = TBClientes!validade
    TBPedido.Update
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizavalores(Cotacao As String, ID_CF As Long, IDCliente As Long, Cliente As String, UF As String, IDempresa As Integer, IDCFOP As Long, Regime As Integer)
On Error GoTo tratar_erro

ProcValorImposto Cotacao, ID_CF, IDCliente, Cliente, UF, IDempresa, False, IDCFOP, Regime
ProcControleImposto IDCFOP, IDCliente
ProcCalculaValores IIf(IsNull(TBCompras_Lista!quantidade), 0, TBCompras_Lista!quantidade), IIf(IsNull(TBCompras_Lista!preco_unitario_desconto), 0, TBCompras_Lista!preco_unitario_desconto), IDempresa, IDCFOP, IIf(IsNull(TBCompras_Lista!preco_unitario_desconto), 0, TBCompras_Lista!preco_unitario_desconto) * IIf(IsNull(TBCompras_Lista!quantidade), 0, TBCompras_Lista!quantidade), IDCliente, ID_CF

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCalculaValores(quantidade As Double, Valorunitario As Double, IDempresa As Integer, IDCFOP As Long, ValorTotal As Double, IDCliente As Long, ID_CF As Long)
On Error GoTo tratar_erro

If quantidade = 0 Or Valorunitario = 0 Then Exit Sub

'Zera valores
SumICMS = 0
SumIPI = 0
SumTotNota = 0
SumTotProdutos = 0
VlrIPI = 0

'Se tem IPI
If TemIPI = "SIM" Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Clientes_Impostos where IDCliente = " & IDCliente & " and ID_CF = " & ID_CF, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        VlrIPI = Valorunitario
        If TBFIltro!PorcentagemIPI <> 0 Then VlrIPI = VlrIPI / TBFIltro!PorcentagemIPI
        VlrIPI = Format((VlrIPI - Valorunitario) * quantidade, "###,##0.00")
    Else
        VlrIPI = Format((Valorunitario * quantidade) * IntIPI / 100, "###,##0.00")
    End If
    TBFIltro.Close
End If
TBCompras_Lista!int_IPI = IntIPI
TBCompras_Lista!dbl_ValorIPI = Format(VlrIPI, "###,##0.00")
   
'Se tem icms
If TemICMS = "SIM" Then
    TBCompras_Lista!IntICMS = IntICMS
    
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select id_CFOP from tbl_NaturezaOperacao where IDCountCfop =  " & TBCompras_Lista!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Familiatext = TBCFOP!ID_CFOP
    End If
    TBCFOP.Close
    
    ProcCalculaBC IDempresa, Familiatext, 0, ValorTotal, VlrIPI, SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBCompras_Lista!txt_CST), "", TBCompras_Lista!txt_CST), "P", 0, ""
    TBCompras_Lista!dbl_Valor_ICMS = Format((BC * IntICMS) / 100, "###,##0.00")
Else
    TBCompras_Lista!IntICMS = 0
    TBCompras_Lista!dbl_Valor_ICMS = 0
End If
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcGravarTotaisPedido()
On Error GoTo tratar_erro

'PRODUTOS
TotalProduto = 0
TotalDesconto = 0
SumIPI = 0
TotalICMSCST = 0
BASECALCULO = 0
TotalICMS = 0
TotalBCICMSCST = 0
Valor_PIS_Prod = 0
Valor_Cofins_Prod = 0
Valor_CSLL_Prod = 0
Valor_IRPJ_Prod = 0
Valor_DAS = 0
Valor_Retencao_PIS = 0
Valor_Retencao_Cofins = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
CamposFiltro = "Sum(VC.preco_unitario * VC.Quantidade) as TotalProduto, Sum(VC.ValorDesconto * VC.Quantidade) as TotalDesconto, Sum(VC.dbl_valoripi) as SumIPI, Sum(VC.Valor_ICMS_ST) as TotalICMSCST, Sum(VC.BC_ICMS) as BASECALCULO, Sum(VC.dbl_Valor_ICMS) as TotalICMS, Sum(VC.BC_ICMS_ST) as TotalBCICMSCST, Sum(VC.Total_PIS_prod) as Valor_PIS_Prod, Sum(VC.Total_Cofins_prod) as Valor_Cofins_Prod, Sum(VC.Total_CSLL_prod) as Valor_CSLL_Prod, Sum(VC.Total_IRPJ_prod) as Valor_IRPJ_Prod, Sum(VC.Total_DAS) as Valor_DAS, Sum(VC.Valor_Retencao_PIS) as Valor_Retencao_PIS, Sum(VC.Valor_Retencao_Cofins) as Valor_Retencao_Cofins"
TBTotaisnota.Open "Select " & CamposFiltro & " from vendas_carteira VC INNER JOIN vendas_proposta VP on VP.Cotacao = VC.Cotacao where VP.cotacao = " & IDpedido & " and VC.Tipo = 'P' and VC.Retorno = 'False' and (VP.Tipo = 'PE' or VP.Tipo = 'PRPE') and (VC.liberacao = 'VENDIDA' or VC.liberacao = 'REVISADA' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    TotalProduto = IIf(IsNull(TBTotaisnota!TotalProduto), 0, TBTotaisnota!TotalProduto)
    TotalDesconto = IIf(IsNull(TBTotaisnota!TotalDesconto), 0, TBTotaisnota!TotalDesconto)
    SumIPI = IIf(IsNull(TBTotaisnota!SumIPI), 0, TBTotaisnota!SumIPI)
    TotalICMSCST = IIf(IsNull(TBTotaisnota!TotalICMSCST), 0, TBTotaisnota!TotalICMSCST)
    BASECALCULO = IIf(IsNull(TBTotaisnota!BASECALCULO), 0, TBTotaisnota!BASECALCULO)
    TotalICMS = IIf(IsNull(TBTotaisnota!TotalICMS), 0, TBTotaisnota!TotalICMS)
    TotalBCICMSCST = IIf(IsNull(TBTotaisnota!TotalBCICMSCST), 0, TBTotaisnota!TotalBCICMSCST)
    Valor_PIS_Prod = IIf(IsNull(TBTotaisnota!Valor_PIS_Prod), 0, TBTotaisnota!Valor_PIS_Prod)
    Valor_Cofins_Prod = IIf(IsNull(TBTotaisnota!Valor_Cofins_Prod), 0, TBTotaisnota!Valor_Cofins_Prod)
    Valor_CSLL_Prod = IIf(IsNull(TBTotaisnota!Valor_CSLL_Prod), 0, TBTotaisnota!Valor_CSLL_Prod)
    Valor_IRPJ_Prod = IIf(IsNull(TBTotaisnota!Valor_IRPJ_Prod), 0, TBTotaisnota!Valor_IRPJ_Prod)
    Valor_DAS = IIf(IsNull(TBTotaisnota!Valor_DAS), 0, TBTotaisnota!Valor_DAS)
    Valor_Retencao_PIS = IIf(IsNull(TBTotaisnota!Valor_Retencao_PIS), 0, TBTotaisnota!Valor_Retencao_PIS)
    Valor_Retencao_Cofins = IIf(IsNull(TBTotaisnota!Valor_Retencao_Cofins), 0, TBTotaisnota!Valor_Retencao_Cofins)
End If
'Retorno
VlrTotalRetorno = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select Sum(VC.preco_unitario * VC.Quantidade) as VlrTotalRetorno from vendas_carteira VC INNER JOIN vendas_proposta VP on VP.Cotacao = VC.Cotacao where VP.cotacao = " & IDpedido & " and VC.Tipo = 'P' and VC.Retorno = 'False' and (VP.Tipo = 'PE' or VP.Tipo = 'PRPE') and (VC.liberacao = 'VENDIDA' or VC.liberacao = 'REVISADA' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    VlrTotalRetorno = IIf(IsNull(TBTotaisnota!VlrTotalRetorno), 0, TBTotaisnota!VlrTotalRetorno)
End If

'SERVIÇOS
TotalServicos = 0
TotalDescontoServico = 0
Valor_PIS_Serv = 0
Valor_Cofins_Serv = 0
Valor_CSLL_Serv = 0
TotalISS = 0
Valor_INSS_Serv = 0
Valor_IRPJ_Serv = 0
Valor_IRRF_Serv = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
CamposFiltro = "Sum(VC.preco_unitario * VC.Quantidade) as TotalServicos, Sum(VC.ValorDesconto * VC.Quantidade) as TotalDescontoServico, Sum(VC.Total_PIS_serv) as Valor_PIS_Serv, Sum(VC.Total_Cofins_serv) as Valor_Cofins_Serv, Sum(VC.Total_CSLL_serv) as Valor_CSLL_Serv, Sum(VC.vlriss) as TotalISS, Sum(VC.Total_INSS_serv) as Valor_INSS_Serv, Sum(VC.Total_IRPJ_serv) as Valor_IRPJ_Serv, Sum(VC.Total_IRRF_serv) as Valor_IRRF_Serv"
TBTotaisnota.Open "Select " & CamposFiltro & " from vendas_carteira VC INNER JOIN vendas_proposta VP on VP.Cotacao = VC.Cotacao where VP.cotacao = " & IDpedido & " and VC.Tipo = 'S' and VC.Retorno = 'False' and (VP.Tipo = 'PE' or VP.Tipo = 'PRPE') and (VC.liberacao = 'VENDIDA' or VC.liberacao = 'REVISADA' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    TotalServicos = IIf(IsNull(TBTotaisnota!TotalServicos), 0, TBTotaisnota!TotalServicos)
    TotalDescontoServico = IIf(IsNull(TBTotaisnota!TotalDescontoServico), 0, TBTotaisnota!TotalDescontoServico)
    Valor_PIS_Serv = IIf(IsNull(TBTotaisnota!Valor_PIS_Serv), 0, TBTotaisnota!Valor_PIS_Serv)
    Valor_Cofins_Serv = IIf(IsNull(TBTotaisnota!Valor_Cofins_Serv), 0, TBTotaisnota!Valor_Cofins_Serv)
    Valor_CSLL_Serv = IIf(IsNull(TBTotaisnota!Valor_CSLL_Serv), 0, TBTotaisnota!Valor_CSLL_Serv)
    TotalISS = IIf(IsNull(TBTotaisnota!TotalISS), 0, TBTotaisnota!TotalISS)
    Valor_INSS_Serv = IIf(IsNull(TBTotaisnota!Valor_INSS_Serv), 0, TBTotaisnota!Valor_INSS_Serv)
    Valor_IRPJ_Serv = IIf(IsNull(TBTotaisnota!Valor_IRPJ_Serv), 0, TBTotaisnota!Valor_IRPJ_Serv)
    Valor_IRRF_Serv = IIf(IsNull(TBTotaisnota!Valor_IRRF_Serv), 0, TBTotaisnota!Valor_IRRF_Serv)
End If
TBTotaisnota.Close

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from vendas_proposta where cotacao = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = True Then TBPedido.AddNew
TBPedido!dbl_Base_ICMS = Format(BASECALCULO, "###,##0.00")
TBPedido!dbl_Valor_ICMS = Format(TotalICMS, "###,##0.00")
TBPedido!dbl_Base_ICMS_Subst = Format(TotalBCICMSCST, "###,##0.00")
TBPedido!dbl_Valor_ICMS_Subst = Format(TotalICMSCST, "###,##0.00")
TBPedido!dbl_Valor_Total_Produtos = Format(TotalProduto, "###,##0.00")
TBPedido!dbl_valor_total_servicos = Format(TotalServicos, "###,##0.00")
TBPedido!TotalDesconto = Format(TotalDesconto + TotalDescontoServico, "###,##0.00")
TBPedido!dbl_Valor_Total_IPI = Format(SumIPI, "###,##0.00")

'Impostos produtos
TBPedido!Total_PIS_prod = Format(Valor_PIS_Prod, "###,##0.00")
TBPedido!Total_Cofins_prod = Format(Valor_Cofins_Prod, "###,##0.00")
TBPedido!Total_CSLL_prod = Format(Valor_CSLL_Prod, "###,##0.00")
TBPedido!Total_IRPJ_prod = Format(Valor_IRPJ_Prod, "###,##0.00")

'Impostos serviços
TBPedido!Total_PIS_serv = Format(Valor_PIS_Serv, "###,##0.00")
TBPedido!Total_Cofins_serv = Format(Valor_Cofins_Serv, "###,##0.00")
TBPedido!Total_CSLL_serv = Format(Valor_CSLL_Serv, "###,##0.00")
TBPedido!VlrTotaliss = Format(TotalISS)
TBPedido!Total_INSS_serv = Format(Valor_INSS_Serv, "###,##0.00")
TBPedido!Total_IRPJ_serv = Format(Valor_IRPJ_Serv, "###,##0.00")
TBPedido!Total_IRRF_serv = Format(Valor_IRRF_Serv, "###,##0.00")

SubTotal = Format(TotalProduto + TotalServicos - (TotalDesconto + TotalDescontoServico), "###,##0.00")
TBPedido!SubTotal = Format(SubTotal, "###,##0.00")

'Impostos faturamento
TBPedido!Total_DAS = Format(Valor_DAS, "###,##0.00")

'Retenção de PIS/Cofins
TBPedido!Total_retencao_PIS = Format(Valor_Retencao_PIS, "###,##0.00")
TBPedido!Total_retencao_Cofins = Format(Valor_Retencao_Cofins, "###,##0.00")

TBPedido!dbl_valor_total = Format(SubTotal + TotalIPI + TotalICMSCST, "###,##0.00")
TBPedido!Total_retorno = Format(VlrTotalRetorno, "###,##0.00")
TBPedido.Update
TBPedido.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerifExcluirPedido(TextoFiltro As String)
On Error GoTo tratar_erro

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select VP.cotacao, VP.dbl_valor_total, VC.Codigo from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.Cotacao = VP.Cotacao where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
    ProcExcluirEmpenhos Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBPedido!CODIGO, True
    Conexao.Execute "DELETE vendas_carteira where Codigo = " & TBPedido!CODIGO
    Conexao.Execute "DELETE vendas_carteira_alteracoes where ID_carteira = " & TBPedido!CODIGO
    Conexao.Execute "DELETE vendas_carteira_composicao where ID_carteira = " & TBPedido!CODIGO
    
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select Codigo from vendas_carteira where Cotacao = " & TBPedido!Cotacao, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = True Then
        Conexao.Execute "DELETE vendas_proposta where Cotacao = " & TBPedido!Cotacao
        Conexao.Execute "DELETE vendas_comercial where Cotacao = " & TBPedido!Cotacao
    Else
        IDpedido = TBPedido!Cotacao
        ProcGravarTotaisPedido
        FunAtualizaStatusPropPI (IDpedido)
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregalista_Prog()
On Error GoTo tratar_erro

lista_prog.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Vendas_programacao where id = " & txtId & " and Id_item = " & txtID_item & " order by Ordenar, data_inicio desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista2.Min = 0
    PBLista2.Max = TBLISTA.RecordCount
    PBLista2.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    With lista_prog.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID_prog
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data_inicio), "", Format(TBLISTA!Data_inicio, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data_fim), "", Format(TBLISTA!Data_fim, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!quantidade), "0,0000", Format(TBLISTA!quantidade, "###,##0.0000"))
            .Item(.Count).SubItems(4) = IIf(TBLISTA!Firme = True, "Sim", "Não")
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Status_prog), "", TBLISTA!Status_prog)
            
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select Vendas_proposta.Ncotacao, vendas_carteira.PCCliente from Vendas_proposta INNER JOIN vendas_carteira ON Vendas_proposta.cotacao = vendas_carteira.cotacao where vendas_carteira.ID_programacao = " & TBLISTA!ID_prog, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                .Item(.Count).SubItems(6) = IIf(IsNull(TBPedido!Ncotacao), "", TBPedido!Ncotacao)
            End If
            TBPedido.Close
            
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!PCCliente), "", TBLISTA!PCCliente)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!N_item), "", TBLISTA!N_item)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Data_negociada), "", Format(TBLISTA!Data_negociada, "dd/mm/yy"))
            
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista2.Value = Contador
        Loop
    End With
End If
TBLISTA.Close

Txt_qtde_total_estoque = Format(FunVerificaQtdeEstoque(txtCodigo, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "") * FunVerificaTabelaConversaoUnidade(Txt_un, Txt_un_com), "###,##0.0000")
Txt_qtde_total_prev = Format(FunVerificaQtdeTotalPrev(txtID_item), "###,##0.0000")
valor = FunVerificaQtdeTotalVend(txtID_item)
Txt_qtde_total_vendida = Format(valor, "###,##0.0000")
Valor1 = FunVerificaQtdeTotalFaturada(txtID_item)
Txt_qtde_total_faturada_rev = Format(FunVerificaQtdeTotalFaturada(txtID_item), "###,##0.0000")
Txt_saldo = Format(valor - Valor1, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaListaNF()
On Error GoTo tratar_erro

valor = 0
lista_nota.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Sum(NFPP.Quantidade) as Qtde, NF.ID, NF.int_NotaFiscal, NF.dt_DataEmissao from ((((tbl_Detalhes_Nota_pedidos NFPP INNER JOIN vendas_carteira VC ON NFPP.ID_carteira = VC.codigo and NFPP.Codinterno = VC.Desenho) INNER JOIN Vendas_Programacao VPR ON VPR.ID_prog = VC.ID_programacao) INNER JOIN Vendas_programa_item VPI ON VPI.ID_item = VPR.ID_item) INNER JOIN Vendas_programa VP ON VP.ID = VPI.ID) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFPP.ID_nota where VPI.codigo = '" & txtCodigo & "' and VP.Programatexto = '" & txtPrograma & "' and NF.Int_status = 1 group by NF.ID, NF.int_NotaFiscal, NF.dt_DataEmissao order by NF.dt_DataEmissao desc, NF.int_NotaFiscal desc", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista3.Min = 0
    PBLista3.Max = TBLISTA.RecordCount
    PBLista3.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With lista_nota.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!int_NotaFiscal), "", TBLISTA!int_NotaFiscal)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!dt_DataEmissao), "", Format(TBLISTA!dt_DataEmissao, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Qtde), "0,0000", Format(TBLISTA!Qtde, "###,##0.0000"))
            valor = valor + IIf(IsNull(TBLISTA!Qtde), 0, Format(TBLISTA!Qtde, "###,##0.0000"))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista3.Value = Contador
    Loop
End If
TBLISTA.Close
Txt_qtde_total_faturada = Format(valor, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function FunVerificaQtdeTotalPrev(Id_Item As Long) As Double
On Error GoTo tratar_erro

FunVerificaQtdeTotalPrev = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select Qtde from Vendas_programacao_qtde_prevista where Id_item = " & Id_Item, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    FunVerificaQtdeTotalPrev = IIf(IsNull(TBFIltro!Qtde), 0, TBFIltro!Qtde)
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Function FunVerificaQtdeTotalVend(Id_Item As Long) As Double
On Error GoTo tratar_erro

FunVerificaQtdeTotalVend = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select Qtde from Vendas_programacao_qtde_vendida where Id_item = " & Id_Item, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    FunVerificaQtdeTotalVend = IIf(IsNull(TBFIltro!Qtde), 0, TBFIltro!Qtde)
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerificaQtdeTotalFaturada(Id_Item As Long) As Double
On Error GoTo tratar_erro

FunVerificaQtdeTotalFaturada = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select Qtdefaturada from Vendas_programacao_qtde_faturada where Id_item = " & Id_Item, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    FunVerificaQtdeTotalFaturada = IIf(IsNull(TBFIltro!QtdeFaturada), 0, TBFIltro!QtdeFaturada)
End If
TBFIltro.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcAltera_Status(ID_prod As Long, ID_programa As Long)
On Error GoTo tratar_erro

'Prooduto
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from Vendas_programa_item where id_item = " & ID_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Vendas_programacao where id_item = " & ID_prod & " and status_prog <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!Status_Item = "ABERTO"
        txtstatus_item = "ABERTO"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_programacao where id_item = " & ID_prod & " and status_prog <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!Status_Item = "PREVISÃO FUTURA"
            txtstatus_item = "PREVISÃO FUTURA"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Vendas_programacao where id_item = " & ID_prod & " and status_prog <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!Status_Item = "FATURADO"
                txtstatus_item = "FATURADO"
            Else
                TBItem!Status_Item = "PARCIAL"
                txtstatus_item = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    TBItem.Update
End If

'Programa
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from Vendas_programa where id = " & ID_programa, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Vendas_programa_item where id = " & ID_programa & " and Status_Item <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!status = "ABERTO"
        txtStatus = "ABERTO"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Vendas_programa_item where id = " & ID_programa & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!status = "PREVISÃO FUTURA"
            txtStatus = "PREVISÃO FUTURA"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Vendas_programa_item where id = " & ID_programa & " and Status_Item <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!status = "FATURADO"
                txtStatus = "FATURADO"
            Else
                TBItem!status = "PARCIAL"
                txtStatus = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    TBItem.Update
End If
TBItem.Close
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If Emitir_PI = False Then ProcFiltrarListaProdutos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaCliente()
On Error GoTo tratar_erro

txtCliente = ""
If txtID_cli = "" Then Exit Sub
Set TBClientes = CreateObject("adodb.recordset")
If Novo_Programacao_Vendas = True Then TextoFiltro = "idcliente = " & txtID_cli & " and status <> 'Bloqueado'" Else TextoFiltro = "idcliente = " & txtID_cli
TBClientes.Open "Select * from Clientes where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    txtCliente = IIf(IsNull(TBClientes!NomeRazao), "", TBClientes!NomeRazao)
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    Case 8: ProcRevisar
    Case 9: ProcEmitirPI
    'Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procNovo_item
    Case 2: procSalvar_item
    Case 3: procExcluir_item
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcEmitirPI_item
    'Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_prog
    Case 2: ProcSalvar_prog
    Case 3: ProcExcluir_prog
    Case 4: ProcImprimir
    Case 5: ProcAnterior_prog
    Case 6: ProcProximo_prog
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procPuxadadosItem()
On Error GoTo tratar_erro

txtID_item = TBProduto!Id_Item
txtCodigo = IIf(IsNull(TBProduto!CODIGO), "", TBProduto!CODIGO)
txtdescricao = lista_item.SelectedItem.ListSubItems(3)
txtstatus_item = lista_item.SelectedItem.ListSubItems(4)
If TBProduto!Validado = "Sim" Then Chk_validado.Value = 1 Else Chk_validado.Value = 0

If IsNull(TBProduto!Referencia) = False And TBProduto!Referencia <> "" Then
    NomeCampo = "o código de referência"
    cmbReferencia.AddItem TBProduto!Referencia
    cmbReferencia = TBProduto!Referencia
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o código de referência deste serviço."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
