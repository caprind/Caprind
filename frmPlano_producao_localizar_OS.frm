VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlano_producao_localizar_OS 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Plano da produção - Localizar OS's"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4500
      Top             =   270
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmPlano_producao_localizar_OS.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   28
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
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
      ButtonCaption2  =   "Adicionar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Adicionar (F3)"
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
      ButtonWidth2    =   52
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   94
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonUseMaskColor3=   0   'False
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   98
      ButtonTop4      =   2
      ButtonWidth4    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   136
      ButtonTop5      =   2
      ButtonWidth5    =   26
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
      ButtonLeft6     =   164
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do material"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   52
      Top             =   1320
      Width           =   11685
      Begin VB.TextBox Txt_un 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10980
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   430
         Width           =   525
      End
      Begin VB.TextBox Txt_descricao 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   430
         Width           =   9135
      End
      Begin VB.TextBox Txt_cod_interno 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   430
         Width           =   1635
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un."
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
         Left            =   11107
         TabIndex        =   55
         Top             =   240
         Width           =   270
      End
      Begin VB.Label Label6 
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
         Left            =   5985
         TabIndex        =   54
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label7 
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
         Left            =   382
         TabIndex        =   53
         Top             =   240
         Width           =   1230
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7680
      Left            =   60
      TabIndex        =   29
      Top             =   990
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   13547
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
      TabCaption(0)   =   "Lista de materiais"
      TabPicture(0)   =   "frmPlano_producao_localizar_OS.frx":2D84
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PBLista_mat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lista_material"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Lista de OS's"
      TabPicture(1)   =   "frmPlano_producao_localizar_OS.frx":2DA0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PBLista"
      Tab(1).Control(1)=   "Lista"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   -74945
         TabIndex        =   48
         Top             =   1230
         Width           =   11685
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   6720
            TabIndex        =   56
            Top             =   210
            WhatsThisHelpID =   210
            Width           =   4785
            Begin VB.OptionButton optIgual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               Height          =   255
               Left            =   3930
               TabIndex        =   7
               Top             =   180
               Width           =   705
            End
            Begin VB.OptionButton Optmeio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               Height          =   255
               Left            =   1470
               TabIndex        =   5
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton Optinicio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início frase"
               Height          =   255
               Left            =   180
               TabIndex        =   4
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Optfim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               Height          =   255
               Left            =   2760
               TabIndex        =   6
               Top             =   180
               Width           =   1155
            End
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
            Left            =   2970
            TabIndex        =   20
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   3675
         End
         Begin VB.ComboBox cmbfiltrarpor 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmPlano_producao_localizar_OS.frx":2DBC
            Left            =   180
            List            =   "frmPlano_producao_localizar_OS.frx":2DD5
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2775
         End
         Begin VB.Label Label5 
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
            Left            =   1147
            TabIndex        =   50
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label1 
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
            Left            =   4072
            TabIndex        =   49
            Top             =   180
            Width           =   1470
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74945
         TabIndex        =   44
         Top             =   6690
         Width           =   11685
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
            Left            =   5670
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
            Left            =   3210
            TabIndex        =   21
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   7890
            TabIndex        =   26
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":2E2C
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
            Left            =   7350
            TabIndex        =   25
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":65D3
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
            Left            =   6240
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
            Left            =   6810
            TabIndex        =   24
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":A0E2
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
            Left            =   8430
            TabIndex        =   27
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":E1D3
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
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de reg.: 0"
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
            TabIndex        =   47
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pág.: 0 de: 0"
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
            Left            =   10290
            TabIndex        =   46
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               reg. p/ pág."
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
            Left            =   2520
            TabIndex        =   45
            Top             =   240
            Width           =   2190
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   55
         TabIndex        =   40
         Top             =   6690
         Width           =   11685
         Begin VB.TextBox txtPagIr_mat 
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
            Left            =   5670
            TabIndex        =   12
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg_mat 
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
            Left            =   3210
            TabIndex        =   11
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx_mat 
            Height          =   315
            Left            =   7890
            TabIndex        =   17
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":11A61
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
         Begin DrawSuite2022.USButton cmdPagAnt_mat 
            Height          =   315
            Left            =   7350
            TabIndex        =   16
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":15208
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
         Begin DrawSuite2022.USButton cmdPagIr_mat 
            Height          =   315
            Left            =   6240
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
         Begin DrawSuite2022.USButton cmdPagPrim_mat 
            Height          =   315
            Left            =   6810
            TabIndex        =   15
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":18D15
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
         Begin DrawSuite2022.USButton cmdPagUlt_mat 
            Height          =   315
            Left            =   8430
            TabIndex        =   18
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmPlano_producao_localizar_OS.frx":1CE07
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
         Begin VB.Label lblRegistros_mat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de reg.: 0"
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
            TabIndex        =   43
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblPaginas_mat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pág.: 0 de: 0"
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
            Left            =   10290
            TabIndex        =   42
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               reg. p/ pág."
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
            Left            =   2520
            TabIndex        =   41
            Top             =   240
            Width           =   2190
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   55
         TabIndex        =   37
         Top             =   1230
         Width           =   11685
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   6720
            TabIndex        =   57
            Top             =   210
            WhatsThisHelpID =   210
            Width           =   4785
            Begin VB.OptionButton optFim_mat 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               Height          =   255
               Left            =   2760
               TabIndex        =   61
               Top             =   180
               Width           =   1155
            End
            Begin VB.OptionButton optInicio_mat 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início frase"
               Height          =   255
               Left            =   180
               TabIndex        =   60
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton optMeio_mat 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               Height          =   255
               Left            =   1470
               TabIndex        =   59
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optIgual_mat 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               Height          =   255
               Left            =   3930
               TabIndex        =   58
               Top             =   180
               Width           =   705
            End
         End
         Begin VB.ComboBox cmbfiltrarpor_mat 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmPlano_producao_localizar_OS.frx":20695
            Left            =   180
            List            =   "frmPlano_producao_localizar_OS.frx":206B1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2565
         End
         Begin VB.TextBox txtTexto_mat 
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
            Left            =   2760
            TabIndex        =   8
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   3855
         End
         Begin VB.ComboBox Cmb_texto_mat 
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
            Left            =   2760
            MousePointer    =   99  'Custom
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   3855
         End
         Begin VB.Label Label3 
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
            Left            =   1042
            TabIndex        =   39
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label2 
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
            Left            =   3982
            TabIndex        =   38
            Top             =   180
            Width           =   1470
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Legenda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   -74940
         TabIndex        =   30
         Top             =   1320
         Width           =   10875
         Begin VB.Image Image1 
            Height          =   720
            Left            =   120
            Picture         =   "frmPlano_producao_localizar_OS.frx":20718
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUTO FINAL"
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
            Left            =   840
            TabIndex        =   34
            Top             =   540
            Width           =   1335
         End
         Begin VB.Image Image2 
            Height          =   720
            Left            =   2910
            Picture         =   "frmPlano_producao_localizar_OS.frx":223E2
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUBCONJUNTO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3630
            TabIndex        =   33
            Top             =   540
            Width           =   1200
         End
         Begin VB.Image Image3 
            Height          =   720
            Left            =   5520
            Picture         =   "frmPlano_producao_localizar_OS.frx":240AC
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMPONENTE"
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
            Height          =   195
            Left            =   6390
            TabIndex        =   32
            Top             =   510
            Width           =   1095
         End
         Begin VB.Image Image4 
            Height          =   720
            Left            =   8370
            Picture         =   "frmPlano_producao_localizar_OS.frx":25D76
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MATÉRIA-PRIMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   9210
            TabIndex        =   31
            Top             =   510
            Width           =   1425
         End
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   -70950
         Top             =   540
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmPlano_producao_localizar_OS.frx":27A40
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74940
         TabIndex        =   35
         Top             =   330
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   1720
         ButtonCount     =   3
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Ajuda"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Ajuda (F1)"
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
         ButtonCaption2  =   "Sair"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Sair (Esc)"
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
         ButtonWidth2    =   26
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonKey3      =   "3"
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   5
         ButtonLeft3     =   68
         ButtonTop3      =   2
         ButtonWidth3    =   24
         ButtonHeight3   =   24
         ButtonUseMaskColor3=   0   'False
      End
      Begin MSComctlLib.ListView Lista_material 
         Height          =   4575
         Left            =   60
         TabIndex        =   10
         Top             =   2100
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   8070
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   0
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
            Object.Width           =   7576
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Dimensão"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Dureza"
            Object.Width           =   2117
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista_mat 
         Height          =   255
         Left            =   60
         TabIndex        =   36
         Top             =   7320
         Width           =   11685
         _ExtentX        =   20611
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
         Height          =   4575
         Left            =   -74940
         TabIndex        =   13
         Top             =   2100
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   8070
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "OS"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Versão"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Fase"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "D"
            Text            =   "Tempo total"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Posto de trab."
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "D"
            Text            =   "Pr. final"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Ordem"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Cód. de ref."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3784
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   -74940
         TabIndex        =   51
         Top             =   7320
         Width           =   11685
         _ExtentX        =   20611
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
End
Attribute VB_Name = "frmPlano_producao_localizar_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSqlLocProdPadraoMat As String 'OK

Private Sub Cmb_texto_mat_Click()
On Error GoTo tratar_erro

ProcLimpaCamposMat

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If txtTexto <> "" And (cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS") Then
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

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.Classe, P.Comprimento, P.Largura, P.Espessura, P.Dureza"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto"
    TextoFiltroPadrao = "P.Tipo = 'P' and P.Compras = 'True' and P.bloqueado = 'False' and P.Subtipoitem = 0 group by " & CamposFiltro & " order by P.Desenho"
    
    If txtTexto_mat.Visible = True And txtTexto_mat <> "" Or Cmb_texto_mat.Visible = True And Cmb_texto_mat <> "" Then
        If cmbfiltrarpor_mat = "Família" Then
            StrSqlLocProdPadraoMat = INNERJOINTEXTO & " where P.classe = '" & Cmb_texto_mat & "' and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor_mat = "Comprimento" Or cmbfiltrarpor_mat = "Largura" Or cmbfiltrarpor_mat = "Espessura" Then
                Select Case cmbfiltrarpor_mat
                    Case "Comprimento": TextoFiltro = "P.Comprimento"
                    Case "Largura": TextoFiltro = "P.Largura"
                    Case "Espessura": TextoFiltro = "P.Espessura"
                End Select
                valor = txtTexto_mat
                NovoValor = Replace(valor, ",", ".")
                StrSqlLocProdPadraoMat = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
            Else
                Select Case cmbfiltrarpor_mat
                    Case "Código interno": TextoFiltro = "P.desenho"
                    Case "Código de referência": TextoFiltro = "IA.N_referencia"
                    Case "Descrição": TextoFiltro = "P.descricao"
                End Select
                StrSqlLocProdPadraoMat = INNERJOINTEXTO & " where " & TextoFiltro & " like " & FunVerifTipoFiltroIMF(optInicio_mat, optMeio_mat, optFim_mat, optIgual_mat, txtTexto_mat) & " and " & TextoFiltroPadrao
        End If
    Else
        StrSqlLocProdPadraoMat = INNERJOINTEXTO & " where " & TextoFiltroPadrao
    End If
    ProcCarregaListaMat
Else
    CamposFiltro = "OS.IDproducao, F.Versao, OS.Fase, OS.Quantidade, OS.TempoTotalLote, OS.Maquina, OS.Prazofinal, P.Ordem, P.Desenho, P.N_Referencia, P.produto"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from (((((Ordemservico OS LEFT JOIN Fases F ON F.IDfase = OS.IDfase) INNER JOIN CadMaquinas CM ON CM.Maquina = OS.Maquina) INNER JOIN Producao P ON P.Ordem = OS.Ordem) LEFT JOIN Producao_pedidos PP ON PP.Ordem = P.Ordem) LEFT JOIN vendas_carteira VC ON VC.Codigo = PP.IDcarteira) LEFT JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao"
    TextoFiltroPadrao = "OS.Maquina = '" & frmPlano_producao.Cmb_posto & "' and P.status <> 'Cancelada' and P.DtValidacao IS NOT NULL and OS.Pronto = 'NÃO' and (OS.ID_apontamento IS NULL or OS.ID_apontamento = 0)"
    
    If Lista_material.ListItems.Count <> 0 Then
        INNERJOINTEXTO = INNERJOINTEXTO & " INNER JOIN Producaomaterial PM ON PM.Ordem = P.Ordem"
        TextoFiltroPadrao = "PM.Codigo = '" & Lista_material.SelectedItem.ListSubItems(1) & "' and " & TextoFiltroPadrao
    End If
    
    If txtTexto <> "" Then
        Select Case cmbfiltrarpor
            Case "Código interno": TextoFiltro = "P.desenho"
            Case "Código de referência": TextoFiltro = "P.N_Referencia"
            Case "Descrição": TextoFiltro = "P.Produto"
            Case "OS": TextoFiltro = "OS.IDproducao"
            Case "Ordem": TextoFiltro = "P.Ordem"
            Case "Grupo": TextoFiltro = "CM.Grupo"
            Case "Pedido interno": TextoFiltro = "VP.Ncotacao"
        End Select
        If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
            StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = " & txtTexto & " and " & TextoFiltroPadrao & " group by " & CamposFiltro & " order by OS.IDproducao"
        Else
            StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao & " group by " & CamposFiltro & " order by OS.IDproducao"
        End If
    Else
        StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao & " group by " & CamposFiltro & " order by OS.IDproducao"
    End If
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAdicionar()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente adicionar esta(s) OS('s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                    If USMsgBox("Alguma OS selecionada será adicionada com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
                    GoTo 1
                Else
                    Exit Sub
                End If
            End If
1:
            Permitido = True
            If Permitido1 = True Then
                IDlista = .ListItems.Item(InitFor)
                Permitido2 = True
                frmPlano_producao_localizar_OS_adicionar.Show 1
                
                If Permitido2 = False Then Exit Sub
            Else
                Conexao.Execute "Update Ordemservico Set ID_apontamento = " & frmPlano_producao.Txt_ID & " where IDproducao = " & .ListItems.Item(InitFor)
                '==================================
                Modulo = "PCP/Plano da produção"
                Evento = "Nova OS"
                ID_documento = .ListItems.Item(InitFor)
                Documento = "Nº plano: " & frmPlano_producao.Txt_numero_plano
                Documento1 = "OS: " & .ListItems.Item(InitFor)
                ProcGravaEvento
                '==================================
            End If
        End If
    Next InitFor
End With

If Permitido = False Then
    USMsgBox ("Informe a(s) OS('s) antes de adicionar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("OS('s) adicionada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    frmPlano_producao.ProcCarregaListaOS
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_mat_Click()
On Error GoTo tratar_erro

ProcLimpaCamposMat
With Cmb_texto_mat
    If cmbfiltrarpor_mat = "Família" Then
        txtTexto_mat.Visible = False
        .Visible = True
        ProcCarregaComboFamilia Cmb_texto_mat, "familia <> 'Null' and Compras = 'True'", True
    Else
        txtTexto_mat.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_mat_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_mat.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao1.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao1.AbsolutePage = -3 Then
        ProcExibePaginaMat (TBLocalizar_produto_padrao1.PageCount - 1)
    Else
        TBLocalizar_produto_padrao1.AbsolutePage = TBLocalizar_produto_padrao1.AbsolutePage - 2
        ProcExibePaginaMat (TBLocalizar_produto_padrao1.AbsolutePage)
    End If
Else
    ProcExibePaginaMat (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_mat_Click()
On Error GoTo tratar_erro

If txtPagIr_mat = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas_mat.Caption, 4))
If Quant <= 1 Or txtPagIr_mat > Quant Then Exit Sub
If txtPagIr_mat.Text >= 1 And txtPagIr_mat.Text <= Quant Then
    TBLocalizar_produto_padrao1.AbsolutePage = txtPagIr_mat.Text
    ProcExibePaginaMat (TBLocalizar_produto_padrao1.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_mat_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_mat.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao1.AbsolutePage = 1
ProcExibePaginaMat (TBLocalizar_produto_padrao1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_mat_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_mat.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao1.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao1.AbsolutePage = 1 Then
        ProcExibePaginaMat (2)
    Else
        ProcExibePaginaMat (TBLocalizar_produto_padrao1.AbsolutePage)
    End If
Else
    ProcExibePaginaMat (TBLocalizar_produto_padrao1.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_mat_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_mat.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao1.AbsolutePage = TBLocalizar_produto_padrao1.PageCount
ProcExibePaginaMat (TBLocalizar_produto_padrao1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_produto_padrao.PageCount - 1)
    Else
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
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
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_produto_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: Lista_DblClick
    Case vbKeyF2: ProcFiltrar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11805, 6, True
SSTab1.Tab = 0
cmbfiltrarpor_mat = "Código interno"
cmbfiltrarpor = "OS"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "OS" Then
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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
With frmPlano_producao
    .Txt_OS = Lista.SelectedItem
    .ProcCarregaDadosOS True
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaMat()
On Error GoTo tratar_erro

lblRegistros_mat.Caption = "Nº de reg.: 0"
lblPaginas_mat.Caption = "Página: 0 de: 0"
ProcLimpaCamposMat
If StrSqlLocProdPadraoMat = "" Then Exit Sub
Set TBLocalizar_produto_padrao1 = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao1.Open StrSqlLocProdPadraoMat, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao1.EOF = False Then ProcExibePaginaMat (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaMat(Pagina)
On Error GoTo tratar_erro

ProcLimpaCamposMat
TBLocalizar_produto_padrao1.PageSize = IIf(txtNreg_mat = "", 30, txtNreg_mat)
TBLocalizar_produto_padrao1.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao1.PageSize
ContadorReg = 1

If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
    PBLista_mat.Min = 0
    PBLista_mat.Max = TBLocalizar_produto_padrao1.PageSize
    PBLista_mat.Value = 1
    Contador = 0
End If
Do While TBLocalizar_produto_padrao1.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_material.ListItems
        .Add , , TBLocalizar_produto_padrao1!Codproduto
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao1!Desenho), "", TBLocalizar_produto_padrao1!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao1!Descricao), "", TBLocalizar_produto_padrao1!Descricao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao1!Unidade), "", TBLocalizar_produto_padrao1!Unidade)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao1!Classe), "", TBLocalizar_produto_padrao1!Classe)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao1!Espessura), "", Format(TBLocalizar_produto_padrao1!Espessura, "###,##0.00")) & "X" & IIf(IsNull(TBLocalizar_produto_padrao1!Largura), "", Format(TBLocalizar_produto_padrao1!Largura, "###,##0.00")) & "X" & IIf(IsNull(TBLocalizar_produto_padrao1!Comprimento), "", Format(TBLocalizar_produto_padrao1!Comprimento, "###,##0.00"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao1!Dureza), "", TBLocalizar_produto_padrao1!Dureza)
    End With
    TBLocalizar_produto_padrao1.MoveNext
    ContadorReg = ContadorReg + 1
    
    If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
        Contador = Contador + 1
        PBLista_mat.Value = Contador
    End If
Loop
lblRegistros_mat.Caption = "Nº de reg.: " & TBLocalizar_produto_padrao1.RecordCount
If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
   lblPaginas_mat.Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao1.PageCount
ElseIf TBLocalizar_produto_padrao1.AbsolutePage = adPosEOF Then
        lblPaginas_mat.Caption = "Pág.: " & TBLocalizar_produto_padrao1.PageCount & " de: " & TBLocalizar_produto_padrao1.PageCount
    Else
        lblPaginas_mat.Caption = "Pág.: " & TBLocalizar_produto_padrao1.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao1.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLocalizar_produto_padrao!IDProducao
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!versao), "", TBLocalizar_produto_padrao!versao)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Fase), "", TBLocalizar_produto_padrao!Fase)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!quantidade), "", TBLocalizar_produto_padrao!quantidade)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!TempoTotalLote), "00:00:00", TBLocalizar_produto_padrao!TempoTotalLote)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!maquina), "", TBLocalizar_produto_padrao!maquina)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!PrazoFinal), "", Format(TBLocalizar_produto_padrao!PrazoFinal, "dd/mm/yy"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLocalizar_produto_padrao!Ordem), "", TBLocalizar_produto_padrao!Ordem)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao!N_referencia), "", TBLocalizar_produto_padrao!N_referencia)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLocalizar_produto_padrao!Produto), "", TBLocalizar_produto_padrao!Produto)
    End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_material_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Lista.ListItems.Clear
If Lista_material.ListItems.Count = 0 Then Exit Sub
With Lista_material.SelectedItem
    Txt_cod_interno = .ListSubItems(1)
    Txt_descricao = .ListSubItems(2)
    Txt_un = .ListSubItems(3)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_mat_Click()
On Error GoTo tratar_erro

ProcLimpaCamposMat

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_mat_Click()
On Error GoTo tratar_erro

ProcLimpaCamposMat

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_mat_Click()
On Error GoTo tratar_erro

ProcLimpaCamposMat

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_mat_Change()
On Error GoTo tratar_erro

If txtNreg_mat <> "" Then
    VerifNumero = txtNreg_mat
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg_mat = ""
        txtNreg_mat.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_mat_Change()
On Error GoTo tratar_erro

If txtPagIr_mat <> "" Then
    VerifNumero = txtPagIr_mat
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr_mat = ""
        txtPagIr_mat.SetFocus
        Exit Sub
    End If
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

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If txtTexto <> "" And (cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS") Then
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

Private Sub txtTexto_mat_Change()
On Error GoTo tratar_erro

ProcLimpaCamposMat

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcAdicionar
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposMat()
On Error GoTo tratar_erro

Txt_cod_interno = ""
Txt_descricao = ""
Txt_un = ""
Lista_material.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

