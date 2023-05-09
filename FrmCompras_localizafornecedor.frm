VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCompras_localizafornecedor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Administrativo | Compras - Fornecedor - Localizar"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   31
      Top             =   8280
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   767
      DibPicture      =   "FrmCompras_localizafornecedor.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "FrmCompras_localizafornecedor.frx":3650
   End
   Begin VB.Frame Frame9 
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
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   270
      TabIndex        =   25
      Top             =   7290
      Width           =   10785
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
         Left            =   2910
         TabIndex        =   12
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
         Left            =   5550
         TabIndex        =   13
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   7770
         TabIndex        =   17
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmCompras_localizafornecedor.frx":396A
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
         Left            =   7230
         TabIndex        =   16
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmCompras_localizafornecedor.frx":7111
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
         Left            =   6120
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
         Left            =   6690
         TabIndex        =   15
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmCompras_localizafornecedor.frx":AC1E
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
         Left            =   8310
         TabIndex        =   18
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "FrmCompras_localizafornecedor.frx":ED0F
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
         Left            =   2220
         TabIndex        =   28
         Top             =   240
         Width           =   2190
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
         TabIndex        =   27
         Top             =   240
         Width           =   945
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
         Left            =   8970
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1680
      TabIndex        =   19
      Top             =   1560
      Width           =   9375
      Begin VB.Frame Frame4 
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
         Height          =   510
         Left            =   4380
         TabIndex        =   29
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   9
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   7
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   8
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   10
            Top             =   180
            Width           =   705
         End
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
         ItemData        =   "FrmCompras_localizafornecedor.frx":1259C
         Left            =   180
         List            =   "FrmCompras_localizafornecedor.frx":125B5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4125
      End
      Begin VB.TextBox txtTexto 
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
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1110
         Width           =   8985
      End
      Begin VB.ComboBox cmbFamilia 
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Família."
         Top             =   1110
         Width           =   8985
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Número do CPF."
         Top             =   1110
         Visible         =   0   'False
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtcnpj 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Número do CNPJ."
         Top             =   1110
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
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
         Left            =   3937
         TabIndex        =   22
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1815
         TabIndex        =   21
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   270
      TabIndex        =   20
      Top             =   1560
      Width           =   1305
      Begin VB.OptionButton optFisica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Física"
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
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optJuridica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jurídica"
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
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   270
      TabIndex        =   23
      Top             =   7920
      Width           =   10785
      _ExtentX        =   19024
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   24
      Top             =   450
      Width           =   11325
      _ExtentX        =   19976
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
         Name            =   "MS Sans Serif"
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
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9120
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "FrmCompras_localizafornecedor.frx":1261A
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4125
      Left            =   270
      TabIndex        =   11
      Top             =   3150
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   7276
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Cód"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "CNPJ/CPF"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Razão social"
         Object.Width           =   4736
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Endereço"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cidade"
         Object.Width           =   2663
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "E-mail"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "UF"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "FrmCompras_localizafornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia.Text <> "" Then
    txtTexto.Text = ""
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Razão social" Or cmbfiltrarpor = "Nome fantasia" Or cmbfiltrarpor = "Cidade" Or cmbfiltrarpor = "Código do fornecedor" Then
    txtTexto.Visible = True
    cmbfamilia.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Regime tributário" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    txtcnpj.Visible = False
    txtCpf.Visible = False
    If cmbfiltrarpor = "Regime tributário" Then ProcCarregaComboRegime
End If
If cmbfiltrarpor = "CNPJ/CPF" And optJuridica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    txtcnpj.Visible = True
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optFisica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If optFisica.Value = True Then
    TipoPessoa = "CF.pessoa = 'FÍSICA'"
    CPFCNPJ = "CF.cpf_cnpj = '" & txtCpf.Text & "'"
Else
    TipoPessoa = "CF.pessoa = 'JURÍDICA'"
    CPFCNPJ = "CF.cpf_cnpj = '" & txtcnpj.Text & "'"
End If
If Compras_Cotacao = True Then ProspectoFiltro = "(CF.Prospecto = 'False' or CF.Prospecto = 'True')" Else ProspectoFiltro = "CF.Prospecto = 'False'"
If Faturamento = True Then NFTexto = " and CF.Enviar_NF = 'True'" Else NFTexto = ""

If cmbfiltrarpor = "Regime tributário" Then
    If cmbfamilia = "Lucro presumido" Then
        TextoRegime = "CF.Presumido = 'True'"
    ElseIf cmbfamilia = "Simples nacional" Then
            TextoRegime = "CF.Simples = 'True'"
        ElseIf cmbfamilia = "Lucro real" Then
                TextoRegime = "CF.Real = 'True'"
            Else
                TextoRegime = "CF.MEI = 'True'"
    End If
End If

CamposFiltro = "CF.IDCliente, CF.CPF_CNPJ, CF.Nome_Razao, CF.Tipo_endereco, CF.Endereco, CF.Cidade, CF.Email, CF.Estado, CF.idTipoEmpresa, CF.Pessoa"
INNERJOINTEXTO = "Select " & CamposFiltro & " from Compras_fornecedores CF LEFT JOIN compras_fornecedores_familia CFF ON CF.IDCliente = CFF.IDCliente"
TextoFiltroPadrao = TipoPessoa & " and CF.DtValidacao IS NOT NULL and CF.Status <> 'Bloqueado' and " & ProspectoFiltro & " group by " & CamposFiltro & " order by CF.nome_razao"

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Or txtcnpj.Visible = True And txtcnpj <> "__.___.___/____-__" Or txtCpf.Visible = True And txtCpf <> "___.___.___-__" Then
    If cmbfiltrarpor = "Família" Then
        StrSqlLocFornPadrao = INNERJOINTEXTO & " where CFF.Familia = '" & cmbfamilia & "' and CFF.tipo = 'F' and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor = "CNPJ/CPF" Then
                StrSqlLocFornPadrao = INNERJOINTEXTO & " where " & CPFCNPJ & " and " & TextoFiltroPadrao
            ElseIf cmbfiltrarpor = "Regime tributário" Then
                StrSqlLocFornPadrao = INNERJOINTEXTO & " where " & TextoRegime & " and " & TextoFiltroPadrao
        Else
            Select Case cmbfiltrarpor
                Case "Razão social": TextoFiltro = "CF.nome_razao"
                Case "Nome fantasia": TextoFiltro = "CF.nomefantasia"
                Case "Cidade": TextoFiltro = "CF.cidade"
                Case "Código do fornecedor": TextoFiltro = "CF.IDCliente"
            End Select
            StrSqlLocFornPadrao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLocFornPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcaCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_fornecedor_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_fornecedor_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_fornecedor_padrao.PageCount - 1)
    Else
        TBLocalizar_fornecedor_padrao.AbsolutePage = TBLocalizar_fornecedor_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_fornecedor_padrao.AbsolutePage)
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
    TBLocalizar_fornecedor_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_fornecedor_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_fornecedor_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_fornecedor_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_fornecedor_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_fornecedor_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_fornecedor_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_fornecedor_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_fornecedor_padrao.AbsolutePage = TBLocalizar_fornecedor_padrao.PageCount
ProcExibePagina (TBLocalizar_fornecedor_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10785, 5, True
With USToolBar1
    .ButtonState(2) = 5
    .Refresh
End With

ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
If Estoque_Consignacao = True Then Caption = "Estoque - Recebimento - Consignação - Localizar fornecedor"
If Estoque_Inventario = True Then Caption = "Estoque - Inventário - Localizar fornecedor"
If Compras_Pedido = True Then Caption = "Administrativo - Compras - Pedido - Localizar fornecedor"
If Vendas_Proposta = True Then Caption = "Administrativo - Vendas - Proposta comercial - Localizar fornecedor"
If Vendas_PI = True Then Caption = "Administrativo - Vendas - Pedido interno - Localizar fornecedor"
If Financeiro_Contas_Pagar = True Then Caption = "Administrativo - Financeiro - Contas à pagar - Localizar fornecedor"
If Financeiro_Contas_Pagas = True Then Caption = "Administrativo - Financeiro - Contas pagas - Localizar fornecedor"
If Financeiro_Contas_Receber = True Then Caption = "Administrativo - Financeiro - Contas à receber - Localizar fornecedor"
If Financeiro_Contas_Recebidas = True Then Caption = "Administrativo - Financeiro - Contas recebidas - Localizar fornecedor"
If Faturamento = True Then
    If Sit_REG = 4 Then
        Caption = "Administrativo - Faturamento - Minuta de despacho - Localizar fornecedor"
    Else
        If Formulario = "Faturamento/Nota fiscal/Própria" Then
            Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Localizar fornecedor"
        ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
                Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Localizar fornecedor"
            ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                    Caption = "Estoque - Ordem de faturamento - Localizar fornecedor"
                Else
                    Caption = "Estoque - Nota fiscal - Localizar fornecedor"
        End If
    End If
End If
If Engenharia_Produtos = True Then TextoCaption = "Engenharia"
If Compras_Produtos = True Then TextoCaption = "Compras"
If Vendas_Produtos = True Then TextoCaption = "Vendas"
If Engenharia_Localfornecedor = True Then Caption = TextoCaption & " - Produtos e serviços - Cadastro de códigos de referência - Localizar fornecedor"
If Engenharia_Localfornecedor1 = True Then Caption = TextoCaption & " - Produtos e serviços - Localizar fornecedor"
If Clientes = True Then Caption = "Administrativo - Vendas - Clientes - Localizar fornecedor"
If Compras_Programacao = True Then Caption = "Administrativo - Compras - Programação - Localizar fornecedor"

If Compras_Cotacao = True Then
    Caption = "Administrativo - Compras - Cotação - Localizar fornecedor"
    If Sit_REG = 1 Then
        With USToolBar1
            .ButtonState(2) = 0
            .Refresh
        End With
        ListView1.ColumnHeaders(1).Text = ""
        ListView1.CheckBoxes = True
    End If
End If
    
If OpcoesGerais = True Then Caption = "Opções gerais - Dados para criar contas a pagar - Localizar fornecedor"
If Compras_Fornecedores = True Then Caption = "Compras - Fornecedores - Localizar fornecedor"
If Qualidade_PPAP_FMEA = True Then Caption = "Qualidade - PPAP - FMEA - Localizar fornecedor"
If RH_Funcionarios = True Then Caption = "RH - Cadastro de funcionários - Agência"
cmbfiltrarpor = "Razão social"

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

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optFisica_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Regime tributário" Then ProcCarregaComboRegime

ListView1.ListItems.Clear
If optFisica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    txtcnpj.Visible = False
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optJuridica_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Regime tributário" Then ProcCarregaComboRegime

ListView1.ListItems.Clear
If optJuridica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    txtcnpj.Visible = True
    txtCpf.Visible = False
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

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

ListView1.ListItems.Clear
If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcnpj_Change()
On Error GoTo tratar_erro
  
ListView1.ListItems.Clear
If txtcnpj.Text <> "__.___.___/____-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCpf_Change()
On Error GoTo tratar_erro
  
ListView1.ListItems.Clear
If txtCpf.Text <> "___.___.___-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: If USToolBar1.ButtonState(2) = 0 Then ProcAdicionar
    Case vbKeyReturn: ListView1_DblClick
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Sit_REG = 1 Then
    If Compras_Pedido = True Then Permitido2 = False
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Or Compras_Cotacao = True And Sit_REG = 1 Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * FROM Compras_fornecedores WHERE idcliente = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    If Estoque_Consignacao = True Then
        With frmEstoque_Recebimento_consignacao
            .txtCliente.Text = ListView1.SelectedItem.SubItems(2)
            .txtid_cliente.Text = ListView1.SelectedItem
            .Txt_tipodest = "F"
        End With
        Unload Me
        Exit Sub
    End If
    If Estoque_Inventario = True Then
        With frmestoque_fisico
            .Cmb_tipo_cli_forn = "Fornecedor"
            .Txt_ID_cli_forn = ListView1.SelectedItem
            .Txt_cli_forn = ListView1.SelectedItem.SubItems(2)
        End With
        Unload Me
        Exit Sub
    End If
    If Compras_Pedido = True Then
        With frmCompras_Pedido
            If Sit_REG < 3 Then
                Cont = IIf(Sit_REG = 1, .Cmb_empresa_carteira.ItemData(.Cmb_empresa_carteira.ListIndex), .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex))
                If FunVerifValidadeCertForn(Cont, .txtData, True) = False Then Exit Sub
                If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
                    If FunVerifRegimeTribCliForn(Cont, True, True) = False Then Exit Sub
                End If
                .txtIDfornecedor = ""
                .txtIDfornecedor = TBFornecedor!IDCliente
                .txtcnpj = TBFornecedor!CPF_CNPJ
            Else
                .cmbtransporte = ListView1.SelectedItem.ListSubItems(2)
            End If
        End With
        Unload Me
        Exit Sub
    End If
    If Financeiro_Contas_Pagar = True Then
        With frmContas_Pagar
            .Cmb_tipo = "Fornecedor"
            .txtIDFornec = TBFornecedor!IDCliente
        End With
        Unload Me
        Exit Sub
    End If
    If Financeiro_Contas_Pagas = True Then
        With frmContas_Pagas
            .Cmb_tipo = "Fornecedor"
            .txtIDFornec = TBFornecedor!IDCliente
        End With
        Unload Me
        Exit Sub
    End If
    If Financeiro_Contas_Receber = True Then
        With frmContas_Receber
            .Cmb_tipo = "Fornecedor"
            .txtIDcliente = TBFornecedor!IDCliente
        End With
        Unload Me
        Exit Sub
    End If
    If Financeiro_Contas_Recebidas = True Then
        With frmContas_recebidas
            .Cmb_tipo = "Fornecedor"
            .txtIDcliente = TBFornecedor!IDCliente
        End With
        Unload Me
        Exit Sub
    End If
    If Faturamento = True Then
        If Sit_REG < 4 Then
        
        If Formulario <> "Estoque/Ordem de faturamento" Then
            With frmFaturamento_Prod_Serv
                If FunVerifValidadeCertForn(IDempresa, .txt_DtEmissao, True) = False Then Exit Sub
                If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
                    If FunVerifRegimeTribCliForn(IDempresa, True, True) = False Then Exit Sub
                End If
                
                Select Case Sit_REG
                    Case 1:
                        IDCliente = TBFornecedor!IDCliente
                        .txt_Razao.Text = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        .txt_Endereco.Text = Endereco
                        .txtNumero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
                        If IsNull(TBFornecedor!Tipo_bairro) = False And TBFornecedor!Tipo_bairro <> "" Then
                            Bairro = TBFornecedor!Tipo_bairro & ": " & IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
                        Else
                            Bairro = IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
                        End If
                        .txt_Bairro.Text = Bairro
                        If TBFornecedor!Pessoa = "JURÍDICA" Then
                            .txttipocliente = "J"
                            .txt_IE.Text = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
                        Else
                            .txttipocliente = "F"
                            .txt_IE.Text = IIf(IsNull(TBFornecedor!RG_IM), "", TBFornecedor!RG_IM)
                        End If
                        If TBFornecedor!idTipoEmpresa = 1 Then .txt_CNPJ_CPF.Text = IIf(IsNull(TBFornecedor!CPF_CNPJ), "", TBFornecedor!CPF_CNPJ)
                        .Txt_CEP.Text = IIf(IsNull(TBFornecedor!CEP), "", TBFornecedor!CEP)
                        .txt_Municipio.Text = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                        .cbo_UF.Text = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
                        .txt_FoneFAX.Text = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
                        Suframa = False
                        .txtIDcliente.Text = IDCliente
                    Case 2:
                        .txtidinttransp = ListView1.SelectedItem
                        .TxtTransp_nome.Text = ListView1.SelectedItem.ListSubItems(2)
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        .txtTransp_endereco = Endereco
                        .txtTransp_numero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
                        .txtTransp_municipio = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                        .txtTransp_uf_Transportadora = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
                        If TBFornecedor!idTipoEmpresa = 1 Then
                            If IsNull(TBFornecedor!CPF_CNPJ) = True Or TBFornecedor!CPF_CNPJ = "__.___.___/____-__" Or TBFornecedor!CPF_CNPJ = "" Then .txtTransp_cnpj = "" Else .txtTransp_cnpj = TBFornecedor!CPF_CNPJ
                        End If
                        
                        If TBFornecedor!Pessoa = "JURÍDICA" Then
                            .txtTransp_IE = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
                            .txtTransp_IE.Locked = False
                            .txtTransp_IE.TabStop = True
                        Else
                            .txtTransp_IE.Locked = True
                            .txtTransp_IE.TabStop = False
                        End If
                            
                        'If TBFornecedor!Sedex = True Then .chkSedex.Value = 1 Else .chkSedex.Value = 0
                    Case 3:
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        Redespacho = "Nome: " & ListView1.SelectedItem.ListSubItems(2) & " - Endereço: " & Endereco & " - Número: " & IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero) & " - Cidade: " & IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade) & " - UF: " & IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado) & " - CNPJ: " & IIf(TBFornecedor!idTipoEmpresa = 1, IIf(IsNull(TBFornecedor!CPF_CNPJ), "", TBFornecedor!CPF_CNPJ), "") & " - IE: " & IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
                        If .txtDados_DadosAdicionais <> "" Then
                            .txtDados_DadosAdicionais = .txtDados_DadosAdicionais & " | REDESPACHO: " & Redespacho
                        Else
                            .txtDados_DadosAdicionais = Redespacho
                        End If
                End Select
            End With
            
    Else
            With frmEstoque_Ordem_Faturamento
                If FunVerifValidadeCertForn(IDempresa, .txt_DtEmissao, True) = False Then Exit Sub
                If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
                    If FunVerifRegimeTribCliForn(IDempresa, True, True) = False Then Exit Sub
                End If
                
                Select Case Sit_REG
                    Case 1:
                        IDCliente = TBFornecedor!IDCliente
                        .txt_Razao.Text = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        .txt_Endereco.Text = Endereco
                        .txtNumero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
                        If IsNull(TBFornecedor!Tipo_bairro) = False And TBFornecedor!Tipo_bairro <> "" Then
                            Bairro = TBFornecedor!Tipo_bairro & ": " & IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
                        Else
                            Bairro = IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
                        End If
                        .txt_Bairro.Text = Bairro
                        If TBFornecedor!Pessoa = "JURÍDICA" Then
                            .txttipocliente = "J"
                            .txt_IE.Text = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
                        Else
                            .txttipocliente = "F"
                            .txt_IE.Text = IIf(IsNull(TBFornecedor!RG_IM), "", TBFornecedor!RG_IM)
                        End If
                        If TBFornecedor!idTipoEmpresa = 1 Then .txt_CNPJ_CPF.Text = IIf(IsNull(TBFornecedor!CPF_CNPJ), "", TBFornecedor!CPF_CNPJ)
                        .Txt_CEP.Text = IIf(IsNull(TBFornecedor!CEP), "", TBFornecedor!CEP)
                        .txt_Municipio.Text = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                        .cbo_UF.Text = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
                        .txt_FoneFAX.Text = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
                        Suframa = False
                        .txtIDcliente.Text = IDCliente
                    Case 2:
                        .txtidinttransp = ListView1.SelectedItem
                        .TxtTransp_nome.Text = ListView1.SelectedItem.ListSubItems(2)
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        .txtTransp_endereco = Endereco
                        .txtTransp_numero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
                        .txtTransp_municipio = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                        .txtTransp_uf_Transportadora = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
                        If TBFornecedor!idTipoEmpresa = 1 Then
                            If IsNull(TBFornecedor!CPF_CNPJ) = True Or TBFornecedor!CPF_CNPJ = "__.___.___/____-__" Or TBFornecedor!CPF_CNPJ = "" Then .txtTransp_cnpj = "" Else .txtTransp_cnpj = TBFornecedor!CPF_CNPJ
                        End If
                        
                        If TBFornecedor!Pessoa = "JURÍDICA" Then
                            .txtTransp_IE = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
                            .txtTransp_IE.Locked = False
                            .txtTransp_IE.TabStop = True
                        Else
                            .txtTransp_IE.Locked = True
                            .txtTransp_IE.TabStop = False
                        End If
                            
                        'If TBFornecedor!Sedex = True Then .chkSedex.Value = 1 Else .chkSedex.Value = 0
                    Case 3:
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        Redespacho = "Nome: " & ListView1.SelectedItem.ListSubItems(2) & " - Endereço: " & Endereco & " - Número: " & IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero) & " - Cidade: " & IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade) & " - UF: " & IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado) & " - CNPJ: " & IIf(TBFornecedor!idTipoEmpresa = 1, IIf(IsNull(TBFornecedor!CPF_CNPJ), "", TBFornecedor!CPF_CNPJ), "") & " - IE: " & IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
'                        If .txtDados_DadosAdicionais <> "" Then
'                            .txtDados_DadosAdicionais = .txtDados_DadosAdicionais & " | REDESPACHO: " & Redespacho
'                        Else
'                            .txtDados_DadosAdicionais = Redespacho
'                        End If
                End Select
            End With
    End If
    
    
        ElseIf Sit_REG = 4 Then
                With frmMinuta
                    If FunVerifValidadeCertForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), .txtData, True) = False Then Exit Sub
                    If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
                        If FunVerifRegimeTribCliForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), True, True) = False Then Exit Sub
                    End If
                    
                    .txtID_transp = ListView1.SelectedItem
                    .txtTranportadora = ListView1.SelectedItem.ListSubItems(2)
                    If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                        Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                    Else
                        Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                    End If
                    .txtendereco = Endereco
                    .txtCidade = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                    .cmbuf = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
                    .txttelefone = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
                    .txtFax = IIf(IsNull(TBFornecedor!Fax), "", TBFornecedor!Fax)
                    If TBFornecedor!idTipoEmpresa = 1 Then
                        If IsNull(TBFornecedor!CPF_CNPJ) = True Or TBFornecedor!CPF_CNPJ = "__.___.___/____-__" Or TBFornecedor!CPF_CNPJ = "" Then .txtcnpj = "" Else .txtcnpj = TBFornecedor!CPF_CNPJ
                    End If
                    .txtIE = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
                End With
            Else
            If Formulario <> "Estoque/Ordem de faturamento" Then
                With frmFaturamento_Prod_Serv_DI
                    If FunVerifValidadeCertForn(IDempresa, .txtData, True) = False Then Exit Sub
                    If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
                        If FunVerifRegimeTribCliForn(IDempresa, True, True) = False Then Exit Sub
                    End If
                    .Txt_fornecedor = ListView1.SelectedItem.ListSubItems(2)
                    .Txt_ID_fornecedor = ListView1.SelectedItem
                End With
            End If
            
        End If
        Unload Me
        Exit Sub
    End If
    If Engenharia_Localfornecedor = True Then
        With frmproj_produto_referencia
            .Txt_ID_cliente_forn = ListView1.SelectedItem
            .Txt_tipo = "F"
            .txtAplicacao = ListView1.SelectedItem.ListSubItems(2)
            Unload Me
            Exit Sub
        End With
    End If
    If Engenharia_Localfornecedor1 = True Then
        With frmproj_produto
            .txtIDfornecedor = ListView1.SelectedItem
            .cmbFornecedor.AddItem ListView1.SelectedItem.ListSubItems(2)
            .cmbFornecedor = ListView1.SelectedItem.ListSubItems(2)
            Unload Me
            Exit Sub
        End With
    End If
    If Clientes = True Then
        With frmVendas_cliente
            .cmbtransportadora = ListView1.SelectedItem.ListSubItems(2)
            Unload Me
            Exit Sub
        End With
    End If
    If Compras_Programacao = True Then
        With frmCompras_programacao
            If FunVerifValidadeCertForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), .txtData, True) = False Then Exit Sub
            If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
                If FunVerifRegimeTribCliForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), True, True) = False Then Exit Sub
            End If
            .txtID_forn.Text = ListView1.SelectedItem
            .txtFornecedor.Text = ListView1.SelectedItem.ListSubItems(2)
            Unload Me
            Exit Sub
        End With
    End If
    If Compras_Cotacao = True Then
        With frmcompras_reqcot
            If FunVerifValidadeCertForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), .txtDataemissao, True) = False Then Exit Sub
            If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
                If FunVerifRegimeTribCliForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), True, True) = False Then Exit Sub
            End If
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Cotacao_fornecedor where idcot = " & .txtidcot.Text & " and idforn =  " & TBFornecedor!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                USMsgBox ("Este fornecedor já está adicionado na lista."), vbExclamation, "CAPRIND v5.0"
                TBFornecedor.Close
                TBAbrir.Close
                Exit Sub
            End If
            TBAbrir.Close
            
            .txtIDforn.Text = IIf(IsNull(TBFornecedor!IDCliente), "", TBFornecedor!IDCliente)
            .txtforn.Text = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
            .txttelforn.Text = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
            .txtfaxforn.Text = IIf(IsNull(TBFornecedor!Fax), "", TBFornecedor!Fax)
            Unload Me
            Exit Sub
        End With
    End If
    If RNC = True Then
        With frmQualidade_RNC
            .txtID_forn = ListView1.SelectedItem
            .txtFornecedor = ListView1.SelectedItem.ListSubItems(2)
            .txttipo = "F"
            Unload Me
            Exit Sub
        End With
    End If
    If OpcoesGerais = True Then
        With frmOpcoesGeral_PC
            .txtIDforn = ListView1.SelectedItem
            .txtFornecedor = ListView1.SelectedItem.ListSubItems(2)
            Unload Me
            Exit Sub
        End With
    End If
    If Compras_Fornecedores = True Then
        frmCompras_fornecedores.cmbtransportadora = ListView1.SelectedItem.ListSubItems(2)
        Unload Me
        Exit Sub
    End If
    If Vendas_Proposta = True Or Vendas_PI = True Then
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            .cmbtransportadora = ListView1.SelectedItem.ListSubItems(2)
            Unload Me
            Exit Sub
        End With
    End If
    If Qualidade_PPAP_FMEA = True Then
        With frmQualidadePPAP_FMEA
            .txtIDforn = ListView1.SelectedItem
            .txtFornecedor = ListView1.SelectedItem.ListSubItems(2)
        End With
        Unload Me
        Exit Sub
    End If
    If RH_Funcionarios = True Then
        With frmRH_Funcionarios
            .txtIDfornecedor = ListView1.SelectedItem
            .txtFornecedor = ListView1.SelectedItem.SubItems(2)
        End With
    End If
End If
TBFornecedor.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcaCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocFornPadrao = "" Then Exit Sub
Set TBLocalizar_fornecedor_padrao = CreateObject("adodb.recordset")
TBLocalizar_fornecedor_padrao.Open StrSqlLocFornPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_fornecedor_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_fornecedor_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_fornecedor_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_fornecedor_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_fornecedor_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_fornecedor_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_fornecedor_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_fornecedor_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_fornecedor_padrao!IDCliente
        If TBLocalizar_fornecedor_padrao!idTipoEmpresa = 1 And IsNull(TBLocalizar_fornecedor_padrao!CPF_CNPJ) = False And TBLocalizar_fornecedor_padrao!CPF_CNPJ <> "__.___.___/____-__" Then .Item(.Count).SubItems(1) = TBLocalizar_fornecedor_padrao!CPF_CNPJ
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_fornecedor_padrao!Nome_Razao), "", TBLocalizar_fornecedor_padrao!Nome_Razao)
        If IsNull(TBLocalizar_fornecedor_padrao!Tipo_endereco) = False And TBLocalizar_fornecedor_padrao!Tipo_endereco <> "" Then
            Endereco = TBLocalizar_fornecedor_padrao!Tipo_endereco & ": " & IIf(IsNull(TBLocalizar_fornecedor_padrao!Endereco), "", TBLocalizar_fornecedor_padrao!Endereco)
        Else
            Endereco = IIf(IsNull(TBLocalizar_fornecedor_padrao!Endereco), "", TBLocalizar_fornecedor_padrao!Endereco)
        End If
        .Item(.Count).SubItems(3) = Endereco
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_fornecedor_padrao!Cidade), "", TBLocalizar_fornecedor_padrao!Cidade)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_fornecedor_padrao!Email), "", TBLocalizar_fornecedor_padrao!Email)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_fornecedor_padrao!Estado), "", TBLocalizar_fornecedor_padrao!Estado)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLocalizar_fornecedor_padrao!Pessoa), "", TBLocalizar_fornecedor_padrao!Pessoa)
    End With
    TBLocalizar_fornecedor_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLocalizar_fornecedor_padrao.RecordCount
If TBLocalizar_fornecedor_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_fornecedor_padrao.PageCount
ElseIf TBLocalizar_fornecedor_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_fornecedor_padrao.PageCount & " de: " & TBLocalizar_fornecedor_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_fornecedor_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_fornecedor_padrao.PageCount
End If


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
    Case 5: ProcSair
End Select

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
Permitido2 = False
With ListView1
    For InitFor1 = 1 To .ListItems.Count
        If .ListItems.Item(InitFor1).Checked = True Then
            If Permitido2 = False Then
                If USMsgBox("Deseja realmente adicionar este(s) fornecedor(es)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido2 = True
            
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from Cotacao_fornecedor where idcot = " & frmcompras_reqcot.txtidcot & " and IDForn = " & .ListItems.Item(InitFor1), Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = True Then TBFornecedor.AddNew
            TBFornecedor!idcot = frmcompras_reqcot.txtidcot
            TBFornecedor!IDforn = .ListItems.Item(InitFor1)
            TBFornecedor!forn = .ListItems.Item(InitFor1).ListSubItems(2)
            TBFornecedor!aprovadoforn = False
            TBFornecedor!naprovadoforn = False
            TBFornecedor.Update
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Novo fornecedor"
            ID_documento = TBFornecedor!IDforn
            Documento = "Nº cotação: " & frmcompras_reqcot.txtidcotacao
            Documento1 = "Fornecedor: " & TBFornecedor!forn
            ProcGravaEvento
            '==================================
            TBFornecedor.Close
        End If
    Next InitFor1
End With
If Permitido2 = False Then
    USMsgBox ("Informe o(s) fornecedor(es) antes de adicionar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    USMsgBox ("Fornecedor(es) adicionado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    frmcompras_reqcot.ProcCarregaListaForn True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboRegime()
On Error GoTo tratar_erro

With cmbfamilia
    .Clear
    If optJuridica.Value = True Then
        .AddItem "Lucro presumido"
        .AddItem "Lucro real"
        .AddItem "Simples nacional"
    Else
        .AddItem "MEI"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
