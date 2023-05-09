VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVendas_LocalizarTransporte 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Administrativo | Vendas - Localizar transportadora"
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10965
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
   ScaleHeight     =   8175
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   714
      DibPicture      =   "FrmVendas_localizatransporte.frx":0000
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
      Icon            =   "FrmVendas_localizatransporte.frx":5A65
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
      Left            =   60
      TabIndex        =   25
      Top             =   7200
      Width           =   10875
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
         DibPicture      =   "FrmVendas_localizatransporte.frx":5D7F
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
         DibPicture      =   "FrmVendas_localizatransporte.frx":9526
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
         DibPicture      =   "FrmVendas_localizatransporte.frx":D033
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
         DibPicture      =   "FrmVendas_localizatransporte.frx":11124
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
      Left            =   1470
      TabIndex        =   19
      Top             =   1470
      Width           =   9465
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
            TabIndex        =   9
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
            TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   10
            Top             =   180
            Width           =   705
         End
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
         ItemData        =   "FrmVendas_localizatransporte.frx":149B1
         Left            =   180
         List            =   "FrmVendas_localizatransporte.frx":149C4
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1822
         TabIndex        =   21
         Top             =   180
         Width           =   840
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
      Left            =   55
      TabIndex        =   20
      Top             =   1470
      Width           =   1395
      Begin VB.OptionButton optFisica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Física"
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
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optJuridica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jurídica"
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
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   7830
      Width           =   10875
      _ExtentX        =   19182
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
      Height          =   1035
      Left            =   0
      TabIndex        =   24
      Top             =   390
      Width           =   10965
      _ExtentX        =   19341
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
      ButtonHeight3   =   58
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
         Img1            =   "FrmVendas_localizatransporte.frx":14A0D
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4125
      Left            =   60
      TabIndex        =   11
      Top             =   3060
      Width           =   10875
      _ExtentX        =   19182
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
Attribute VB_Name = "FrmVendas_LocalizarTransporte"
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

ProcCarregaToolBar1 Me, 10965, 5, True
With USToolBar1
    .ButtonState(2) = 5
    .Refresh
End With

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

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * FROM Compras_fornecedores WHERE idcliente = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then

If Vendas_Proposta = True Then
        With frmVendas_proposta
            If Transporte1 = True Then
            .txtTransportadora.Text = ListView1.SelectedItem.SubItems(2)
            .txtidTransportadora.Text = ListView1.SelectedItem
            Else
            .txtRedespacho.Text = ListView1.SelectedItem.SubItems(2)
        End If
        End With
End If
If Vendas_PI = True Then
        With frmVendas_PI
            If Transporte1 = True Then
            .txtTransportadora.Text = ListView1.SelectedItem.SubItems(2)
            .txtidTransportadora.Text = ListView1.SelectedItem
            Else
            .txtRedespacho.Text = ListView1.SelectedItem.SubItems(2)
        End If
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
