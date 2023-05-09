VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpcoesGeral2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Configurar banco de dados e atualização automática"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   -90
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpcoesGeral2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   23
      Top             =   6540
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   820
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView listaBancos 
      Height          =   1530
      Left            =   180
      TabIndex        =   9
      Top             =   4620
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2699
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Local dos relatórios"
         Object.Width           =   5954
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Nome da instância SQL"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Nome do banco de dados"
         Object.Width           =   4233
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   450
      Width           =   9105
      _ExtentX        =   16060
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
      ButtonLeft4     =   118
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
      ButtonLeft5     =   122
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
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
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
      ButtonState7    =   5
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin MSComDlg.CommonDialog CD1 
         Left            =   5760
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   4980
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmOpcoesGeral2.frx":000C
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados para conexão ao banco de dados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   180
      TabIndex        =   10
      Top             =   1530
      Width           =   8780
      Begin VB.ComboBox Cmb_nome_banco 
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
         Left            =   4950
         Sorted          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Nome do banco de dados."
         Top             =   540
         Width           =   3650
      End
      Begin VB.ComboBox Cmb_servidor 
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
         TabIndex        =   0
         ToolTipText     =   "Nome da instância SQL."
         Top             =   540
         Width           =   4755
      End
      Begin VB.TextBox Txt_senha 
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
         IMEMode         =   3  'DISABLE
         Left            =   4950
         Locked          =   -1  'True
         MaxLength       =   100
         PasswordChar    =   "*"
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Senha do banco de dados."
         Top             =   1830
         Width           =   3650
      End
      Begin VB.TextBox Txt_usuario 
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
         MaxLength       =   100
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Usuário do banco de dados."
         Top             =   1830
         Width           =   4755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário do banco de dados"
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
         Left            =   1740
         TabIndex        =   19
         Top             =   1620
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha do banco de dados"
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
         Left            =   5865
         TabIndex        =   18
         Top             =   1620
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do banco de dados*"
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
         Left            =   5820
         TabIndex        =   12
         Top             =   330
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da instância SQL*"
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
         Left            =   1695
         TabIndex        =   11
         Top             =   330
         Width           =   1725
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   6150
      Width           =   8775
      _ExtentX        =   15478
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Configuração das pastas..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   180
      TabIndex        =   15
      Top             =   2580
      Width           =   8780
      Begin VB.TextBox txtLocalrel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         MaxLength       =   100
         TabIndex        =   21
         ToolTipText     =   "Local dos relatórios."
         Top             =   420
         Width           =   8085
      End
      Begin VB.TextBox txtlocalantigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         MaxLength       =   100
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Local dos arquivos antigos."
         Top             =   1020
         Width           =   8085
      End
      Begin VB.TextBox txtlocalnovo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         MaxLength       =   100
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Local dos arquivos atualizados."
         Top             =   1590
         Width           =   8085
      End
      Begin DrawSuite2022.USButton Cmd_localizar_rel 
         Height          =   315
         Left            =   8280
         TabIndex        =   2
         ToolTipText     =   "Buscar item cadastrado ou vendido"
         Top             =   420
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmOpcoesGeral2.frx":33EE
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
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton cmdLocalnovo 
         Height          =   315
         Left            =   8280
         TabIndex        =   4
         ToolTipText     =   "Buscar item cadastrado ou vendido"
         Top             =   1590
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmOpcoesGeral2.frx":3FAA
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
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton cmdLocalantigo 
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         ToolTipText     =   "Buscar item cadastrado ou vendido"
         Top             =   1020
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmOpcoesGeral2.frx":4B66
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
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pasta de localização dos relatórios Caprind*"
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
         Left            =   2647
         TabIndex        =   22
         Top             =   210
         Width           =   3150
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pasta da instalação Caprind*"
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
         Left            =   3135
         TabIndex        =   17
         Top             =   810
         Width           =   2085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pasta da atualização Caprind*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   3090
         TabIndex        =   16
         Top             =   1395
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmOpcoesGeral2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_LocalBD    As Boolean 'OK

Private Sub Cmb_servidor_Change()
On Error GoTo tratar_erro

Cmb_nome_banco.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_servidor_Click()
On Error GoTo tratar_erro

Cmb_nome_banco.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_servidor_LostFocus()
On Error GoTo tratar_erro

If Cmb_servidor <> "" Then
    With Cmb_nome_banco
        .Clear
        For Each vDb In EnumSqlDbAdo(Cmb_servidor.Text, IIf(Txt_usuario = "", "Procam", Txt_usuario), IIf(Txt_senha = "", "PRO0902loc$?", Txt_senha))
            .AddItem vDb
        Next
    End With
End If

Exit Sub
tratar_erro:
    If Err.Number = 13 Then
        USMsgBox ("Não foi encontrado nenhum banco de dados ao efetuar a conexão nessa instância."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listaBancos_DblClick()
On Error GoTo tratar_erro

If listaBancos.ListItems.Count = 0 Then Exit Sub
Select Case listaBancos.SelectedItem
    Case 1:
        frmabertura.cmbBanco.Text = Nome_banco
    Case 2:
        frmabertura.cmbBanco.Text = Nome_banco1
    Case 3:
        frmabertura.cmbBanco.Text = Nome_banco2
    Case 4:
        frmabertura.cmbBanco.Text = Nome_banco3
End Select
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_senha_Change()
On Error GoTo tratar_erro

Cmb_nome_banco.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_usuario_Change()
On Error GoTo tratar_erro

Cmb_nome_banco.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_rel_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Localizar local dos relatórios"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtLocalrel.Text = sBuffer
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

With listaBancos
    If .ListItems.Count = 0 Then
        USMsgBox ("Informe o local do banco de dados que deseja excluir."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If USMsgBox("Deseja realmente excluir esta configuração? " & vbCrLf & "Nome da instância SQL: " & .SelectedItem.ListSubItems(2) & vbCrLf & "Nome do banco de dados: " & .SelectedItem.ListSubItems(3), vbYesNo) = vbYes Then
        If .SelectedItem.ListSubItems(1) = Localrel And .SelectedItem.ListSubItems(2) = NomeServidor And .SelectedItem.ListSubItems(3) = Nome_banco Then
            DeleteSetting "Procam", "CaprindSQL", "NomeServidor"
            DeleteSetting "Procam", "CaprindSQL", "LocalRel"
            DeleteSetting "Procam", "CaprindSQL", "Nome_banco"
            If Usuario_banco <> "" Then DeleteSetting "Procam", "CaprindSQL", "Usuario_banco"
            If Senha_banco <> "" Then DeleteSetting "Procam", "CaprindSQL", "Senha_banco"
            Nome_banco = ""
            Localrel = ""
        ElseIf .SelectedItem.ListSubItems(1) = Localrel1 And .SelectedItem.ListSubItems(2) = NomeServidor1 And .SelectedItem.ListSubItems(3) = Nome_banco1 Then
                DeleteSetting "Procam", "CaprindSQL", "NomeServidor1"
                DeleteSetting "Procam", "CaprindSQL", "LocalRel1"
                DeleteSetting "Procam", "CaprindSQL", "Nome_banco1"
                If Usuario_banco1 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Usuario_banco1"
                If Senha_banco1 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Senha_banco1"
                Nome_banco1 = ""
                Localrel1 = ""
            ElseIf .SelectedItem.ListSubItems(1) = Localrel2 And .SelectedItem.ListSubItems(2) = NomeServidor2 And .SelectedItem.ListSubItems(3) = Nome_banco2 Then
                    DeleteSetting "Procam", "CaprindSQL", "NomeServidor2"
                    DeleteSetting "Procam", "CaprindSQL", "LocalRel2"
                    DeleteSetting "Procam", "CaprindSQL", "Nome_banco2"
                    If Usuario_banco2 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Usuario_banco2"
                    If Senha_banco2 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Senha_banco2"
                    Nome_banco2 = ""
                    Localrel2 = ""
                ElseIf .SelectedItem.ListSubItems(1) = Localrel3 And .SelectedItem.ListSubItems(2) = NomeServidor3 And .SelectedItem.ListSubItems(3) = Nome_banco3 Then
                        DeleteSetting "Procam", "CaprindSQL", "NomeServidor3"
                        DeleteSetting "Procam", "CaprindSQL", "LocalRel3"
                        DeleteSetting "Procam", "CaprindSQL", "Nome_banco3"
                        If Usuario_banco3 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Usuario_banco3"
                        If Senha_banco3 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Senha_banco3"
                        Nome_banco3 = ""
                        Localrel3 = ""
                    Else
                        USMsgBox ("Configuração não encontrada nos registros do windows."), vbExclamation, "CAPRIND v5.0"
            End If
        USMsgBox ("Configuração excluída com sucesso."), vbInformation, "CAPRIND v5.0"
        ProcLimpaCampos
        ProcCarregaBancoDados
        ProcCarregaListaBancos
        Novo_LocalBD = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravar()
On Error GoTo tratar_erro

If Novo_LocalBD = False Then
    ProcVerificaSalvar
    Exit Sub
End If

Acao = "salvar"
If txtLocalrel.Text = "" Then
    NomeCampo = "o local dos relatórios"
    ProcVerificaAcao
    Cmd_localizar_rel.SetFocus
    Exit Sub
End If
If Cmb_servidor = "" Then
    NomeCampo = "o nome da instância SQL"
    ProcVerificaAcao
    Cmb_servidor.SetFocus
    Exit Sub
End If
If Cmb_nome_banco = "" Then
    NomeCampo = "o nome do banco de dados"
    ProcVerificaAcao
    Cmb_nome_banco.SetFocus
    Exit Sub
End If
If txtlocalantigo.Text = "" Then
    NomeCampo = "o local onde esta armazenado os arquivos antigos"
    ProcVerificaAcao
    cmdLocalantigo.SetFocus
    Exit Sub
End If
If txtlocalnovo.Text = "" Then
    NomeCampo = "o local onde esta armazenado os novos arquivos"
    ProcVerificaAcao
    cmdLocalnovo.SetFocus
    Exit Sub
End If
Caprind = "\Caprind.exe"
Gerprod = "\Gerprod.exe"

If Cmb_servidor = NomeServidor And Cmb_nome_banco = Nome_banco Or Cmb_servidor = NomeServidor1 And Cmb_nome_banco = Nome_banco1 Or Cmb_servidor = NomeServidor2 And Cmb_nome_banco = Nome_banco2 Then
    USMsgBox ("Essa configuração de instância SQL e banco de dados já foi cadastrada, favor alterar."), vbExclamation, "CAPRIND v5.0"
    Cmb_servidor.SetFocus
    Exit Sub
End If

Permitido = True
ProcFunAbreBD_Configuracao Cmb_servidor, Cmb_nome_banco, IIf(Txt_usuario = "", "Procam", Txt_usuario), IIf(Txt_senha = "", "PRO0902loc$?", Txt_senha)
If Permitido = False Then
    USMsgBox "Não foi possivel salvar pois não foi econtrado essa instância e banco de dados.", vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Localrel = "" Then
    NomeServidor = Cmb_servidor
    SaveSetting "Procam", "CaprindSQL", "NomeServidor", NomeServidor
    
    Localrel = txtLocalrel.Text
    SaveSetting "Procam", "CaprindSQL", "LocalRel", Localrel
    
    Nome_banco = Cmb_nome_banco
    SaveSetting "Procam", "CaprindSQL", "Nome_banco", Nome_banco
    
    Usuario_banco = Txt_usuario
    SaveSetting "Procam", "CaprindSQL", "Usuario_banco", Usuario_banco
    
    Senha_banco = Txt_senha
    SaveSetting "Procam", "CaprindSQL", "Senha_banco", Senha_banco
    
    LocalAntigoCaprind = txtlocalantigo.Text & Caprind
    LocalAntigoGerprod = txtlocalantigo.Text & Gerprod
    LocalNovoCaprind = txtlocalnovo.Text & Caprind
    LocalNovoGerprod = txtlocalnovo.Text & Gerprod
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind", LocalAntigoCaprind
    SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind", LocalNovoCaprind
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod", LocalAntigoGerprod
    SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod", LocalNovoGerprod
    USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
ElseIf Localrel1 = "" Then
        NomeServidor1 = Cmb_servidor
        SaveSetting "Procam", "CaprindSQL", "NomeServidor1", NomeServidor1
        
        Localrel1 = txtLocalrel.Text
        SaveSetting "Procam", "CaprindSQL", "LocalRel1", Localrel1
        
        Nome_banco1 = Cmb_nome_banco
        SaveSetting "Procam", "CaprindSQL", "Nome_banco1", Nome_banco1
        
        Usuario_banco1 = Txt_usuario
        SaveSetting "Procam", "CaprindSQL", "Usuario_banco1", Usuario_banco1
        
        Senha_banco1 = Txt_senha
        SaveSetting "Procam", "CaprindSQL", "Senha_banco1", Senha_banco1
        
        LocalAntigoCaprind1 = txtlocalantigo.Text & Caprind
        LocalAntigoGerprod1 = txtlocalantigo.Text & Gerprod
        LocalNovoCaprind1 = txtlocalnovo.Text & Caprind
        LocalNovoGerprod1 = txtlocalnovo.Text & Gerprod
        SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind1", LocalAntigoCaprind1
        SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind1", LocalNovoCaprind1
        SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod1", LocalAntigoGerprod1
        SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod1", LocalNovoGerprod1
        USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
    ElseIf Localrel2 = "" Then
            NomeServidor2 = Cmb_servidor
            SaveSetting "Procam", "CaprindSQL", "NomeServidor2", NomeServidor2
            
            Localrel2 = txtLocalrel.Text
            SaveSetting "Procam", "CaprindSQL", "LocalRel2", Localrel2
            
            Nome_banco2 = Cmb_nome_banco
            SaveSetting "Procam", "CaprindSQL", "Nome_banco2", Nome_banco2
            
            Usuario_banco2 = Txt_usuario
            SaveSetting "Procam", "CaprindSQL", "Usuario_banco2", Usuario_banco2
            
            Senha_banco2 = Txt_senha
            SaveSetting "Procam", "CaprindSQL", "Senha_banco2", Senha_banco2
            
            LocalAntigoCaprind2 = txtlocalantigo.Text & Caprind
            LocalAntigoGerprod2 = txtlocalantigo.Text & Gerprod
            LocalNovoCaprind2 = txtlocalnovo.Text & Caprind
            LocalNovoGerprod2 = txtlocalnovo.Text & Gerprod
            SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind2", LocalAntigoCaprind2
            SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind2", LocalNovoCaprind2
            SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod2", LocalAntigoGerprod2
            SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod2", LocalNovoGerprod2
            USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
        ElseIf Localrel3 = "" Then
                NomeServidor3 = Cmb_servidor
                SaveSetting "Procam", "CaprindSQL", "NomeServidor3", NomeServidor3
                
                Localrel3 = txtLocalrel.Text
                SaveSetting "Procam", "CaprindSQL", "LocalRel3", Localrel3
                
                Nome_banco3 = Cmb_nome_banco
                SaveSetting "Procam", "CaprindSQL", "Nome_banco3", Nome_banco3
                
                Usuario_banco3 = Txt_usuario
                SaveSetting "Procam", "CaprindSQL", "Usuario_banco3", Usuario_banco3
                
                Senha_banco3 = Txt_senha
                SaveSetting "Procam", "CaprindSQL", "Senha_banco3", Senha_banco3
                
                LocalAntigoCaprind3 = txtlocalantigo.Text & Caprind
                LocalAntigoGerprod3 = txtlocalantigo.Text & Gerprod
                LocalNovoCaprind3 = txtlocalnovo.Text & Caprind
                LocalNovoGerprod3 = txtlocalnovo.Text & Gerprod
                SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind3", LocalAntigoCaprind3
                SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind3", LocalNovoCaprind3
                SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod3", LocalAntigoGerprod3
                SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod3", LocalNovoGerprod3
                USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
            Else
                 USMsgBox ("Você só pode armazenar quatro configurações diferentes."), vbExclamation, "CAPRIND v5.0"
End If

ProcBloqueiaCampos
Salvarrel = True
Main
FunAbreBD
Salvarrel = False
Novo_LocalBD = False
Unload Me
ProcCarregaListaBancos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcVerifInstancia Cmb_servidor
Novo_LocalBD = True
ProcLimpaCampos
ProcHabilitaCampos
'Cmd_localizar_rel_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtLocalrel = ""
Txt_usuario = ""
Txt_senha = ""
Cmb_servidor = ""
Cmb_nome_banco = ""
txtlocalantigo = App.Path
txtlocalnovo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaCampos()
On Error GoTo tratar_erro

Cmd_localizar_rel.Enabled = True
cmdLocalantigo.Enabled = True
cmdLocalnovo.Enabled = True
With Txt_usuario
    .Locked = False
    .TabStop = True
End With
With Txt_senha
    .Locked = False
    .TabStop = True
End With
With Cmb_servidor
    .Locked = False
    .TabStop = True
End With
With Cmb_nome_banco
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

Cmd_localizar_rel.Enabled = False
cmdLocalantigo.Enabled = False
cmdLocalnovo.Enabled = False
With Txt_usuario
    .Locked = True
    .TabStop = False
End With
With Txt_senha
    .Locked = True
    .TabStop = False
End With
With Cmb_servidor
    .Locked = True
    .TabStop = False
End With
With Cmb_nome_banco
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaBancos()
On Error GoTo tratar_erro

listaBancos.ListItems.Clear
If Localrel <> "" Then
    With listaBancos.ListItems
        .Add , , 1
        .Item(.Count).SubItems(1) = Localrel
        If NomeServidor <> "" Then .Item(.Count).SubItems(2) = NomeServidor
        If Nome_banco <> "" Then .Item(.Count).SubItems(3) = Nome_banco
    End With
End If
If Localrel1 <> "" Then
    With listaBancos.ListItems
        .Add , , 2
        .Item(.Count).SubItems(1) = Localrel1
        If NomeServidor1 <> "" Then .Item(.Count).SubItems(2) = NomeServidor1
        If Nome_banco1 <> "" Then .Item(.Count).SubItems(3) = Nome_banco1
    End With
End If
If Localrel2 <> "" Then
    With listaBancos.ListItems
        .Add , , 3
        .Item(.Count).SubItems(1) = Localrel2
        If NomeServidor2 <> "" Then .Item(.Count).SubItems(2) = NomeServidor2
        If Nome_banco2 <> "" Then .Item(.Count).SubItems(3) = Nome_banco2
    End With
End If
If Localrel3 <> "" Then
    With listaBancos.ListItems
        .Add , , 4
        .Item(.Count).SubItems(1) = Localrel3
        If NomeServidor3 <> "" Then .Item(.Count).SubItems(2) = NomeServidor3
        If Nome_banco3 <> "" Then .Item(.Count).SubItems(3) = Nome_banco3
    End With
End If
If Localrel <> "" Or Localrel1 <> "" Or Localrel2 <> "" Or Localrel3 <> "" Then PBLista.Value = 100

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalantigo_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Localizar local arquivos antigos"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtlocalantigo.Text = sBuffer
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalnovo_Click()
On Error GoTo tratar_erro

szTitle = vbCr & vbCr & "Localizar local novos arquivos"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtlocalnovo.Text = sBuffer
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
    Case vbKeyF3: ProcGravar
    Case vbKeyF4: ProcExcluir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Novo_LocalBD = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9100, 7, True

txtLocalrel = Localrel
Txt_usuario = Usuario_banco
Txt_senha = Senha_banco
Cmb_servidor = NomeServidor
Cmb_nome_banco = Nome_banco

If LocalAntigoCaprind <> "" And LocalNovoCaprind <> "" Then
    txtlocalantigo.Text = Left(LocalAntigoCaprind, Len(LocalAntigoCaprind) - 12)
    txtlocalnovo.Text = Left(LocalNovoCaprind, Len(LocalNovoCaprind) - 12)
End If
    
1:
    ProcCarregaListaBancos
    ProcBloqueiaCampos

Exit Sub
tratar_erro:
    If Err.Number = 383 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listaBancos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listaBancos.ListItems.Count = 0 Then Exit Sub
Select Case listaBancos.SelectedItem
    Case 1:
        txtLocalrel = Localrel
        Txt_usuario = Usuario_banco
        Txt_senha = Senha_banco
        NomeCampo = NomeServidor
        Cmb_servidor = NomeServidor
        Cmb_nome_banco = Nome_banco
        If LocalAntigoCaprind <> "" And LocalNovoCaprind <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind, Len(LocalAntigoCaprind) - 12)
            txtlocalnovo = Left(LocalNovoCaprind, Len(LocalNovoCaprind) - 12)
        End If
    Case 2:
        txtLocalrel = Localrel1
        Txt_usuario = Usuario_banco1
        Txt_senha = Senha_banco1
        NomeCampo = NomeServidor1
        Cmb_servidor = NomeServidor1
        Cmb_nome_banco = Nome_banco1
        If LocalAntigoCaprind1 <> "" And LocalNovoCaprind1 <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind1, Len(LocalAntigoCaprind1) - 12)
            txtlocalnovo = Left(LocalNovoCaprind1, Len(LocalNovoCaprind1) - 12)
        End If
    Case 3:
        txtLocalrel = Localrel2
        Txt_usuario = Usuario_banco2
        Txt_senha = Senha_banco2
        NomeCampo = NomeServidor2
        Cmb_servidor = NomeServidor2
        Cmb_nome_banco = Nome_banco2
        If LocalAntigoCaprind2 <> "" And LocalNovoCaprind2 <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind2, Len(LocalAntigoCaprind2) - 12)
            txtlocalnovo = Left(LocalNovoCaprind2, Len(LocalNovoCaprind2) - 12)
        End If
    Case 4:
        txtLocalrel = Localrel3
        Txt_usuario = Usuario_banco3
        Txt_senha = Senha_banco3
        NomeCampo = NomeServidor3
        Cmb_servidor = NomeServidor3
        Cmb_nome_banco = Nome_banco3
        If LocalAntigoCaprind3 <> "" And LocalNovoCaprind3 <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind3, Len(LocalAntigoCaprind3) - 12)
            txtlocalnovo = Left(LocalNovoCaprind3, Len(LocalNovoCaprind3) - 12)
        End If
End Select
Novo_LocalBD = False
ProcBloqueiaCampos

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("A instância " & NomeCampo & " não está disponível."), vbExclamation, "CAPRIND v5.0"
        txtLocalrel = ""
        Txt_usuario = ""
        Txt_senha = ""
        Cmb_servidor = ""
        Cmb_nome_banco = ""
        txtlocalantigo = ""
        txtlocalnovo = ""
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcGravar
    Case 3: ProcExcluir
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
