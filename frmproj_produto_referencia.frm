VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmproj_produto_referencia 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Produtos e serviços - Cadastro de códigos de referência"
   ClientHeight    =   10035
   ClientLeft      =   1695
   ClientTop       =   1500
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
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
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
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
      Left            =   2790
      MaxLength       =   50
      MouseIcon       =   "frmproj_produto_referencia.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Text            =   "0"
      ToolTipText     =   "Tipo."
      Top             =   4050
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_tipo 
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
      Left            =   2160
      MaxLength       =   50
      MouseIcon       =   "frmproj_produto_referencia.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Tipo."
      Top             =   4050
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_ID_cliente_forn 
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
      Left            =   1530
      MaxLength       =   50
      MouseIcon       =   "frmproj_produto_referencia.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Text            =   "0"
      ToolTipText     =   "ID do cliente/fornecedor."
      Top             =   4050
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtid 
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
      Left            =   900
      MaxLength       =   50
      MouseIcon       =   "frmproj_produto_referencia.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Text            =   "0"
      ToolTipText     =   "Id."
      Top             =   4050
      Visible         =   0   'False
      Width           =   615
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   21
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   20
      Top             =   0
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
      ButtonCaption4  =   "Anterior"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Registro anterior."
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
      ButtonWidth4    =   47
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Próximo"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Próximo registro."
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
      ButtonLeft5     =   167
      ButtonTop5      =   2
      ButtonWidth5    =   46
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Atualizar"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Utilizado pelo administrador do sistema."
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
      ButtonLeft6     =   215
      ButtonTop6      =   2
      ButtonWidth6    =   50
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
      ButtonLeft7     =   267
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   271
      ButtonTop8      =   2
      ButtonWidth8    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft9     =   309
      ButtonTop9      =   2
      ButtonWidth9    =   26
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
      ButtonLeft10    =   337
      ButtonTop10     =   2
      ButtonWidth10   =   24
      ButtonHeight10  =   24
      ButtonUseMaskColor10=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7230
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmproj_produto_referencia.frx":0C28
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1455
      Left            =   55
      TabIndex        =   7
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtdescricaoproduto 
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
         Height          =   345
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Cliente/Fornecedor."
         Top             =   390
         Width           =   8265
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
         ItemData        =   "frmproj_produto_referencia.frx":6109
         Left            =   180
         List            =   "frmproj_produto_referencia.frx":6119
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Grupo de clientes."
         Top             =   990
         Width           =   6690
      End
      Begin VB.TextBox txtaplicacao 
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
         Left            =   6870
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Cliente/Fornecedor."
         Top             =   990
         Width           =   7485
      End
      Begin VB.TextBox Txt_revisao 
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
         Height          =   345
         Left            =   14550
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Revisão."
         Top             =   390
         Width           =   495
      End
      Begin VB.TextBox Txt_numero_desenho 
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
         Height          =   345
         Left            =   12630
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Número do desenho."
         Top             =   390
         Width           =   1905
      End
      Begin VB.TextBox txtreferencia 
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
         Height          =   345
         Left            =   10440
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Código de referência."
         Top             =   390
         Width           =   2175
      End
      Begin VB.TextBox txtdesenhoproduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   1965
      End
      Begin DrawSuite2022.USButton cmdLocalizarCliente 
         Height          =   315
         Left            =   14370
         TabIndex        =   22
         ToolTipText     =   "Localizar cliente | Fornecedor."
         Top             =   990
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         DibPicture      =   "frmproj_produto_referencia.frx":6169
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
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         ForeColor       =   0
         ForeColorOver   =   0
         ForeColorDown   =   0
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         PicAlign        =   8
         Theme           =   1
      End
      Begin DrawSuite2022.USButton Cmd_limpar_ClieForn 
         Height          =   315
         Left            =   14730
         TabIndex        =   23
         ToolTipText     =   "Localizar cliente | Fornecedor."
         Top             =   990
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         DibPicture      =   "frmproj_produto_referencia.frx":97B9
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
         BorderColor     =   5263559
         BorderColorDown =   2039646
         BorderColorOver =   3026574
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorOver1=   3026574
         GradientColorOver2=   3026574
         GradientColorOver3=   3026574
         GradientColorOver4=   3026574
         GradientColorDown1=   2039646
         GradientColorDown2=   2039646
         GradientColorDown3=   2039646
         GradientColorDown4=   2039646
         PicAlign        =   8
         Theme           =   4
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente | Fornecedor"
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
         Left            =   9877
         TabIndex        =   18
         Top             =   780
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo"
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
         Left            =   3308
         TabIndex        =   17
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label3 
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
         Left            =   14610
         TabIndex        =   16
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código do projeto"
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
         Left            =   12795
         TabIndex        =   13
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de referência"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   10657
         TabIndex        =   11
         Top             =   180
         Width           =   1740
      End
      Begin VB.Label Label18 
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
         Left            =   547
         TabIndex        =   10
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do item"
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
         Left            =   5940
         TabIndex        =   9
         Top             =   180
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   7290
      Left            =   60
      TabIndex        =   6
      Top             =   2460
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12859
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
         Text            =   "Cód. de referência"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "N° do desenho"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9529
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cliente/Fornecedor"
         Object.Width           =   9529
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4680
      TabIndex        =   8
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "frmproj_produto_referencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Referencia As Boolean 'OK

Private Sub cmbGrupo_Click()
On Error GoTo tratar_erro

If cmbGrupo <> "" Then
    If cmbGrupo.ListIndex <> -1 And cmbGrupo.ListIndex <> 0 Then
        txtIDGrupo = cmbGrupo.ItemData(cmbGrupo.ListIndex)
        cliente_excluir = txtaplicacao
        txtaplicacao = ""
        Txt_ID_cliente_forn = "0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtid = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from item_aplicacoes where codproduto = " & frmproj_produto.txtcodproduto & " order by iditem", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("iditem = " & txtid)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtid = TBLISTA!IDitem
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where iditem = " & txtid, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcLimpaCampos
            ProcCarregaDados
        End If
    Else
        USMsgBox ("Fim dos cadastros de código de referência."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_ClieForn_Click()
On Error GoTo tratar_erro

txtIDGrupo = 0
Clientes_Grupos = False
cmbGrupo.ListIndex = -1
Txt_ID_cliente_forn = 0
txtaplicacao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarCliente_Click()
On Error GoTo tratar_erro

If Clientes_Grupos = True Then Exit Sub
frmproj_produto_referencia_cliente_forn.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtid = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from item_aplicacoes where codproduto = " & frmproj_produto.txtcodproduto & " order by iditem", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("iditem = " & txtid)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtid = TBLISTA!IDitem
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where iditem = " & txtid, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcLimpaCampos
            ProcCarregaDados
        End If
    Else
        USMsgBox ("Fim dos cadastros de código de referência."), vbInformation, "CAPRIND v5.0"
    End If
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
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 10, True

If Engenharia_Produtos = True Then
    Caption = "Engenharia - Produtos e serviços - Cadastro de códigos de referência"
    Formulario = "Engenharia/Produtos e serviços/Cadastro de códigos de referência"
End If
If Compras_Produtos = True Then
    Caption = "Compras - Produtos e serviços - Cadastro de códigos de referência"
    Formulario = "Compras/Produtos e serviços/Cadastro de códigos de referência"
End If
If Vendas_Produtos = True Then
    Caption = "Vendas - Produtos e serviços - Cadastro de códigos de referência"
    Formulario = "Vendas/Produtos e serviços/Cadastro de códigos de referência"
End If
txtdesenhoproduto.Text = frmproj_produto.txtdesenhoproduto.Text
txtdescricaoproduto.Text = frmproj_produto.txtdescricaoproduto
ProcLimpaVariaveisPrincipais
AtualizaLista

ProcCarregaComboGrupoCliente cmbGrupo, True

ProcRemoveObjetosResize Me
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub AtualizaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from item_aplicacoes where codproduto = " & frmproj_produto.txtcodproduto.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
       With Lista.ListItems
            .Add , , TBLISTA!IDitem
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!N_referencia), "", (TBLISTA!N_referencia))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Rev), "", (TBLISTA!Rev))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Desenho), "", (TBLISTA!Desenho))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Descricao), "", (TBLISTA!Descricao))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Aplicacao), "", (TBLISTA!Aplicacao))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Engenharia_Produtos = True Then Formulario = "Engenharia/Produtos e serviços/Cadastro de códigos de referência"
If Compras_Produtos = True Then Formulario = "Compras/Produtos e serviços/Cadastro de códigos de referência"
If Vendas_Produtos = True Then Formulario = "Vendas/Produtos e serviços/Cadastro de códigos de referência"
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362R" Then
    If USMsgBox("Deseja realmente atualizar os dados dos códigos de referência?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where aplicacao <> 'Null' order by aplicacao", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBAbrir.MoveFirst
            Do While TBAbrir.EOF = False
                Set TBClientes = CreateObject("adodb.recordset")
                TBClientes.Open "Select * from Clientes where NomeRazao = '" & IIf(TBAbrir!Aplicacao = "", 0, TBAbrir!Aplicacao) & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then
                    TBAbrir!ID_cliente_forn = TBClientes!IDCliente
                    TBAbrir!Tipo = "C"
                    TBAbrir.Update
                End If
                TBClientes.Close
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
        Evento = "Atualizar"
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

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) código(s) de referência?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from item_aplicacoes WHERE iditem = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = Formulario
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Cód. interno: " & txtdesenhoproduto
                Documento1 = "Cód. referência: " & TBFI!N_referencia
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from item_aplicacoes where iditem = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) código(s) de referência antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Código(s) de referência excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    AtualizaLista
    Novo_Referencia = False
    Frame2.Enabled = False
    Caption = "Engenharia - Produtos e serviços - Cadastro de códigos de referência"
    
    With frmproj_produto
        .ProcSalvarUltimaAlteracao .txtcodproduto
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtid.Text = 0
txtIDGrupo = 0
cmbGrupo.ListIndex = -1
Txt_ID_cliente_forn = 0
Txt_tipo = ""
txtaplicacao.Text = ""
txtreferencia.Text = ""
Txt_revisao = ""
Txt_numero_desenho = ""
CodigoLista = 0
Caption = "Engenharia - Produtos e serviços - Cadastro de códigos de referência"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtreferencia.Text = "" Then
    NomeCampo = "o código de referência"
    ProcVerificaAcao
    txtreferencia.SetFocus
    Exit Sub
End If
If Clientes_Grupos = True And cmbGrupo = "" Then
    NomeCampo = "o grupo do cliente"
    ProcVerificaAcao
    cmbGrupo.SetFocus
    Exit Sub
End If

'Verifica se já existe código de ref. cadastrado
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select IDitem, Codproduto from item_aplicacoes where n_referencia = '" & txtreferencia & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select Desenho from projproduto where codproduto = " & TBItem!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBComponente.EOF = False Then
        If TBComponente!Desenho <> txtdesenhoproduto Then
            If USMsgBox("Este código de referência está sendo utilizado em outro(s) produto(s), deseja excluir para ser salvo no produto " & txtdesenhoproduto & "?", vbYesNo) = vbYes Then
                If USMsgBox("Deseja realmente excluir o código de referência " & txtreferencia & " no(s) outro(s) produto(s)?", vbYesNo) = vbYes Then
                    Conexao.Execute "DELETE from item_aplicacoes where n_referencia = '" & txtreferencia & "'"
                    TBComponente.Close
                    GoTo Referencia
                End If
            End If
        ElseIf Novo_Referencia = True Or TBItem!IDitem <> txtid Then
                'If usMsgbox("Este código de referência já foi cadastrado para este produto, deseja prosseguir mesmo assim?", vbyesno, "CAPRIND v5.0") = vbNo Then
                USMsgBox ("Este código de referência já foi cadastrado para este produto, favor alterar."), vbExclamation, "CAPRIND v5.0"
                TBComponente.Close
                Exit Sub
                'End If
        End If
    End If
    TBComponente.Close
End If
TBItem.Close
Referencia:

    If Txt_ID_cliente_forn <> 0 Then
Alterar:
        If Novo_Referencia = True Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from item_aplicacoes where iditem = " & txtid.Text, Conexao, adOpenKeyset, adLockOptimistic
            USMsgBox "Novo código de referência cadatrado com sucesso.", vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = Formulario
            Evento = "Novo"
            ID_documento = txtid
            Documento = "Cód. interno: " & txtdesenhoproduto
            Documento1 = "Cód. referência: " & txtreferencia
            ProcGravaEvento
            '==================================
        Else
            If FunVerifValidacaoRegistro("alterar", frmproj_produto.txtDtValidacao, "registro", "este código de referência", True) = False Then Exit Sub
            
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from item_aplicacoes where iditem = " & txtid.Text & " and n_referencia = '" & Lista.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = Formulario
            Evento = "Alterar"
            ID_documento = txtid
            Documento = "Cód. interno: " & txtdesenhoproduto
            Documento1 = "Cód. referência: " & Lista.SelectedItem.ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
        If TBProduto.EOF = False Then
            If txtreferencia <> TBProduto!N_referencia Then
                Conexao.Execute "Update Compras_pedido_lista Set n_referencia = '" & txtreferencia & "' where Desenho = '" & txtdesenhoproduto & "' and n_referencia = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update Estoque_Controle Set ref = '" & txtreferencia & "' where Desenho = '" & txtdesenhoproduto & "' and ref = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update Producao Set n_referencia = '" & txtreferencia & "' where Desenho = '" & txtdesenhoproduto & "' and n_referencia = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update QualidadePPAP Set n_referencia = '" & txtreferencia & "' where CodInterno = '" & txtdesenhoproduto & "' and n_referencia = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update QualidadePPAP_PlanoControle Set n_referencia = '" & txtreferencia & "' where CodInterno = '" & txtdesenhoproduto & "' and  n_referencia = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update tbl_Detalhes_Nota Set n_referencia = '" & txtreferencia & "' where int_Cod_Produto = '" & txtdesenhoproduto & "' and  n_referencia = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update Vendas_analise Set n_referencia = '" & txtreferencia & "' where CodInterno = '" & txtdesenhoproduto & "' and n_referencia = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update Vendas_analise_setores Set n_referencia = '" & txtreferencia & "' where CodInterno = '" & txtdesenhoproduto & "' and n_referencia = '" & TBProduto!N_referencia & "'"
                Conexao.Execute "Update Vendas_carteira Set n_referencia = '" & txtreferencia & "' where Desenho = '" & txtdesenhoproduto & "' and n_referencia = '" & TBProduto!N_referencia & "'"
            End If
        Else
            TBProduto.AddNew
        End If
        TBProduto!Codproduto = frmproj_produto.txtcodproduto
        TBProduto!ID_cliente_forn = Txt_ID_cliente_forn
        TBProduto!Tipo = Txt_tipo
        If txtreferencia.Text <> "" Then TBProduto!N_referencia = txtreferencia.Text
        TBProduto!Rev = Txt_revisao
        TBProduto!Desenho = Txt_numero_desenho
        If txtaplicacao.Text <> "" Then TBProduto!Aplicacao = txtaplicacao Else TBProduto!Aplicacao = Null
        If txtdescricaoproduto.Text <> "" Then TBProduto!Descricao = txtdescricaoproduto.Text
        TBProduto!idgrupo = txtIDGrupo
        TBProduto.Update
        txtid = TBProduto!IDitem
        TBProduto.Close
    Else
        With frmproj_produto
            If Clientes_Grupos = True And .Chk_vendas.Value = 1 And txtIDGrupo = "0" Then ProcAtualizaGrupo Else GoTo Alterar
        End With
    End If
    
    With frmproj_produto
        .ProcSalvarUltimaAlteracao .txtcodproduto
    End With
    
    AtualizaLista
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
    Novo_Referencia = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If FunVerifValidacaoRegistro("criar novo", frmproj_produto.txtDtValidacao, "registro", "código de referência", True) = False Then Exit Sub

ProcLimpaCampos
Novo_Referencia = True
Frame2.Enabled = True
With cmbGrupo
    If USMsgBox("Deseja criar o código de referência por grupo de clientes? ", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Clientes_Grupos = True
        .Locked = False
        .TabStop = True
    Else
        Clientes_Grupos = False
        .Locked = True
        .TabStop = False
    End If
End With
txtreferencia.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Referencia = True Then
    If USMsgBox("O código de referência ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Referencia = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Referencia = False
Unload Me

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
                If FunVerificaRegistroValidadoSemMsg("projproduto", "Codproduto = " & frmproj_produto.txtcodproduto, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido_lista", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Producao", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "QualidadePPAP", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "QualidadePPAP_PlanoControle", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_analise", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_analise_setores", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_carteira", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Estoque_Controle", "ref = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                If Permitido = False Then
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
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("projproduto", "Codproduto = " & frmproj_produto.txtcodproduto, "registro", "código de referência", "excluir este", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            Mensagem = "Não é permitido excluir este código de referência, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "Compras_pedido_lista", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and Codproduto = " & frmproj_produto.txtcodproduto, "Compras"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Producao", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and desenho = '" & frmproj_produto.txtdesenhoproduto & "'", "PCP/Gerenciamento de ordem"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "QualidadePPAP", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and codinterno = '" & frmproj_produto.txtdesenhoproduto & "'", "Qualidade/PPAP/PSW"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "QualidadePPAP_PlanoControle", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and codinterno = '" & frmproj_produto.txtdesenhoproduto & "'", "Qualidade/PPAP/PSW"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Detalhes_Nota", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and Codproduto = " & frmproj_produto.txtcodproduto, "Faturamento/Nota fiscal"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Vendas_analise", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and IDproduto = " & frmproj_produto.txtcodproduto, "Outros/Análise crítica"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Vendas_analise_setores", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and IDproduto = " & frmproj_produto.txtcodproduto, "Outros/Análise crítica"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Vendas_carteira", "n_referencia = '" & .ListItems(InitFor).ListSubItems(1) & "' and desenho = '" & frmproj_produto.txtdesenhoproduto & "'", "Vendas"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Estoque_Controle", "ref = '" & .ListItems(InitFor).ListSubItems(1) & "' and desenho = '" & frmproj_produto.txtdesenhoproduto & "'", "Estoque/Movimentação"
            If Permitido = False Then
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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from item_aplicacoes where iditem = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtid.Text = TBAbrir!IDitem
Txt_tipo = IIf(IsNull(TBAbrir!Tipo), "", TBAbrir!Tipo)
Caption = "Engenharia - Produtos e serviços - Cadastro de códigos de referência - (Cód. de referência : " & IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia) & ")"
txtreferencia.Text = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
Txt_revisao = IIf(IsNull(TBAbrir!Rev), "", TBAbrir!Rev)
Txt_numero_desenho = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
Clientes_Grupos = False
With cmbGrupo
    .Locked = True
    .TabStop = False
    If IsNull(TBAbrir!idgrupo) = False And TBAbrir!idgrupo <> "0" Then
        txtIDGrupo = TBAbrir!idgrupo
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "select * from clientes_grupos where id = " & TBAbrir!idgrupo, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            .Text = TBFIltro!Texto
            .Locked = False
            .TabStop = True
            Clientes_Grupos = True
        End If
        TBFIltro.Close
    End If
End With
Txt_ID_cliente_forn = IIf(IsNull(TBAbrir!ID_cliente_forn), 0, TBAbrir!ID_cliente_forn)
txtaplicacao.Text = IIf(IsNull(TBAbrir!Aplicacao), "", TBAbrir!Aplicacao)
Novo_Referencia = False
Frame2.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaGrupo()
On Error GoTo tratar_erro

Conexao.Execute "DELETE from item_aplicacoes where idgrupo = " & txtIDGrupo & " and n_referencia = '" & txtreferencia & "'"
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "select * from clientes where idgrupo = " & txtIDGrupo, Conexao, adOpenKeyset, adLockOptimistic
Do While TBClientes.EOF = False
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "select * from item_aplicacoes", Conexao, adOpenKeyset, adLockOptimistic
    TBFIltro.AddNew
    TBFIltro!Codproduto = frmproj_produto.txtcodproduto
    TBFIltro!ID_cliente_forn = TBClientes!IDCliente
    TBFIltro!Tipo = "C"
    TBFIltro!N_referencia = IIf(txtreferencia.Text = "", Null, txtreferencia)
    TBFIltro!Rev = Txt_revisao
    TBFIltro!Desenho = Txt_numero_desenho
    TBFIltro!Aplicacao = IIf(IsNull(TBClientes!NomeRazao), "", TBClientes!NomeRazao)
    TBFIltro!Descricao = IIf(txtdescricaoproduto = "", Null, txtdescricaoproduto)
    TBFIltro!idgrupo = txtIDGrupo
    TBFIltro.Update
    TBFIltro.Close
    TBClientes.MoveNext
Loop
TBClientes.Close
USMsgBox "Novo código de referência cadatrado com sucesso.", vbInformation, "CAPRIND v5.0"
'==================================
Modulo = Formulario
Evento = "Novo"
ID_documento = txtid
Documento = "Cód. interno: " & txtdesenhoproduto
Documento1 = "Cód. referência: " & txtreferencia
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    Case 6: ProcAtualizar
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
