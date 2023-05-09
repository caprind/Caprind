VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_Requisicao_Lista 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Outros - Follow up de compras"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alterar status para"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   12940
      TabIndex        =   38
      Top             =   2515
      Width           =   2310
      Begin VB.ComboBox Cmb_alterar_status 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmCompras_Requisicao_Lista.frx":0000
         Left            =   180
         List            =   "frmCompras_Requisicao_Lista.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   195
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   34
      Top             =   9120
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
         TabIndex        =   5
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
         TabIndex        =   4
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   9
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Requisicao_Lista.frx":0022
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
         TabIndex        =   8
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Requisicao_Lista.frx":37C9
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
         TabIndex        =   6
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
         TabIndex        =   7
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Requisicao_Lista.frx":72E0
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
         TabIndex        =   10
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Requisicao_Lista.frx":B3D7
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
         Left            =   4410
         TabIndex        =   39
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label18 
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
         Index           =   0
         Left            =   3090
         TabIndex        =   35
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.CheckBox Chk_prazo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo de entrega"
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
      Left            =   6600
      TabIndex        =   20
      Top             =   2775
      Width           =   1755
   End
   Begin VB.CheckBox Chk_pedido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pedido"
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
      Left            =   5355
      TabIndex        =   19
      Top             =   2775
      Width           =   885
   End
   Begin VB.CheckBox Chk_cotacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cotação"
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
      Left            =   4005
      TabIndex        =   18
      Top             =   2775
      Width           =   1005
   End
   Begin VB.CheckBox Chk_requisicao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requisição"
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
      Left            =   2430
      TabIndex        =   17
      Top             =   2775
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   645
      Left            =   55
      TabIndex        =   24
      Top             =   2505
      Width           =   12855
      Begin MSComCtl2.DTPicker txtinicio 
         Height          =   315
         Left            =   9510
         TabIndex        =   21
         ToolTipText     =   "Data início para pesquisa."
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   198508545
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txtfim 
         Height          =   315
         Left            =   11400
         TabIndex        =   22
         ToolTipText     =   "Data final para pesquisa."
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   198508545
         CurrentDate     =   39057
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
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
         Left            =   9150
         TabIndex        =   26
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         Left            =   10950
         TabIndex        =   25
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   27
      Top             =   990
      Width           =   15195
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   10230
         TabIndex        =   41
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
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
            TabIndex        =   15
            Top             =   180
            Width           =   705
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
            TabIndex        =   13
            Top             =   180
            Width           =   1275
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
            TabIndex        =   12
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
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
            TabIndex        =   14
            Top             =   180
            Width           =   1155
         End
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
         ItemData        =   "frmCompras_Requisicao_Lista.frx":EC89
         Left            =   180
         List            =   "frmCompras_Requisicao_Lista.frx":EC8B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   5415
      End
      Begin VB.CheckBox Chk_remessa 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Incluir remessa"
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
         Left            =   13380
         TabIndex        =   16
         Top             =   1118
         Width           =   1635
      End
      Begin VB.ComboBox cmbstatus 
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
         ItemData        =   "frmCompras_Requisicao_Lista.frx":EC8D
         Left            =   10350
         List            =   "frmCompras_Requisicao_Lista.frx":EC8F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Status. "
         Top             =   1050
         Width           =   2925
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
         Height          =   330
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   10155
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
         ItemData        =   "frmCompras_Requisicao_Lista.frx":EC91
         Left            =   5610
         List            =   "frmCompras_Requisicao_Lista.frx":ECB6
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4515
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmCompras_Requisicao_Lista.frx":ED3E
         Left            =   180
         List            =   "frmCompras_Requisicao_Lista.frx":ED40
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   10155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   11580
         TabIndex        =   33
         Top             =   840
         Width           =   465
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
         Left            =   2520
         TabIndex        =   32
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label45 
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
         Left            =   7447
         TabIndex        =   29
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label13 
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
         Left            =   4522
         TabIndex        =   28
         Top             =   840
         Width           =   1470
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   30
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
      TabIndex        =   31
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
      ButtonCaption3  =   "Status"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Status (F7)"
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
      ButtonLeft4     =   134
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
      ButtonLeft5     =   138
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
      ButtonLeft6     =   176
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
      ButtonLeft7     =   204
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   8370
         Top             =   90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_Requisicao_Lista.frx":ED42
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5940
      Left            =   60
      TabIndex        =   3
      Top             =   3165
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10478
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
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Prioridade"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4410
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
         Text            =   "Vlr. unit."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Solicitação"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Cotação"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Pedido"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1729
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Qtde. req."
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Qtde. compr."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Qtde. receb."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Object.Tag             =   "D"
         Text            =   "Prazo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Tag             =   "T"
         Text            =   "Aprovado"
         Object.Width           =   2293
      EndProperty
   End
End
Attribute VB_Name = "frmCompras_Requisicao_Lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql_Followup_Compras As String 'OK
Dim FormulaRel_Followup_Compras As String 'OK
Dim TBLISTA_Followup_Compras As ADODB.Recordset 'OK

Private Sub Chk_cotacao_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If Chk_cotacao.Value = 1 Then
    Chk_requisicao.Value = 0
    Chk_pedido.Value = 0
    Chk_prazo.Value = 0
    Frame4.Enabled = True
    txtinicio.SetFocus
Else
    Frame4.Enabled = False
    txtinicio.Value = Date
    txtfim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_pedido_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If Chk_pedido.Value = 1 Then
    Chk_requisicao.Value = 0
    Chk_cotacao.Value = 0
    Chk_prazo.Value = 0
    Frame4.Enabled = True
    txtinicio.SetFocus
Else
    Frame4.Enabled = False
    txtinicio.Value = Date
    txtfim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_prazo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If Chk_prazo.Value = 1 Then
    Chk_requisicao.Value = 0
    Chk_cotacao.Value = 0
    Chk_pedido.Value = 0
    Frame4.Enabled = True
    txtinicio.SetFocus
Else
    Frame4.Enabled = False
    txtinicio.Value = Date
    txtfim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_remessa_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_requisicao_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If Chk_requisicao.Value = 1 Then
    Chk_cotacao.Value = 0
    Chk_pedido.Value = 0
    Chk_prazo.Value = 0
    Frame4.Enabled = True
    txtinicio.SetFocus
Else
    Frame4.Enabled = False
    txtinicio.Value = Date
    txtfim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_alterar_status_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
txtTexto.Visible = True
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Setor" Or cmbfiltrarpor = "Requisitante" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbTexto, "familia is not null and compras = 'True'", True
    ElseIf cmbfiltrarpor = "Grupo" Then
            ProcCarregaComboGrupoFamilia cmbTexto, "Grupo is not null", True
        Else
            If cmbfiltrarpor = "Setor" Then
                TabelaFiltro = "Usuarios_Setor"
                TextoFiltro = "Setor"
            Else
                TabelaFiltro = "Usuarios"
                TextoFiltro = "Usuario"
            End If
            With cmbTexto
                .Clear
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select " & TextoFiltro & " as NomeCampo from " & TabelaFiltro & " Group by " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    .AddItem ""
                    Do While TBLISTA.EOF = False
                        If IsNull(TBLISTA!NomeCampo) = False And TBLISTA!NomeCampo <> "" Then .AddItem TBLISTA!NomeCampo
                        TBLISTA.MoveNext
                    Loop
                End If
                TBLISTA.Close
            End With
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbstatus_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbStatus = "Requisitado" Then
    Chk_requisicao.Enabled = True
    Chk_cotacao.Enabled = False
    Chk_cotacao.Value = 0
    Chk_pedido.Enabled = False
    Chk_pedido.Value = 0
ElseIf cmbStatus = "Cotando" Or cmbStatus = "Não aprovado" Then
        Chk_requisicao.Enabled = True
        Chk_cotacao.Enabled = True
        Chk_pedido.Enabled = False
        Chk_pedido.Value = 0
    Else
        Chk_requisicao.Enabled = True
        Chk_cotacao.Enabled = True
        Chk_pedido.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o status desse(s) produto(s)/serviços(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select IDpedido, Status_Item from compras_pedido_lista where idlista = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = False Then
                If Cmb_alterar_status = "Comprado" Then
                    TBCompras!Status_Item = "N_RECEBIDO"
                    TBCompras.Update
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from compras_pedido_lista where IDPedido = " & TBCompras!IDpedido & " and (status_item = 'PARCIAL' or status_item = 'RECEBIDO')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then StatusTexto = "PARCIAL" Else StatusTexto = "ABERTO"
                    TBAbrir.Close
                    
                    Evento = "Alterar status p/ comprado"
                Else
                    TBCompras!Status_Item = "RECEBIDO"
                    TBCompras.Update
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from compras_pedido_lista where IDPedido = " & TBCompras!IDpedido & " and (status_item = 'PARCIAL' or status_item = 'N_RECEBIDO')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then StatusTexto = "PARCIAL" Else StatusTexto = "ENCERRADO"
                    TBAbrir.Close
                    
                    Evento = "Alterar status p/ recebido"
                End If
                Conexao.Execute "Update Compras_pedido Set Status_pedido = '" & StatusTexto & "' where IDPedido = " & TBCompras!IDpedido
                
                '==================================
                Modulo = "Outros/Follow up de compras"
                ID_documento = .ListItems(InitFor)
                Documento = "Cód. interno: " & .ListItems(InitFor).ListSubItems(2)
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBCompras.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviços(s) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Status do(s) produto(s)/serviços(s) alterado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With txtfim
    If FunVerificaDataFinal(txtinicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

'Verifica o filtro do status
Select Case cmbStatus
    Case "Requisitado":
        TextoFiltroStatus = "status_item = 'REQUISIT.'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'REQUISIT.'"
        TextoFiltroEmpresa = "ID_empresa_req"
    Case "Cotando":
        TextoFiltroStatus = "status_item = 'COTANDO'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'COTANDO'"
        TextoFiltroEmpresa = "ID_empresa_cot"
    Case "Aguardando aprovação":
        TextoFiltroStatus = "status_item = 'AGUARDANDO APROVAÇÃO'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'AGUARDANDO APROVAÇÃO'"
        TextoFiltroEmpresa = "ID_empresa_ped"
    Case "Aprovado":
        TextoFiltroStatus = "status_item = 'APROVADO'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'APROVADO'"
        TextoFiltroEmpresa = "ID_empresa_ped"
    Case "Comprado":
        TextoFiltroStatus = "status_item = 'N_RECEBIDO'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'N_RECEBIDO'"
        TextoFiltroEmpresa = "ID_empresa_ped"
    Case "Recebido":
        TextoFiltroStatus = "status_item = 'RECEBIDO'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'RECEBIDO'"
        TextoFiltroEmpresa = "ID_empresa_ped"
    Case "Recebido Parcial":
        TextoFiltroStatus = "status_item = 'PARCIAL'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'PARCIAL'"
        TextoFiltroEmpresa = "ID_empresa_ped"
    Case "Comprado e Recebido Parcial"
        TextoFiltroStatus = "(status_item = 'N_RECEBIDO' or status_item = 'PARCIAL')"
        TextoFiltroStatusRel = "({Follow_up_compras.status_item} = 'N_RECEBIDO' or {Follow_up_compras.status_item} = 'PARCIAL')"
        TextoFiltroEmpresa = "ID_empresa_ped"
    Case "Não aprovado (cotação)":
        TextoFiltroStatus = "status_item = 'CANCELADO'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'CANCELADO'"
        TextoFiltroEmpresa = "ID_empresa_cot"
    Case "Cancelado":
        TextoFiltroStatus = "status_item = 'CANCELADO'"
        TextoFiltroStatusRel = "{Follow_up_compras.status_item} = 'CANCELADO'"
        TextoFiltroEmpresa = "ID_empresa_ped"
End Select

TextoFiltroRemessa = ""
TextoFiltroRemessaRel = ""
If Chk_remessa.Value = 0 Then
    TextoFiltroRemessa = " and Remessa = 'False'"
    TextoFiltroRemessaRel = " and {Follow_up_compras.Remessa} = False"
End If

DataFiltro = ""
DataFiltroRel = ""
If Chk_requisicao.Value = 1 Or Chk_cotacao.Value = 1 Or Chk_pedido.Value = 1 Or Chk_prazo.Value = 1 Then
    If Chk_requisicao.Value = 1 Then
        DataFiltroTexto = "Data_Solicitacao"
    ElseIf Chk_cotacao.Value = 1 Then
            DataFiltroTexto = "dataemissao"
        ElseIf Chk_pedido.Value = 1 Then
                DataFiltroTexto = "data"
            ElseIf Chk_prazo.Value = 1 Then
                    DataFiltroTexto = "prazo"
    End If
    DataFiltro = " and " & DataFiltroTexto & " Between '" & Format(txtinicio.Value, "Short Date") & "' And '" & Format(txtfim.Value, "Short Date") & "'"
    DataFiltroRel = " and {Follow_up_compras." & DataFiltroTexto & "} >= Date(" & Year(txtinicio.Value) & "," & Month(txtinicio.Value) & "," & Day(txtinicio.Value) & ") and {Follow_up_compras." & DataFiltroTexto & "} <= Date(" & Year(txtfim.Value) & "," & Month(txtfim.Value) & "," & Day(txtfim.Value) & ")"
End If

Select Case cmbfiltrarpor
    Case "Código interno": TextoFiltro = "Desenho"
    Case "Descrição": TextoFiltro = "Descricao"
    Case "Família": TextoFiltro = "familia"
    Case "Grupo": TextoFiltro = "Grupo"
    Case "Solicitação": TextoFiltro = "Requisicaotexto"
    Case "Cotação": TextoFiltro = "Cotacaotexto"
    Case "Pedido de compra": TextoFiltro = "Pedido"
    Case "Pedido interno": TextoFiltro = "PI"
    Case "Setor":
        TextoFiltro = "Setorsolic"
        TextoFiltro1 = "setor"
    Case "Requisitante":
        TextoFiltro = "Solicitado"
        TextoFiltro1 = "requisitante"
    Case "Fornecedor":
        TextoFiltro = "Fornecedor"
        If cmbStatus = "Cotando" Then TextoFiltro = "forn"
End Select
CamposFiltro = "Status_Item, Prazo, Prioridade, Desenho, Descricao, Un, Unidade_com, Valor_unit_ped, Requisicaotexto, Cotacaotexto, Pedido, Data, Fornecedor, quant_req, Quant_Comp, IDpedido, IDlista"
CamposFiltro = CamposFiltro & ", prazoreq, Prazo, Autorizado_req, ID_cotacao, ID_requisicao"
If cmbStatus = "Cotando" Or cmbStatus = "" Then CamposFiltro1 = ", Valor_unit_cot, forn, prazoentregaforn, Autorizado_cot" Else CamposFiltro1 = ""
CamposFiltro = CamposFiltro & CamposFiltro1

If cmbStatus = "" Then
    TextoFiltroEmpresa1 = "(ID_empresa_req = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " or ID_empresa_cot = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " or ID_empresa_ped = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ")"
    TextoFiltroEmpresa1Rel = "({Follow_up_compras.ID_empresa_req} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " or {Follow_up_compras.ID_empresa_cot} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " or {Follow_up_compras.ID_empresa_ped} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ")"
    TextoFiltroStatus = ""
    TextoFiltroStatusRel = ""
Else
    TextoFiltroEmpresa1 = TextoFiltroEmpresa & " = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TextoFiltroEmpresa1Rel = "{Follow_up_compras." & TextoFiltroEmpresa & "} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TextoFiltroStatus = " and " & TextoFiltroStatus
    TextoFiltroStatusRel = " and " & TextoFiltroStatusRel
End If
TextoFiltroPadrao = TextoFiltroEmpresa1 & TextoFiltroStatus & DataFiltro & TextoFiltroRemessa & " group by " & CamposFiltro & " order by IDpedido, ID_cotacao, ID_requisicao, IDlista"
TextoFiltroPadraoRel = TextoFiltroEmpresa1Rel & TextoFiltroStatusRel & DataFiltroRel & TextoFiltroRemessaRel & " and {Empresa.Codigo} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Setor" Or cmbfiltrarpor = "Requisitante" Then
        If cmbfiltrarpor = "Setor" Or cmbfiltrarpor = "Requisitante" Then
            Sql_Followup_Compras = "Select " & CamposFiltro & " from Follow_up_compras where (" & TextoFiltro & " = '" & cmbTexto & "' or " & TextoFiltro1 & " = '" & cmbTexto & "') and " & TextoFiltroPadrao
            FormulaRel_Followup_Compras = "({Follow_up_compras." & TextoFiltro & "} = '" & cmbTexto & "' or {Follow_up_compras." & TextoFiltro1 & "} = '" & cmbTexto & "') and " & TextoFiltroPadraoRel
        Else
            Sql_Followup_Compras = "Select " & CamposFiltro & " from Follow_up_compras where " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
            FormulaRel_Followup_Compras = "{Follow_up_compras." & TextoFiltro & "} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
        End If
    Else
        Sql_Followup_Compras = "Select " & CamposFiltro & " from Follow_up_compras where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
        FormulaRel_Followup_Compras = "{Follow_up_compras." & TextoFiltro & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
    End If
Else
    Sql_Followup_Compras = "Select " & CamposFiltro & " from Follow_up_compras where " & TextoFiltroPadrao
    FormulaRel_Followup_Compras = "" & TextoFiltroPadraoRel
End If
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
NomeRel = "Compras_follow up de compras.rpt"
ProcImprimirRel FormulaRel_Followup_Compras, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Followup_Compras.AbsolutePage <> 2 Then
    If TBLISTA_Followup_Compras.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Followup_Compras.PageCount - 1)
    Else
        TBLISTA_Followup_Compras.AbsolutePage = TBLISTA_Followup_Compras.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Followup_Compras.AbsolutePage)
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
    TBLISTA_Followup_Compras.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Followup_Compras.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Followup_Compras.AbsolutePage = 1
ProcExibePagina (TBLISTA_Followup_Compras.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Followup_Compras.AbsolutePage <> -3 Then
    If TBLISTA_Followup_Compras.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Followup_Compras.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Followup_Compras.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Followup_Compras.AbsolutePage = TBLISTA_Followup_Compras.PageCount
ProcExibePagina (TBLISTA_Followup_Compras.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcStatus
    Case vbKeyEscape: ProcSair
    'Case vbKeyF1: ProcAjuda
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 7, True

Formulario = "Outros/Follow up de compras"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Solicitação"
txtinicio.Value = Date
txtfim.Value = Date
With cmbStatus
    .AddItem ""
    .AddItem "Aguardando aprovação"
    .AddItem "Aprovado"
    .AddItem "Cancelado"
    .AddItem "Comprado"
    .AddItem "Cotando"
    .AddItem "Não Aprovado(cotação)"
    .AddItem "Recebido"
    .AddItem "Recebido Parcial"
    .AddItem "Requisitado"
    .AddItem "Comprado e Recebido Parcial"
End With
Cmb_alterar_status = "Comprado"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Outros/Follow up de compras"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

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
                If .ListItems(InitFor).ListSubItems(1) = "REQUISIT." Then
                    'Se o status estiver requisitado, verifica se a solicitação esta liberada
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select ID_Requisicao from Compras_requisicao where Requisicaotexto = '" & .ListItems(InitFor).ListSubItems(6) & "' and Autorizado is null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then GoTo Proximo
                    TBCompras.Close
                ElseIf Cmb_alterar_status = "Comprado" Then
                        If .ListItems(InitFor).ListSubItems(1) = "COMPRADO" Or .ListItems(InitFor).ListSubItems(1) = "CANCELADO" Then GoTo Proximo
                    Else
                        If .ListItems(InitFor).ListSubItems(1) = "RECEBIDO" Or .ListItems(InitFor).ListSubItems(1) = "CANCELADO" Then GoTo Proximo
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
            If .ListItems(InitFor).ListSubItems(1) = "REQUISIT." Then
                    'Se o status estiver requisitado, verifica se a solicitação esta liberada
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select ID_Requisicao from Compras_requisicao where Requisicaotexto = '" & .ListItems(InitFor).ListSubItems(6) & "' and Autorizado is null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then
                        USMsgBox ("Não é permitido alterar o status deste produto/serviço, pois o mesmo não esta com a solicitação liberada."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        TBCompras.Close
                        Exit Sub
                    End If
                    TBCompras.Close
            Else
                Permitido = True
                If Cmb_alterar_status = "Comprado" Then
                    If .ListItems(InitFor).ListSubItems(1) = "COMPRADO" Or .ListItems(InitFor).ListSubItems(1) = "CANCELADO" Then
                        Permitido = False
                        If .ListItems(InitFor).ListSubItems(1) = "COMPRADO" Then StatusTexto = "comprado" Else StatusTexto = "cancelado"
                    End If
                Else
                    If .ListItems(InitFor).ListSubItems(1) = "RECEBIDO" Or .ListItems(InitFor).ListSubItems(1) = "CANCELADO" Then
                        Permitido = False
                        If .ListItems(InitFor).ListSubItems(1) = "RECEBIDO" Then StatusTexto = "recebido" Else StatusTexto = "cancelado"
                    End If
                End If
                If Permitido = False Then
                    USMsgBox ("Não é permitido alterar o status deste produto/serviço, pois o mesmo está com o status " & StatusTexto & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtfim_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
IDlista = 0
Set TBLISTA_Followup_Compras = CreateObject("adodb.recordset")
'Debug.print Sql_Followup_Compras
TBLISTA_Followup_Compras.Open Sql_Followup_Compras, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Followup_Compras.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Dataini = 0
Lista.ListItems.Clear
TBLISTA_Followup_Compras.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Followup_Compras.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Followup_Compras.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Followup_Compras.RecordCount - IIf(Pagina > 1, (TBLISTA_Followup_Compras.PageSize * (Pagina - 1)), 0), TBLISTA_Followup_Compras.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Followup_Compras.EOF = False And (ContadorReg <= TamanhoPagina)
    Cor = vbBlack
    With Lista.ListItems.Add(, , TBLISTA_Followup_Compras!IDlista)
        If TBLISTA_Followup_Compras!Status_Item = "N_RECEBIDO" Then
            If IsNull(TBLISTA_Followup_Compras!Prazo) = False Then
                Dataini = Format(TBLISTA_Followup_Compras!Prazo, "dd/mm/yy")
                If Date = (Dataini - 1) Or Date = Dataini Then  '1 dia antes do prazo ou no dia
                    Cor = &H80FF&
                ElseIf Date > Dataini Then 'Atraso
                        Cor = vbRed
                End If
            End If
            .SubItems(1) = "COMPRADO"
        Else
            .SubItems(1) = TBLISTA_Followup_Compras!Status_Item
        End If
        .SubItems(2) = IIf(IsNull(TBLISTA_Followup_Compras!Prioridade), "", TBLISTA_Followup_Compras!Prioridade)
        .SubItems(3) = IIf(IsNull(TBLISTA_Followup_Compras!Desenho), "", TBLISTA_Followup_Compras!Desenho)
        .SubItems(4) = IIf(IsNull(TBLISTA_Followup_Compras!Descricao), "", TBLISTA_Followup_Compras!Descricao)
        .SubItems(5) = IIf(IsNull(TBLISTA_Followup_Compras!Un), "", TBLISTA_Followup_Compras!Un)
        .SubItems(6) = IIf(IsNull(TBLISTA_Followup_Compras!Unidade_com), "", TBLISTA_Followup_Compras!Unidade_com)
        
        If TBLISTA_Followup_Compras!Status_Item = "COTANDO" Then
            .SubItems(7) = IIf(IsNull(TBLISTA_Followup_Compras!Valor_unit_cot), "", Format(TBLISTA_Followup_Compras!Valor_unit_cot, "###,##0.00"))
        Else
            .SubItems(7) = IIf(IsNull(TBLISTA_Followup_Compras!Valor_unit_ped), "", Format(TBLISTA_Followup_Compras!Valor_unit_ped, "###,##0.00"))
        End If
        
        .SubItems(8) = IIf(IsNull(TBLISTA_Followup_Compras!Requisicaotexto), "", TBLISTA_Followup_Compras!Requisicaotexto)
        .SubItems(9) = IIf(IsNull(TBLISTA_Followup_Compras!Cotacaotexto), "", TBLISTA_Followup_Compras!Cotacaotexto)
        .SubItems(10) = IIf(IsNull(TBLISTA_Followup_Compras!Pedido), "", TBLISTA_Followup_Compras!Pedido)
        .SubItems(11) = IIf(IsNull(TBLISTA_Followup_Compras!Data), "", Format(TBLISTA_Followup_Compras!Data, "dd/mm/yy"))
        If TBLISTA_Followup_Compras!Status_Item = "COTANDO" Then
            .SubItems(12) = IIf(IsNull(TBLISTA_Followup_Compras!forn), "", TBLISTA_Followup_Compras!forn)
        Else
            .SubItems(12) = IIf(IsNull(TBLISTA_Followup_Compras!Fornecedor), "", TBLISTA_Followup_Compras!Fornecedor)
        End If
        .SubItems(13) = IIf(IsNull(TBLISTA_Followup_Compras!quant_req), "", Format(TBLISTA_Followup_Compras!quant_req, "0.000"))
        .SubItems(14) = IIf(IsNull(TBLISTA_Followup_Compras!Quant_Comp), "", Format(TBLISTA_Followup_Compras!Quant_Comp, "0.000"))
        
        'Verifica qtde. recebida
        Qtde = 0
        If TBLISTA_Followup_Compras!Status_Item = "RECEBIDO" Or TBLISTA_Followup_Compras!Status_Item = "PARCIAL" Then
            If TBLISTA_Followup_Compras!Status_Item = "RECEBIDO" Then Cor = vbBlue Else Cor = &HC000&
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Recebido) as Qtde from Estoque_controle_recebimento where idpedido = " & IIf(IsNull(TBLISTA_Followup_Compras!IDpedido), 0, TBLISTA_Followup_Compras!IDpedido) & " and IDlista = " & TBLISTA_Followup_Compras!IDlista & " and Desenho = '" & TBLISTA_Followup_Compras!Desenho & "' and Programacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
            End If
            TBAbrir.Close
        End If
        .SubItems(15) = Format(Qtde, "###,##0.0000")
        
        If TBLISTA_Followup_Compras!Status_Item = "REQUISIT." Then
            .SubItems(16) = IIf(IsNull(TBLISTA_Followup_Compras!prazoreq), "", Format(TBLISTA_Followup_Compras!prazoreq, "dd/mm/yy"))
        ElseIf TBLISTA_Followup_Compras!Status_Item = "COTANDO" Then
                .SubItems(16) = IIf(IsNull(TBLISTA_Followup_Compras!prazoentregaforn), "", Format(TBLISTA_Followup_Compras!prazoentregaforn, "dd/mm/yy"))
            Else
                .SubItems(16) = IIf(IsNull(TBLISTA_Followup_Compras!Prazo), "", Format(TBLISTA_Followup_Compras!Prazo, "dd/mm/yy"))
        End If
        
        If TBLISTA_Followup_Compras!Status_Item = "REQUISIT." Then
            'Status solicitação
            If IsNull(TBLISTA_Followup_Compras!Autorizado_req) = False Then .SubItems(17) = "APROVADO" Else .SubItems(17) = "NÃO APROVADO"
        ElseIf TBLISTA_Followup_Compras!Status_Item = "COTANDO" Then
                'Status cotação
                If IsNull(TBLISTA_Followup_Compras!Autorizado_cot) = False Then .SubItems(17) = "APROVADO" Else .SubItems(17) = "NÃO APROVADO"
        End If
        
        .ForeColor = Cor
        .ListSubItems(1).ForeColor = Cor
        .ListSubItems(2).ForeColor = Cor
        .ListSubItems(3).ForeColor = Cor
        .ListSubItems(4).ForeColor = Cor
        .ListSubItems(5).ForeColor = Cor
        .ListSubItems(6).ForeColor = Cor
        .ListSubItems(7).ForeColor = Cor
        .ListSubItems(8).ForeColor = Cor
        .ListSubItems(9).ForeColor = Cor
        .ListSubItems(10).ForeColor = Cor
        .ListSubItems(11).ForeColor = Cor
        .ListSubItems(12).ForeColor = Cor
        .ListSubItems(13).ForeColor = Cor
        .ListSubItems(14).ForeColor = Cor
        .ListSubItems(15).ForeColor = Cor
        .ListSubItems(16).ForeColor = Cor
        If TBLISTA_Followup_Compras!Status_Item = "REQUISIT." Or TBLISTA_Followup_Compras!Status_Item = "COTANDO" Then .ListSubItems(17).ForeColor = Cor
    End With
    TBLISTA_Followup_Compras.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Followup_Compras.RecordCount
If TBLISTA_Followup_Compras.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Followup_Compras.PageCount
ElseIf TBLISTA_Followup_Compras.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Followup_Compras.PageCount & " de: " & TBLISTA_Followup_Compras.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Followup_Compras.AbsolutePage - 1 & " de: " & TBLISTA_Followup_Compras.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtinicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

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

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 3: ProcStatus
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
