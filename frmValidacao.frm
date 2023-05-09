VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmValidacao 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Outros - Validação de procedimentos"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1360
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12990
      TabIndex        =   52
      Top             =   1950
      Width           =   2265
      Begin VB.ComboBox Cmb_filtrar 
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
         ItemData        =   "frmValidacao.frx":0000
         Left            =   180
         List            =   "frmValidacao.frx":000D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         ToolTipText     =   "Filtrar."
         Top             =   210
         Width           =   1905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sem versão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10590
      TabIndex        =   51
      Top             =   1950
      Width           =   2385
      Begin VB.CheckBox Chk_estrutura_versao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estrutura"
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
         TabIndex        =   32
         Top             =   300
         Width           =   975
      End
      Begin VB.CheckBox Chk_processo_versao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Processo"
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
         Left            =   1270
         TabIndex        =   33
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.Frame Frame_versao 
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
      Height          =   615
      Left            =   10410
      TabIndex        =   48
      Top             =   7620
      Width           =   4845
      Begin VB.ComboBox Cmb_versao_processo 
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
         ItemData        =   "frmValidacao.frx":003A
         Left            =   3270
         List            =   "frmValidacao.frx":003C
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Versão do processo."
         Top             =   210
         Width           =   915
      End
      Begin VB.ComboBox Cmb_versao_estrutura 
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
         ItemData        =   "frmValidacao.frx":003E
         Left            =   990
         List            =   "frmValidacao.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Versão da estrutura."
         Top             =   180
         Width           =   915
      End
      Begin DrawSuite2022.USButton Cmd_salvar_versao 
         Height          =   405
         Left            =   4200
         TabIndex        =   22
         ToolTipText     =   "Salvar (F3)"
         Top             =   150
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   714
         DibPicture      =   "frmValidacao.frx":0042
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton Cmd_visualizar_desc_versao 
         Height          =   405
         Left            =   1920
         TabIndex        =   20
         ToolTipText     =   "Visualizar descrição da versão da estrutura."
         Top             =   150
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   714
         DibPicture      =   "frmValidacao.frx":07D4
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Estrutura:"
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
         Left            =   210
         TabIndex        =   50
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Processo:"
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
         Left            =   2490
         TabIndex        =   49
         Top             =   210
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Não validado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   47
      Top             =   1950
      Width           =   10515
      Begin VB.CheckBox Chk_nota 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nota fiscal"
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
         Left            =   7470
         TabIndex        =   31
         Top             =   300
         Width           =   1065
      End
      Begin VB.CheckBox Chk_ordem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordem"
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
         Left            =   6470
         TabIndex        =   30
         Top             =   300
         Width           =   795
      End
      Begin VB.CheckBox Chk_compra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compra"
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
         Left            =   5410
         TabIndex        =   29
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox Chk_plano 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Plano de inspeção"
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
         Left            =   3600
         TabIndex        =   28
         Top             =   300
         Width           =   1605
      End
      Begin VB.CheckBox Chk_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Processo"
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
         Left            =   2450
         TabIndex        =   27
         Top             =   300
         Width           =   945
      End
      Begin VB.CheckBox Chk_estrutura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estrutura"
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
         Left            =   1270
         TabIndex        =   26
         Top             =   300
         Width           =   975
      End
      Begin VB.CheckBox Chk_produto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto"
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
         TabIndex        =   25
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.CheckBox optperiodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo final"
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
      Left            =   13470
      TabIndex        =   8
      Top             =   1020
      Width           =   1065
   End
   Begin VB.CheckBox Chk_data_venda 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. venda"
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
      Left            =   12330
      TabIndex        =   7
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   915
      Left            =   12270
      TabIndex        =   44
      Top             =   1020
      Width           =   2985
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   1500
         TabIndex        =   10
         ToolTipText     =   "Data final para pesquisa."
         Top             =   450
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
         Format          =   479920129
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         ToolTipText     =   "Data início para pesquisa."
         Top             =   450
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
         Format          =   479920131
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "De"
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
         Left            =   750
         TabIndex        =   46
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2025
         TabIndex        =   45
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   60
      TabIndex        =   41
      Top             =   1020
      Width           =   12195
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   2490
         TabIndex        =   54
         Top             =   270
         WhatsThisHelpID =   210
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
            TabIndex        =   3
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
            TabIndex        =   1
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
            TabIndex        =   2
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
            TabIndex        =   4
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
         ItemData        =   "frmValidacao.frx":2988
         Left            =   180
         List            =   "frmValidacao.frx":29A4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   450
         Width           =   2235
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
         Left            =   7380
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   450
         Width           =   4635
      End
      Begin VB.ComboBox cmbfamilia 
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
         Left            =   7380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   450
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         TabIndex        =   43
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   877
         TabIndex        =   42
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   37
      Top             =   7620
      Width           =   10335
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
         Left            =   5130
         TabIndex        =   13
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Centralizar
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
         TabIndex        =   12
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   7350
         TabIndex        =   17
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmValidacao.frx":2A27
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   6810
         TabIndex        =   16
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmValidacao.frx":61CB
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   5700
         TabIndex        =   14
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   6270
         TabIndex        =   15
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmValidacao.frx":9CD4
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   7890
         TabIndex        =   18
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmValidacao.frx":DDC3
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   3420
         TabIndex        =   53
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   8640
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         TabIndex        =   39
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   2100
         TabIndex        =   38
         Top             =   240
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5025
      Left            =   60
      TabIndex        =   11
      Top             =   2580
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   8864
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      NumItems        =   30
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Ped. int./SPR"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cod. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Dimensões"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Observações do produto"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Observações da venda"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Inspeção"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Embalagem"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Gravação"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "T"
         Text            =   "Novo projeto"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "Prioridade"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Object.Tag             =   "D"
         Text            =   "Dt. venda"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Object.Tag             =   "D"
         Text            =   "Prazo final"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Tag             =   "T"
         Text            =   "Validação do pedido interno"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Tag             =   "T"
         Text            =   "Validação do produto"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Tag             =   "T"
         Text            =   "Validação da estrutura"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Tag             =   "T"
         Text            =   "Validação do processo"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Tag             =   "T"
         Text            =   "Validação do plano de inspeção"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Tag             =   "T"
         Text            =   "Validação da compra"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Tag             =   "T"
         Text            =   "Validação da ordem"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Object.Tag             =   "T"
         Text            =   "Data da inspeção final"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Object.Tag             =   "T"
         Text            =   "Data de entrada no estoque"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Object.Tag             =   "T"
         Text            =   "Validação da nota fiscal"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Object.Tag             =   "T"
         Text            =   "Data da expedição"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   28
         Object.Tag             =   "T"
         Text            =   "Versão estrut."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   29
         Object.Tag             =   "T"
         Text            =   "Versão proc."
         Object.Width           =   1940
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   35
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
      TabIndex        =   36
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
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
      ButtonLeft3     =   93
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
      ButtonLeft4     =   97
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
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
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
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13050
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmValidacao.frx":1164F
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista_arquivos 
      Height          =   1500
      Left            =   60
      TabIndex        =   23
      Top             =   8250
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   2646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "T"
         Text            =   "Caminho"
         Object.Width           =   24051
      EndProperty
   End
   Begin DrawSuite2022.USButton Cmd_visualizar 
      Height          =   1500
      Left            =   14145
      TabIndex        =   24
      ToolTipText     =   "Visualizar arquivo."
      Top             =   8250
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   2646
      DibPicture      =   "frmValidacao.frx":14443
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      Caption         =   "Visualizar arquivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   16777215
      GradientColor3  =   16777215
      GradientColorDown2=   16246986
      GradientColorDown3=   15189380
      GradientColorDown4=   14596208
      GradientColorOver1=   16643560
      GradientColorOver2=   16576988
      GradientColorOver3=   16441780
      GradientColorOver4=   16178091
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
End
Attribute VB_Name = "frmValidacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Validacao As String 'OK
Dim FormulaRel_Validacao As String 'OK
Dim TBLISTA_Validacao As ADODB.Recordset 'OK

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "Validacao_procedimentos.rpt"
ProcImprimirRel FormulaRel_Validacao, ""

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

Private Sub Chk_compra_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_data_venda_Click()
On Error GoTo tratar_erro

ProcLimparListas
If Chk_data_venda.Value = 1 Then
    optPeriodo.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltFim.Value = Date
    msk_fltInicio.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_estrutura_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_estrutura_versao_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_nota_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_plano_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_processo_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_processo_versao_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_produto_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcLimparListas
If cmbfiltrarpor.Text = "Família" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    ProcCarregaComboFamilia cmbfamilia, "Vendas = 'True'", True
Else
    cmbfamilia.Visible = False
    txtTexto.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_versao_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista.ListItems.Count = 0 Then Exit Sub
Conexao.Execute "Update Vendas_Carteira Set Versao_estrutura = '" & IIf(Cmb_versao_estrutura = "", Null, Cmb_versao_estrutura) & "', Versao_processo = '" & IIf(Cmb_versao_processo = "", Null, Cmb_versao_processo) & "' where Codigo = " & Lista.SelectedItem

Conexao.Execute "DELETE from Producaomaterial where ID_carteira = " & Lista.SelectedItem
If Cmb_versao_estrutura <> "" Then
    'Verifica qtde empenhada
    QuantSolicitado = 0
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select VC.Qtde_produzir - ISNULL(ECEV.Qtde_empenhada, 0) AS Quantsolicitado from Vendas_carteira VC LEFT JOIN Estoque_Controle_Empenho_Vendas ECEV ON ECEV.ID_carteira = VC.Codigo where VC.Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        ProcAcertaRequisicao Lista.SelectedItem.ListSubItems(3), 0, Lista.SelectedItem, Cmb_versao_estrutura, IIf(IsNull(TBLISTA!QuantSolicitado), 0, TBLISTA!QuantSolicitado), True
    End If
    TBLISTA.Close
End If
USMsgBox ("Versões salvas com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Outros/Validação de procedimentos"
Acao = "Salvar versões"
ID_documento = Lista.SelectedItem
Documento = "Pedido int.: " & Lista.SelectedItem.ListSubItems.Item(1) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems.Item(3)
Documento1 = ""
ProcGravaEvento
'==================================
ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_Click()
On Error GoTo tratar_erro

If Lista_arquivos.ListItems.Count = 0 Then Exit Sub
ProcAbrirArquivo Lista_arquivos.SelectedItem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_desc_versao_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If Cmb_versao_estrutura = "" Then
    USMsgBox ("Informe a versão da estrutura antes de visualizar a descrição."), vbExclamation, "CAPRIND v5.0"
    Cmb_versao_estrutura.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select PCDV.Descricao from Projproduto P INNER JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto where P.Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and PCDV.Versao = '" & Cmb_versao_estrutura & "' and PCDV.Descricao IS NOT NULL and PCDV.Descricao <> N''", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Call USMsgBox(TBProduto!Descricao, vbOKOnly, "Descrição da versão da estrutura")
Else
    USMsgBox ("Não existe descrição cadastrada para esta versão da estrutura."), vbInformation, "CAPRIND v5.0"
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Validacao.AbsolutePage <> 2 Then
    If TBLISTA_Validacao.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Validacao.PageCount - 1)
    Else
        TBLISTA_Validacao.AbsolutePage = TBLISTA_Validacao.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Validacao.AbsolutePage)
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
    TBLISTA_Validacao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Validacao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Validacao.AbsolutePage = 1
ProcExibePagina (TBLISTA_Validacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Validacao.AbsolutePage <> -3 Then
    If TBLISTA_Validacao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Validacao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Validacao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Validacao.AbsolutePage = TBLISTA_Validacao.PageCount
ProcExibePagina (TBLISTA_Validacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: If Frame_versao.Enabled = True Then Cmd_salvar_versao_Click
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 6, True

Formulario = "Outros/Validação de procedimentos"
Direitos
ProcLimpaVariaveisPrincipais
ProcVerifColunas
cmbfiltrarpor = "Código interno"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
Cmb_filtrar = "Com necessidade"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Outros/Validação de procedimentos"
Direitos
ProcLimpaVariaveisPrincipais
ProcVerifColunas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifColunas()
On Error GoTo tratar_erro

ProcCorrigeColunasForm Lista, Formulario, 29, False, 0, 1200, 550, 1200, 550, 3600, 1000, 2000, 2500, 2500, 2500, 2500, 2500, 1300, 1200, 1000, 1000, 2800, 2800, 2800, 2800, 2800, 2800, 2800, 2800, 2800, 2800, 2800, 1300, 1100, 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_arquivos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_arquivos, ColumnHeader

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

Private Sub ProcAtualizalista(Pagina As Long)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ProcLimparListas
If StrSql_Validacao = "" Then Exit Sub
Set TBLISTA_Validacao = CreateObject("adodb.recordset")
TBLISTA_Validacao.Open StrSql_Validacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Validacao.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaValidacao(Campos As String, Tabela As String, Condicao As String)
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "SELECT " & Campos & " from " & Tabela & " where " & Condicao & "", Conexao, adOpenKeyset, adLockOptimistic

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
DataFiltro = ""
DataFiltroRel = ""
If Chk_data_venda.Value = 1 Or optPeriodo.Value = 1 Then
    If Chk_data_venda.Value = 1 Then DataTexto = "VC.Datavendas" Else DataTexto = "VC.prazofinal"
    DataFiltro = " and " & DataTexto & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    DataFiltroRel = " and {" & Replace(DataTexto, "VC.", "Vendas_carteira.") & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & Replace(DataTexto, "VC.", "Vendas_carteira.") & "} <= Date(" & Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
End If
CamposFiltro = "VP.ordenarproposta, VP.cotacao, VP.Ncotacao, VP.Revisao, OSPCP.Requisicaotexto, VP.DtValidacaoPI, VP.RespValidacaoPI, VC.Codigo, VC.Desenho, VC.Rev_codinterno, VC.Descricao, VC.Quantidade, VC.Espessura, VC.Largura, VC.Comprimento, P.observacoes AS Obs_produto, VC.observacoes, VC.Inspecao, VC.Embalagem, VC.Gravacao, VC.Novo_projeto, VC.Prioridade, VC.Datavendas, VC.prazofinal, VC.Versao_estrutura, VC.Versao_processo"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (((((((((((vendas_carteira VC INNER JOIN Vendas_Proposta VP ON VC.Cotacao = VP.Cotacao) LEFT JOIN Outros_SolicitacaoPCP OSPCP ON OSPCP.ID = VC.ID_Solicitacao) LEFT JOIN Projproduto P ON P.Desenho = VC.Desenho) LEFT JOIN Processos PROCE ON PROCE.Codproduto = P.Codproduto) LEFT JOIN Plano PL ON PL.Desenho = VC.Desenho) LEFT JOIN Compras_pedido_lista_empenhos CPLE ON CPLE.IDcarteira = VC.Codigo) LEFT JOIN Compras_pedido_lista CPL ON CPL.IDlista = CPLE.IDlista) LEFT JOIN Compras_pedido CPE ON CPE.IDpedido = CPL.IDPedido) LEFT JOIN Producao PROD RIGHT OUTER JOIN Producao_pedidos PP ON PROD.Ordem = PP.Ordem ON PROD.desenho = VC.Desenho AND PP.IDcarteira = VC.Codigo) LEFT JOIN tbl_Detalhes_Nota_pedidos DNP ON DNP.ID_carteira = VC.Codigo and DNP.Codinterno = VC.Desenho) LEFT JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = DNP.ID_nota) LEFT JOIN carteira_producao CP ON CP.Codigo = VC.Codigo"

TextoFiltroValid = ""
TextoFiltroValidRel = ""
If Chk_produto.Value = 1 Then
    TextoFiltroValid = "and P.DtValidacao IS NULL"
    TextoFiltroValidRel = "and ISNULL({Projproduto.DtValidacao}) = True"
End If
If Chk_estrutura.Value = 1 Then
    If TextoFiltroValid = "" Then
        TextoFiltroValid = "and P.DtValidacaoConj IS NULL"
        TextoFiltroValidRel = "and ISNULL({Projproduto.DtValidacaoConj}) = True"
    Else
        TextoFiltroValid = TextoFiltroValid & " and P.DtValidacaoConj IS NULL"
        TextoFiltroValidRel = TextoFiltroValidRel & " and ISNULL({Projproduto.DtValidacaoConj}) = True"
    End If
End If
If Chk_processo.Value = 1 Then
    If TextoFiltroValid = "" Then
        TextoFiltroValid = "and PROCE.DtValidacao IS NULL"
        TextoFiltroValidRel = "and ISNULL({Processos.DtValidacao}) = True"
    Else
        TextoFiltroValid = TextoFiltroValid & " and PROCE.DtValidacao IS NULL"
        TextoFiltroValidRel = TextoFiltroValidRel & " and ISNULL({Processos.DtValidacao}) = True"
    End If
End If
If Chk_plano.Value = 1 Then
    If TextoFiltroValid = "" Then
        TextoFiltroValid = "and P.DtValidacaoPlano IS NULL"
        TextoFiltroValidRel = "and ISNULL({Projproduto.DtValidacaoPlano}) = True"
    Else
        TextoFiltroValid = TextoFiltroValid & " and P.DtValidacaoPlano IS NULL"
        TextoFiltroValidRel = TextoFiltroValidRel & " and ISNULL({Projproduto.DtValidacaoPlano}) = True"
    End If
End If
If Chk_compra.Value = 1 Then
    If TextoFiltroValid = "" Then
        TextoFiltroValid = "and CPE.DtValidacao IS NULL"
        TextoFiltroValidRel = "and ISNULL({Compras_pedido.DtValidacao}) = True"
    Else
        TextoFiltroValid = TextoFiltroValid & " and CPE.DtValidacao IS NULL"
        TextoFiltroValidRel = TextoFiltroValidRel & " and ISNULL({Compras_pedido.DtValidacao}) = True"
    End If
End If
If Chk_ordem.Value = 1 Then
    If TextoFiltroValid = "" Then
        TextoFiltroValid = "and PROD.DtValidacao IS NULL"
        TextoFiltroValidRel = "and ISNULL({Producao.DtValidacao}) = True"
    Else
        TextoFiltroValid = TextoFiltroValid & " and PROD.DtValidacao IS NULL"
        TextoFiltroValidRel = TextoFiltroValidRel & " and ISNULL({Producao.DtValidacao}) = True"
    End If
End If
If Chk_nota.Value = 1 Then
    If TextoFiltroValid = "" Then
        TextoFiltroValid = "and NF.DtValidacao IS NULL"
        TextoFiltroValidRel = "and ISNULL({tbl_Dados_Nota_Fiscal.DtValidacao}) = True"
    Else
        TextoFiltroValid = TextoFiltroValid & " and NF.DtValidacao IS NULL"
        TextoFiltroValidRel = TextoFiltroValidRel & " and ISNULL({tbl_Dados_Nota_Fiscal.DtValidacao}) = True"
    End If
End If

TextoFiltroVersao = ""
TextoFiltroVersaoRel = ""
If Chk_estrutura.Value = 1 Then
    TextoFiltroVersao = " and VC.Versao_estrutura IS NULL"
    TextoFiltroVersaoRel = " and ISNULL({Vendas_carteira.Versao_estrutura}) = True"
End If
If Chk_processo.Value = 1 Then
    If TextoFiltroVersao = "" Then
        TextoFiltroVersao = " and VC.Versao_processo IS NULL"
        TextoFiltroVersaoRel = " and ISNULL({Vendas_carteira.Versao_processo}) = True"
    Else
        TextoFiltroVersao = TextoFiltroVersao & " and VC.Versao_processo IS NULL"
        TextoFiltroVersaoRel = TextoFiltroVersaoRel & " and ISNULL({Vendas_carteira.Versao_processo}) = True"
    End If
End If

FiltroMRP = ""
FiltroMRPRel = ""
Select Case Cmb_filtrar
    Case "Com necessidade":
        FiltroMRP = " and CP.Necessidade > 0"
        FiltroMRPRel = " and {carteira_producao.Necessidade} > 0"
    Case "Sem necessidade":
        FiltroMRP = " and CP.Necessidade <= 0"
        FiltroMRPRel = " and {carteira_producao.Necessidade} <= 0"
End Select

TexotFiltroPadrao = "VP.DtValidacaoPI IS NOT NULL " & TextoFiltroValid & TextoFiltroVersao & DataFiltro & FiltroMRP & " group by VP.ordenarproposta, VP.cotacao, VP.Ncotacao, VP.Revisao, OSPCP.Requisicaotexto, VP.DtValidacaoPI, VP.RespValidacaoPI, VC.Codigo, VC.Desenho, VC.Rev_codinterno, VC.Descricao, VC.Quantidade, VC.Espessura, VC.Largura, VC.Comprimento, P.observacoes, VC.observacoes, VC.Inspecao, VC.Embalagem, VC.Gravacao, VC.Novo_projeto, VC.Prioridade, VC.Datavendas, VC.prazofinal, VC.Versao_estrutura, VC.Versao_processo order by VP.ordenarproposta desc, VP.cotacao desc"
TexotFiltroPadraoRel = "NOT(ISNULL({vendas_proposta.DtValidacaoPI})) " & TextoFiltroValidRel & TextoFiltroVersaoRel & DataFiltroRel & FiltroMRPRel

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor.Text = "Família" Then
        StrSql_Validacao = INNERJOINTEXTO & " where VC.Familia = '" & cmbfamilia & "' and " & TexotFiltroPadrao
        FormulaRel_Validacao = "{Vendas_carteira.Familia} = '" & cmbfamilia & "' and " & TexotFiltroPadraoRel
    Else
        Select Case cmbfiltrarpor
            Case "Cliente": TextoFiltro = "VP.Cliente"
            Case "Código de referência": TextoFiltro = "VC.n_referencia"
            Case "Código interno": TextoFiltro = "VC.Desenho"
            Case "Descrição": TextoFiltro = "VC.Descricao_tecnica"
            Case "Pedido do cliente": TextoFiltro = "VC.PCcliente"
            Case "Pedido interno": TextoFiltro = "VP.Ncotacao"
            Case "Solicitação de produção": TextoFiltro = "OSPCP.Requisicaotexto"
        End Select
        StrSql_Validacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TexotFiltroPadrao
        FormulaRel_Validacao = "{" & IIf(Left(TextoFiltro, 3) = "VP.", Replace(TextoFiltro, "VP.", "Vendas_proposta."), Replace(TextoFiltro, "VC.", "Vendas_carteira.")) & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TexotFiltroPadraoRel
    End If
Else
    StrSql_Validacao = INNERJOINTEXTO & " where " & TexotFiltroPadrao
    FormulaRel_Validacao = TexotFiltroPadraoRel
End If
ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ProcLimparListas
TBLISTA_Validacao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Validacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Validacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Validacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Validacao.PageSize * (Pagina - 1)), 0), TBLISTA_Validacao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Validacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Validacao!CODIGO
        If IsNull(TBLISTA_Validacao!Cotacao) = True Or TBLISTA_Validacao!Cotacao = 0 Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Validacao!Requisicaotexto), "", TBLISTA_Validacao!Requisicaotexto)
            .Item(.Count).SubItems(2) = ""
        Else
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Validacao!Ncotacao), "", TBLISTA_Validacao!Ncotacao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Validacao!Revisao), "", TBLISTA_Validacao!Revisao)
        End If
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Validacao!Desenho), "", TBLISTA_Validacao!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Validacao!Rev_codinterno), "", TBLISTA_Validacao!Rev_codinterno)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Validacao!Descricao), "", TBLISTA_Validacao!Descricao)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Validacao!quantidade), "", Format(TBLISTA_Validacao!quantidade, "###,##0.0000"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Validacao!Espessura), "", Format(TBLISTA_Validacao!Espessura, "###,##0.00")) & "X" & IIf(IsNull(TBLISTA_Validacao!Largura), "", Format(TBLISTA_Validacao!Largura, "###,##0.00")) & "X" & IIf(IsNull(TBLISTA_Validacao!Comprimento), "", Format(TBLISTA_Validacao!Comprimento, "###,##0.00"))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Validacao!Obs_produto), "", TBLISTA_Validacao!Obs_produto)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Validacao!observacoes), "", TBLISTA_Validacao!observacoes)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Validacao!Inspecao), "", TBLISTA_Validacao!Inspecao)
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Validacao!Embalagem), "", TBLISTA_Validacao!Embalagem)
        .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA_Validacao!Gravacao), "", TBLISTA_Validacao!Gravacao)
        .Item(.Count).SubItems(13) = IIf(TBLISTA_Validacao!Novo_projeto = True, "Sim", "Não")
        .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_Validacao!Prioridade), "", TBLISTA_Validacao!Prioridade)
        .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA_Validacao!Datavendas), "", Format(TBLISTA_Validacao!Datavendas, "dd/mm/yy"))
        .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA_Validacao!PrazoFinal), "", Format(TBLISTA_Validacao!PrazoFinal, "dd/mm/yy"))
        .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA_Validacao!DtValidacaoPI), "", TBLISTA_Validacao!DtValidacaoPI & " - " & TBLISTA_Validacao!RespValidacaoPI)
        
        'PRODUTO, ESTRUTURA E PLANO DE INSPEÇÃO
        ProcVerificaValidacao "DtValidacao, RespValidacao, DtValidacaoConj, RespValidacaoConj, DtValidacaoPlano, RespValidacaoPlano", "projproduto", "desenho = '" & TBLISTA_Validacao!Desenho & "'"
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(18) = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao & " - " & TBProduto!RespValidacao)
            .Item(.Count).SubItems(19) = IIf(IsNull(TBProduto!DtValidacaoConj), "", TBProduto!DtValidacaoConj & " - " & TBProduto!RespValidacaoConj)
            .Item(.Count).SubItems(21) = IIf(IsNull(TBProduto!DtValidacaoPlano), "", TBProduto!DtValidacaoPlano & " - " & TBProduto!RespValidacaoPlano)
        End If
        TBProduto.Close
        
        'PROCESSOS
        ProcVerificaValidacao "PRO.DTValidacao, PRO.RespValidacao", "projproduto PP INNER JOIN Processos PRO ON PP.CodProduto = PRO.CodProduto", "PP.desenho = '" & TBLISTA_Validacao!Desenho & "'"
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(20) = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao & " - " & TBProduto!RespValidacao)
        End If
        TBProduto.Close
        
        'COMPRAS
        ProcVerificaValidacao "CP.DTValidacao, CP.RespValidacao", "(Compras_pedido_lista_empenhos CPLE INNER JOIN Compras_pedido_lista CPL ON CPL.IDlista = CPLE.Idlista) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDpedido", "CPLE.IDCarteira = " & TBLISTA_Validacao!CODIGO
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(22) = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao & " - " & TBProduto!RespValidacao)
        End If
        TBProduto.Close
        
        'ORDEM DE PRODUÇÃO
        ProcVerificaValidacao "P.DTValidacao, P.RespValidacao", "Producao_Pedidos PP INNER JOIN Producao P ON PP.Ordem = P.Ordem", "PP.IDCarteira = " & TBLISTA_Validacao!CODIGO & " and P.Desenho = '" & TBLISTA_Validacao!Desenho & "'"
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(23) = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao & " - " & TBProduto!RespValidacao)
        End If
        TBProduto.Close
        
        'INSPEÇÃO FINAL
        ProcVerificaValidacao "M.Data, M.Inspetor", "(Producao_Pedidos PP INNER JOIN Producao P ON PP.Ordem = P.Ordem) INNER JOIN Medicao M ON M.Peca = P.Ordem", "PP.IDCarteira = " & TBLISTA_Validacao!CODIGO & " and P.Desenho = '" & TBLISTA_Validacao!Desenho & "'"
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(24) = IIf(IsNull(TBProduto!data), "", TBProduto!data & " - " & TBProduto!Inspetor)
        End If
        TBProduto.Close
        
        'ESTOQUE
        ProcVerificaValidacao "EC.Data, EC.Responsavel", "((Producao_Pedidos PP INNER JOIN Producao P ON PP.Ordem = P.Ordem) INNER JOIN Ordens_texto_SA OTSA ON OTSA.Ordem = P.Ordem) INNER JOIN Estoque_Controle EC ON OTSA.Ordem = EC.Lote", "PP.IDCarteira = " & TBLISTA_Validacao!CODIGO & " and P.Desenho = '" & TBLISTA_Validacao!Desenho & "' and EC.IDestoque = (Select MAX(EC1.IDestoque) from Estoque_Controle EC1 where EC1.IDestoque = EC.IDestoque)"
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(25) = IIf(IsNull(TBProduto!data), "", TBProduto!data & " - " & TBProduto!Responsavel)
        End If
        TBProduto.Close
        
        'FATURAMENTO
        ProcVerificaValidacao "NF.dtValidacao, NF.RespValidacao", "tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Detalhes_Nota_pedidos DNP ON DNP.ID_Nota = NF.ID", "DNP.ID_Carteira = " & TBLISTA_Validacao!CODIGO & " and DNP.CodInterno = '" & TBLISTA_Validacao!Desenho & "' and NF.ID = (Select MAX(NF1.ID) from tbl_Dados_Nota_Fiscal NF1 where NF1.ID = NF.ID)"
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(26) = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao & " - " & TBProduto!RespValidacao)
        End If
        TBProduto.Close
        
        'EXPEDIÇÃO
        ProcVerificaValidacao "EM.Data, EM.Responsavel", "tbl_Detalhes_Nota_pedidos DNP INNER JOIN Estoque_movimentacao EM ON EM.ID_prod_NF = DNP.ID_prod_NF", "DNP.ID_Carteira = " & TBLISTA_Validacao!CODIGO & " and DNP.CodInterno = '" & TBLISTA_Validacao!Desenho & "' and EM.Idoperacao = (Select MAX(EM1.Idoperacao) from Estoque_movimentacao EM1 where EM1.IDestoque = EM.IDestoque)"
        If TBProduto.EOF = False Then
            .Item(.Count).SubItems(27) = IIf(IsNull(TBProduto!data), "", TBProduto!data & " - " & TBProduto!Responsavel)
        End If
        TBProduto.Close
        
        .Item(.Count).SubItems(28) = IIf(IsNull(TBLISTA_Validacao!Versao_estrutura), "", TBLISTA_Validacao!Versao_estrutura)
        .Item(.Count).SubItems(29) = IIf(IsNull(TBLISTA_Validacao!Versao_processo), "", TBLISTA_Validacao!Versao_processo)
        
    End With
    TBLISTA_Validacao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Validacao.RecordCount
If TBLISTA_Validacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Validacao.PageCount
ElseIf TBLISTA_Validacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Validacao.PageCount & " de: " & TBLISTA_Validacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Validacao.AbsolutePage - 1 & " de: " & TBLISTA_Validacao.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparListas()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_arquivos.ListItems.Clear
Frame_versao.Enabled = False
Cmb_versao_estrutura.ListIndex = -1
Cmb_versao_processo.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.SelectedItem.ListSubItems.Count = 0 Then Exit Sub

Frame_versao.Enabled = True

ProcCarregaComboVersao Cmb_versao_estrutura, True, True, True, False, Lista.SelectedItem.ListSubItems(3)
ProcCarregaComboVersao Cmb_versao_processo, True, True, False, True, Lista.SelectedItem.ListSubItems(3)

Cmb_versao_estrutura.ListIndex = -1
Cmb_versao_processo.ListIndex = -1
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Versao_estrutura, Versao_processo from vendas_carteira where Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    If IsNull(TBLISTA!Versao_estrutura) = False And TBLISTA!Versao_estrutura <> "" Then Cmb_versao_estrutura = TBLISTA!Versao_estrutura
    If IsNull(TBLISTA!Versao_processo) = False And TBLISTA!Versao_processo <> "" Then Cmb_versao_processo = TBLISTA!Versao_processo
End If
TBLISTA.Close

1:
With Lista_arquivos.ListItems
    .Clear
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select imagem from projproduto where Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and Imagem IS NOT NULL and Imagem <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        .Add , , IIf(IsNull(TBLISTA!imagem), "", TBLISTA!imagem)
    End If
    TBLISTA.Close

    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select PD.Caminho from Projproduto P INNER JOIN projproduto_documentos PD ON P.Codproduto = PD.Codproduto where P.Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' order by ID", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBLISTA.RecordCount
        PBLista.Value = 1
        Contador = 0
        Do While TBLISTA.EOF = False
            .Add , , IIf(IsNull(TBLISTA!caminho), "", TBLISTA!caminho)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
        
    End If
    TBLISTA.Close
End With

Exit Sub
tratar_erro:
    If Err.Number = 383 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ProcLimparListas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

ProcLimparListas
If optPeriodo.Value = 1 Then
    Chk_data_venda.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = True
    msk_fltFim.Value = Date
    msk_fltInicio.Value = Date
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

ProcLimparListas

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
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
