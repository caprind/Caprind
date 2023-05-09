VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_Relatorios_Indice_Atraso 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Relatórios - Índice de atraso"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
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
   Begin VB.OptionButton optPrazoFinal_Orig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo final original"
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
      Left            =   11220
      TabIndex        =   9
      Top             =   1080
      Width           =   1875
   End
   Begin VB.OptionButton optPrazoFinal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo final"
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
      Left            =   9930
      TabIndex        =   8
      Top             =   1080
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.OptionButton optFaturamento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data faturamento"
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
      Left            =   13230
      TabIndex        =   10
      Top             =   1080
      Width           =   1845
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
      Height          =   1005
      Left            =   60
      TabIndex        =   35
      Top             =   1260
      Width           =   1905
      Begin VB.OptionButton Opt_prazo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pr. final"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1155
      End
      Begin VB.OptionButton Opt_prazo_qtde 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pr. final e qtde."
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
         Top             =   330
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   14040
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_Relatorios_Indice_Atraso.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   55
      TabIndex        =   26
      Top             =   9180
      Width           =   15195
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total vendida."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_indice 
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
         Left            =   12975
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Índice."
         Top             =   390
         Width           =   2010
      End
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
         Left            =   4445
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total faturada."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_qtde_total_atraso 
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
         Left            =   8710
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total atraso."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total vendida"
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
         Left            =   375
         TabIndex        =   30
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4610
         TabIndex        =   29
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Índice"
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
         Left            =   13718
         TabIndex        =   28
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total atraso"
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
         Left            =   8965
         TabIndex        =   27
         Top             =   180
         Width           =   1500
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   33
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
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   34
      Top             =   8910
      Width           =   11775
      _ExtentX        =   20770
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
   Begin VB.Frame Frame4 
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
      Height          =   1005
      Left            =   5160
      TabIndex        =   22
      Top             =   1260
      Width           =   7965
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmVendas_Relatorios_Indice_Atraso.frx":2DFA
         Left            =   2310
         List            =   "frmVendas_Relatorios_Indice_Atraso.frx":2DFC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Width           =   5475
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
         ItemData        =   "frmVendas_Relatorios_Indice_Atraso.frx":2DFE
         Left            =   180
         List            =   "frmVendas_Relatorios_Indice_Atraso.frx":2E11
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Opções para filtro."
         Top             =   480
         Width           =   2115
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
         Left            =   817
         TabIndex        =   24
         Top             =   270
         Width           =   840
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
         Left            =   4312
         TabIndex        =   23
         Top             =   270
         Width           =   1470
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1005
      Left            =   3690
      TabIndex        =   25
      Top             =   1260
      Width           =   1455
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
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
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
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
         TabIndex        =   5
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   1005
      Left            =   1980
      TabIndex        =   31
      Top             =   1260
      Width           =   1695
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
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
         TabIndex        =   2
         ToolTipText     =   "0"
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1425
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   13140
      TabIndex        =   19
      Top             =   1260
      Width           =   2115
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   12
         ToolTipText     =   "Data final."
         Top             =   630
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
         Format          =   488570881
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   11
         ToolTipText     =   "Data inicio."
         Top             =   270
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
         Format          =   488570881
         CurrentDate     =   39057
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   240
         TabIndex        =   21
         Top             =   330
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   180
         TabIndex        =   20
         Top             =   690
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6615
      Left            =   60
      TabIndex        =   13
      Top             =   2280
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11668
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Prazo final"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. faturamento"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Pedido int."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Qtde. vendida"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Qtde. faturada"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Qtde. atraso"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6615
      Left            =   60
      TabIndex        =   14
      Top             =   2280
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11668
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
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Object.Width           =   16942
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Qtde. vendida"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Qtde. faturada"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Qtde. atraso"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11940
      TabIndex        =   32
      Top             =   8940
      Width           =   3315
   End
End
Attribute VB_Name = "frmVendas_Relatorios_Indice_Atraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataBanco As String 'OK
Dim DataTexto As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=FOVIOhJT6Dw&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=31&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then
        If Lista.ListItems.Count = 0 Then Exit Sub
    Else
        If Lista1.ListItems.Count = 0 Then Exit Sub
    End If
Else
    If Lista1.ListItems.Count = 0 Then Exit Sub
End If
Vendas_Relatorio_Historico = False
Vendas_Relatorio_IndiceAtraso = True
Vendas_Relatorio_Comissao = False
Compras_Relatorio_IndiceAtraso = False
PCP_relatorios_indice_atraso = False
Manutencao_Relatorio_Historico = False
FrmMenu_impressao_padrao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcLocalizar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Posicao = 0
Lista.ListItems.Clear
Lista1.ListItems.Clear
If TBLISTA.EOF = False Then
    Posicao = TBLISTA.RecordCount
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil)
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select VP.NCotacao, VP.Revisao, VP.Cliente, VC.* from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.Cotacao = VP.Cotacao where VC.Codigo = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Ncotacao), "", TBAbrir!Ncotacao)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Revisao), "", TBAbrir!Revisao)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!descricao_tecnica), "", TBAbrir!descricao_tecnica)
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
                End If
                TBAbrir.Close
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!QtdePrev), "", TBLISTA!QtdePrev)
                .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!qtdeOK), "", TBLISTA!qtdeOK)
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!qtdeNC), "", TBLISTA!qtdeNC)
                .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%")
            End With
        Else
            With Lista1.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!QtdePrev), "", TBLISTA!QtdePrev)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!qtdeOK), "", TBLISTA!qtdeOK)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!qtdeNC), "", TBLISTA!qtdeNC)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%")
            End With
        End If
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_qtde_total_vendida = TBLISTA!QtdePrevista
    Txt_qtde_total_faturada = TBLISTA!QtdeProduzida
    Txt_qtde_total_atraso = TBLISTA!qtdeNC
    Txt_indice = Format(TBLISTA!TotalEficiencia, "###,##0.00") & "%"
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
Txt_qtde_total_vendida = ""
Txt_qtde_total_faturada = ""
Txt_qtde_total_atraso = ""
Txt_indice = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
Formulario = "Vendas/Relatórios/Índice de atraso"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
cmbfiltrarpor.Text = "Código interno"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Relatórios/Índice de atraso"
Direitos
ProcLimpaVariaveisPrincipais

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

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_individual.Value = True Then ProcCarregaComboTextoFiltro
Lista1.ColumnHeaders(2).Text = cmbfiltrarpor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTextoFiltro()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Cliente": Texto = "VP.Cliente"
    Case "Código de referência": Texto = "VC.N_Referencia"
    Case "Código interno": Texto = "VC.desenho"
    Case "Descrição": Texto = "VC.descricao_tecnica"
    Case "Família": Texto = "VC.Familia"
End Select
With cmbTexto
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & Texto & " as NomeCampo1 from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.Cotacao = VP.Cotacao where " & Texto & " is not null and " & Texto & " <> N'' group by " & Texto, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .AddItem ""
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!NomeCampo1
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

Acao = "filtrar"
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
IDlista = 0
Inicio = Time
ProcLimpaCamposTotais
ProcAbrirTabelas
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True And optDetalhado.Value = True Then
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data, Maquina", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Maquina", Conexao, adOpenKeyset, adLockOptimistic
End If
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

Select Case cmbfiltrarpor
    Case "Código interno": TextoFiltro = "VC.desenho"
    Case "Código de referência": TextoFiltro = "VC.n_referencia"
    Case "Descrição": TextoFiltro = "VC.Descricao_tecnica"
    Case "Família": TextoFiltro = "VC.familia"
    Case "Cliente": TextoFiltro = "VP.Cliente"
End Select
TextoFiltro1 = " "
If Opt_individual.Value = True Then
    If cmbTexto <> "" Then TextoFiltro1 = TextoFiltro & " = '" & cmbTexto & "' and "
    Ordenar = "VC.prazofinal"
Else
    Ordenar = TextoFiltro & ", VC.CODIGO"
End If
CamposFiltro = "VC.CODIGO, VC.desenho, VC.n_referencia, VC.Descricao_tecnica, VC.familia, VC.Prazofinal, VC.Prazo_original, VC.DataFaturamento, VC.Quantidade, VP.Cliente"

If Opt_prazo_qtde.Value = True Then TextoFiltro2 = "and NF.Int_status = 1" Else TextoFiltro2 = ""
If Opt_prazo.Value = True And (optPrazoFinal.Value = True Or optPrazoFinal_Orig.Value = True) Then
    INNERJOINTEXTO = "vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.Cotacao = VP.Cotacao"
Else
    INNERJOINTEXTO = "((vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.Cotacao = VP.Cotacao) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = VC.Codigo and NFPP.Codinterno = VC.Desenho) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFPP.ID_nota"
End If
If optPrazoFinal.Value = True Then
    TextoFiltro3 = "VC.prazofinal"
ElseIf optPrazoFinal_Orig.Value = True Then
        TextoFiltro3 = "VC.Prazo_original"
    Else
        TextoFiltro3 = "NF.dt_DataEmissao"
End If

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select " & CamposFiltro & " from " & INNERJOINTEXTO & " where " & TextoFiltro1 & TextoFiltro3 & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' " & TextoFiltro2 & " and (VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL') group by " & CamposFiltro & " order by " & Ordenar, Conexao, adOpenKeyset, adLockReadOnly
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If TBCarteira.EOF = False Then
    Permitido = True
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBCarteira.EOF = False
        If Opt_prazo_qtde.Value = True Then
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select NFPP.*, NF.dt_dataemissao from tbl_Detalhes_Nota_pedidos NFPP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFPP.ID_nota where NF.Int_status = 1 and NFPP.ID_carteira = " & TBCarteira!CODIGO & " and NFPP.Codinterno = '" & TBCarteira!Desenho & "' order by NFPP.ID_carteira", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    ProcSalvarDados
                    IDlista = TBLISTA!ID_carteira
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        Else
            ProcSalvarDados
        End If
        TBCarteira.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarDados()
On Error GoTo tratar_erro

Set TBProdutividade = CreateObject("adodb.recordset")
If Opt_individual.Value = True And optDetalhado.Value = True Then
    TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
    ProcEnviaDadosDetalhado
Else
    Select Case cmbfiltrarpor
        Case "Código interno": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Desenho & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        Case "Código de referência": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!N_referencia & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        Case "Descrição": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!descricao_tecnica & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        Case "Família": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Familia & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        Case "Cliente": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Cliente & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
    End Select
    ProcEnviaDadosResumido
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

TBProdutividade.AddNew
TBProdutividade!Ordem = TBCarteira!CODIGO
TBProdutividade!Data = TBCarteira!PrazoFinal

If Opt_prazo_qtde.Value = True Then
    If IDlista <> TBLISTA!ID_carteira Then TBProdutividade!QtdePrev = TBCarteira!quantidade 'Qtde. vendida
    TBProdutividade!qtdeOK = TBLISTA!quantidade 'Qtde. faturada
    TBProdutividade!Totalhsutil = Format(TBLISTA!dt_DataEmissao, "dd/mm/yy") 'Data faturamento
    If TBLISTA!dt_DataEmissao > TBCarteira!PrazoFinal Then TBProdutividade!qtdeNC = TBLISTA!quantidade Else TBProdutividade!qtdeNC = 0 'Qtde. em atraso
    TBProdutividade!Eficiencia = (TBProdutividade!qtdeNC / TBCarteira!quantidade) * 100
Else
    'Linha do pedido
    TBProdutividade!QtdePrev = 1 'Qtde. vendida
    TBProdutividade!qtdeOK = 1 'Qtde. faturada
    TBProdutividade!Totalhsutil = Format(TBCarteira!DataFaturamento, "dd/mm/yy") 'Data faturamento
    If TBCarteira!DataFaturamento > TBCarteira!PrazoFinal Then TBProdutividade!qtdeNC = 1 Else TBProdutividade!qtdeNC = 0 'Qtde. em atraso
    TBProdutividade!Eficiencia = Format((TBProdutividade!qtdeNC / 1) * 100, "###,##0.00")
End If

TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = cmbTexto
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

If TBProdutividade.EOF = True Then TBProdutividade.AddNew
If Opt_individual.Value = True Then
    Texto = cmbTexto
Else
    Select Case cmbfiltrarpor
        Case "Código interno": Texto = TBCarteira!Desenho
        Case "Código de referência": Texto = TBCarteira!N_referencia
        Case "Descrição": Texto = TBCarteira!descricao_tecnica
        Case "Família": Texto = TBCarteira!Familia
        Case "Cliente": Texto = TBCarteira!Cliente
    End Select
End If
TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario

If Opt_prazo_qtde.Value = True Then
    If IDlista <> TBLISTA!ID_carteira Then TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + TBCarteira!quantidade 'Qtde. vendida
    TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + TBLISTA!quantidade 'Qtde. faturada
    If TBLISTA!dt_DataEmissao > TBCarteira!PrazoFinal Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TBLISTA!quantidade 'Qtde. em atraso
Else
    'Linha do pedido
    TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + 1 'Qtde. vendida
    TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + 1 'Qtde. faturada
    If TBCarteira!DataFaturamento > TBCarteira!PrazoFinal Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + 1 'Qtde. em atraso
End If

If TBProdutividade!qtdeOK <> 0 Then TBProdutividade!Eficiencia = Format((TBProdutividade!qtdeNC / TBProdutividade!QtdePrev) * 100, "###,##0.00") Else TBProdutividade!Eficiencia = 0
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

quantidade = 0
QTLOTE = 0
Quant = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew

TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
If Opt_individual.Value = True Then TBAbrir!Texto = cmbfiltrarpor & ") : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(QtdePrev) as quantidade, Sum(QtdeOK) as QTLOTE, Sum(QtdeNC) as Quant from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    quantidade = IIf(IsNull(TBproducao!quantidade), 0, TBproducao!quantidade) 'Qtde. vendida
    QTLOTE = IIf(IsNull(TBproducao!QTLOTE), 0, TBproducao!QTLOTE) 'Qtde. faturada
    Quant = IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant) 'Qtde. em atraso
End If
TBproducao.Close
TBAbrir!QtdePrevista = quantidade 'Qtde. vendida
TBAbrir!QtdeProduzida = QTLOTE 'Qtde. faturada
TBAbrir!qtdeNC = Quant 'Qtde. em atraso
If TBAbrir!QtdeProduzida <> 0 Then TBAbrir!TotalEficiencia = (TBAbrir!qtdeNC / TBAbrir!QtdePrevista) * 100 Else TBAbrir!TotalEficiencia = 0
TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_prazo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_prazo_qtde_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    Lista.ListItems.Clear
    Lista.Visible = True
    Lista1.ListItems.Clear
    Lista1.Visible = False
    ProcLimpaCamposTotais
    ProcCarregaComboTextoFiltro
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optFaturamento_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPrazoFinal_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPrazoFinal_Orig_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    Lista.ListItems.Clear
    Lista.Visible = False
    Lista1.ListItems.Clear
    Lista1.Visible = True
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcLocalizar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
