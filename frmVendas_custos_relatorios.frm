VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_custos_relatorios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Custos - Relatórios - Detalhado"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin DrawSuite2022.USProgressBar PBLista1 
      Height          =   255
      Left            =   60
      TabIndex        =   132
      Top             =   8670
      Visible         =   0   'False
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custo acumulado por peça"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   180
      TabIndex        =   112
      Top             =   4110
      Visible         =   0   'False
      Width           =   3525
      Begin VB.TextBox Txt_compras_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Compras por peça."
         Top             =   510
         Width           =   2965
      End
      Begin VB.TextBox Txt_terceiros_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Terceiros por peça."
         Top             =   2805
         Width           =   2965
      End
      Begin VB.TextBox Txt_impostos_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Impostos por peça."
         Top             =   5115
         Width           =   2965
      End
      Begin VB.TextBox Txt_mao_de_obra_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Mão de obra por peça."
         Top             =   1275
         Width           =   2965
      End
      Begin VB.TextBox Txt_material_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Material por peça."
         Top             =   2040
         Width           =   2965
      End
      Begin VB.TextBox Txt_total_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Total por peça."
         Top             =   4350
         Width           =   2965
      End
      Begin VB.TextBox Txt_outras_peca 
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
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Outras despesas por peça."
         Top             =   3585
         Width           =   2965
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
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
         Left            =   1387
         TabIndex        =   119
         Top             =   300
         Width           =   750
      End
      Begin VB.Label txtProducao_peca 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Mão de obra"
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
         Left            =   1072
         TabIndex        =   118
         Top             =   1065
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Terceiros"
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
         Left            =   1200
         TabIndex        =   117
         Top             =   2595
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Impostos"
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
         Left            =   1185
         TabIndex        =   116
         Top             =   4905
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Material"
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
         Left            =   1245
         TabIndex        =   115
         Top             =   1845
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Total"
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
         Left            =   1380
         TabIndex        =   114
         Top             =   4140
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Outras despesas"
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
         Left            =   915
         TabIndex        =   113
         Top             =   3390
         Width           =   1740
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resultado por peça"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   3780
      TabIndex        =   107
      Top             =   4110
      Visible         =   0   'False
      Width           =   3525
      Begin VB.TextBox Txt_valor_venda_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0,0000"
         ToolTipText     =   "Valor de venda por peça."
         Top             =   975
         Width           =   2965
      End
      Begin VB.TextBox Txt_custo_acumulado_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Custo acumulado por peça."
         Top             =   2255
         Width           =   2965
      End
      Begin VB.TextBox Txt_resultado_peca 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Resultado por peça."
         Top             =   3545
         Width           =   2965
      End
      Begin VB.TextBox Txt_resultado_peca_porcento 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Resultado por peça em porcentagem."
         Top             =   4815
         Width           =   2965
      End
      Begin VB.Label label_peca_porcento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Prejuízo (%)"
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
         Left            =   280
         TabIndex        =   111
         Top             =   4605
         Width           =   2970
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da venda"
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
         Left            =   1125
         TabIndex        =   110
         Top             =   765
         Width           =   1260
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Custo acumulado"
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
         Left            =   900
         TabIndex        =   109
         Top             =   2025
         Width           =   1725
      End
      Begin VB.Label label_peca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Prejuízo"
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
         Left            =   280
         TabIndex        =   108
         Top             =   3315
         Width           =   2965
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   75
      TabIndex        =   83
      Top             =   330
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Relatório"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Relatório (F5)"
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
      ButtonWidth1    =   60
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   64
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   68
      ButtonTop3      =   2
      ButtonWidth3    =   41
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
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
      ButtonLeft4     =   111
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   143
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   10470
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_custos_relatorios.frx":0000
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   85
      Top             =   9700
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados técnicos do produto"
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
      Height          =   975
      Left            =   75
      TabIndex        =   72
      Top             =   2955
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox txtCod_ref 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Código de referência."
         Top             =   510
         Width           =   2385
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   510
         Width           =   10455
      End
      Begin VB.TextBox txtCod_int 
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
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   510
         Width           =   1965
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   14790
         Picture         =   "frmVendas_custos_relatorios.frx":22C5
         Top             =   210
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgFolder 
         Height          =   240
         Left            =   14520
         Picture         =   "frmVendas_custos_relatorios.frx":284F
         Top             =   210
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
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
         Left            =   712
         TabIndex        =   75
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label25 
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
         Height          =   195
         Left            =   9442
         TabIndex        =   74
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label26 
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
         Height          =   195
         Left            =   2677
         TabIndex        =   73
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resultado final"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   11610
      TabIndex        =   102
      Top             =   4110
      Visible         =   0   'False
      Width           =   3525
      Begin VB.TextBox Txt_resultado_porcento 
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
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Resultado em porcentagem."
         Top             =   4815
         Width           =   2965
      End
      Begin VB.TextBox Txt_valor_venda 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "0,0000"
         ToolTipText     =   "Valor de venda."
         Top             =   975
         Width           =   2965
      End
      Begin VB.TextBox Txt_custo_acumulado 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Custo acumulado."
         Top             =   2255
         Width           =   2965
      End
      Begin VB.TextBox Txt_resultado 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Resultado."
         Top             =   3545
         Width           =   2965
      End
      Begin VB.Label label_resultado_porcento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Prejuízo (%)"
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
         Left            =   280
         TabIndex        =   106
         Top             =   4605
         Width           =   2965
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da venda"
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
         Left            =   1132
         TabIndex        =   105
         Top             =   765
         Width           =   1260
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Custo acumulado"
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
         Left            =   900
         TabIndex        =   104
         Top             =   2025
         Width           =   1725
      End
      Begin VB.Label label_resultado 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Prejuízo"
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
         Left            =   280
         TabIndex        =   103
         Top             =   3315
         Width           =   2965
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custo total acumulado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   8010
      TabIndex        =   120
      Top             =   4110
      Visible         =   0   'False
      Width           =   3525
      Begin VB.TextBox Txt_compras 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Compras."
         Top             =   510
         Width           =   2965
      End
      Begin VB.TextBox Txt_material 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Material."
         Top             =   1980
         Width           =   2965
      End
      Begin VB.TextBox Txt_mao_de_obra 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Mão de obra."
         Top             =   1215
         Width           =   2965
      End
      Begin VB.TextBox Txt_impostos 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Impostos."
         Top             =   5055
         Width           =   2965
      End
      Begin VB.TextBox Txt_terceiros 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Terceiros."
         Top             =   2745
         Width           =   2965
      End
      Begin VB.TextBox Txt_total 
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
         Left            =   280
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Total."
         Top             =   4290
         Width           =   2965
      End
      Begin VB.TextBox Txt_outras 
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Outras despesas."
         Top             =   3525
         Width           =   2965
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
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
         Left            =   1387
         TabIndex        =   127
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Material"
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
         Left            =   1245
         TabIndex        =   126
         Top             =   1785
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Impostos"
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
         Left            =   1185
         TabIndex        =   125
         Top             =   4845
         Width           =   1140
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Terceiros"
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
         Left            =   1200
         TabIndex        =   124
         Top             =   2535
         Width           =   1125
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Mão de obra"
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
         Left            =   1072
         TabIndex        =   123
         Top             =   1005
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Total"
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
         Left            =   1380
         TabIndex        =   122
         Top             =   4080
         Width           =   765
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Outras despesas"
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
         Left            =   975
         TabIndex        =   121
         Top             =   3330
         Width           =   1740
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados comerciais do produto"
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
      Height          =   1635
      Left            =   75
      TabIndex        =   65
      Top             =   1320
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox Txt_prazo 
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
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Prazo final."
         Top             =   1140
         Width           =   1425
      End
      Begin VB.CommandButton Cmd_abrir_nf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14700
         Picture         =   "frmVendas_custos_relatorios.frx":2DD9
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Abrir lista de notas fiscais."
         Top             =   1140
         Width           =   315
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   12750
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Status do pedido."
         Top             =   495
         Width           =   2265
      End
      Begin VB.TextBox txtRev 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Revisão."
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox txtPedido_cliente 
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
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Pedido do cliente."
         Top             =   1140
         Width           =   7575
      End
      Begin VB.TextBox txtdata_venda 
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
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Data da venda."
         Top             =   1140
         Width           =   1425
      End
      Begin VB.TextBox txtQtd_faturada 
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
         Left            =   12660
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade faturada."
         Top             =   1140
         Width           =   2025
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Cliente."
         Top             =   495
         Width           =   9405
      End
      Begin VB.TextBox txtQtd_vendida 
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
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade vendida."
         Top             =   1140
         Width           =   1995
      End
      Begin VB.TextBox txtPedido 
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
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Pedido interno."
         Top             =   495
         Width           =   2445
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   1957
         TabIndex        =   82
         Top             =   930
         Width           =   750
      End
      Begin VB.Label Label44 
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
         Left            =   6240
         TabIndex        =   77
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label34 
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
         Height          =   195
         Left            =   13650
         TabIndex        =   76
         Top             =   285
         Width           =   465
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. faturada"
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
         Left            =   13110
         TabIndex        =   71
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. vendida"
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
         Left            =   11130
         TabIndex        =   70
         Top             =   930
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. da venda"
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
         Left            =   420
         TabIndex        =   69
         Top             =   930
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev."
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
         Left            =   2805
         TabIndex        =   68
         Top             =   285
         Width           =   345
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido interno"
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
         Left            =   885
         TabIndex        =   67
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label13 
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
         Height          =   195
         Left            =   7785
         TabIndex        =   66
         Top             =   285
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Pedidos"
      TabPicture(0)   =   "frmVendas_custos_relatorios.frx":2EDB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame15"
      Tab(0).Control(1)=   "txtIdcarteira"
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(3)=   "Lista_filtro"
      Tab(0).Control(4)=   "USToolBar2"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Compras"
      TabPicture(1)   =   "frmVendas_custos_relatorios.frx":2EF7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Produção"
      TabPicture(2)   =   "frmVendas_custos_relatorios.frx":2F13
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Grid1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Resultados"
      TabPicture(3)   =   "frmVendas_custos_relatorios.frx":2F2F
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Valor por família"
      TabPicture(4)   =   "frmVendas_custos_relatorios.frx":2F4B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Lista_resultados"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Valor por OS"
      TabPicture(5)   =   "frmVendas_custos_relatorios.frx":2F67
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Lista_OS"
      Tab(5).ControlCount=   1
      Begin FlexCell.Grid Grid1 
         Height          =   4335
         Left            =   90
         TabIndex        =   130
         Top             =   4320
         Visible         =   0   'False
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   7646
         Cols            =   2
         DefaultFontSize =   8.25
         GridColor       =   12632256
         ReadOnly        =   -1  'True
         Rows            =   2
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   5985
         Left            =   75
         TabIndex        =   90
         Top             =   3930
         Width           =   15195
         Begin VB.CheckBox chkNivel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Carregar lista em forma de niveis"
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
            TabIndex        =   131
            Top             =   120
            Width           =   3225
         End
         Begin VB.TextBox Txt_custo_terceiros 
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
            Left            =   7614
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total de terceiros."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox txtCusto_previsto 
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
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total previsto de mão de obra."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox txtCusto_real 
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
            Left            =   2658
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total real de mão de obra."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox Txt_custo_material 
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
            Left            =   5136
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total de material."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox Txt_custo_total 
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
            Left            =   12570
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox Txt_custo_outras 
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
            Height          =   315
            Left            =   10080
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total de outras despesas."
            Top             =   5415
            Width           =   2415
         End
         Begin MSComctlLib.ListView lista_fabricacao 
            Height          =   4365
            Left            =   0
            TabIndex        =   33
            Top             =   375
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   7699
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
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Text            =   "ID"
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
               Text            =   "Cód. de ref."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   3890
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Ordem"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Tipo"
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Qtde. prod."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Tempo total real"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Ct. prev. (mo)"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Ct. real (mo)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Ct. material"
               Object.Width           =   1677
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Object.Tag             =   "N"
               Text            =   "Ct. terceiros"
               Object.Width           =   1676
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   13
               Object.Tag             =   "N"
               Text            =   "Ct. outras desp."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   14
               Object.Tag             =   "N"
               Text            =   "Ct. total"
               Object.Width           =   1587
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Custo total terceiros"
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
            Left            =   7786
            TabIndex        =   96
            Top             =   5205
            Width           =   2070
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo total real (mo)"
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
            Left            =   2980
            TabIndex        =   95
            Top             =   5205
            Width           =   1770
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo total previsto (mo)"
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
            Left            =   315
            TabIndex        =   94
            Top             =   5205
            Width           =   2145
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Custo total material"
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
            Left            =   5331
            TabIndex        =   93
            Top             =   5205
            Width           =   2025
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(=) Custo total"
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
            Left            =   13147
            TabIndex        =   92
            Top             =   5205
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Custo total outras desp."
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
            Left            =   10122
            TabIndex        =   91
            Top             =   5205
            Width           =   2355
         End
      End
      Begin VB.Frame Frame9 
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
         Height          =   6045
         Left            =   -74940
         TabIndex        =   128
         Top             =   3930
         Width           =   15195
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   6045
         Left            =   -74925
         TabIndex        =   97
         Top             =   3930
         Width           =   15195
         Begin VB.TextBox Txt_custo_total_IPI_compras 
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
            Height          =   315
            Left            =   7620
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total de IPI."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox Txt_custo_total_serv_compras 
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
            Height          =   315
            Left            =   10092
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total de serviços."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox Txt_custo_total_compras 
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
            Left            =   12570
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total."
            Top             =   5415
            Width           =   2415
         End
         Begin VB.TextBox Txt_custo_total_prod_compras 
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
            Left            =   5130
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Custo total de produtos."
            Top             =   5415
            Width           =   2415
         End
         Begin MSComctlLib.ListView Lista_compras 
            Height          =   4725
            Left            =   0
            TabIndex        =   28
            Top             =   15
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   8334
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
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Cód. interno"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Cód. de ref."
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   7064
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Ped. de compra"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Tipo"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Vlr. unit."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Custo total"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Vlr. IPI"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Custo total"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Custo total IPI"
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
            Left            =   8145
            TabIndex        =   101
            Top             =   5205
            Width           =   1560
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(+) Custo total serviços"
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
            Left            =   10395
            TabIndex        =   100
            Top             =   5205
            Width           =   2010
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(=) Custo total"
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
            Left            =   13170
            TabIndex        =   99
            Top             =   5205
            Width           =   1260
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo total produtos"
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
            Left            =   5565
            TabIndex        =   98
            Top             =   5205
            Width           =   1740
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   86
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
            DibPicture      =   "frmVendas_custos_relatorios.frx":2F83
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
            Left            =   11220
            TabIndex        =   8
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_custos_relatorios.frx":672A
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
            Left            =   10110
            TabIndex        =   6
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
            Left            =   10680
            TabIndex        =   7
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_custos_relatorios.frx":A234
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
            Left            =   12300
            TabIndex        =   10
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_custos_relatorios.frx":E325
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
            TabIndex        =   89
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
            TabIndex        =   88
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label14 
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
            Left            =   3090
            TabIndex        =   87
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.TextBox txtIdcarteira 
         Height          =   285
         Left            =   -73230
         TabIndex        =   78
         Text            =   "0"
         Top             =   7650
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Frame Frame11 
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
         Left            =   -74925
         TabIndex        =   79
         Top             =   1320
         Width           =   15195
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   10230
            TabIndex        =   129
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
               TabIndex        =   14
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
               TabIndex        =   12
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
               TabIndex        =   11
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
               TabIndex        =   13
               Top             =   180
               Width           =   1155
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
            ItemData        =   "frmVendas_custos_relatorios.frx":11BB3
            Left            =   180
            List            =   "frmVendas_custos_relatorios.frx":11BCC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   9975
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
            Left            =   150
            TabIndex        =   1
            ToolTipText     =   "Texto para pesquisa."
            Top             =   1110
            Width           =   14865
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
            Left            =   150
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Família."
            Top             =   1110
            Width           =   14865
         End
         Begin VB.Label Label33 
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
            Left            =   4747
            TabIndex        =   81
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label32 
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
            Left            =   6847
            TabIndex        =   80
            Top             =   900
            Width           =   1470
         End
      End
      Begin MSComctlLib.ListView Lista_filtro 
         Height          =   6165
         Left            =   -74925
         TabIndex        =   3
         ToolTipText     =   "Lista de propostas cadastradas."
         Top             =   2910
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10874
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Pedido int."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   6448
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Cód. de ref."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   9525
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "D"
            Text            =   "Prazo final"
            Object.Width           =   2117
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   84
         Top             =   330
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   1720
         ButtonCount     =   5
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
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonAlignment2=   2
         ButtonType2     =   1
         ButtonStyle2    =   -1
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState2    =   -1
         ButtonLeft2     =   46
         ButtonTop2      =   4
         ButtonWidth2    =   2
         ButtonHeight2   =   54
         ButtonCaption3  =   "Ajuda"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Ajuda (F1)"
         ButtonKey3      =   "3"
         ButtonAlignment3=   2
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   50
         ButtonTop3      =   2
         ButtonWidth3    =   41
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Sair"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Sair (Esc)"
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
         ButtonLeft4     =   93
         ButtonTop4      =   2
         ButtonWidth4    =   30
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonKey5      =   "5"
         ButtonAlignment5=   2
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   5
         ButtonLeft5     =   125
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   10770
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_custos_relatorios.frx":11C26
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_OS 
         Height          =   5745
         Left            =   -74925
         TabIndex        =   63
         Top             =   3945
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10134
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483628
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
         NumItems        =   22
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "OS"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Fase"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Gupo/op."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Posto de trab."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Ct. hr. prep. prev."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Prep. prev. pç"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Prep. utiliz. pç"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Ct. hr. exec. prev."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Exec. prev. pç"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Exec. utiliz. pç"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Qtde. OK"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Qtde. NC"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Object.Tag             =   "N"
            Text            =   "Qtde. prod."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Object.Tag             =   "N"
            Text            =   "Ct. prep. prev. lote"
            Object.Width           =   2522
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Object.Tag             =   "N"
            Text            =   "Ct. exec. prev. lote"
            Object.Width           =   2522
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Object.Tag             =   "N"
            Text            =   "Ct. prep. real lote"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Object.Tag             =   "N"
            Text            =   "Ct. exec. real lote"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Object.Tag             =   "N"
            Text            =   "Ct. total lote"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   19
            Object.Tag             =   "N"
            Text            =   "Efic. prep."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   20
            Object.Tag             =   "N"
            Text            =   "Efic. exec."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   21
            Object.Tag             =   "N"
            Text            =   "Efic. média"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_resultados 
         Height          =   5745
         Left            =   -74925
         TabIndex        =   62
         Top             =   3945
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10134
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   22939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   3175
         EndProperty
      End
   End
End
Attribute VB_Name = "frmVendas_custos_relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLISTA_Custos_Detalhado As ADODB.Recordset 'OK

Dim TBOrdem1        As ADODB.Recordset
Dim TBOrdem2        As ADODB.Recordset
Dim TBOrdem3        As ADODB.Recordset
Dim TBOrdem4        As ADODB.Recordset
Dim TBOrdem5        As ADODB.Recordset
Dim TBOrdem6        As ADODB.Recordset
Dim TBOrdem7        As ADODB.Recordset
Dim TBOrdem8        As ADODB.Recordset
Dim TBOrdem9        As ADODB.Recordset
Dim TBOrdem10        As ADODB.Recordset
Dim TBOrdem11        As ADODB.Recordset
Dim TBOrdem12        As ADODB.Recordset

'GridEstrutura
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String

Private Sub chkNivel_Click()
On Error GoTo tratar_erro

If chkNivel.Value = 1 Then
    lista_fabricacao.Visible = False
    Grid1.Visible = True
    ProcCarregaListaProducaoGrid
Else
    lista_fabricacao.Visible = True
    Grid1.Visible = False
    ProcCarregaListaProducao
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

Lista_filtro.ListItems.Clear
txtTexto = ""
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista_filtro.ListItems.Clear
If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    cmbfamilia.Clear
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", False
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
    If cmbfiltrarpor = "Ordem" And txtTexto <> "" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_abrir_nf_Click()
On Error GoTo tratar_erro

Qtde = IIf(txtQtd_faturada = "", 0, txtQtd_faturada)
If Qtde = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select NF.dt_DataEmissao, NFP.ID_nota, NFP.int_NotaFiscal, NFP.int_Qtd as Qtde FROM (((tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo) INNER JOIN vendas_carteira VC ON VC.Codigo = NFPP.ID_carteira and VC.Desenho = NFPP.Codinterno) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NF.Int_TipoNota = 1 and NFPP.ID_carteira = " & txtidcarteira & " and NFPP.Codinterno = '" & txtCod_int & "' and (CFOP.Vendas = 'True' or CFOP.MaoObra = 'True') order by NFP.Int_NotaFiscal", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    frmVendas_custos_relatorios_NF.Show 1
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista_filtro.ListItems.Clear

StatusFiltro = "(VC.liberacao = 'VENDIDA' or VC.liberacao = 'VENDIDA PARCIAL' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL')"
Set TBLISTA_Custos_Detalhado = CreateObject("adodb.recordset")
If txtTexto <> "" Or cmbfamilia <> "" Then
    If cmbfiltrarpor = "Família" Then
        TBLISTA_Custos_Detalhado.Open "Select VC.*, VP.Ncotacao, VP.Revisao, VP.Cliente from vendas_carteira VC INNER JOIN Vendas_proposta VP ON VC.Cotacao = VP.Cotacao where VC.familia = '" & cmbfamilia & "' and " & StatusFiltro & " order by VC.cotacao, VC.codigo", Conexao, adOpenKeyset, adLockReadOnly
    ElseIf cmbfiltrarpor = "Ordem" Then
            TBLISTA_Custos_Detalhado.Open "Select VC.*, VP.Ncotacao, VP.Revisao, VP.Cliente from (vendas_carteira VC INNER JOIN Vendas_proposta VP ON VC.Cotacao = VP.Cotacao) INNER JOIN producao_pedidos PP ON VC.codigo = PP.Idcarteira where PP.Ordem = " & txtTexto & " and " & StatusFiltro & " order by VC.cotacao, VC.codigo", Conexao, adOpenKeyset, adLockReadOnly
        Else
            Select Case cmbfiltrarpor
                Case "Pedido interno": NomeCampo = "VP.Ncotacao"
                Case "Cliente": NomeCampo = "VP.cliente"
                Case "Cód. interno": NomeCampo = "VC.desenho"
                Case "Cód. de referência": NomeCampo = "VC.N_referencia"
                Case "Descrição": NomeCampo = "VC.Descricao_tecnica"
            End Select
            If Optinicio.Value = True Then Texto = " like '" & txtTexto & "%'"
            If Optmeio.Value = True Then Texto = " like '%" & txtTexto & "%'"
            If Optfim.Value = True Then Texto = " like '%" & txtTexto & "'"
            If optIgual.Value = True Then Texto = " = '" & txtTexto & "'"
            TBLISTA_Custos_Detalhado.Open "Select VC.*, VP.Ncotacao, VP.Revisao, VP.Cliente from vendas_carteira VC INNER JOIN Vendas_proposta VP ON VC.Cotacao = VP.Cotacao where " & NomeCampo & Texto & " and " & StatusFiltro & " order by VC.cotacao, VC.codigo", Conexao, adOpenKeyset, adLockReadOnly
    End If
Else
    TBLISTA_Custos_Detalhado.Open "Select VC.*, VP.Ncotacao, VP.Revisao, VP.Cliente from vendas_carteira VC INNER JOIN Vendas_proposta VP ON VC.Cotacao = VP.Cotacao where " & StatusFiltro & " order by VC.cotacao, VC.codigo", Conexao, adOpenKeyset, adLockReadOnly
End If
If TBLISTA_Custos_Detalhado.EOF = False Then ProcExibePagina (1)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista_filtro.ListItems.Clear
TBLISTA_Custos_Detalhado.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Custos_Detalhado.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Custos_Detalhado.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Custos_Detalhado.RecordCount - IIf(Pagina > 1, (TBLISTA_Custos_Detalhado.PageSize * (Pagina - 1)), 0), TBLISTA_Custos_Detalhado.PageSize)
PBLista.Value = 1
contador = 0
Do While TBLISTA_Custos_Detalhado.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_filtro.ListItems
        .Add , , TBLISTA_Custos_Detalhado!CODIGO
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Custos_Detalhado!Ncotacao), "", TBLISTA_Custos_Detalhado!Ncotacao)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Custos_Detalhado!Revisao), "", TBLISTA_Custos_Detalhado!Revisao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Custos_Detalhado!Cliente), "", TBLISTA_Custos_Detalhado!Cliente)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Custos_Detalhado!Desenho), "", TBLISTA_Custos_Detalhado!Desenho)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Custos_Detalhado!N_referencia), "", TBLISTA_Custos_Detalhado!N_referencia)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Custos_Detalhado!descricao_tecnica), "", TBLISTA_Custos_Detalhado!descricao_tecnica)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Custos_Detalhado!PrazoFinal), "", Format(TBLISTA_Custos_Detalhado!PrazoFinal, "dd/mm/yy"))
    End With
    TBLISTA_Custos_Detalhado.MoveNext
    ContadorReg = ContadorReg + 1
    contador = contador + 1
    PBLista.Value = contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Custos_Detalhado.RecordCount
If TBLISTA_Custos_Detalhado.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Custos_Detalhado.PageCount
ElseIf TBLISTA_Custos_Detalhado.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Custos_Detalhado.PageCount & " de: " & TBLISTA_Custos_Detalhado.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Custos_Detalhado.AbsolutePage - 1 & " de: " & TBLISTA_Custos_Detalhado.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Custos_Detalhado.AbsolutePage <> 2 Then
    If TBLISTA_Custos_Detalhado.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Custos_Detalhado.PageCount - 1)
    Else
        TBLISTA_Custos_Detalhado.AbsolutePage = TBLISTA_Custos_Detalhado.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Custos_Detalhado.AbsolutePage)
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
    TBLISTA_Custos_Detalhado.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Custos_Detalhado.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Custos_Detalhado.AbsolutePage = 1
ProcExibePagina (TBLISTA_Custos_Detalhado.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Custos_Detalhado.AbsolutePage <> -3 Then
    If TBLISTA_Custos_Detalhado.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Custos_Detalhado.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Custos_Detalhado.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Custos_Detalhado.AbsolutePage = TBLISTA_Custos_Detalhado.PageCount
ProcExibePagina (TBLISTA_Custos_Detalhado.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
        'Case vbKeyF1: ProcAjuda
        Case vbKeyF2: ProcFiltrar
    End Select
Else
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
        'Case vbKeyF1: ProcAjuda
        Case vbKeyF5: ProcImprimir
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
ProcCarregaToolBar2 Me, 15195, 5, True
Formulario = "Custos/Relatórios/Detalhado"
Direitos
ProcLimpaVariaveisPrincipais
USToolBar1.Visible = False
SSTab1.Tab = 0
cmbfiltrarpor = "Pedido interno"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Custos/Relatórios/Detalhado"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

Conexao.Execute "DELETE from custos where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"
Conexao.Execute "DELETE from Custos_familias where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Custos", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
ProcEnviaDadosResultados
TBGravar.Update

'Verifica valor total por família no pedido
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPP.Familia, Sum(CPP.preco_total) as Valores FROM (compras_pedido_lista CPP INNER JOIN Producao P ON CPP.Ordem = P.Ordem) INNER JOIN Producao_pedidos PP ON CPP.Ordem = PP.Ordem where PP.IDCarteira = " & txtidcarteira & " and CPP.Remessa = 'False' and (CPP.OS IS NULL or CPP.OS = 0) and (CPP.status_item = 'N_RECEBIDO' or CPP.status_item = 'RECEBIDO' or CPP.status_item = 'PARCIAL') group by CPP.familia", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        ProcGravarDadosFamilia TBLISTA!Familia, IIf(IsNull(TBLISTA!Valores), 0, TBLISTA!Valores)
        TBLISTA.MoveNext
    Loop
End If
'Verifica valor total por família no pedido (Empenho)
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPL.Familia, Sum(ROUND(CPL.preco_unitario_desconto * CPLE.Qtde_empenho,2)) as Valores FROM (compras_pedido_lista CPL INNER JOIN Compras_pedido_lista_empenhos CPLE ON CPLE.IDlista = CPL.IDLista) INNER JOIN Projproduto P ON P.Desenho = CPL.Desenho where CPLE.IDCarteira = " & txtidcarteira & " and P.estoque = 'False' and CPL.Remessa = 'False' and (CPL.OS IS NULL or CPL.OS = 0) group by CPL.familia", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        ProcGravarDadosFamilia TBLISTA!Familia, IIf(IsNull(TBLISTA!Valores), 0, TBLISTA!Valores)
        TBLISTA.MoveNext
    Loop
End If
'Verifica no estoque
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PROD.Classe, Sum(E.VlrTotal) as Valores, Sum(P.Quant) as Qt, Sum(P.Quantprod) as Qtde FROM (((Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem) INNER JOIN Producaomaterial PM ON PM.Ordem = P.Ordem) INNER JOIN Projproduto PROD ON PROD.Desenho = PM.Codigo) INNER JOIN Estoque_movimentacao E ON E.Ordem = P.Ordem and E.Familia = PROD.Classe where PP.IDCarteira = " & txtidcarteira & " and (PM.Saida = 'SIM' or PM.Saida = 'NÃO' or PM.Saida = 'PARCIAL') and (E.Operacao = 'SAIDA_ORDEM' or E.Operacao = 'SAIDA_ORDEM_PARCIAL') group by PROD.Classe", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        Valor1 = txtQtd_vendida
        If IsNull(TBLISTA!Qtde) = False And TBLISTA!Qtde <> 0 Then Quant = TBLISTA!Qtde Else Quant = TBLISTA!qt
        valor = (IIf(IsNull(TBLISTA!Valores), 0, TBLISTA!Valores) / Quant) * Valor1
        ProcGravarDadosFamilia TBLISTA!Classe, Format(valor, "###,##0.00")
        
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close
TBGravar.Close

NomeRel = "Custos_detalhado.rpt"
ProcImprimirRel "{Custos.IDcarteira} = " & txtidcarteira, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResultados()
On Error GoTo tratar_erro

TBGravar!IDcarteira = txtidcarteira
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
TBGravar!Custo_total_prev = txtCusto_previsto

'Por peça
TBGravar!Custo_total_compras_peca = Txt_compras_peca
TBGravar!Custo_total_real_peca = Txt_mao_de_obra_peca
TBGravar!Custo_total_material_peca = Txt_material_peca
TBGravar!Custo_total_terceiros_peca = Txt_terceiros_peca
TBGravar!Custo_total_outras_peca = Txt_outras_peca
TBGravar!ICMS_peca = Txt_impostos_peca
TBGravar!Valor_venda_peca = Txt_valor_venda_peca
TBGravar!Custo_acumulado_peca = Txt_custo_acumulado_peca
TBGravar!lucro_preju_peca = Txt_resultado_peca

'Total
TBGravar!Custo_total_compras = Txt_compras
TBGravar!Custo_total_real = Txt_mao_de_obra
TBGravar!Custo_total_material = Txt_material
TBGravar!Custo_total_terceiros = Txt_terceiros
TBGravar!Custo_total_outras = Txt_outras
TBGravar!ICMS = Txt_impostos
TBGravar!Valor_venda = Txt_valor_venda
TBGravar!Custo_acumulado = Txt_custo_acumulado
TBGravar!lucro_preju = Txt_resultado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarDadosFamilia(Familia As String, Valor_familia As Double)
On Error GoTo tratar_erro

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Custos_familias", Conexao, adOpenKeyset, adLockOptimistic
TBCompras.AddNew
TBCompras!IDCustos = TBGravar!ID
TBCompras!Responsavel = pubUsuario
TBCompras!Modulo = Formulario
TBCompras!Familia = Familia
TBCompras!Valor_familia = Valor_familia
TBCompras.Update
TBCompras.Close

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

Private Sub Lista_filtro_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_filtro, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_filtro_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_filtro.ListItems.Count = 0 Then Exit Sub
txtidcarteira = Lista_filtro.SelectedItem
ProcCarregaDadosProduto
ProcCarregaListaCompras
If chkNivel.Value = 1 Then ProcCarregaListaProducaoGrid Else ProcCarregaListaProducao
ProcCarregaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosProduto()
On Error GoTo tratar_erro

valor = 0
ValorTotal = 0
Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select VP.Ncotacao, VP.revisao, VP.Cliente, VP.dbl_valor_total, VC.* FROM vendas_proposta VP INNER JOIN vendas_carteira VC ON VP.cotacao = VC.cotacao where VC.codigo = " & txtidcarteira, Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = False Then
    txtPedido = IIf(IsNull(TBCarteira!Ncotacao), "", TBCarteira!Ncotacao)
    txtRev = IIf(IsNull(TBCarteira!Revisao), "", TBCarteira!Revisao)
    txtCliente = IIf(IsNull(TBCarteira!Cliente), "", TBCarteira!Cliente)
    txtStatus = IIf(IsNull(TBCarteira!Liberacao), "", TBCarteira!Liberacao)
    txtdata_venda = IIf(IsNull(TBCarteira!Datavendas), "", Format(TBCarteira!Datavendas, "dd/mm/yy"))
    Txt_prazo = IIf(IsNull(TBCarteira!PrazoFinal), "", Format(TBCarteira!PrazoFinal, "dd/mm/yy"))
    txtPedido_cliente = IIf(IsNull(TBCarteira!PCCliente), "", TBCarteira!PCCliente)
    txtQtd_vendida = IIf(IsNull(TBCarteira!quantidade), "", Format(TBCarteira!quantidade, "###,##0.0000"))
    txtQtd_faturada = IIf(IsNull(TBCarteira!QtdeFaturada), "", Format(TBCarteira!QtdeFaturada, "###,##0.0000"))
    txtCod_int = IIf(IsNull(TBCarteira!Desenho), "", TBCarteira!Desenho)
    txtCod_ref = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    txtdescricao = IIf(IsNull(TBCarteira!descricao_tecnica), "", TBCarteira!descricao_tecnica)
    
    Qtde = IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade) 'Qtde vendida
    Txt_valor_venda_peca = Format(IIf(IsNull(TBCarteira!preco_unitario_desconto), 0, TBCarteira!preco_unitario_desconto), "###,##0.0000000000")
    valor = IIf(IsNull(TBCarteira!preco_unitario_desconto), 0, TBCarteira!preco_unitario_desconto) * Qtde
    Txt_valor_venda = Format(valor, "###,##0.00")
    
    If TBCarteira!Tipo = "P" Then
        'Calcula o ICMS sem IPI
        'ICMS = IIf(IsNull(TBCarteira!IntICMS), 0, TBCarteira!IntICMS)
        'Valortotal = Txt_valor_venda
        'ValorICMS = (Valortotal * ICMS) / 100
        
        'Busca valor do ICMS do pedido
        ValorICMS = IIf(IsNull(TBCarteira!dbl_Valor_ICMS), 0, TBCarteira!dbl_Valor_ICMS)
        
        ValorTotal = Format(ValorICMS + IIf(IsNull(TBCarteira!Total_PIS_prod), 0, TBCarteira!Total_PIS_prod) + IIf(IsNull(TBCarteira!Total_Cofins_prod), 0, TBCarteira!Total_Cofins_prod) + IIf(IsNull(TBCarteira!Total_CSLL_prod), 0, TBCarteira!Total_CSLL_prod) + IIf(IsNull(TBCarteira!Total_IRPJ_prod), 0, TBCarteira!Total_IRPJ_prod) + IIf(IsNull(TBCarteira!Total_DAS), 0, TBCarteira!Total_DAS), "###,##0.0000000000")
    Else
        ValorTotal = Format(IIf(IsNull(TBCarteira!VlrISS), 0, TBCarteira!VlrISS) + IIf(IsNull(TBCarteira!Total_IRRF_serv), 0, TBCarteira!Total_IRRF_serv) + IIf(IsNull(TBCarteira!Total_IRPJ_serv), 0, TBCarteira!Total_IRPJ_serv) + IIf(IsNull(TBCarteira!Total_PIS_serv), 0, TBCarteira!Total_PIS_serv) + IIf(IsNull(TBCarteira!Total_Cofins_serv), 0, TBCarteira!Total_Cofins_serv) + IIf(IsNull(TBCarteira!Total_CSLL_serv), 0, TBCarteira!Total_CSLL_serv) + IIf(IsNull(TBCarteira!Total_DAS), 0, TBCarteira!Total_DAS), "###,##0.0000000000") 'Valor total de impostos por peça
    End If
    Txt_impostos_peca = Format(ValorTotal / Qtde, "###,##0.0000000000")
    Txt_impostos = Format(ValorTotal, "###,##0.00")
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaCompras()
On Error GoTo tratar_erro

valor = 0 'Produtos/Serviços
Valor1 = 0 'IPI
Valor2 = 0 'Total produtos
ValorIPI = 0 'Total IPI
Valor3 = 0 'Total serviços
Qtde = 0

Lista_compras.ListItems.Clear
CamposFiltro = "CPL.IDlista, CPL.Desenho, CPL.N_Referencia, CPL.Descricao, CPL.Tipo, CPL.preco_unitario_desconto, ISNULL(CPL.IPI, 0) AS IPI, CP.Pedido, CPLE.Qtde_empenho"
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select " & CamposFiltro & " from ((Compras_pedido_lista_empenhos CPLE INNER JOIN Compras_pedido_lista CPL ON CPL.IDlista = CPLE.IDlista) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDpedido) INNER JOIN Projproduto P ON P.Desenho = CPL.Desenho where CPLE.IDCarteira = " & txtidcarteira & " and P.estoque = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    PBLista11.Min = 0
    PBLista11.Max = TBCompras.RecordCount
    PBLista1.Value = 1
    contador = 0
    Do While TBCompras.EOF = False
        With Lista_compras.ListItems
            .Add , , TBCompras!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBCompras!Desenho), "", TBCompras!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBCompras!N_referencia), "", TBCompras!N_referencia)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBCompras!Descricao), "", TBCompras!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBCompras!Pedido), "", TBCompras!Pedido)
            .Item(.Count).SubItems(5) = IIf(TBCompras!Tipo = "P", "Prod.", "Serv.")
            .Item(.Count).SubItems(6) = IIf(IsNull(TBCompras!preco_unitario_desconto), "0,00000", Format(TBCompras!preco_unitario_desconto, "###,##0.0000000000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBCompras!Qtde_empenho), "0,0000", Format(TBCompras!Qtde_empenho, "###,##0.0000"))
            
            valor = TBCompras!preco_unitario_desconto * TBCompras!Qtde_empenho
            .Item(.Count).SubItems(8) = Format(valor, "###,##0.00")
            
            Valor1 = Format((valor * TBCompras!IPI) / 100, "###,##0.00")
            .Item(.Count).SubItems(9) = Format(Valor1, "###,##0.00")
            
            .Item(.Count).SubItems(10) = Format(valor + Valor1, "###,##0.00")
            
            If TBCompras!Tipo = "P" Then Valor2 = Valor2 + valor Else Valor3 = Valor3 + valor
            ValorIPI = ValorIPI + Valor1
        End With
        TBCompras.MoveNext
        contador = contador + 1
        PBLista1.Value = contador
    Loop
End If
TBCompras.Close

'Custo total
Txt_custo_total_prod_compras = Format(Valor2, "###,##0.00")
Txt_custo_total_IPI_compras = Format(ValorIPI, "###,##0.00")
Txt_custo_total_serv_compras = Format(Valor3, "###,##0.00")
Txt_custo_total_compras = Format(Valor2 + Valor3 + ValorIPI, "###,##0.00")

'Resultados por peça
Qtde = txtQtd_vendida
Txt_compras_peca = Format((Valor2 + Valor3 + ValorIPI) / Qtde, "###,##0.0000000000")

'Resultados
Txt_compras = Txt_custo_total_compras

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaProducaoGrid()
On Error GoTo tratar_erro
''ReDim arrNodes(2000)

Valor3 = 0 'Mão de obra
Valor_Cofins_Prod = 0 'Mão de obra por peça
Valor_Produto = 0 'Material
Valor_Cofins_Serv = 0 'Material por peça
ValorPagar = 0 'Terceiros
Valor_CSLL_Prod = 0 'Terceiros por peça
Valor_DAS = 0 'Outras desp.
Valor_ICMS_SN = 0 'Outras desp. por peça
Qtde = 0
valor = 0
ValorTotal = 0
Valor1 = 0
Valor2 = 0
Valor_INSS_Serv = 0

Call m_Tree.Nodes.Clear
Grid1.rows = 1
m_Row = 1
m_Col = 1
Contador1 = -1

Permitido = False

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select P.* FROM Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem where PP.IDCarteira = " & txtidcarteira & " AND P.desenho = '" & txtCod_int & "' order by P.Ordem", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    Permitido = True
Else
    'Verifica ordem do estoque que não tem vinculado o ID da carteira na ordem baixada do estoque
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select P.* FROM ((Producao P INNER JOIN Ordens_texto_SA OT ON OT.Ordem = P.Ordem) INNER JOIN Estoque_movimentacao EM ON EM.Lote = OT.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = EM.ID_prod_NF where NFPP.ID_carteira = " & txtidcarteira & " and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL') order by P.Ordem", Conexao, adOpenKeyset, adLockReadOnly
    If TBproducao.EOF = False Then
        Permitido = True
    End If
End If
If Permitido = True Then
    PBLista1.Min = 0
    PBLista1.Max = TBproducao.RecordCount
    PBLista1.Value = 1
    contador = 0
    Do While Not TBproducao.EOF
    
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 0
        
        qtdeGrid = IIf(IsNull(TBproducao!Quant), "", Format(TBproducao!Quant, "###,##0.0000"))
        qtdeProdGrid = IIf(IsNull(TBproducao!QuantProd), "", Format(TBproducao!QuantProd, "###,##0.0000"))
        TTTRealGrid = IIf(IsNull(TBproducao!TTTReal), "", Format(TBproducao!TTTReal, "hh:mm:ss"))
        CTTPrevGrid = IIf(IsNull(TBproducao!CTTPrev), "", Format(TBproducao!CTTPrev, "###,##0.00"))
        CTTRealGrid = IIf(IsNull(TBproducao!CTTReal), "", Format(TBproducao!CTTReal, "###,##0.00"))
        CTMaterialGrid = IIf(IsNull(TBproducao!CTMaterial), "", Format(TBproducao!CTMaterial, "###,##0.00"))
        CTServicoGrid = IIf(IsNull(TBproducao!CTServico), "", Format(TBproducao!CTServico, "###,##0.00"))
        CTOutrasGrid = IIf(IsNull(TBproducao!CTOutras), "", Format(TBproducao!CTOutras, "###,##0.00"))
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBproducao!Ordem & "' and desenho = '" & TBproducao!Desenho & "' and documento = '" & TBproducao!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
        End If
        TBItem.Close
        
        QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        
        'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBproducao!QuantProd), 0, Format(TBproducao!QuantProd, "0.0000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        'Calcula totalização===========================================================================================
        Qtde = Qtde + IIf(IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd) = 0, TBproducao!Quant, TBproducao!QuantProd)
        valor = valor + IIf(IsNull(TBproducao!CTTPrev), 0, TBproducao!CTTPrev) 'Custo previsto
        
        'Custo total
        Valor3 = IIf(IsNull(TBproducao!CTTReal), 0, TBproducao!CTTReal)
        Valor_Produto = IIf(IsNull(TBproducao!CTMaterial), 0, TBproducao!CTMaterial)
        ValorPagar = IIf(IsNull(TBproducao!CTServico), 0, TBproducao!CTServico)
        Valor_DAS = IIf(IsNull(TBproducao!CTOutras), 0, TBproducao!CTOutras)
        ValorTotalGrid = Format(Valor3 + Valor_Produto + ValorPagar + Valor_DAS, "###,##0.00")
        
        ValorTotal = ValorTotal + Valor3
        Valor1 = Valor1 + Valor_Produto
        Valor2 = Valor2 + ValorPagar
        Valor_INSS_Serv = Valor_INSS_Serv + Valor_DAS
        '=================================================================================================================

        arrNodes(Contador1).Text = TBproducao!Desenho & vbTab & "" & vbTab & TBproducao!N_referencia & vbTab & TBproducao!Produto & vbTab & TBproducao!Ordem & vbTab & IIf(TBproducao!Tipo = "E", "P", TBproducao!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & ValorTotalGrid & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & VermelhoGrid

        ProcNivel1Ordem TBproducao!Ordem
        
        TBproducao.MoveNext
        contador = contador + 1
        PBLista1.Value = contador
    Loop
    
    With Grid1
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 22
        .rows = m_Row
        .Cell(0, 1).Text = "Cód. interno"
        .Cell(0, 2).Text = "Pos."
        .Cell(0, 3).Text = "Cód. de ref."
        .Cell(0, 4).Text = "Descrição"
        .Cell(0, 5).Text = "Ordem"
        .Cell(0, 6).Text = "Tipo"
        .Cell(0, 7).Text = "Qtde."
        .Cell(0, 8).Text = "Qtde. prod."
        .Cell(0, 9).Text = "Qtde. entrada"
        .Cell(0, 10).Text = "Tempo real"
        .Cell(0, 11).Text = "CT. prev."
        .Cell(0, 12).Text = "CT. real"
        .Cell(0, 13).Text = "CT. material"
        .Cell(0, 14).Text = "CT. terceiros"
        .Cell(0, 15).Text = "CT. outras desp."
        .Cell(0, 16).Text = "CT. total"
        .Cell(0, 17).Text = "Nota"
        .Cell(0, 18).Text = "Vlr. unit."
        .Cell(0, 19).Text = "Qtde. requisit."
        .Cell(0, 20).Text = "Qtde. saida"
        .Cell(0, 21).Text = "Vlr. total"
        .Range(0, 1, 0, 21).Alignment = cellCenterCenter
        .Column(1).Width = 150
        .Column(2).Width = 30
        .Column(3).Width = 100
        .Column(4).Width = 250
        .Column(5).Width = 80
        .Column(6).Width = 30
        .Column(7).Width = 90
        .Column(7).Alignment = cellRightCenter
        .Column(8).Width = 90
        .Column(8).Alignment = cellRightCenter
        .Column(9).Width = 90
        .Column(9).Alignment = cellRightCenter
        .Column(10).Width = 90
        .Column(10).Alignment = cellCenterCenter
        .Column(11).Width = 90
        .Column(11).Alignment = cellRightCenter
        .Column(12).Width = 90
        .Column(12).Alignment = cellRightCenter
        .Column(13).Width = 90
        .Column(13).Alignment = cellRightCenter
        .Column(14).Width = 90
        .Column(14).Alignment = cellRightCenter
        .Column(15).Width = 90
        .Column(15).Alignment = cellRightCenter
        .Column(16).Width = 90
        .Column(16).Alignment = cellRightCenter
        .Column(17).Width = 80
        .Column(18).Width = 90
        .Column(18).Alignment = cellRightCenter
        .Column(19).Width = 90
        .Column(19).Alignment = cellRightCenter
        .Column(20).Width = 90
        .Column(20).Alignment = cellRightCenter
        .Column(21).Width = 90
        .Column(21).Alignment = cellRightCenter

        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        '.AddItem arrNodes(0).Text este é o correto, estou usando o de baixo porque existe um campo no final que não quero mostrar
        .AddItem Left(arrNodes(0).Text, Len(arrNodes(0).Text) - 1)
        If Right(arrNodes(intIndex).Text, 1) = 1 Then
            .Cell(1, 1).BackColor = Red
            .Cell(1, 2).BackColor = Red
            .Cell(1, 3).BackColor = Red
            .Cell(1, 4).BackColor = Red
            .Cell(1, 5).BackColor = Red
            .Cell(1, 6).BackColor = Red
            .Cell(1, 7).BackColor = Red
            .Cell(1, 8).BackColor = Red
            .Cell(1, 9).BackColor = Red
            .Cell(1, 10).BackColor = Red
            .Cell(1, 11).BackColor = Red
            .Cell(1, 12).BackColor = Red
            .Cell(1, 13).BackColor = Red
            .Cell(1, 14).BackColor = Red
            .Cell(1, 15).BackColor = Red
            .Cell(1, 16).BackColor = Red
            .Cell(1, 17).BackColor = Red
            .Cell(1, 18).BackColor = Red
            .Cell(1, 19).BackColor = Red
            .Cell(1, 20).BackColor = Red
            .Cell(1, 21).BackColor = Red
        End If
        
        'Other nodes
        For intIndex = 1 To Contador1 'UBound(arrNodes)
            If arrNodes(intIndex).Level = arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Parent.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level > arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level < arrNodes(intIndex - 1).Level Then
                For i = arrNodes(intIndex).Level To arrNodes(intIndex - 1).Level
                    Set tempNode = tempNode.Parent
                Next
                Set tempNode = tempNode.Nodes.Add("")
            End If
            '.AddItem arrNodes(intIndex).Text
            
            .AddItem Left(arrNodes(intIndex).Text, Len(arrNodes(intIndex).Text) - 1)
            If Right(arrNodes(intIndex).Text, 1) = 1 Then
                .Cell(intIndex + 1, 1).BackColor = Red
                .Cell(intIndex + 1, 2).BackColor = Red
                .Cell(intIndex + 1, 3).BackColor = Red
                .Cell(intIndex + 1, 4).BackColor = Red
                .Cell(intIndex + 1, 5).BackColor = Red
                .Cell(intIndex + 1, 6).BackColor = Red
                .Cell(intIndex + 1, 7).BackColor = Red
                .Cell(intIndex + 1, 8).BackColor = Red
                .Cell(intIndex + 1, 9).BackColor = Red
                .Cell(intIndex + 1, 10).BackColor = Red
                .Cell(intIndex + 1, 11).BackColor = Red
                .Cell(intIndex + 1, 12).BackColor = Red
                .Cell(intIndex + 1, 13).BackColor = Red
                .Cell(intIndex + 1, 14).BackColor = Red
                .Cell(intIndex + 1, 15).BackColor = Red
                .Cell(intIndex + 1, 16).BackColor = Red
                .Cell(intIndex + 1, 17).BackColor = Red
                .Cell(intIndex + 1, 17).BackColor = Red
                .Cell(intIndex + 1, 18).BackColor = Red
                .Cell(intIndex + 1, 19).BackColor = Red
                .Cell(intIndex + 1, 20).BackColor = Red
                .Cell(intIndex + 1, 21).BackColor = Red
            End If
        Next
        
        .AutoRedraw = True
        .Refresh
    End With
    
End If
TBproducao.Close

txtCusto_previsto = Format(valor, "###,##0.00")

'Custo total
txtCusto_real = Format(ValorTotal, "###,##0.00")
Txt_custo_material = Format(Valor1, "###,##0.00")
Txt_custo_terceiros = Format(Valor2, "###,##0.00")
Txt_custo_outras = Format(Valor_INSS_Serv, "###,##0.00")
Txt_custo_total = Format(ValorTotal + Valor1 + Valor2 + Valor_INSS_Serv, "###,##0.00")

'Resultados por peça
If Qtde <> 0 Then
    Txt_mao_de_obra_peca = Format(ValorTotal / Qtde, "###,##0.0000000000")
    Txt_material_peca = Format(Valor1 / Qtde, "###,##0.0000000000")
    Txt_terceiros_peca = Format(Valor2 / Qtde, "###,##0.0000000000")
    Txt_outras_peca = Format(Valor_INSS_Serv / Qtde, "###,##0.0000000000")
Else
    Txt_mao_de_obra_peca = Format(0, "###,##0.0000000000")
    Txt_material_peca = Format(0, "###,##0.0000000000")
    Txt_terceiros_peca = Format(0, "###,##0.0000000000")
    Txt_outras_peca = Format(0, "###,##0.0000000000")
End If

'Resultados
Valor_Cofins_Prod = Txt_mao_de_obra_peca
Valor_Cofins_Serv = Txt_material_peca
Valor_CSLL_Prod = Txt_terceiros_peca
Valor_ICMS_SN = Txt_outras_peca
Qtde = txtQtd_vendida

Txt_mao_de_obra = Format(Valor_Cofins_Prod * Qtde, "###,##0.00")
Txt_material = Format(Valor_Cofins_Serv * Qtde, "###,##0.00")
Txt_terceiros = Format(Valor_CSLL_Prod * Qtde, "###,##0.00")
Txt_outras = Format(Valor_ICMS_SN * Qtde, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaProducao()
On Error GoTo tratar_erro

Valor3 = 0 'Mão de obra
Valor_Cofins_Prod = 0 'Mão de obra por peça
Valor_Produto = 0 'Material
Valor_Cofins_Serv = 0 'Material por peça
ValorPagar = 0 'Terceiros
Valor_CSLL_Prod = 0 'Terceiros por peça
Valor_DAS = 0 'Outras desp.
Valor_ICMS_SN = 0 'Outras desp. por peça
Qtde = 0
valor = 0
ValorTotal = 0
Valor1 = 0
Valor2 = 0
Valor_INSS_Serv = 0
lista_fabricacao.ListItems.Clear

Permitido = False
'INNERJOINTEXTO = "Select P.* FROM Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem where PP.IDCarteira = " & txtidcarteira & " and"
'Set TBproducao = CreateObject("adodb.recordset")
'TBproducao.Open INNERJOINTEXTO & " P.desenho = '" & txtCod_int & "' and P.Tipo = 'E' order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
'If TBproducao.EOF = False Then
'    Permitido = True
'Else
'    Set TBproducao = CreateObject("adodb.recordset")
'    TBproducao.Open INNERJOINTEXTO & " P.Tipo = 'E' order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
'    If TBproducao.EOF = False Then
'        Permitido = True
'    Else
'        Set TBproducao = CreateObject("adodb.recordset")
'        TBproducao.Open INNERJOINTEXTO & " P.Tipo <> 'E' order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
'        If TBproducao.EOF = False Then
'            Permitido = True
'        Else
'            Set TBproducao = CreateObject("adodb.recordset")
'            TBproducao.Open "Select P.* FROM ((Producao P INNER JOIN Ordens_texto_SA OT ON OT.Ordem = P.Ordem) INNER JOIN Estoque_movimentacao EM ON EM.Lote = OT.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = EM.ID_prod_NF where NFPP.ID_carteira = " & txtidcarteira & " and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL') order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
'            If TBproducao.EOF = False Then
'                Permitido = True
'            End If
'        End If
'    End If
'End If
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select P.* FROM Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem where PP.IDCarteira = " & txtidcarteira & " order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
'TBproducao.Open "Select P.* FROM Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem where PP.IDCarteira = " & txtIdcarteira & " AND P.desenho = '" & txtCod_int & "' order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Permitido = True
Else
    'Verifica ordem do estoque que não tem vinculado o ID da carteira na ordem baixada do estoque
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select P.* FROM ((Producao P INNER JOIN Ordens_texto_SA OT ON OT.Ordem = P.Ordem) INNER JOIN Estoque_movimentacao EM ON EM.Lote = OT.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = EM.ID_prod_NF where NFPP.ID_carteira = " & txtidcarteira & " and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL') order by P.Ordem", Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        Permitido = True
    End If
End If
If Permitido = True Then
    PBLista1.Min = 0
    PBLista1.Max = TBproducao.RecordCount
    PBLista1.Value = 1
    contador = 0
    Do While TBproducao.EOF = False
        With lista_fabricacao.ListItems.Add(, , TBproducao!NOF)
            .SubItems(1) = IIf(IsNull(TBproducao!Desenho), "", TBproducao!Desenho)
            .SubItems(2) = IIf(IsNull(TBproducao!N_referencia), "", TBproducao!N_referencia)
            .SubItems(3) = IIf(IsNull(TBproducao!Produto), "", TBproducao!Produto)
            .SubItems(4) = IIf(IsNull(TBproducao!Ordem), "", TBproducao!Ordem)
            .SubItems(5) = IIf(IsNull(TBproducao!Tipo), "", IIf(TBproducao!Tipo = "E", "P", TBproducao!Tipo))
            .SubItems(6) = IIf(IsNull(TBproducao!Quant), "0,0000", Format(TBproducao!Quant, "###,##0.0000"))
            .SubItems(7) = IIf(IsNull(TBproducao!QuantProd), "0,0000", Format(TBproducao!QuantProd, "###,##0.0000"))
            .SubItems(8) = IIf(IsNull(TBproducao!TTTReal), "00:00:00", Format(TBproducao!TTTReal, "hh:mm:ss"))
            .SubItems(9) = IIf(IsNull(TBproducao!CTTPrev), "0,00", Format(TBproducao!CTTPrev, "###,##0.00"))
            .SubItems(10) = IIf(IsNull(TBproducao!CTTReal), "0,00", Format(TBproducao!CTTReal, "###,##0.00"))
            .SubItems(11) = IIf(IsNull(TBproducao!CTMaterial), "0,00", Format(TBproducao!CTMaterial, "###,##0.00"))
            .SubItems(12) = IIf(IsNull(TBproducao!CTServico), "0,00", Format(TBproducao!CTServico, "###,##0.00"))
            .SubItems(13) = IIf(IsNull(TBproducao!CTOutras), "0,00", Format(TBproducao!CTOutras, "###,##0.00"))
            
            If TBproducao!Desenho = txtCod_int Then
                Qtde = Qtde + IIf(IIf(IsNull(TBproducao!QuantProd), 0, TBproducao!QuantProd) = 0, TBproducao!Quant, TBproducao!QuantProd)
                valor = valor + IIf(IsNull(TBproducao!CTTPrev), 0, TBproducao!CTTPrev) 'Custo previsto
                
                'Custo total
                Valor3 = IIf(IsNull(TBproducao!CTTReal), 0, TBproducao!CTTReal)
                Valor_Produto = IIf(IsNull(TBproducao!CTMaterial), 0, TBproducao!CTMaterial)
                ValorPagar = IIf(IsNull(TBproducao!CTServico), 0, TBproducao!CTServico)
                Valor_DAS = IIf(IsNull(TBproducao!CTOutras), 0, TBproducao!CTOutras)
                .SubItems(14) = Format(Valor3 + Valor_Produto + ValorPagar + Valor_DAS, "###,##0.00")
                
                ValorTotal = ValorTotal + Valor3
                Valor1 = Valor1 + Valor_Produto
                Valor2 = Valor2 + ValorPagar
                Valor_INSS_Serv = Valor_INSS_Serv + Valor_DAS
                
                .ForeColor = vbBlue
                .ListSubItems(1).ForeColor = vbBlue
                .ListSubItems(2).ForeColor = vbBlue
                .ListSubItems(3).ForeColor = vbBlue
                .ListSubItems(4).ForeColor = vbBlue
                .ListSubItems(5).ForeColor = vbBlue
                .ListSubItems(6).ForeColor = vbBlue
                .ListSubItems(7).ForeColor = vbBlue
                .ListSubItems(8).ForeColor = vbBlue
                .ListSubItems(9).ForeColor = vbBlue
                .ListSubItems(10).ForeColor = vbBlue
                .ListSubItems(11).ForeColor = vbBlue
                .ListSubItems(12).ForeColor = vbBlue
                .ListSubItems(13).ForeColor = vbBlue
            End If
        End With
        TBproducao.MoveNext
        contador = contador + 1
        PBLista1.Value = contador
    Loop
End If
TBproducao.Close

txtCusto_previsto = Format(valor, "###,##0.00")

'Custo total
txtCusto_real = Format(ValorTotal, "###,##0.00")
Txt_custo_material = Format(Valor1, "###,##0.00")
Txt_custo_terceiros = Format(Valor2, "###,##0.00")
Txt_custo_outras = Format(Valor_INSS_Serv, "###,##0.00")
Txt_custo_total = Format(ValorTotal + Valor1 + Valor2 + Valor_INSS_Serv, "###,##0.00")

'Resultados por peça
If Qtde <> 0 Then
    Txt_mao_de_obra_peca = Format(ValorTotal / Qtde, "###,##0.0000000000")
    Txt_material_peca = Format(Valor1 / Qtde, "###,##0.0000000000")
    Txt_terceiros_peca = Format(Valor2 / Qtde, "###,##0.0000000000")
    Txt_outras_peca = Format(Valor_INSS_Serv / Qtde, "###,##0.0000000000")
Else
    Txt_mao_de_obra_peca = Format(0, "###,##0.0000000000")
    Txt_material_peca = Format(0, "###,##0.0000000000")
    Txt_terceiros_peca = Format(0, "###,##0.0000000000")
    Txt_outras_peca = Format(0, "###,##0.0000000000")
End If

'Resultados
Valor_Cofins_Prod = Txt_mao_de_obra_peca
Valor_Cofins_Serv = Txt_material_peca
Valor_CSLL_Prod = Txt_terceiros_peca
Valor_ICMS_SN = Txt_outras_peca
Qtde = txtQtd_vendida

Txt_mao_de_obra = Format(Valor_Cofins_Prod * Qtde, "###,##0.00")
Txt_material = Format(Valor_Cofins_Serv * Qtde, "###,##0.00")
Txt_terceiros = Format(Valor_CSLL_Prod * Qtde, "###,##0.00")
Txt_outras = Format(Valor_ICMS_SN * Qtde, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaResultados()
On Error GoTo tratar_erro

Lista_resultados.ListItems.Clear
'Verifica no pedido
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPL.Familia, Sum(CPL.preco_total) as Valores FROM (compras_pedido_lista CPL INNER JOIN Producao P ON CPL.Ordem = P.Ordem) INNER JOIN Producao_pedidos PP ON CPL.Ordem = PP.Ordem where PP.IDCarteira = " & txtidcarteira & " and CPL.Remessa = 'False' and (CPL.OS IS NULL or CPL.OS = 0) and (CPL.status_item = 'N_RECEBIDO' or CPL.status_item = 'RECEBIDO' or CPL.status_item = 'PARCIAL') group by CPL.familia", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista_resultados.ListItems
            .Add , , TBLISTA!Familia
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Valores), 0, Format(TBLISTA!Valores, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
'Verifica no pedido (Empenho)
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPL.Familia, Sum(ROUND(CPL.preco_unitario_desconto * CPLE.Qtde_empenho,2)) as Valores FROM (compras_pedido_lista CPL INNER JOIN Compras_pedido_lista_empenhos CPLE ON CPLE.IDlista = CPL.IDLista) INNER JOIN Projproduto P ON P.Desenho = CPL.Desenho where CPLE.IDCarteira = " & txtidcarteira & " and P.estoque = 'False' and CPL.Remessa = 'False' and (CPL.OS IS NULL or CPL.OS = 0) group by CPL.familia", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista_resultados.ListItems
            .Add , , TBLISTA!Familia
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Valores), 0, Format(TBLISTA!Valores, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

'Verifica no estoque
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PROD.Classe, Sum(E.VlrTotal) as Valores, Sum(P.Quant) as Qt, Sum(P.Quantprod) as Qtde FROM (((Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem) INNER JOIN Producaomaterial PM ON PM.Ordem = P.Ordem) INNER JOIN Projproduto PROD ON PROD.Desenho = PM.Codigo) INNER JOIN Estoque_movimentacao E ON E.Ordem = P.Ordem and E.Familia = PROD.Classe where PP.IDCarteira = " & txtidcarteira & " and (PM.Saida = 'SIM' or PM.Saida = 'NÃO' or PM.Saida = 'PARCIAL') and (E.Operacao = 'SAIDA_ORDEM' or E.Operacao = 'SAIDA_ORDEM_PARCIAL') group by PROD.Classe", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista_resultados.ListItems
            .Add , , TBLISTA!Classe
            
            Valor1 = txtQtd_vendida
            If IsNull(TBLISTA!Qtde) = False And TBLISTA!Qtde <> 0 Then Quant = TBLISTA!Qtde Else Quant = TBLISTA!qt
            valor = (IIf(IsNull(TBLISTA!Valores), 0, TBLISTA!Valores) / Quant) * Valor1
            
            .Item(.Count).SubItems(1) = Format(valor, "###,##0.00")
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotais()
On Error GoTo tratar_erro

Valor3 = 0
Valores = 0
Valor1 = 0
valor = 0
Valor_DAS = 0
ValorICMS = 0
Valor2 = 0
Total = 0

'Por peça
Valor3 = Txt_compras_peca
Valores = Txt_mao_de_obra_peca
Valor1 = Txt_material_peca
valor = Txt_terceiros_peca
Valor_DAS = Txt_outras_peca
ValorICMS = Txt_impostos_peca
Txt_total_peca = Format(Valor3 + Valores + Valor1 + valor + Valor_DAS, "###,##0.0000000000")

ValorTotal = Valor3 + Valores + Valor1 + valor + Valor_DAS + ValorICMS
Txt_custo_acumulado_peca = Format(ValorTotal, "###,##0.0000000000")

Valores = Txt_valor_venda_peca
Total = Valores - ValorTotal
Txt_resultado_peca = Format(Total, "###,##0.0000000000")
If Valores <> 0 Then Txt_resultado_peca_porcento = Format(((Total / Valores) * 100) * 1, "###,##0.00") & "%" Else Txt_resultado_peca_porcento = "0,00" & "%"

If Total < 0 Then
    'Peça
    Txt_resultado_peca.ForeColor = &HFF&
    label_peca.Caption = "(=) Prejuízo"
        
    Txt_resultado_peca_porcento.ForeColor = &HFF&
    label_peca_porcento.Caption = "(=) Prejuízo"
        
    'Total
    Txt_resultado.ForeColor = &HFF&
    label_resultado.Caption = "(=) Prejuízo"
    
    Txt_resultado_porcento.ForeColor = &HFF&
    label_resultado_porcento.Caption = "(=) Prejuízo"
ElseIf Total > 0 Then
        'Peça
        Txt_resultado_peca.ForeColor = &HFF0000
        label_peca.Caption = "(=) Lucro"
                
        Txt_resultado_peca_porcento.ForeColor = &HFF0000
        label_peca_porcento.Caption = "(=) Lucro"
        
        'Total
        Txt_resultado.ForeColor = &HFF0000
        label_resultado.Caption = "(=) Lucro"
        
        Txt_resultado_porcento.ForeColor = &HFF0000
        label_resultado_porcento.Caption = "(=) Lucro"
    Else
        'Peça
        Txt_resultado_peca.ForeColor = &H0&
        label_peca.Caption = "(=) Empate"
        
        Txt_resultado_peca_porcento.ForeColor = &H0&
        label_peca_porcento.Caption = "(=) Empate"
        
        'Total
        Txt_resultado.ForeColor = &H0&
        label_resultado.Caption = "(=) Empate"
        
        Txt_resultado_porcento.ForeColor = &H0&
        label_resultado_porcento.Caption = "(=) Empate"
End If

'Total
Valor3 = Txt_compras
Valores = Txt_mao_de_obra
Valor1 = Txt_material
valor = Txt_terceiros
Valor_DAS = Txt_outras
ValorICMS = Txt_impostos
Txt_total = Format(Valor3 + Valores + Valor1 + valor + Valor_DAS, "###,##0.0000000000")

ValorTotal = Valor3 + Valores + Valor1 + valor + Valor_DAS + ValorICMS
Txt_custo_acumulado = Format(ValorTotal, "###,##0.00")

Valores = Txt_valor_venda
Total = Valores - ValorTotal
Txt_resultado = Format(Total, "###,##0.00")
If Valores <> 0 Then Txt_resultado_porcento = Format(((Total / Valores) * 100) * 1, "###,##0.00") & "%" Else Txt_resultado_porcento = "0,00" & "%"

If Total < 0 Then
    Txt_resultado.ForeColor = &HFF&
    label_resultado.Caption = "(=) Prejuízo"
    
    Txt_resultado_porcento.ForeColor = &HFF&
    label_resultado_porcento.Caption = "(=) Prejuízo"
ElseIf Total > 0 Then
        Txt_resultado.ForeColor = &HFF0000
        label_resultado.Caption = "(=) Lucro"
        
        Txt_resultado_porcento.ForeColor = &HFF0000
        label_resultado_porcento.Caption = "(=) Lucro"
    Else
        Txt_resultado.ForeColor = &H0&
        label_resultado.Caption = "(=) Empate"
        
        Txt_resultado_porcento.ForeColor = &H0&
        label_resultado_porcento.Caption = "(=) Empate"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista_filtro.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista_filtro.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista_filtro.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab > 0 And (Lista_filtro.ListItems.Count = 0 Or txtidcarteira = 0) Then
    SSTab1.Tab = 0
    Exit Sub
End If
If SSTab1.Tab = 0 Or SSTab1.Tab = 4 Or SSTab1.Tab = 5 Then
    PBLista.Visible = True
    PBLista1.Visible = False
ElseIf SSTab1.Tab = 3 Then
    PBLista.Visible = False
    PBLista1.Visible = False
Else
    PBLista.Visible = False
    PBLista1.Visible = True
End If

USToolBar1.Visible = True
Frame7.Visible = True
Frame1.Visible = True
Frame5.Visible = False
Frame6.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Select Case SSTab1.Tab
    Case 0:
        USToolBar1.Visible = False
        Frame7.Visible = False
        Frame1.Visible = False
        If txtTexto.Visible = True Then txtTexto.SetFocus
    Case 1:
        Lista_compras.SetFocus
        ProcCarregaListaCompras
    Case 2:
        If chkNivel.Value = 1 Then ProcCarregaListaProducaoGrid Else ProcCarregaListaProducao
    Case 3:
        Frame5.Visible = True
        Frame6.Visible = True
        Frame4.Visible = True
        Frame3.Visible = True
        Cmd_abrir_nf.SetFocus
        ProcCarregaTotais
    Case 4:
        Lista_resultados.SetFocus
        ProcCarregaListaResultados
    Case 5:
        Lista_OS.SetFocus
        ProcCarregaListaOS
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaOS()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatoriosTotal

Permitido = False
Lista_OS.ListItems.Clear
INNERJOINTEXTO = "Select OS.*, CM.Descricao FROM ((Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem) INNER JOIN ordemservico OS ON OS.Ordem = P.Ordem) INNER JOIN CadMaquinas CM ON OS.Maquina = CM.Maquina where PP.IDCarteira = " & txtidcarteira & " and"
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open INNERJOINTEXTO & " P.desenho = '" & txtCod_int & "' and P.Tipo = 'E' order by OS.Ordem, OS.fase", Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    Permitido = True
Else
    Set TBOrdem = CreateObject("adodb.recordset")
    TBOrdem.Open INNERJOINTEXTO & " P.Tipo = 'E' order by OS.Ordem, OS.fase", Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdem.EOF = False Then
        Permitido = True
    Else
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open INNERJOINTEXTO & " P.Tipo <> 'E' order by OS.Ordem, OS.fase", Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            Permitido = True
        Else
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select OS.*, CM.Descricao FROM ((((Producao P INNER JOIN Ordens_texto_SA OT ON OT.Ordem = P.Ordem) INNER JOIN Estoque_movimentacao EM ON EM.Lote = OT.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = EM.ID_prod_NF) INNER JOIN ordemservico OS ON OS.Ordem = P.Ordem) INNER JOIN CadMaquinas CM ON OS.Maquina = CM.Maquina where NFPP.ID_carteira = " & txtidcarteira & " and (EM.Operacao = 'SAIDA_NOTA' or EM.Operacao = 'SAIDA_NOTA_PARCIAL') order by OS.Ordem, OS.fase", Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = False Then
                Permitido = True
            End If
        End If
    End If
End If
If Permitido = True Then
    PBLista.Min = 0
    PBLista.Max = TBOrdem.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBOrdem.EOF = False
        Set TBFIltro = CreateObject("adodb.recordset")
        CamposFiltro = "OS.Ordem, OS.IDProducao, OS.Fase, OS.Grupo_op, OS.maquina, OS.TempoPreparacao, OS.TPUTIL, OS.TempoExecucao, OS.TEUTIL, OS.QTOK, OS.QTNC, OS.Totalprod, OS.CPPECA, OS.CPLOTE, OS.TPUSEG, OS.Custos, OS.CRLOTE, OS.CTServico, OS.Eficiencia_prep, OS.Eficiencia_exec, OS.Eficiencia, OS.TempoPrepUtilPeca, OS.CTPrepPrevPeca, CM.Descricao, OS.Valor_hs_prep, OS.Valor_hs_exec, CM.Descricao"
        TBFIltro.Open "Select " & CamposFiltro & " from ordemservico OS INNER JOIN CadMaquinas CM on OS.Maquina = CM.Maquina where OS.Ordem = " & TBOrdem!Ordem & " order by OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Set TBFI = CreateObject("adodb.recordset")
            CamposFiltro = "OSMU.ID, OSMU.maquina, OSMU.TPUTIL, OSMU.TEUTIL, OSMU.QTOK, OSMU.QTNC, OSMU.Totalprod, OSMU.TPUSEG, OSMU.CRLOTE, OSMU.CRPECA, OSMU.Eficiencia_prep, OSMU.Eficiencia_exec, OSMU.Eficiencia, OSMU.TempoPrepUtilPeca, OSMU.CTPrepPrevPeca, CM.Descricao"
            TBFI.Open "Select " & CamposFiltro & "  from Ordemservico_maq_utilizadas OSMU INNER JOIN CadMaquinas CM on OSMU.Maquina = CM.Maquina where OSMU.OS = " & TBOrdem!IDProducao & " order by OSMU.ID", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    With Lista_OS.ListItems
                        .Add , , TBOrdem!IDProducao
                        .Item(.Count).SubItems(1) = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
                        .Item(.Count).SubItems(2) = IIf(IsNull(TBOrdem!Grupo_op), "", TBOrdem!Grupo_op)
                        .Item(.Count).SubItems(3) = IIf(IsNull(TBFI!maquina), "", TBFI!maquina)
                        .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!Descricao), "", TBFI!Descricao)
                        .Item(.Count).SubItems(5) = IIf(IsNull(TBOrdem!Valor_hs_prep), "", Format(TBOrdem!Valor_hs_prep, "###,##0.00"))
                        .Item(.Count).SubItems(6) = IIf(IsNull(TBOrdem!TempoPreparacao), "00:00:00", Format(TBOrdem!TempoPreparacao, "hh:mm:ss"))
                        
                        'Tempo de preparação real p/ peça
                        TPUSEG = IIf(IsNull(TBFI!TPUSEG), 0, TBFI!TPUSEG)
                        s = TPUSEG
                        .Item(.Count).SubItems(7) = FormataTempo(s)
                         
                        .Item(.Count).SubItems(8) = IIf(IsNull(TBOrdem!Valor_hs_exec), "", Format(TBOrdem!Valor_hs_exec, "###,##0.00"))
                        .Item(.Count).SubItems(9) = IIf(IsNull(TBOrdem!TempoExecucao), "00:00:00", Format(TBOrdem!TempoExecucao, "hh:mm:ss"))
                         
                        'Tempo de execucao real p/ peça
                        .Item(.Count).SubItems(10) = IIf(IsNull(TBFI!TEUTIL), "00:00:00", Format(TBFI!TEUTIL, "hh:mm:ss"))
                         
                        .Item(.Count).SubItems(11) = IIf(IsNull(TBFI!QTOK), "", Format(TBFI!QTOK, "###,##0.00"))
                        .Item(.Count).SubItems(12) = IIf(IsNull(TBFI!QTNC), "", Format(TBFI!QTNC, "###,##0.00"))
                        .Item(.Count).SubItems(13) = IIf(IsNull(TBFI!Totalprod), "", Format(TBFI!Totalprod, "###,##0.00"))
                                 
                        valor = IIf(IsNull(TBOrdem!Valor_hs_prep), 0, TBOrdem!Valor_hs_prep)
                        valor = valor / 3600
                                     
                        'Tempo de preparação previsto por peça
                        If TBOrdem!TempoPreparacao <= "23:59:59" Then
                            TempoTotalPrep = IIf(IsNull(TBOrdem!TempoPreparacao), 0, TBOrdem!TempoPreparacao)
                            ElapsedTime (TempoTotalPrep)
                        Else
                            ProcFormataHora (IIf(IsNull(TBOrdem!TempoPreparacao), 0, TBOrdem!TempoPreparacao))
                        End If
                                     
                        'Custo de preparação previsto por peça
                        .Item(.Count).SubItems(14) = Format(s * valor, "###,##0.00")
                        
                        'Custo previsto por peça x qtde. total prod. por máquina
                        .Item(.Count).SubItems(15) = Format(IIf(IsNull(TBOrdem!CPPECA), "", TBOrdem!CPPECA) * IIf(IsNull(TBFI!Totalprod), "", TBFI!Totalprod), "###,##0.00")
                        
                        'Tempo de preparação real do lote
                        quantidade = IIf(IsNull(TBFI!TPUSEG), 0, TBFI!TPUSEG)
                        'Custo de preparação real do lote
                        Valor1 = Format(quantidade * valor, "###,##0.00")
                        .Item(.Count).SubItems(16) = Format(Valor1, "###,##0.00")
                        
                        'Custo de preparação real do lote
                        .Item(.Count).SubItems(17) = Format(IIf(IsNull(TBFI!CRLOTE), 0, TBFI!CRLOTE) - Valor1, "###,##0.00")
                        
                        'Custo de execução real do lote
                        If TBOrdem!custos = True Then
                            .Item(.Count).SubItems(18) = IIf(IsNull(TBFI!CRLOTE), "", Format(TBFI!CRLOTE, "###,##0.00"))
                        Else
                            .Item(.Count).SubItems(18) = IIf(IsNull(TBOrdem!CTServico), "", Format(TBOrdem!CTServico, "###,##0.00"))
                        End If
                        .Item(.Count).SubItems(19) = IIf(IsNull(TBFI!Eficiencia_prep), "", TBFI!Eficiencia_prep & " %")
                        .Item(.Count).SubItems(20) = IIf(IsNull(TBFI!Eficiencia_exec), "", TBFI!Eficiencia_exec & " %")
                        .Item(.Count).SubItems(21) = IIf(IsNull(TBFI!Eficiencia), "", TBFI!Eficiencia & " %")
                    End With
                    TBFI.MoveNext
                Loop
            Else
                With Lista_OS.ListItems
                    .Add , , TBOrdem!IDProducao
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBOrdem!Grupo_op), "", TBOrdem!Grupo_op)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBOrdem!maquina), "", TBOrdem!maquina)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBOrdem!Descricao), "", TBOrdem!Descricao)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBOrdem!Valor_hs_prep), "", Format(TBOrdem!Valor_hs_prep, "###,##0.00"))
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBOrdem!TempoPreparacao), "00:00:00", Format(TBOrdem!TempoPreparacao, "hh:mm:ss"))
                    
                    'Tempo de preparação real p/ peça
                    TPUSEG = IIf(IsNull(TBOrdem!TPUSEG), 0, TBOrdem!TPUSEG)
                    s = TPUSEG
                    .Item(.Count).SubItems(7) = FormataTempo(s)
                     
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBOrdem!Valor_hs_exec), "", Format(TBOrdem!Valor_hs_exec, "###,##0.00"))
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBOrdem!TempoExecucao), "00:00:00", Format(TBOrdem!TempoExecucao, "hh:mm:ss"))
                     
                    'Tempo de execucao real p/ peça
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBOrdem!TEUTIL), "00:00:00", Format(TBOrdem!TEUTIL, "hh:mm:ss"))
                     
                    .Item(.Count).SubItems(11) = IIf(IsNull(TBOrdem!QTOK), "", Format(TBOrdem!QTOK, "###,##0.00"))
                    .Item(.Count).SubItems(12) = IIf(IsNull(TBOrdem!QTNC), "", Format(TBOrdem!QTNC, "###,##0.00"))
                    .Item(.Count).SubItems(13) = IIf(IsNull(TBOrdem!Totalprod), "", Format(TBOrdem!Totalprod, "###,##0.00"))
                             
                    valor = IIf(IsNull(TBOrdem!Valor_hs_prep), 0, TBOrdem!Valor_hs_prep)
                    valor = valor / 3600
                                 
                    'Tempo de preparação previsto por peça
                    If TBOrdem!TempoPreparacao <= "23:59:59" Then
                        TempoTotalPrep = IIf(IsNull(TBOrdem!TempoPreparacao), 0, TBOrdem!TempoPreparacao)
                        ElapsedTime (TempoTotalPrep)
                    Else
                        ProcFormataHora (IIf(IsNull(TBOrdem!TempoPreparacao), 0, TBOrdem!TempoPreparacao))
                    End If
                                 
                    'Custo de preparação previsto por peça
                    .Item(.Count).SubItems(14) = Format(s * valor, "###,##0.00")
                     
                    'Custo previsto por peça x qtde. total prod. por máquina
                    .Item(.Count).SubItems(15) = Format(IIf(IsNull(TBOrdem!CPLOTE), "", TBOrdem!CPLOTE), "###,##0.00")
                     
                    'Tempo de preparação real do lote
                    quantidade = IIf(IsNull(TBOrdem!TPUSEG), 0, TBOrdem!TPUSEG)
                    'Custo de preparação real do lote
                    .Item(.Count).SubItems(16) = Format(quantidade * valor, "###,##0.00")
                    Valor1 = quantidade * valor
                     
                    'Custo de preparação real do lote
                    .Item(.Count).SubItems(17) = Format(IIf(IsNull(TBOrdem!CRLOTE), 0, TBOrdem!CRLOTE) - Valor1, "###,##0.00")
                     
                    'Custo de execução real do lote
                    If TBOrdem!custos = True Then
                        .Item(.Count).SubItems(18) = IIf(IsNull(TBOrdem!CRLOTE), "", Format(TBOrdem!CRLOTE, "###,##0.00"))
                    Else
                        .Item(.Count).SubItems(18) = IIf(IsNull(TBOrdem!CTServico), "", Format(TBOrdem!CTServico, "###,##0.00"))
                    End If
                    .Item(.Count).SubItems(19) = IIf(IsNull(TBOrdem!Eficiencia_prep), "", TBOrdem!Eficiencia_prep & " %")
                    .Item(.Count).SubItems(20) = IIf(IsNull(TBOrdem!Eficiencia_exec), "", TBOrdem!Eficiencia_exec & " %")
                    .Item(.Count).SubItems(21) = IIf(IsNull(TBOrdem!Eficiencia), "", TBOrdem!Eficiencia & " %")
                End With
            End If
            TBFI.Close
        End If
        TBFIltro.Close
        TBOrdem.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
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

Lista_filtro.ListItems.Clear
If txtTexto <> "" Then cmbfamilia.ListIndex = -1
If cmbfiltrarpor = "Ordem" And txtTexto <> "" Then
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNivel1Ordem(OrdemGrid As Long)
On Error GoTo tratar_erro

Set TBNivel1 = CreateObject("adodb.recordset")
TBNivel1.Open "Select * from Producaomaterial where ordem = " & OrdemGrid & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockReadOnly
If TBNivel1.EOF = False Then
    Do While TBNivel1.EOF = False
        PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel1!Posicao, 3)
        
        CodRef = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select I.n_referencia from item_aplicacoes I INNER JOIN projproduto P ON I.codproduto = P.codproduto where P.desenho = '" & TBNivel1!CODIGO & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            CodRef = TBItem!N_referencia
        End If
        TBItem.Close
        
        NFGrid = ""
        VlrunitGrid = ""
        SaidaGrid = ""
        VlrTotalGrid = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "SELECT * from RE_Entrada_RM_Ordem where desenho = '" & TBNivel1!CODIGO & "' and ordem = " & TBNivel1!Ordem, Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            NFGrid = IIf(IsNull(TBItem!Nota_fiscal), "", TBItem!Nota_fiscal)
            VlrunitGrid = IIf(IsNull(TBItem!VlrUnit), "", Format(TBItem!VlrUnit, "###,##0.0000"))
            VlrTotalGrid = IIf(IsNull(TBItem!VlrTotal), "", Format(TBItem!VlrTotal, "###,##0.00"))
        End If
        TBItem.Close
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Saida) as Saida from estoque_movimentacao where oe = '" & OrdemGrid & "' and desenho = '" & TBNivel1!CODIGO & "' and documento = '" & OrdemGrid & "' AND operacao IN ('SAIDA_ORDEM', 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Saida), 0, Format(TBItem!Saida, "0.000"))
        End If
        TBItem.Close
        
        'verifica se tem diferença entre qtde de saida e requisitada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBNivel1!Requisitado), 0, Format(TBNivel1!Requisitado, "0.000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        SaidaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        RequisitGrid = IIf(Qtd = 0, "", Format(Qtd, "###,##0.000"))
        
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 1
        arrNodes(Contador1).Text = TBNivel1!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel1!Descricao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & vbTab & VlrTotalGrid & VermelhoGrid
        
        'Verifica se existe uma ordem abaixo
        Set TBOrdem1 = CreateObject("adodb.recordset")
        TBOrdem1.Open "SELECT Lote FROM CustosProducao_OrdensAbaixo WHERE OE = '" & OrdemGrid & "' AND desenho = '" & TBNivel1!CODIGO & "' AND documento = '" & OrdemGrid & "' ORDER BY oe", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem1.EOF = False Then
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "SELECT P.Ordem, Tipo, Quant, QuantProd, TTTReal, CTTPrev, CTTReal, CTMaterial, CTServico, CTOutras FROM Producao P INNER JOIN Producao_Pedidos PP ON P.Ordem = PP.Ordem WHERE P.Ordem = " & TBOrdem1!LOTE & " AND PP.IDcarteira = " & txtidcarteira, Conexao, adOpenKeyset, adLockReadOnly
            If TBCFOP.EOF = False Then
                qtdeGrid = IIf(IsNull(TBCFOP!Quant), "", Format(TBCFOP!Quant, "###,##0.0000"))
                qtdeProdGrid = IIf(IsNull(TBCFOP!QuantProd), "", Format(TBCFOP!QuantProd, "###,##0.0000"))
                TTTRealGrid = IIf(IsNull(TBCFOP!TTTReal), "", Format(TBCFOP!TTTReal, "hh:mm:ss"))
                CTTPrevGrid = IIf(IsNull(TBCFOP!CTTPrev), "", Format(TBCFOP!CTTPrev, "###,##0.00"))
                CTTRealGrid = IIf(IsNull(TBCFOP!CTTReal), "", Format(TBCFOP!CTTReal, "###,##0.00"))
                CTMaterialGrid = IIf(IsNull(TBCFOP!CTMaterial), "", Format(TBCFOP!CTMaterial, "###,##0.00"))
                CTServicoGrid = IIf(IsNull(TBCFOP!CTServico), "", Format(TBCFOP!CTServico, "###,##0.00"))
                CTOutrasGrid = IIf(IsNull(TBCFOP!CTOutras), "", Format(TBCFOP!CTOutras, "###,##0.00"))
                
                Qtde = 0
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBCFOP!Ordem & "' and desenho = '" & TBNivel1!CODIGO & "' and documento = '" & TBCFOP!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBItem.EOF = False Then
                    Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
                End If
                TBItem.Close
                
                QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
                
                'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
                If VermelhoGrid = 0 Then 'Se já tiver pra ficar vermelhor não precisa ver de novo
                    Qtd = IIf(IsNull(TBCFOP!QuantProd), 0, Format(TBCFOP!QuantProd, "0.0000"))
                    If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
                End If

                arrNodes(Contador1).Text = TBNivel1!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel1!Descricao & vbTab & TBCFOP!Ordem & vbTab & IIf(TBCFOP!Tipo = "E", "P", TBCFOP!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & VermelhoGrid
                ProcNivel2Ordem TBCFOP!Ordem
            End If
            TBCFOP.Close
        End If
        TBOrdem1.Close
            
        TBNivel1.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcNivel2Ordem(OrdemGrid As Long)
On Error GoTo tratar_erro

Set TBNivel2 = CreateObject("adodb.recordset")
TBNivel2.Open "Select * from Producaomaterial where ordem = " & OrdemGrid & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockReadOnly
If TBNivel2.EOF = False Then
    Do While TBNivel2.EOF = False
        PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel2!Posicao, 3)
        
        CodRef = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select I.n_referencia from item_aplicacoes I INNER JOIN projproduto P ON I.codproduto = P.codproduto where P.desenho = '" & TBNivel2!CODIGO & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            CodRef = TBItem!N_referencia
        End If
        TBItem.Close
        
        NFGrid = ""
        VlrunitGrid = ""
        SaidaGrid = ""
        VlrTotalGrid = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "SELECT * from RE_Entrada_RM_Ordem where desenho = '" & TBNivel2!CODIGO & "' and ordem = " & TBNivel2!Ordem, Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            NFGrid = IIf(IsNull(TBItem!Nota_fiscal), "", TBItem!Nota_fiscal)
            VlrunitGrid = IIf(IsNull(TBItem!VlrUnit), "", Format(TBItem!VlrUnit, "###,##0.0000"))
            VlrTotalGrid = IIf(IsNull(TBItem!VlrTotal), "", Format(TBItem!VlrTotal, "###,##0.00"))
        End If
        TBItem.Close
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Saida) as Saida from estoque_movimentacao where oe = '" & OrdemGrid & "' and desenho = '" & TBNivel2!CODIGO & "' and documento = '" & OrdemGrid & "' AND operacao IN ('SAIDA_ORDEM', 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Saida), 0, Format(TBItem!Saida, "0.000"))
        End If
        TBItem.Close
        
        'verifica se tem diferença entre qtde de saida e requisitada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBNivel2!Requisitado), 0, Format(TBNivel2!Requisitado, "0.000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        SaidaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        RequisitGrid = IIf(Qtd = 0, "", Format(Qtd, "###,##0.000"))
        
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 2
        arrNodes(Contador1).Text = TBNivel2!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel2!Descricao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & vbTab & VlrTotalGrid & VermelhoGrid
        
        'Verifica se existe uma ordem abaixo
        Set TBOrdem2 = CreateObject("adodb.recordset")
        TBOrdem2.Open "SELECT Lote FROM CustosProducao_OrdensAbaixo WHERE OE = '" & OrdemGrid & "' AND desenho = '" & TBNivel2!CODIGO & "' AND documento = '" & OrdemGrid & "' ORDER BY oe", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem2.EOF = False Then
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "SELECT P.Ordem, Tipo, Quant, QuantProd, TTTReal, CTTPrev, CTTReal, CTMaterial, CTServico, CTOutras FROM Producao P INNER JOIN Producao_Pedidos PP ON P.Ordem = PP.Ordem WHERE P.Ordem = " & TBOrdem2!LOTE & " AND PP.IDcarteira = " & txtidcarteira, Conexao, adOpenKeyset, adLockReadOnly
            If TBCFOP.EOF = False Then
                qtdeGrid = IIf(IsNull(TBCFOP!Quant), "", Format(TBCFOP!Quant, "###,##0.0000"))
                qtdeProdGrid = IIf(IsNull(TBCFOP!QuantProd), "", Format(TBCFOP!QuantProd, "###,##0.0000"))
                TTTRealGrid = IIf(IsNull(TBCFOP!TTTReal), "", Format(TBCFOP!TTTReal, "hh:mm:ss"))
                CTTPrevGrid = IIf(IsNull(TBCFOP!CTTPrev), "", Format(TBCFOP!CTTPrev, "###,##0.00"))
                CTTRealGrid = IIf(IsNull(TBCFOP!CTTReal), "", Format(TBCFOP!CTTReal, "###,##0.00"))
                CTMaterialGrid = IIf(IsNull(TBCFOP!CTMaterial), "", Format(TBCFOP!CTMaterial, "###,##0.00"))
                CTServicoGrid = IIf(IsNull(TBCFOP!CTServico), "", Format(TBCFOP!CTServico, "###,##0.00"))
                CTOutrasGrid = IIf(IsNull(TBCFOP!CTOutras), "", Format(TBCFOP!CTOutras, "###,##0.00"))
                
                Qtde = 0
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBCFOP!Ordem & "' and desenho = '" & TBNivel2!CODIGO & "' and documento = '" & TBCFOP!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBItem.EOF = False Then
                    Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
                End If
                TBItem.Close
                
                QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
                
                'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
                If VermelhoGrid = 0 Then 'Se já tiver pra ficar vermelhor não precisa ver de novo
                    Qtd = IIf(IsNull(TBCFOP!QuantProd), 0, Format(TBCFOP!QuantProd, "0.0000"))
                    If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
                End If

                arrNodes(Contador1).Text = TBNivel2!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel2!Descricao & vbTab & TBCFOP!Ordem & vbTab & IIf(TBCFOP!Tipo = "E", "P", TBCFOP!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & VermelhoGrid
                ProcNivel3Ordem TBCFOP!Ordem
            End If
            
        End If
        TBOrdem2.Close
            
        TBNivel2.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcNivel3Ordem(OrdemGrid As Long)
On Error GoTo tratar_erro

Set TBNivel3 = CreateObject("adodb.recordset")
TBNivel3.Open "Select * from Producaomaterial where ordem = " & OrdemGrid & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockReadOnly
If TBNivel3.EOF = False Then
    Do While TBNivel3.EOF = False
        PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel3!Posicao, 3)
        
        CodRef = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select I.n_referencia from item_aplicacoes I INNER JOIN projproduto P ON I.codproduto = P.codproduto where P.desenho = '" & TBNivel3!CODIGO & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            CodRef = TBItem!N_referencia
        End If
        TBItem.Close
        
        NFGrid = ""
        VlrunitGrid = ""
        SaidaGrid = ""
        VlrTotalGrid = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "SELECT * from RE_Entrada_RM_Ordem where desenho = '" & TBNivel3!CODIGO & "' and ordem = " & TBNivel3!Ordem, Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            NFGrid = IIf(IsNull(TBItem!Nota_fiscal), "", TBItem!Nota_fiscal)
            VlrunitGrid = IIf(IsNull(TBItem!VlrUnit), "", Format(TBItem!VlrUnit, "###,##0.0000"))
            VlrTotalGrid = IIf(IsNull(TBItem!VlrTotal), "", Format(TBItem!VlrTotal, "###,##0.00"))
        End If
        TBItem.Close
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Saida) as Saida from estoque_movimentacao where oe = '" & OrdemGrid & "' and desenho = '" & TBNivel3!CODIGO & "' and documento = '" & OrdemGrid & "' AND operacao IN ('SAIDA_ORDEM', 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Saida), 0, Format(TBItem!Saida, "0.000"))
        End If
        TBItem.Close
        
        'verifica se tem diferença entre qtde de saida e requisitada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBNivel3!Requisitado), 0, Format(TBNivel3!Requisitado, "0.000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        SaidaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        RequisitGrid = IIf(Qtd = 0, "", Format(Qtd, "###,##0.000"))
        
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 3
        arrNodes(Contador1).Text = TBNivel3!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & vbTab & VlrTotalGrid & VermelhoGrid
        
        'Verifica se existe uma ordem abaixo
        Set TBOrdem3 = CreateObject("adodb.recordset")
        TBOrdem3.Open "SELECT Lote FROM CustosProducao_OrdensAbaixo WHERE OE = '" & OrdemGrid & "' AND desenho = '" & TBNivel3!CODIGO & "' AND documento = '" & OrdemGrid & "' ORDER BY oe", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem3.EOF = False Then
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "SELECT P.Ordem, Tipo, Quant, QuantProd, TTTReal, CTTPrev, CTTReal, CTMaterial, CTServico, CTOutras FROM Producao P INNER JOIN Producao_Pedidos PP ON P.Ordem = PP.Ordem WHERE P.Ordem = " & TBOrdem3!LOTE & " AND PP.IDcarteira = " & txtidcarteira, Conexao, adOpenKeyset, adLockReadOnly
            If TBCFOP.EOF = False Then
                qtdeGrid = IIf(IsNull(TBCFOP!Quant), "", Format(TBCFOP!Quant, "###,##0.0000"))
                qtdeProdGrid = IIf(IsNull(TBCFOP!QuantProd), "", Format(TBCFOP!QuantProd, "###,##0.0000"))
                TTTRealGrid = IIf(IsNull(TBCFOP!TTTReal), "", Format(TBCFOP!TTTReal, "hh:mm:ss"))
                CTTPrevGrid = IIf(IsNull(TBCFOP!CTTPrev), "", Format(TBCFOP!CTTPrev, "###,##0.00"))
                CTTRealGrid = IIf(IsNull(TBCFOP!CTTReal), "", Format(TBCFOP!CTTReal, "###,##0.00"))
                CTMaterialGrid = IIf(IsNull(TBCFOP!CTMaterial), "", Format(TBCFOP!CTMaterial, "###,##0.00"))
                CTServicoGrid = IIf(IsNull(TBCFOP!CTServico), "", Format(TBCFOP!CTServico, "###,##0.00"))
                CTOutrasGrid = IIf(IsNull(TBCFOP!CTOutras), "", Format(TBCFOP!CTOutras, "###,##0.00"))
                
                Qtde = 0
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBCFOP!Ordem & "' and desenho = '" & TBNivel3!CODIGO & "' and documento = '" & TBCFOP!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBItem.EOF = False Then
                    Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
                End If
                TBItem.Close
                
                QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
                
                'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
                If VermelhoGrid = 0 Then 'Se já tiver pra ficar vermelhor não precisa ver de novo
                    Qtd = IIf(IsNull(TBCFOP!QuantProd), 0, Format(TBCFOP!QuantProd, "0.0000"))
                    If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
                End If

                arrNodes(Contador1).Text = TBNivel3!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel3!Descricao & vbTab & TBCFOP!Ordem & vbTab & IIf(TBCFOP!Tipo = "E", "P", TBCFOP!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & VermelhoGrid
                ProcNivel4Ordem TBCFOP!Ordem
            End If
            
        End If
        TBOrdem3.Close
            
        TBNivel3.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcNivel4Ordem(OrdemGrid As Long)
On Error GoTo tratar_erro

Set TBNivel4 = CreateObject("adodb.recordset")
TBNivel4.Open "Select * from Producaomaterial where ordem = " & OrdemGrid & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockReadOnly
If TBNivel4.EOF = False Then
    Do While TBNivel4.EOF = False
        PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel4!Posicao, 3)
        
        CodRef = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select I.n_referencia from item_aplicacoes I INNER JOIN projproduto P ON I.codproduto = P.codproduto where P.desenho = '" & TBNivel4!CODIGO & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            CodRef = TBItem!N_referencia
        End If
        TBItem.Close
        
        NFGrid = ""
        VlrunitGrid = ""
        SaidaGrid = ""
        VlrTotalGrid = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "SELECT * from RE_Entrada_RM_Ordem where desenho = '" & TBNivel4!CODIGO & "' and ordem = " & TBNivel4!Ordem, Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            NFGrid = IIf(IsNull(TBItem!Nota_fiscal), "", TBItem!Nota_fiscal)
            VlrunitGrid = IIf(IsNull(TBItem!VlrUnit), "", Format(TBItem!VlrUnit, "###,##0.0000"))
            VlrTotalGrid = IIf(IsNull(TBItem!VlrTotal), "", Format(TBItem!VlrTotal, "###,##0.00"))
        End If
        TBItem.Close
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Saida) as Saida from estoque_movimentacao where oe = '" & OrdemGrid & "' and desenho = '" & TBNivel4!CODIGO & "' and documento = '" & OrdemGrid & "' AND operacao IN ('SAIDA_ORDEM', 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Saida), 0, Format(TBItem!Saida, "0.000"))
        End If
        TBItem.Close
        
        'verifica se tem diferença entre qtde de saida e requisitada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBNivel4!Requisitado), 0, Format(TBNivel4!Requisitado, "0.000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        SaidaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        RequisitGrid = IIf(Qtd = 0, "", Format(Qtd, "###,##0.000"))
        
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 4
        arrNodes(Contador1).Text = TBNivel4!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel4!Descricao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & vbTab & VlrTotalGrid & VermelhoGrid
        
        'Verifica se existe uma ordem abaixo
        Set TBOrdem4 = CreateObject("adodb.recordset")
        TBOrdem4.Open "SELECT Lote FROM CustosProducao_OrdensAbaixo WHERE OE = '" & OrdemGrid & "' AND desenho = '" & TBNivel4!CODIGO & "' AND documento = '" & OrdemGrid & "' ORDER BY oe", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem4.EOF = False Then
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "SELECT P.Ordem, Tipo, Quant, QuantProd, TTTReal, CTTPrev, CTTReal, CTMaterial, CTServico, CTOutras FROM Producao P INNER JOIN Producao_Pedidos PP ON P.Ordem = PP.Ordem WHERE P.Ordem = " & TBOrdem4!LOTE & " AND PP.IDcarteira = " & txtidcarteira, Conexao, adOpenKeyset, adLockReadOnly
            If TBCFOP.EOF = False Then
                qtdeGrid = IIf(IsNull(TBCFOP!Quant), "", Format(TBCFOP!Quant, "###,##0.0000"))
                qtdeProdGrid = IIf(IsNull(TBCFOP!QuantProd), "", Format(TBCFOP!QuantProd, "###,##0.0000"))
                TTTRealGrid = IIf(IsNull(TBCFOP!TTTReal), "", Format(TBCFOP!TTTReal, "hh:mm:ss"))
                CTTPrevGrid = IIf(IsNull(TBCFOP!CTTPrev), "", Format(TBCFOP!CTTPrev, "###,##0.00"))
                CTTRealGrid = IIf(IsNull(TBCFOP!CTTReal), "", Format(TBCFOP!CTTReal, "###,##0.00"))
                CTMaterialGrid = IIf(IsNull(TBCFOP!CTMaterial), "", Format(TBCFOP!CTMaterial, "###,##0.00"))
                CTServicoGrid = IIf(IsNull(TBCFOP!CTServico), "", Format(TBCFOP!CTServico, "###,##0.00"))
                CTOutrasGrid = IIf(IsNull(TBCFOP!CTOutras), "", Format(TBCFOP!CTOutras, "###,##0.00"))
                
                Qtde = 0
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBCFOP!Ordem & "' and desenho = '" & TBNivel4!CODIGO & "' and documento = '" & TBCFOP!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBItem.EOF = False Then
                    Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
                End If
                TBItem.Close
                
                QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
                
                'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
                If VermelhoGrid = 0 Then 'Se já tiver pra ficar vermelhor não precisa ver de novo
                    Qtd = IIf(IsNull(TBCFOP!QuantProd), 0, Format(TBCFOP!QuantProd, "0.0000"))
                    If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
                End If

                arrNodes(Contador1).Text = TBNivel4!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel4!Descricao & vbTab & TBCFOP!Ordem & vbTab & IIf(TBCFOP!Tipo = "E", "P", TBCFOP!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & VermelhoGrid
                ProcNivel5Ordem TBCFOP!Ordem
            End If
            
        End If
        TBOrdem4.Close
            
        TBNivel4.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcNivel5Ordem(OrdemGrid As Long)
On Error GoTo tratar_erro

Set TBNivel5 = CreateObject("adodb.recordset")
TBNivel5.Open "Select * from Producaomaterial where ordem = " & OrdemGrid & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockReadOnly
If TBNivel5.EOF = False Then
    Do While TBNivel5.EOF = False
        PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel5!Posicao, 3)
        
        CodRef = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select I.n_referencia from item_aplicacoes I INNER JOIN projproduto P ON I.codproduto = P.codproduto where P.desenho = '" & TBNivel5!CODIGO & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            CodRef = TBItem!N_referencia
        End If
        TBItem.Close
        
        NFGrid = ""
        VlrunitGrid = ""
        SaidaGrid = ""
        VlrTotalGrid = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "SELECT * from RE_Entrada_RM_Ordem where desenho = '" & TBNivel5!CODIGO & "' and ordem = " & TBNivel5!Ordem, Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            NFGrid = IIf(IsNull(TBItem!Nota_fiscal), "", TBItem!Nota_fiscal)
            VlrunitGrid = IIf(IsNull(TBItem!VlrUnit), "", Format(TBItem!VlrUnit, "###,##0.0000"))
            VlrTotalGrid = IIf(IsNull(TBItem!VlrTotal), "", Format(TBItem!VlrTotal, "###,##0.00"))
        End If
        TBItem.Close
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Saida) as Saida from estoque_movimentacao where oe = '" & OrdemGrid & "' and desenho = '" & TBNivel5!CODIGO & "' and documento = '" & OrdemGrid & "' AND operacao IN ('SAIDA_ORDEM', 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Saida), 0, Format(TBItem!Saida, "0.000"))
        End If
        TBItem.Close
        
        'verifica se tem diferença entre qtde de saida e requisitada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBNivel5!Requisitado), 0, Format(TBNivel5!Requisitado, "0.000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        SaidaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        RequisitGrid = IIf(Qtd = 0, "", Format(Qtd, "###,##0.000"))
        
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 5
        arrNodes(Contador1).Text = TBNivel5!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel5!Descricao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & vbTab & VlrTotalGrid & VermelhoGrid
        
        'Verifica se existe uma ordem abaixo
        Set TBOrdem5 = CreateObject("adodb.recordset")
        TBOrdem5.Open "SELECT Lote FROM CustosProducao_OrdensAbaixo WHERE OE = '" & OrdemGrid & "' AND desenho = '" & TBNivel5!CODIGO & "' AND documento = '" & OrdemGrid & "' ORDER BY oe", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem5.EOF = False Then
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "SELECT P.Ordem, Tipo, Quant, QuantProd, TTTReal, CTTPrev, CTTReal, CTMaterial, CTServico, CTOutras FROM Producao P INNER JOIN Producao_Pedidos PP ON P.Ordem = PP.Ordem WHERE P.Ordem = " & TBOrdem5!LOTE & " AND PP.IDcarteira = " & txtidcarteira, Conexao, adOpenKeyset, adLockReadOnly
            If TBCFOP.EOF = False Then
                qtdeGrid = IIf(IsNull(TBCFOP!Quant), "", Format(TBCFOP!Quant, "###,##0.0000"))
                qtdeProdGrid = IIf(IsNull(TBCFOP!QuantProd), "", Format(TBCFOP!QuantProd, "###,##0.0000"))
                TTTRealGrid = IIf(IsNull(TBCFOP!TTTReal), "", Format(TBCFOP!TTTReal, "hh:mm:ss"))
                CTTPrevGrid = IIf(IsNull(TBCFOP!CTTPrev), "", Format(TBCFOP!CTTPrev, "###,##0.00"))
                CTTRealGrid = IIf(IsNull(TBCFOP!CTTReal), "", Format(TBCFOP!CTTReal, "###,##0.00"))
                CTMaterialGrid = IIf(IsNull(TBCFOP!CTMaterial), "", Format(TBCFOP!CTMaterial, "###,##0.00"))
                CTServicoGrid = IIf(IsNull(TBCFOP!CTServico), "", Format(TBCFOP!CTServico, "###,##0.00"))
                CTOutrasGrid = IIf(IsNull(TBCFOP!CTOutras), "", Format(TBCFOP!CTOutras, "###,##0.00"))
                
                Qtde = 0
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBCFOP!Ordem & "' and desenho = '" & TBNivel5!CODIGO & "' and documento = '" & TBCFOP!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBItem.EOF = False Then
                    Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
                End If
                TBItem.Close
                
                QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
                
                'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
                If VermelhoGrid = 0 Then 'Se já tiver pra ficar vermelhor não precisa ver de novo
                    Qtd = IIf(IsNull(TBCFOP!QuantProd), 0, Format(TBCFOP!QuantProd, "0.0000"))
                    If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
                End If

                arrNodes(Contador1).Text = TBNivel5!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel5!Descricao & vbTab & TBCFOP!Ordem & vbTab & IIf(TBCFOP!Tipo = "E", "P", TBCFOP!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & VermelhoGrid
                ProcNivel6Ordem TBCFOP!Ordem
            End If
            
        End If
        TBOrdem5.Close
            
        TBNivel5.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcNivel6Ordem(OrdemGrid As Long)
On Error GoTo tratar_erro

Set TBNivel6 = CreateObject("adodb.recordset")
TBNivel6.Open "Select * from Producaomaterial where ordem = " & OrdemGrid & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockReadOnly
If TBNivel6.EOF = False Then
    Do While TBNivel6.EOF = False
        PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel6!Posicao, 3)
        
        CodRef = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select I.n_referencia from item_aplicacoes I INNER JOIN projproduto P ON I.codproduto = P.codproduto where P.desenho = '" & TBNivel6!CODIGO & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            CodRef = TBItem!N_referencia
        End If
        TBItem.Close
        
        NFGrid = ""
        VlrunitGrid = ""
        SaidaGrid = ""
        VlrTotalGrid = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "SELECT * from RE_Entrada_RM_Ordem where desenho = '" & TBNivel6!CODIGO & "' and ordem = " & TBNivel6!Ordem, Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            NFGrid = IIf(IsNull(TBItem!Nota_fiscal), "", TBItem!Nota_fiscal)
            VlrunitGrid = IIf(IsNull(TBItem!VlrUnit), "", Format(TBItem!VlrUnit, "###,##0.0000"))
            VlrTotalGrid = IIf(IsNull(TBItem!VlrTotal), "", Format(TBItem!VlrTotal, "###,##0.00"))
        End If
        TBItem.Close
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Saida) as Saida from estoque_movimentacao where oe = '" & OrdemGrid & "' and desenho = '" & TBNivel6!CODIGO & "' and documento = '" & OrdemGrid & "' AND operacao IN ('SAIDA_ORDEM', 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Saida), 0, Format(TBItem!Saida, "0.000"))
        End If
        TBItem.Close
        
        'verifica se tem diferença entre qtde de saida e requisitada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBNivel6!Requisitado), 0, Format(TBNivel6!Requisitado, "0.000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        SaidaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        RequisitGrid = IIf(Qtd = 0, "", Format(Qtd, "###,##0.000"))
        
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 6
        arrNodes(Contador1).Text = TBNivel6!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel6!Descricao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & vbTab & VlrTotalGrid & VermelhoGrid
        
        'Verifica se existe uma ordem abaixo
        Set TBOrdem6 = CreateObject("adodb.recordset")
        TBOrdem6.Open "SELECT Lote FROM CustosProducao_OrdensAbaixo WHERE OE = '" & OrdemGrid & "' AND desenho = '" & TBNivel6!CODIGO & "' AND documento = '" & OrdemGrid & "' ORDER BY oe", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem6.EOF = False Then
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "SELECT P.Ordem, Tipo, Quant, QuantProd, TTTReal, CTTPrev, CTTReal, CTMaterial, CTServico, CTOutras FROM Producao P INNER JOIN Producao_Pedidos PP ON P.Ordem = PP.Ordem WHERE P.Ordem = " & TBOrdem6!LOTE & " AND PP.IDcarteira = " & txtidcarteira, Conexao, adOpenKeyset, adLockReadOnly
            If TBCFOP.EOF = False Then
                qtdeGrid = IIf(IsNull(TBCFOP!Quant), "", Format(TBCFOP!Quant, "###,##0.0000"))
                qtdeProdGrid = IIf(IsNull(TBCFOP!QuantProd), "", Format(TBCFOP!QuantProd, "###,##0.0000"))
                TTTRealGrid = IIf(IsNull(TBCFOP!TTTReal), "", Format(TBCFOP!TTTReal, "hh:mm:ss"))
                CTTPrevGrid = IIf(IsNull(TBCFOP!CTTPrev), "", Format(TBCFOP!CTTPrev, "###,##0.00"))
                CTTRealGrid = IIf(IsNull(TBCFOP!CTTReal), "", Format(TBCFOP!CTTReal, "###,##0.00"))
                CTMaterialGrid = IIf(IsNull(TBCFOP!CTMaterial), "", Format(TBCFOP!CTMaterial, "###,##0.00"))
                CTServicoGrid = IIf(IsNull(TBCFOP!CTServico), "", Format(TBCFOP!CTServico, "###,##0.00"))
                CTOutrasGrid = IIf(IsNull(TBCFOP!CTOutras), "", Format(TBCFOP!CTOutras, "###,##0.00"))
                
                Qtde = 0
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBCFOP!Ordem & "' and desenho = '" & TBNivel6!CODIGO & "' and documento = '" & TBCFOP!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBItem.EOF = False Then
                    Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
                End If
                TBItem.Close
                
                QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
                
                'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
                If VermelhoGrid = 0 Then 'Se já tiver pra ficar vermelhor não precisa ver de novo
                    Qtd = IIf(IsNull(TBCFOP!QuantProd), 0, Format(TBCFOP!QuantProd, "0.0000"))
                    If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
                End If

                arrNodes(Contador1).Text = TBNivel6!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel6!Descricao & vbTab & TBCFOP!Ordem & vbTab & IIf(TBCFOP!Tipo = "E", "P", TBCFOP!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & VermelhoGrid
                ProcNivel7Ordem TBCFOP!Ordem
            End If
            
        End If
        TBOrdem6.Close
            
        TBNivel6.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcNivel7Ordem(OrdemGrid As Long)
On Error GoTo tratar_erro

Set TBNivel7 = CreateObject("adodb.recordset")
TBNivel7.Open "Select * from Producaomaterial where ordem = " & OrdemGrid & " order by Posicao, Codigo", Conexao, adOpenKeyset, adLockReadOnly
If TBNivel7.EOF = False Then
    Do While TBNivel7.EOF = False
        PosicaoTexto = FunTamanhoTextoZeroEsq(TBNivel7!Posicao, 3)
        
        CodRef = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select I.n_referencia from item_aplicacoes I INNER JOIN projproduto P ON I.codproduto = P.codproduto where P.desenho = '" & TBNivel7!CODIGO & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            CodRef = TBItem!N_referencia
        End If
        TBItem.Close
        
        NFGrid = ""
        VlrunitGrid = ""
        SaidaGrid = ""
        VlrTotalGrid = ""
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "SELECT * from RE_Entrada_RM_Ordem where desenho = '" & TBNivel7!CODIGO & "' and ordem = " & TBNivel7!Ordem, Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            NFGrid = IIf(IsNull(TBItem!Nota_fiscal), "", TBItem!Nota_fiscal)
            VlrunitGrid = IIf(IsNull(TBItem!VlrUnit), "", Format(TBItem!VlrUnit, "###,##0.0000"))
            VlrTotalGrid = IIf(IsNull(TBItem!VlrTotal), "", Format(TBItem!VlrTotal, "###,##0.00"))
        End If
        TBItem.Close
        
        Qtde = 0
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Sum(Saida) as Saida from estoque_movimentacao where oe = '" & OrdemGrid & "' and desenho = '" & TBNivel7!CODIGO & "' and documento = '" & OrdemGrid & "' AND operacao IN ('SAIDA_ORDEM', 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
        If TBItem.EOF = False Then
            Qtde = IIf(IsNull(TBItem!Saida), 0, Format(TBItem!Saida, "0.000"))
        End If
        TBItem.Close
        
        'verifica se tem diferença entre qtde de saida e requisitada, se tiver pinta a linha
        Qtd = IIf(IsNull(TBNivel7!Requisitado), 0, Format(TBNivel7!Requisitado, "0.000"))
        If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
        
        SaidaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
        RequisitGrid = IIf(Qtd = 0, "", Format(Qtd, "###,##0.000"))
        
        Contador1 = Contador1 + 1
        arrNodes(Contador1).Level = 7
        arrNodes(Contador1).Text = TBNivel7!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel7!Descricao & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & vbTab & VlrTotalGrid & VermelhoGrid
        
        'Verifica se existe uma ordem abaixo
        Set TBOrdem7 = CreateObject("adodb.recordset")
        TBOrdem7.Open "SELECT Lote FROM CustosProducao_OrdensAbaixo WHERE OE = '" & OrdemGrid & "' AND desenho = '" & TBNivel7!CODIGO & "' AND documento = '" & OrdemGrid & "' ORDER BY oe", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem7.EOF = False Then
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "SELECT P.Ordem, Tipo, Quant, QuantProd, TTTReal, CTTPrev, CTTReal, CTMaterial, CTServico, CTOutras FROM Producao P INNER JOIN Producao_Pedidos PP ON P.Ordem = PP.Ordem WHERE P.Ordem = " & TBOrdem7!LOTE & " AND PP.IDcarteira = " & txtidcarteira, Conexao, adOpenKeyset, adLockReadOnly
            If TBCFOP.EOF = False Then
                qtdeGrid = IIf(IsNull(TBCFOP!Quant), "", Format(TBCFOP!Quant, "###,##0.0000"))
                qtdeProdGrid = IIf(IsNull(TBCFOP!QuantProd), "", Format(TBCFOP!QuantProd, "###,##0.0000"))
                TTTRealGrid = IIf(IsNull(TBCFOP!TTTReal), "", Format(TBCFOP!TTTReal, "hh:mm:ss"))
                CTTPrevGrid = IIf(IsNull(TBCFOP!CTTPrev), "", Format(TBCFOP!CTTPrev, "###,##0.00"))
                CTTRealGrid = IIf(IsNull(TBCFOP!CTTReal), "", Format(TBCFOP!CTTReal, "###,##0.00"))
                CTMaterialGrid = IIf(IsNull(TBCFOP!CTMaterial), "", Format(TBCFOP!CTMaterial, "###,##0.00"))
                CTServicoGrid = IIf(IsNull(TBCFOP!CTServico), "", Format(TBCFOP!CTServico, "###,##0.00"))
                CTOutrasGrid = IIf(IsNull(TBCFOP!CTOutras), "", Format(TBCFOP!CTOutras, "###,##0.00"))
                
                Qtde = 0
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select Sum(Entrada) as Entrada from estoque_movimentacao where Lote = '" & TBCFOP!Ordem & "' and desenho = '" & TBNivel7!CODIGO & "' and documento = '" & TBCFOP!Ordem & "' AND operacao IN ('ENTRADA_ORDEM', 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBItem.EOF = False Then
                    Qtde = IIf(IsNull(TBItem!Entrada), 0, Format(TBItem!Entrada, "0.000"))
                End If
                TBItem.Close
                
                QtdeEntredaGrid = IIf(Qtde = 0, "", Format(Qtde, "###,##0.000"))
                
                'verifica se tem diferença entre qtde produzida e entrada, se tiver pinta a linha
                If VermelhoGrid = 0 Then 'Se já tiver pra ficar vermelhor não precisa ver de novo
                    Qtd = IIf(IsNull(TBCFOP!QuantProd), 0, Format(TBCFOP!QuantProd, "0.0000"))
                    If Qtd <> Qtde Then VermelhoGrid = 1 Else VermelhoGrid = 0
                End If

                arrNodes(Contador1).Text = TBNivel7!CODIGO & vbTab & PosicaoTexto & vbTab & CodRef & vbTab & TBNivel7!Descricao & vbTab & TBCFOP!Ordem & vbTab & IIf(TBCFOP!Tipo = "E", "P", TBCFOP!Tipo) & vbTab & qtdeGrid & vbTab & qtdeProdGrid & vbTab & QtdeEntredaGrid & vbTab & TTTRealGrid & vbTab & CTTPrevGrid & vbTab & CTTRealGrid & vbTab & CTMaterialGrid & vbTab & CTServicoGrid & vbTab & CTOutrasGrid & vbTab & "" & vbTab & NFGrid & vbTab & VlrunitGrid & vbTab & RequisitGrid & vbTab & SaidaGrid & VermelhoGrid
            End If
        End If
        TBOrdem7.Close
            
        TBNivel7.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Grid1_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
On Error GoTo tratar_erro
Dim i As Long, j As Long
Dim x As Long, Y As Long
Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long
Dim lngLevel As Long
Dim blnDrawLine As Boolean
Dim objNode As Node, tmpNode As Node
Dim intWidth As Integer
Dim intAdd As Integer

If Row < m_Row Or Col <> m_Col Then Exit Sub

intWidth = 20
intAdd = 26
    
Set objNode = m_Tree.FindNode(Row - m_Row + 2)
If Not objNode Is Nothing Then
    lngLevel = objNode.Level - 1

    'Tree lines
    hPen = CreatePen(0, 1, RGB(128, 128, 128))
    hOldPen = SelectObject(hdc, hPen)
    For i = 0 To lngLevel
        If i < lngLevel - 1 Then
            blnDrawLine = True
            Set tmpNode = objNode
            For j = i To lngLevel - 2
                Set tmpNode = tmpNode.Parent
            Next
            If tmpNode.NextNode Is Nothing Then
                blnDrawLine = False
            End If
            If blnDrawLine Then
                'All
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel - 1 Then
            'Top
            Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2)
            If Not objNode.NextNode Is Nothing Then
                'Bottom
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel Then
            'Top
            If objNode.VisibleNodesCount > 1 Then
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        End If
        'Horizontal line
        If lngLevel > 0 Then
            Call DrawLine(hdc, Left + intWidth * (lngLevel - 1) + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * (lngLevel - 1) + intAdd + 10, Top + (Bottom - Top) / 2)
        End If
    Next
    
    Call SelectObject(hdc, hOldPen)
    Call DeleteObject(hPen)

    '+/-
    If objNode.ChildrenCount > 0 Then
        hPen = CreatePen(0, 1, 0)
        hOldPen = SelectObject(hdc, hPen)
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        hOldPen = SelectObject(hdc, hBrush)
        
        x = Left + 2 + intWidth * lngLevel
        Y = Top + (Bottom - Top - 9) / 2
        
        Call Rectangle(hdc, x, Y, x + 9, Y + 9)
        If objNode.Expanded Then
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
        Else
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
            Call DrawLine(hdc, x + 4, Y + 2, x + 4, Y + 7)
        End If
    
        Call SelectObject(hdc, hOldPen)
        Call DeleteObject(hPen)
        Call SelectObject(hdc, hOldBrush)
        Call DeleteObject(hBrush)
    End If
    
    'Icon
    If objNode.ChildrenCount > 0 Then
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFolder.Picture, 16, 16, 0, 0, DI_NORMAL
    Else
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFile.Picture, 16, 16, 0, 0, DI_NORMAL
    End If
    
    'Text
    With Grid1.Cell(Row, Col)
        Dim rc As rect
        Call SetRect(rc, Left + intWidth * lngLevel + 37, Top, Right, Bottom)
        Call DrawText(hdc, .Text, -1, rc, DT_SINGLELINE Or DT_VCENTER)
    End With

    Handled = True
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_Click()
On Error GoTo tratar_erro
Dim point As POINTAPI
Dim objCell As FlexCell.Cell
Dim intWidth As Integer

If FunCheckEditStatus() Then Exit Sub
intWidth = 20

Call GetCursorPos(point)
Call ScreenToClient(Grid1.hWnd, point)
Set objCell = Grid1.HitTest(point.x, point.Y)

If Not objCell Is Nothing Then
    If objCell.Row >= m_Row And objCell.Col = m_Col Then
        Dim objNode As Node
        Set objNode = m_Tree.FindNode(objCell.Row - m_Row + 2)
        If Not objNode Is Nothing Then
            Dim i As Long, x As Long, Y As Long
            x = objCell.Left + 2 + (objNode.Level - 1) * intWidth
            Y = objCell.Top + (objCell.Height - 9) / 2
            If point.x >= x And point.x <= x + 9 And point.Y >= Y And point.Y <= Y + 9 Then
                If objNode.Expanded Then
                    objNode.Collapse
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        Grid1.RowHeight(objCell.Row + i) = 0
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                Else
                    objNode.Expand
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        If objNode.FindNode(i + 1).Visible Then
                            Grid1.RowHeight(objCell.Row + i) = -1 'DefaultRowHeight
                        End If
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                End If
            End If
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Function FunCheckEditStatus() As Boolean
On Error GoTo tratar_erro
Dim hWnd As Long
Dim strClassName As String
Dim intPos As Integer

strClassName = Space(256)
hWnd = GetFocus()
Call GetClassName(hWnd, strClassName, 256)
intPos = InStr(1, strClassName, Chr(0))
strClassName = Left(strClassName, intPos - 1)
If strClassName = "ThunderRT6TextBox" Then FunCheckEditStatus = True    'Editing Else    FunCheckEditStatus = False

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
