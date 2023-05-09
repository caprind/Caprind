VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_Relatorios_Indice_Atraso 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Relatórios - Índice de atraso"
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
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
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
      TabIndex        =   31
      Top             =   960
      Width           =   1905
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
   End
   Begin VB.OptionButton optPrazoFinal 
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
      Left            =   13230
      TabIndex        =   8
      Top             =   990
      Value           =   -1  'True
      Width           =   945
   End
   Begin VB.OptionButton optEstoque 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. reb."
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
      Left            =   14250
      TabIndex        =   9
      Top             =   990
      Width           =   915
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   14040
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_Relatorios_Indice_Atraso.frx":0000
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
      TabIndex        =   22
      Top             =   9180
      Width           =   15195
      Begin VB.TextBox txtIndice_antecipado 
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
         Left            =   6270
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total atraso."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox txtIndice_prazo 
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
         Left            =   9930
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total atraso."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_qtde_total_recebida 
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
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total faturada."
         Top             =   390
         Width           =   1920
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
         Left            =   13575
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Índice."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_qtde_total_comprada 
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
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total vendida."
         Top             =   390
         Width           =   1920
      End
      Begin VB.TextBox Txt_qtde_total_antecipado 
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
         Left            =   4275
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total atraso."
         Top             =   390
         Width           =   1980
      End
      Begin VB.TextBox Txt_qtde_total_prazo 
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
         Left            =   7995
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total atraso."
         Top             =   390
         Width           =   1920
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
         Left            =   11640
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total atraso."
         Top             =   390
         Width           =   1920
      End
      Begin VB.Label Label10 
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
         Left            =   6728
         TabIndex        =   42
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label7 
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
         Left            =   10388
         TabIndex        =   41
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total no prazo"
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
         Left            =   8115
         TabIndex        =   40
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total antecipada"
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
         Left            =   4290
         TabIndex        =   39
         Top             =   180
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total comprada"
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
         Left            =   210
         TabIndex        =   26
         Top             =   180
         Width           =   1800
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total recebida"
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
         Left            =   2205
         TabIndex        =   25
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
         Left            =   14033
         TabIndex        =   24
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
         Left            =   11850
         TabIndex        =   23
         Top             =   180
         Width           =   1500
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   29
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
      TabIndex        =   30
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
      TabIndex        =   18
      Top             =   960
      Width           =   7965
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmCompras_Relatorios_Indice_Atraso.frx":2DFB
         Left            =   2310
         List            =   "frmCompras_Relatorios_Indice_Atraso.frx":2DFD
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
         ItemData        =   "frmCompras_Relatorios_Indice_Atraso.frx":2DFF
         Left            =   180
         List            =   "frmCompras_Relatorios_Indice_Atraso.frx":2E12
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
         TabIndex        =   20
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
         TabIndex        =   19
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
      TabIndex        =   21
      Top             =   960
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
      TabIndex        =   27
      Top             =   960
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
      TabIndex        =   15
      Top             =   960
      Width           =   2115
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   11
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
         Format          =   140181505
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   10
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
         Format          =   140181505
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   690
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6915
      Left            =   60
      TabIndex        =   12
      Top             =   1980
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12197
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
      NumItems        =   17
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
         Text            =   "Dt. recebimento"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Pedido"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Qtde. comprada"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Qtde. recebida"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Qtde. antecipada"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Qtde. no prazo"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "Qtde. atraso"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6920
      Left            =   60
      TabIndex        =   13
      Top             =   1980
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12197
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Object.Width           =   7594
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Qtde. comprada"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Qtde. recebida"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Qtde. antecipada"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Qtde. no prazo"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Índice"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Qtde. atraso"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
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
      TabIndex        =   28
      Top             =   8940
      Width           =   3315
   End
End
Attribute VB_Name = "frmCompras_Relatorios_Indice_Atraso"
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
Vendas_Relatorio_IndiceAtraso = False
Vendas_Relatorio_Comissao = False
Compras_Relatorio_IndiceAtraso = True
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
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil)
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select CP.Pedido, CP.Fornecedor, CPL.Desenho, CPL.N_Referencia, CPL.Descricao, CPL.Familia from compras_pedido_lista CPL INNER JOIN Compras_pedido CP ON CPL.IDPedido = CP.IDPedido where CPL.IDLista = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Pedido), "", TBAbrir!Pedido)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Fornecedor), "", TBAbrir!Fornecedor)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
                End If
                TBAbrir.Close
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Terceiros), "", TBLISTA!Terceiros) 'Qtde comprada
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!qtdeOK), "", TBLISTA!qtdeOK) 'Qtde recebida
                .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Numero1), "", TBLISTA!Numero1) 'Qtde antecipada
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Numero3), "", Format(TBLISTA!Numero3, "###,##0.00") & "%") 'Indice antecipado
                .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Numero2), "", TBLISTA!Numero2) 'Qtde no prazo
                .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Numero4), "", Format(TBLISTA!Numero4, "###,##0.00") & "%") 'Indice no prazo
                .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!qtdeNC), "", TBLISTA!qtdeNC) 'Qtde antrasada
                .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%") 'Indice atrasado
            End With
        Else
            With Lista1.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!QtdePrev), "", TBLISTA!QtdePrev) 'Qtde comprada
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!qtdeOK), "", TBLISTA!qtdeOK) 'Qtde recebida
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Numero1), "", TBLISTA!Numero1) 'Qtde antecipada
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Numero3), "", Format(TBLISTA!Numero3, "###,##0.00") & "%") 'Indice antecipado
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Numero2), "", TBLISTA!Numero2) 'Qtde no prazo
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Numero4), "", Format(TBLISTA!Numero4, "###,##0.00") & "%") 'Indice no prazo
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!qtdeNC), "", TBLISTA!qtdeNC) 'Qtde antrasada
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%") 'Indice atrasado
            End With
        End If
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    If optDetalhado.Value = True Then Else
End If
TBLISTA.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_qtde_total_comprada = TBLISTA!QtdePrevista 'Qtde comprada
    Txt_qtde_total_recebida = TBLISTA!QtdeProduzida 'Qtde recebida
    Txt_qtde_total_antecipado = TBLISTA!Numero1 'Qtde antecipada
    txtIndice_antecipado = Format(TBLISTA!Numero3, "###,##0.00") & "%" 'Indice antecipado
    Txt_qtde_total_prazo = TBLISTA!Numero2 'Qtde no prazo
    txtIndice_prazo = Format(TBLISTA!Numero4, "###,##0.00") & "%" 'Indice no prazo
    Txt_qtde_total_atraso = TBLISTA!qtdeNC 'Qtde atrasado
    Txt_indice = Format(TBLISTA!TotalEficiencia, "###,##0.00") & "%" 'Indice atrasado
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
Txt_qtde_total_comprada = ""
Txt_qtde_total_recebida = ""
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
Formulario = "Compras/Relatórios/Índice de atraso"
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
    Case "Fornecedor": Texto = "CP.Fornecedor"
    Case "Código de referência": Texto = "CPL.N_Referencia"
    Case "Código interno": Texto = "CPL.desenho"
    Case "Descrição": Texto = "CPL.descricao"
    Case "Família": Texto = "CPL.Familia"
End Select
cmbTexto.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select " & Texto & " as NomeCampo1 from compras_pedido_lista CPL INNER JOIN Compras_pedido CP ON CPL.IDpedido = CP.IDpedido where " & Texto & " is not null group by " & Texto, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If TBAbrir!NomeCampo1 <> "" Then cmbTexto.AddItem TBAbrir!NomeCampo1
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Opt_individual.Value = True And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If
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
    Case "Código interno": TextoFiltro = "CPL.desenho"
    Case "Código de referência": TextoFiltro = "CPL.n_referencia"
    Case "Descrição": TextoFiltro = "CPL.Descricao"
    Case "Família": TextoFiltro = "CPL.familia"
    Case "Fornecedor": TextoFiltro = "CP.Fornecedor"
End Select
If Opt_individual.Value = True Then
    TextoFiltro1 = TextoFiltro & " = '" & cmbTexto & "' and "
    Ordenar = "CPL.Prazo"
Else
    TextoFiltro1 = ""
    Ordenar = TextoFiltro & ", CPL.Idlista"
End If

CamposFiltro = "CPL.Idlista, CPL.desenho, CPL.n_referencia, CPL.Descricao, CPL.familia, CPL.Prazo, CPL.Quant_Comp, CP.IDpedido, CP.Fornecedor, SUM(ECR.Recebido) AS Recebido, ECR.Data_recebimento"
INNERJOINTEXTO = "(Compras_pedido_lista CPL INNER JOIN Compras_pedido CP ON CPL.IDpedido = CP.IDPedido) INNER JOIN Estoque_controle_recebimento ECR ON ECR.IDlista = CPL.IdLista and ECR.idpedido = CP.IDpedido"
If optPrazoFinal.Value = True Then TextoFiltro2 = "(CPL.prazo)" Else TextoFiltro2 = "(ECR.Data_recebimento)"

Set TBCarteira = CreateObject("adodb.recordset")
StrSql = "Select " & CamposFiltro & ", MAX(ECR.Data_recebimento) as Data_Recebimento from " & INNERJOINTEXTO & " where " & TextoFiltro1 & TextoFiltro2 & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and (CPL.Status_item = 'RECEBIDO' or CPL.Status_item = 'PARCIAL') group by CPL.Idlista, CPL.desenho, CPL.n_referencia, CPL.Descricao, CPL.familia, CPL.Prazo, CPL.Quant_Comp, CP.IDpedido, CP.Fornecedor, ECR.Data_recebimento order by " & Ordenar & ", MAX(ECR.Data_recebimento)"

Debug.Print StrSql


TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

IDlista = 0
If TBCarteira.EOF = False Then
    Permitido = True
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBCarteira.EOF = False
        ProcSalvarDados
        IDlista = TBCarteira!IDlista
        TBCarteira.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
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
        Case "Descrição": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Descricao & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        Case "Família": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Familia & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
        Case "Fornecedor": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBCarteira!Fornecedor & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
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
TBProdutividade!Ordem = TBCarteira!IDlista
TBProdutividade!Data = TBCarteira!Prazo

If Opt_prazo_qtde.Value = True Then
    If IDlista <> TBCarteira!IDlista Then TBProdutividade!QtdePrev = TBCarteira!Quant_Comp 'Qtde. comprada para somar no total
    TBProdutividade!Terceiros = TBCarteira!Quant_Comp 'Qtde. comprada apenas para mostrar na lista e relatorio
    TBProdutividade!qtdeOK = TBCarteira!Recebido 'Qtde. recebida
    TBProdutividade!Totalhsutil = Format(TBCarteira!Data_recebimento, "dd/mm/yy") 'Data recebimento
    If TBCarteira!Data_recebimento < TBCarteira!Prazo Then TBProdutividade!Numero1 = TBCarteira!Recebido Else TBProdutividade!Numero1 = 0 'Qtde. antecipada
    If TBCarteira!Data_recebimento = TBCarteira!Prazo Then TBProdutividade!Numero2 = TBCarteira!Recebido Else TBProdutividade!Numero2 = 0 'Qtde. no prazo
    If TBCarteira!Data_recebimento > TBCarteira!Prazo Then TBProdutividade!qtdeNC = TBCarteira!Recebido Else TBProdutividade!qtdeNC = 0 'Qtde. em atraso
    
    If TBCarteira!Quant_Comp <> 0 Then
        TBProdutividade!Numero3 = (TBProdutividade!Numero1 / TBCarteira!Quant_Comp) * 100 'Indice antecipado
        TBProdutividade!Numero4 = (TBProdutividade!Numero2 / TBCarteira!Quant_Comp) * 100 'Indice no prazo
        TBProdutividade!Eficiencia = (TBProdutividade!qtdeNC / TBCarteira!Quant_Comp) * 100 'Indice atraso
    Else
        TBProdutividade!Numero3 = 0 'Indice antecipado
        TBProdutividade!Numero4 = 0  'Indice no prazo
        TBProdutividade!Eficiencia = 0 'Indice atraso
    End If
Else
    'Linha do pedido
    TBProdutividade!QtdePrev = 1 'comprado para somar no total
    TBProdutividade!Terceiros = 1 'Qtde. comprada apenas para mostrar na lista e relatorio
    TBProdutividade!qtdeOK = 1 'recebido
    TBProdutividade!Totalhsutil = Format(TBCarteira!Data_recebimento, "dd/mm/yy") 'Data recebimento
    If TBCarteira!Data_recebimento < TBCarteira!Prazo Then TBProdutividade!Numero1 = 1 Else TBProdutividade!Numero1 = 0 'Qtde. antecipada
    If TBCarteira!Data_recebimento = TBCarteira!Prazo Then TBProdutividade!Numero2 = 1 Else TBProdutividade!Numero2 = 0 'Qtde. no prazo
    If TBCarteira!Data_recebimento > TBCarteira!Prazo Then TBProdutividade!qtdeNC = 1 Else TBProdutividade!qtdeNC = 0 'em atraso
    
    If TBProdutividade!QtdePrev <> 0 Then
        TBProdutividade!Numero3 = (TBProdutividade!Numero1 / TBProdutividade!QtdePrev) * 100 'Indice antecipado
        TBProdutividade!Numero4 = (TBProdutividade!Numero2 / TBProdutividade!QtdePrev) * 100 'Indice no prazo
        TBProdutividade!Eficiencia = (TBProdutividade!qtdeNC / TBProdutividade!QtdePrev) * 100
    Else
        TBProdutividade!Numero3 = 0 'Indice antecipado
        TBProdutividade!Numero4 = 0  'Indice no prazo
        TBProdutividade!Eficiencia = 0 'Indice atraso
    End If
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
        Case "Descrição": Texto = TBCarteira!Descricao
        Case "Família": Texto = TBCarteira!Familia
        Case "Fornecedor": Texto = TBCarteira!Fornecedor
    End Select
End If
TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario

If Opt_prazo_qtde.Value = True Then
    If IDlista <> TBCarteira!IDlista Then TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + TBCarteira!Quant_Comp 'Qtde. comprada
    TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + TBCarteira!Recebido 'Qtde. recebida
    If TBCarteira!Data_recebimento < TBCarteira!Prazo Then TBProdutividade!Numero1 = TBProdutividade!Numero1 + TBCarteira!Recebido 'Qtde. antecipada
    If TBCarteira!Data_recebimento = TBCarteira!Prazo Then TBProdutividade!Numero2 = TBProdutividade!Numero2 + TBCarteira!Recebido 'Qtde. no prazo
    If TBCarteira!Data_recebimento > TBCarteira!Prazo Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TBCarteira!Recebido 'Qtde. em atraso
Else
    'Linha do pedido
    TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + 1 'comprados
    TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + 1 'recebidos
    If TBCarteira!Data_recebimento < TBCarteira!Prazo Then TBProdutividade!Numero1 = TBProdutividade!Numero1 + 1 'antecipados
    If TBCarteira!Data_recebimento = TBCarteira!Prazo Then TBProdutividade!Numero2 = TBProdutividade!Numero2 + 1 'no prazo
    If TBCarteira!Data_recebimento > TBCarteira!Prazo Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + 1 'em atraso
End If

If TBProdutividade!qtdeOK <> 0 Then
    TBProdutividade!Numero3 = (TBProdutividade!Numero1 / TBProdutividade!QtdePrev) * 100 'Indice antecipado
    TBProdutividade!Numero4 = (TBProdutividade!Numero2 / TBProdutividade!QtdePrev) * 100 'Indice no prazo
    TBProdutividade!Eficiencia = (TBProdutividade!qtdeNC / TBProdutividade!QtdePrev) * 100 'Indice atraso
Else
    TBProdutividade!Numero3 = 0 'Indice antecipado
    TBProdutividade!Numero4 = 0  'Indice no prazo
    TBProdutividade!Eficiencia = 0 'Indice atraso
End If

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
Qtd = 0
Qtde = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew

TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
If Opt_individual.Value = True Then TBAbrir!Texto = cmbfiltrarpor & ") : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(QtdePrev) as quantidade, Sum(QtdeOK) as QTLOTE, Sum(QtdeNC) as Quant, Sum(Numero1) as Quant1, Sum(Numero2) as Quant2 from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    quantidade = IIf(IsNull(TBproducao!quantidade), 0, TBproducao!quantidade) 'Qtde. comprada
    QTLOTE = IIf(IsNull(TBproducao!QTLOTE), 0, TBproducao!QTLOTE) 'Qtde. recebida
    Qtd = IIf(IsNull(TBproducao!Quant1), 0, TBproducao!Quant1) 'Qtde. antecipado
    Qtde = IIf(IsNull(TBproducao!Quant2), 0, TBproducao!Quant2) 'Qtde. no prazo
    Quant = IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant) 'Qtde. em atraso
End If
TBproducao.Close
TBAbrir!QtdePrevista = quantidade 'Qtde. comprada
TBAbrir!QtdeProduzida = QTLOTE 'Qtde. recebida
TBAbrir!Numero1 = Qtd 'Qtde. antecipado
TBAbrir!Numero2 = Qtde 'Qtde. no prazo
TBAbrir!qtdeNC = Quant 'Qtde. em atraso
If TBAbrir!QtdeProduzida <> 0 Then
    TBAbrir!Numero3 = (TBAbrir!Numero1 / TBAbrir!QtdePrevista) * 100 'Indice antecipado
    TBAbrir!Numero4 = (TBAbrir!Numero2 / TBAbrir!QtdePrevista) * 100 'Indice no prazo
    TBAbrir!TotalEficiencia = (TBAbrir!qtdeNC / TBAbrir!QtdePrevista) * 100 'Indice em atraso
Else
    TBAbrir!Numero3 = 0 'Indice antecipado
    TBAbrir!Numero4 = 0 'Indice no prazo
    TBAbrir!TotalEficiencia = 0 'Indice em atraso
End If
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

Private Sub optEstoque_Click()
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
