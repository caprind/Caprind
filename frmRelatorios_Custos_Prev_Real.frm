VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmRelatorios_Custos_Prev_Real 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Custos - Relatórios - Previsto x Realizado"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
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
   ScaleWidth      =   15270
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
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1020
      Left            =   11670
      TabIndex        =   34
      Top             =   990
      Width           =   1275
      Begin VB.TextBox Txt_rev_prev 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         MaxLength       =   9
         TabIndex        =   4
         ToolTipText     =   "Revisão prevista."
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Rev. prev."
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   750
      End
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
      Height          =   1020
      Left            =   5940
      TabIndex        =   29
      Top             =   990
      Width           =   5715
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmRelatorios_Custos_Prev_Real.frx":0000
         Left            =   180
         List            =   "frmRelatorios_Custos_Prev_Real.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   450
         Width           =   5355
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
         Left            =   2122
         TabIndex        =   30
         Top             =   240
         Width           =   1470
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
      Height          =   1020
      Left            =   4230
      TabIndex        =   28
      Top             =   990
      Width           =   1695
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
         TabIndex        =   2
         Top             =   600
         Width           =   1425
      End
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
         TabIndex        =   1
         ToolTipText     =   "0"
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1020
      Left            =   55
      TabIndex        =   26
      Top             =   990
      Width           =   4155
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
         ItemData        =   "frmRelatorios_Custos_Prev_Real.frx":0004
         Left            =   180
         List            =   "frmRelatorios_Custos_Prev_Real.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   450
         Width           =   3780
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
         Left            =   1703
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1020
      Left            =   12960
      TabIndex        =   22
      Top             =   990
      Width           =   2295
      Begin VB.ComboBox Cmb_ano_de 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmRelatorios_Custos_Prev_Real.frx":0008
         Left            =   1350
         List            =   "frmRelatorios_Custos_Prev_Real.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Ano de."
         Top             =   210
         Width           =   765
      End
      Begin VB.ComboBox Cmb_mes_ate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmRelatorios_Custos_Prev_Real.frx":000C
         Left            =   630
         List            =   "frmRelatorios_Custos_Prev_Real.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Mês até."
         Top             =   570
         Width           =   735
      End
      Begin VB.ComboBox Cmb_ano_ate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmRelatorios_Custos_Prev_Real.frx":0075
         Left            =   1350
         List            =   "frmRelatorios_Custos_Prev_Real.frx":0077
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ano até."
         Top             =   570
         Width           =   765
      End
      Begin VB.ComboBox Cmb_mes_de 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmRelatorios_Custos_Prev_Real.frx":0079
         Left            =   630
         List            =   "frmRelatorios_Custos_Prev_Real.frx":00A1
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Mês de."
         Top             =   210
         Width           =   735
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
         TabIndex        =   24
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label10 
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
         TabIndex        =   23
         Top             =   270
         Width           =   300
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   9720
      Width           =   11385
      _ExtentX        =   20082
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
      TabIndex        =   25
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
      ButtonToolTipText2=   "Relatório (F5]"
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
      ButtonCaption3  =   "Justificar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Justificar (F7)"
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
      ButtonWidth3    =   51
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   146
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
      ButtonLeft5     =   150
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
      ButtonLeft6     =   188
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
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
      ButtonLeft7     =   216
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   10680
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmRelatorios_Custos_Prev_Real.frx":00E2
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista_resumido 
      Height          =   2715
      Left            =   60
      TabIndex        =   9
      Top             =   2025
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   4789
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
      NumItems        =   0
   End
   Begin VB.Frame frameTotalReceber 
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
      Left            =   60
      TabIndex        =   16
      Top             =   8880
      Width           =   15195
      Begin VB.TextBox Txt_percentual 
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
         Left            =   13005
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Percentual."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_total_orcado 
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
         Left            =   6935
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Total orçado."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_total_real 
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
         Left            =   8960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Total real."
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox Txt_variacao_total 
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
         Left            =   10985
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Variação total."
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencentual"
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
         Left            =   13530
         TabIndex        =   31
         Top             =   180
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total orçado"
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
         Left            =   7400
         TabIndex        =   19
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total real"
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
         Left            =   9560
         TabIndex        =   18
         Top             =   180
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Variação total"
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
         Left            =   11390
         TabIndex        =   17
         Top             =   180
         Width           =   1200
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   60
      TabIndex        =   32
      Top             =   4770
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Resumido"
      TabPicture(0)   =   "frmRelatorios_Custos_Prev_Real.frx":355E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lista_res_PC"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalhado (previsto e realizado)"
      TabPicture(1)   =   "frmRelatorios_Custos_Prev_Real.frx":357A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lista_det_orcado"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lista_det_real"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSComctlLib.ListView Lista_det_real 
         Height          =   2235
         Left            =   30
         TabIndex        =   11
         Top             =   1860
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   3942
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Centro de custo origem"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Módulo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Referência"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Código contábil"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   4789
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_res_PC 
         Height          =   3765
         Left            =   -74970
         TabIndex        =   33
         Top             =   330
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   6641
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Lista_det_orcado 
         Height          =   1515
         Left            =   30
         TabIndex        =   10
         Top             =   330
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   2672
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Mês"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Ano"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Código contábil"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   11315
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "V"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
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
      Left            =   11610
      TabIndex        =   21
      Top             =   9750
      Width           =   3315
   End
End
Attribute VB_Name = "frmRelatorios_Custos_Prev_Real"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID_CC As Long 'OK

Private Sub ProcJustificar()
On Error GoTo tratar_erro

If Lista_resumido.ListItems.Count = 0 Or Lista_res_PC.ListItems.Count = 0 Then Exit Sub
frmRelatorios_Custos_Prev_Real_Just.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ano_de_Click()
On Error GoTo tratar_erro

Cmb_ano_ate = Cmb_ano_de
ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_ate_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista_resumido.ListItems.Count = 0 Then Exit Sub
frmRelatorios_Custos_Prev_Real_Menu_Impressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
CC_RM = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_de_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

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
    Case vbKeyF7: ProcJustificar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Familiatext = ""
Contador1 = 1
Posicao = 0
Lista_resumido.ListItems.Clear
Lista_det_orcado.ListItems.Clear
Lista_det_real.ListItems.Clear
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        If TBLISTA!maquina <> "" Then
            With Lista_resumido.ListItems
                Contador1 = 2
                Do While Lista_resumido.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                    Contador1 = Contador1 + 1
                Loop
Pula:
                If TBLISTA!maquina <> Familiatext Then
                    .Add , , TBLISTA!Fase
                    .Item(.Count).SubItems(1) = TBLISTA!maquina
                    Posicao = Posicao + 1
                End If
                
                valor = IIf(IsNull(TBLISTA!qtdeOK), 0, TBLISTA!qtdeOK) 'Orçado
                Valor1 = IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC) 'Real
                Valor2 = IIf(IsNull(TBLISTA!Terceiros), 0, TBLISTA!Terceiros) 'Variação
                Valor3 = IIf(IsNull(TBLISTA!impostos), 0, TBLISTA!impostos) 'Percentual
                
                Valor_Cofins_Prod = IIf(IsNull(TBLISTA!Qtdetotalprod), 0, TBLISTA!Qtdetotalprod) 'Total orçado
                Valor_Cofins_Serv = IIf(IsNull(TBLISTA!OS), 0, TBLISTA!OS) 'Total real
                Valor_CSLL_Prod = IIf(IsNull(TBLISTA!Lucro), 0, TBLISTA!Lucro) 'Variação total
                Valor_CSLL_Serv = IIf(IsNull(TBLISTA!material), 0, TBLISTA!material) 'Percentual
                
                .Item(.Count).SubItems(Contador1) = Format(valor, "###,##0.00") & "  |  " & Format(Valor1, "###,##0.00") & "  |  " & Format(Valor2, "###,##0.00") & "  |  " & Format(Valor3, "###,##0.00") & "%"
            
                'Carrega valor total
                Contador1 = 1
                Do While Lista_resumido.ColumnHeaders(Contador1 + 1).Text <> "Vlr. total orçado  |  Real  |  Var.  |  Perc."
                    Contador1 = Contador1 + 1
                Loop
                
                .Item(.Count).SubItems(Contador1) = Format(Valor_Cofins_Prod, "###,##0.00") & "  |  " & Format(Valor_Cofins_Serv, "###,##0.00") & "  |  " & Format(Valor_CSLL_Prod, "###,##0.00") & "  |  " & Format(Valor_CSLL_Serv, "###,##0.00") & "%"
            End With
        End If
        Familiatext = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_total_orcado = IIf(IsNull(TBLISTA!QtdeOrdem), 0, Format(TBLISTA!QtdeOrdem, "###,##0.00")) 'Total orçado
    Txt_total_real = IIf(IsNull(TBLISTA!CustoMat), 0, Format(TBLISTA!CustoMat, "###,##0.00")) 'Total real
    Txt_variacao_total = IIf(IsNull(TBLISTA!CustoObra), 0, Format(TBLISTA!CustoObra, "###,##0.00")) 'Variação Total
    Txt_percentual = IIf(IsNull(TBLISTA!Terceros), 0, Format(TBLISTA!Terceros, "###,##0.00")) & "%" 'Percentual
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaListaeCampos()
On Error GoTo tratar_erro

Lista_resumido.ListItems.Clear
Lista_res_PC.ListItems.Clear
Lista_det_orcado.ListItems.Clear
Lista_det_real.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
Txt_total_orcado = ""
Txt_total_real = ""
Txt_variacao_total = ""
Txt_percentual = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True

Formulario = "Custos/Relatórios/Previsto x Realizado"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_ano_de.Clear
Cmb_ano_ate.Clear
AnoAtual = 2010
Do While AnoAtual <> (Year(Date) + 4)
    Cmb_ano_de.AddItem AnoAtual
    Cmb_ano_ate.AddItem AnoAtual
    AnoAtual = AnoAtual + 1
Loop
Cmb_ano_ate = Year(Date)
Cmb_ano_de = Year(Date)

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    ProcCarregaComboSetor cmbTexto, "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "", False, True, False, "", True, False
Else
    ProcCarregaComboSetor cmbTexto, "US.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "", False, True, True, pubUsuario, True, False
End If
TBAcessos.Close

SSTab1.Tab = 0

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Custos/Relatórios/Previsto x Realizado"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_det_orcado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_det_orcado, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_det_real_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_det_real, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_det_real_DblClick()
On Error GoTo tratar_erro

If Lista_det_real.ListItems.Count = 0 Then Exit Sub
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select ID_estoque, ID_financeiro, operacao from CC_realizado where ID = " & Lista_det_real.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If IsNull(TBTempo!ID_estoque) = False And TBTempo!ID_estoque <> 0 Then
        Set TBAfericao = CreateObject("adodb.recordset")
        TBAfericao.Open "Select Entrada, Documento, Lote from Estoque_movimentacao where Idoperacao = " & TBTempo!ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
        If TBAfericao.EOF = False Then
            If TBAfericao!Entrada > 0 Then
                Formulario = "Compras/Pedido"
                ProcLiberaAcessos True
                If Acessos = False Then Exit Sub
                With frmCompras_Pedido
                    Set TBCompras_Pedido = CreateObject("adodb.recordset")
                    TBCompras_Pedido.Open "Select * from compras_pedido where Pedido = '" & TBAfericao!LOTE & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras_Pedido.EOF = False Then
                        .ProcLimpar
                        .ProcLimpaCamposItem False
                        .ProcLimpaCamposServ False
                        .ProcPuxaDados
                        .Sql_Pedido_Localizar = "Select CP.IDpedido, CP.Data, CP.Pedido, CC.Cotacaotexto, CP.Fornecedor, CP.Status_pedido, CP.DtValidacao, CP.Data_aprovado from Compras_pedido CP LEFT JOIN Compras_cotacao CC ON CC.ID_cotacao = CP.IDcotacao where CP.IDpedido = " & TBCompras_Pedido!IDpedido
                        .ProcAtualizalistapedido (1)
                    End If
                    TBCompras_Pedido.Close
                End With
            Else
                If Left(TBAfericao!Documento, 2) = "RM" Then
                    Formulario = "Estoque/Requisição de materiais"
                    ProcLiberaAcessos True
                    If Acessos = False Then Exit Sub
                    With frmRequisicao_materiais
                        CC_RM = True
                        ID_CC = Lista_resumido.SelectedItem
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Requisicao_materiais where requisicao = '" & TBAfericao!Documento & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .ProcLimpaCampos
                            .ProcPuxaDados
                            CamposFiltro = "RM.ID, RM.requisicao, RM.Data, RM.Responsavel, RM.Status, RM.Dtvalidacao, E.Empresa"
                            .StrSql_Localizar_Requisicao = "Select " & CamposFiltro & " from Requisicao_materiais RM INNER JOIN Empresa E ON E.Codigo = RM.ID_empresa where ID = " & TBAbrir!ID & " group by " & CamposFiltro
                            .ProcCarregaLista (1)
                        End If
                        TBAbrir.Close
                    End With
                End If
            End If
        End If
    Else
        If TBTempo!Operacao = "Crédito" Then
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_Contas_receber where IdIntConta = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If TBContas!Logsit = "N" Then
                    Formulario = "Financeiro/Contas a receber"
                    ProcLiberaAcessos True
                    If Acessos = False Then Exit Sub
                    With frmContas_Receber
                        .Show
                        .ProcLiberaBotao
                        .ProcLimpaCampos
                        .ProcFiltrarContas "IdIntConta = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), "{tbl_Contas_receber.IdIntConta} = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), True, False, False, False, TBContas!Vencimento, TBContas!Vencimento, "vencimento"
                        .ProcCarregaDados
                    End With
                Else
                    Formulario = "Financeiro/Contas recebidas"
                    ProcLiberaAcessos True
                    If Acessos = False Then Exit Sub
                    With frmContas_recebidas
                        .Show
                        .ProcLimpaCampos
                        .ProcFiltrarContas "IdIntConta = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), "{tbl_Contas_receber.IdIntConta} = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), True, False, False, False, False, TBContas!Data_pagamento, TBContas!Data_pagamento, "data_pagamento"
                        .ProcCarregaDados
                    End With
                End If
            End If
        Else
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If TBContas!Logsit = "N" Then
                    Formulario = "Financeiro/Contas a pagar"
                    ProcLiberaAcessos True
                    If Acessos = False Then Exit Sub
                    With frmContas_Pagar
                        .Show
                        .ProcLiberaBotao
                        .ProcLimpaCampos
                        .ProcCarregaDados
                        .ProcFiltrarContas "IdIntConta = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), "{tbl_ContasPagar.IdIntConta} = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), True, False, False, False, TBContas!dt_Pagamento, TBContas!dt_Pagamento, "dt_Pagamento"
                    End With
                Else
                    Formulario = "Financeiro/Contas pagas"
                    ProcLiberaAcessos True
                    If Acessos = False Then Exit Sub
                    With frmContas_Pagas
                        .Show
                        .ProcLimpaCampos
                        .ProcFiltrarContas "IdIntConta = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), "{tbl_ContasPagar.IdIntConta} = " & IIf(IsNull(TBTempo!ID_financeiro), 0, TBTempo!ID_financeiro), True, False, False, False, False, TBContas!DataBaixa, TBContas!DataBaixa, "DataBaixa"
                        .ProcCarregaDados
                    End With
                End If
            End If
        End If
    End If
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_resumido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_resumido, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Opt_individual.Value = True And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If
If Txt_rev_prev = "" Then
    NomeCampo = "a revisão prevista"
    ProcVerificaAcao
    Txt_rev_prev.SetFocus
    Exit Sub
End If
If Cmb_mes_de = "" Then
    NomeCampo = "o mês"
    ProcVerificaAcao
    Cmb_mes_de.SetFocus
    Exit Sub
End If
If Cmb_mes_ate = "" Then
    NomeCampo = "o mês"
    ProcVerificaAcao
    Cmb_mes_ate.SetFocus
    Exit Sub
End If
If Cmb_ano_de = "" Then
    NomeCampo = "o ano"
    ProcVerificaAcao
    Cmb_ano_de.SetFocus
    Exit Sub
End If

Inicio = Time
Desenho = ""
ProcLimpaListaeCampos
ProcAbrirTabelas
    
ProcCriaColunas

'Soma e grava o total geral
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Fase, maquina, Sum(QtdeOK) as Valor, Sum(QtdeNC) as Valor1 from Producao_relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' Group by Fase, Maquina", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        valor = IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor) 'Valor orçado
        Valor1 = IIf(IsNull(TBLISTA!Valor1), 0, TBLISTA!Valor1) 'Valor real
        NovoValor = Replace(valor, ",", ".")
        NovoValor1 = Replace(Valor1, ",", ".")
        
        'Variacao
        Valor2 = valor - Valor1
        NovoValor2 = Replace(Valor2, ",", ".")
        
        'Percentual
        If valor = 0 Then
            Valor_total = -100
        ElseIf valor <> 0 And Valor2 <> 0 Then
                Valor_total = (Valor2 / valor) * 100
            Else
                Valor_total = 0
        End If
        NovoValor3 = Replace(Valor_total, ",", ".")
        
        Conexao.Execute "Update Producao_relatorios Set Qtdetotalprod = " & NovoValor & ", OS = " & NovoValor1 & ", Lucro = " & NovoValor2 & ", Material = " & NovoValor3 & " where Fase = " & TBLISTA!Fase & " and Maquina = '" & TBLISTA!maquina & "'"
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

If Permitido = True Then ProcGravarTotalizacoes

Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True Then
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Nota = 1 order by Maquina", Conexao, adOpenKeyset, adLockReadOnly
Else
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Nota = 1 and Maquina is not Null order by Maquina", Conexao, adOpenKeyset, adLockReadOnly
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

qt = FunVerificaMes(Cmb_mes_de)
Qtd = FunVerificaMes(Cmb_mes_ate)
MesX = qt
MesX1 = Qtd

Par1 = ""
Permitido = False

Do While qt <= Qtd
    If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
    Permitido = True
    qt = qt + 1
Loop
Pesquisa = "Mes >= '" & MesX & "' and Ano >= '" & Cmb_ano_de & "' and Mes <= '" & MesX1 & "' and Ano <= '" & Cmb_ano_ate & "'"
Pesquisa1 = "PIVOT (Sum(Valor) for Mes In (" & Par1 & "))"

Set TBCarteira = CreateObject("adodb.recordset")
Set TBGravar = CreateObject("adodb.recordset")
If Opt_individual.Value = True Then
    'Previsto
    TBCarteira.Open "SELECT ID_CC, Revisao, " & Par1 & " From (Select ID_CC, Mes, Ano, Revisao, Valor from Usuarios_Setor_Previsao Where ID_CC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and Revisao = " & Txt_rev_prev & " and " & Pesquisa & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
    
    'Realizado
    TBGravar.Open "SELECT Operacao, ID_CC, Setor, " & Par1 & " From (Select Operacao, ID_CC, Setor, Mes, Ano, Valor from Centro_de_custo_real_resumido Where ID_CC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and " & Pesquisa & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
Else
    'Previsto
    TBCarteira.Open "SELECT ID_CC, Revisao, " & Par1 & " From (Select ID_CC, Mes, Ano, Revisao, Valor from Usuarios_Setor_Previsao Where Revisao = " & Txt_rev_prev & " and " & Pesquisa & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
    
    TextoFiltro = ""
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Usuarios_setor.ID from Usuarios_setor INNER JOIN Usuarios_Setor_Responsavel ON Usuarios_setor.ID = Usuarios_Setor_Responsavel.ID_CC where Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' and Usuarios_setor.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by Usuarios_setor.Codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TextoFiltro = "" Then TextoFiltro = "ID_CC = " & TBAbrir!ID Else TextoFiltro = TextoFiltro & "or ID_CC = " & TBAbrir!ID
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    Else
        TextoFiltro = "ID_CC is not null"
    End If
    TBAcessos.Close
        
    'Realizado
    TBGravar.Open "SELECT Operacao, ID_CC, Setor, " & Par1 & " From (Select Operacao, ID_CC, Setor, Mes, Ano, Valor from Centro_de_custo_real_resumido Where " & Pesquisa & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
End If
ProcFiltrar1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

FamiliaAntiga = ""
If TBCarteira.EOF = False Then
    Permitido = True
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
        ProcCriarResumidoOrcado
        If FamiliaAntiga <> TBCarteira!ID_CC Then ProcAbrirTabelasPCOrcado
        FamiliaAntiga = TBCarteira!ID_CC
        TBCarteira.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBCarteira.Close

FamiliaAntiga = ""
If TBGravar.EOF = False Then
    Permitido = True
    TBGravar.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBGravar.RecordCount
    PBLista.Value = 1
    contador = 0
    TBGravar.MoveFirst
    Do While TBGravar.EOF = False
        ProcCriarResumidoReal
        If FamiliaAntiga <> TBGravar!ID_CC Then ProcAbrirTabelasPCReal
        FamiliaAntiga = TBGravar!ID_CC
        TBGravar.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarResumidoOrcado()
On Error GoTo tratar_erro

Permitido = True
qt = MesX
Qtd = MesX1
Do While qt <= Qtd
    ProcEnviaDadosResumidoOrcado
    qt = qt + 1
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarResumidoReal()
On Error GoTo tratar_erro

Permitido = True
qt = MesX
Qtd = MesX1
Do While qt <= Qtd
    ProcEnviaDadosResumidoReal
    qt = qt + 1
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumidoOrcado()
On Error GoTo tratar_erro

DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
DataTexto = "01/" & qt & "/" & Cmb_ano_de

'Verifica descrição do centro de custo
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Setor from Usuarios_Setor where Id = " & TBCarteira!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Familiatext = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
End If
TBAbrir.Close

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
TBProdutividade!Nota = 1
TBProdutividade!Fase = TBCarteira!ID_CC
Select Case qt
    Case 1: TotalOrcado = IIf(IsNull(TBCarteira![1]), 0, Format(TBCarteira![1], "###,##0.00"))
    Case 2: TotalOrcado = IIf(IsNull(TBCarteira![2]), 0, Format(TBCarteira![2], "###,##0.00"))
    Case 3: TotalOrcado = IIf(IsNull(TBCarteira![3]), 0, Format(TBCarteira![3], "###,##0.00"))
    Case 4: TotalOrcado = IIf(IsNull(TBCarteira![4]), 0, Format(TBCarteira![4], "###,##0.00"))
    Case 5: TotalOrcado = IIf(IsNull(TBCarteira![5]), 0, Format(TBCarteira![5], "###,##0.00"))
    Case 6: TotalOrcado = IIf(IsNull(TBCarteira![6]), 0, Format(TBCarteira![6], "###,##0.00"))
    Case 7: TotalOrcado = IIf(IsNull(TBCarteira![7]), 0, Format(TBCarteira![7], "###,##0.00"))
    Case 8: TotalOrcado = IIf(IsNull(TBCarteira![8]), 0, Format(TBCarteira![8], "###,##0.00"))
    Case 9: TotalOrcado = IIf(IsNull(TBCarteira![9]), 0, Format(TBCarteira![9], "###,##0.00"))
    Case 10: TotalOrcado = IIf(IsNull(TBCarteira![10]), 0, Format(TBCarteira![10], "###,##0.00"))
    Case 11: TotalOrcado = IIf(IsNull(TBCarteira![11]), 0, Format(TBCarteira![11], "###,##0.00"))
    Case 12: TotalOrcado = IIf(IsNull(TBCarteira![12]), 0, Format(TBCarteira![12], "###,##0.00"))
End Select
TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
TBProdutividade!qtdeOK = IIf(IsNull(TotalOrcado), 0, Format(TotalOrcado, "###,##0.00"))
TBProdutividade!Ordem = qt

TBProdutividade!maquina = Familiatext

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumidoReal()
On Error GoTo tratar_erro

DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
DataTexto = "01/" & qt & "/" & Cmb_ano_de

Familiatext = IIf(IsNull(TBGravar!Setor), "", TBGravar!Setor)

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
TBProdutividade!Nota = 1
TBProdutividade!Fase = TBGravar!ID_CC
Select Case qt
    Case 1: TotalReal = IIf(IsNull(TBGravar![1]), 0, Format(TBGravar![1], "###,##0.00"))
    Case 2: TotalReal = IIf(IsNull(TBGravar![2]), 0, Format(TBGravar![2], "###,##0.00"))
    Case 3: TotalReal = IIf(IsNull(TBGravar![3]), 0, Format(TBGravar![3], "###,##0.00"))
    Case 4: TotalReal = IIf(IsNull(TBGravar![4]), 0, Format(TBGravar![4], "###,##0.00"))
    Case 5: TotalReal = IIf(IsNull(TBGravar![5]), 0, Format(TBGravar![5], "###,##0.00"))
    Case 6: TotalReal = IIf(IsNull(TBGravar![6]), 0, Format(TBGravar![6], "###,##0.00"))
    Case 7: TotalReal = IIf(IsNull(TBGravar![7]), 0, Format(TBGravar![7], "###,##0.00"))
    Case 8: TotalReal = IIf(IsNull(TBGravar![8]), 0, Format(TBGravar![8], "###,##0.00"))
    Case 9: TotalReal = IIf(IsNull(TBGravar![9]), 0, Format(TBGravar![9], "###,##0.00"))
    Case 10: TotalReal = IIf(IsNull(TBGravar![10]), 0, Format(TBGravar![10], "###,##0.00"))
    Case 11: TotalReal = IIf(IsNull(TBGravar![11]), 0, Format(TBGravar![11], "###,##0.00"))
    Case 12: TotalReal = IIf(IsNull(TBGravar![12]), 0, Format(TBGravar![12], "###,##0.00"))
End Select
TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
If TBGravar!Operacao = "Crédito" Then Valor2 = TBProdutividade!qtdeNC - TotalReal Else Valor2 = TBProdutividade!qtdeNC + TotalReal
TBProdutividade!qtdeNC = Valor2

Valor1 = TBProdutividade!qtdeOK 'Orçado

'Variação
Valor_Cofins_Prod = Valor1 - Valor2
TBProdutividade!Terceiros = Valor_Cofins_Prod

'Percentual
If Valor1 = 0 Then
    Valor_total = -100
ElseIf Valor1 <> 0 And Valor_Cofins_Prod <> 0 Then
        Valor_total = (Valor_Cofins_Prod / Valor1) * 100
    Else
        Valor_total = 0
End If
TBProdutividade!impostos = Valor_total

TBProdutividade!Ordem = qt
TBProdutividade!maquina = Familiatext

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaColunas()
On Error GoTo tratar_erro

Lista_resumido.ColumnHeaders.Clear
contador = 1

With Lista_resumido.ColumnHeaders
    .Add
    .Item(contador).Text = "ID_CC"
    .Item(contador).Width = 0
    contador = 2
    .Add
    .Item(contador).Text = "Centro de custo"
    .Item(contador).Width = 3500
    
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    Do While qt <= Qtd
        .Add
        contador = contador + 1
        .Item(contador).Text = qt & "/" & Cmb_ano_de
        .Item(contador).Width = 3800
        .Item(contador).Alignment = lvwColumnRight
        qt = qt + 1
    Loop
    
    .Add
    contador = contador + 1
    .Item(contador).Text = "Vlr. total orçado  |  Real  |  Var.  |  Perc."
    .Item(contador).Width = 3800
    .Item(contador).Alignment = lvwColumnRight
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

MesX = FunVerificaMes(Cmb_mes_de)
MesX1 = FunVerificaMes(Cmb_mes_ate)
If Opt_individual.Value = True Then
    Data_pesquisa = "ID_CC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and Mes >= '" & MesX & "' And Ano >= '" & Cmb_ano_de & "' and Mes <= '" & MesX1 & "' And Ano <= '" & Cmb_ano_ate & "'"
Else
    Data_pesquisa = "Mes >= '" & MesX & "' And Ano >= '" & Cmb_ano_de & "' and Mes <= '" & MesX1 & "' And Ano <= '" & Cmb_ano_ate & "'"
End If
        
'Total orçado
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(Valor) as Valor1 from Usuarios_Setor_Previsao where " & Data_pesquisa & " and Revisao = " & Txt_rev_prev, Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Valor1 = IIf(IsNull(TBproducao!Valor1), 0, TBproducao!Valor1)
End If
        
'Total real
'Débito
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(Valor) as Valor2 from Centro_de_custo_real_resumido where " & Data_pesquisa & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Valor2 = IIf(IsNull(TBproducao!Valor2), 0, TBproducao!Valor2)
End If

'Crédito
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(Valor) as Valor2 from Centro_de_custo_real_resumido where " & Data_pesquisa & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Operacao = 'Crédito'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Valor3 = IIf(IsNull(TBproducao!Valor2), 0, TBproducao!Valor2)
End If
TBproducao.Close
Valor2 = Valor2 - Valor3

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Totalutilizada = Cmb_mes_de & "/" & Cmb_ano_de
TBAbrir!Totalprevista = Cmb_mes_ate & "/" & Cmb_ano_ate
If Opt_individual.Value = True Then TBAbrir!Texto = "Centro de custo) : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
TBAbrir!QtdeOrdem = Valor1 'Total orçado
TBAbrir!QtdeProduzida = Txt_rev_prev 'Rev. orçado
TBAbrir!CustoMat = Valor2 'Total real
'Variação
Valor_Cofins_Prod = Valor1 - Valor2
TBAbrir!CustoObra = Valor_Cofins_Prod

'Percentual
If Valor1 = 0 Then
    Valor_total = -100
ElseIf Valor1 <> 0 And Valor_Cofins_Prod <> 0 Then
        Valor_total = (Valor_Cofins_Prod / Valor1) * 100
    Else
        Valor_total = 0
End If
TBAbrir!Terceros = Valor_total

TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelasPCOrcado()
On Error GoTo tratar_erro

qt = FunVerificaMes(Cmb_mes_de)
Qtd = FunVerificaMes(Cmb_mes_ate)
MesX = qt
MesX1 = Qtd

Par1 = ""
Permitido = False

Do While qt <= Qtd
    If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
    Permitido = True
    qt = qt + 1
Loop
Pesquisa = "Mes >= '" & MesX & "' and Ano >= '" & Cmb_ano_de & "' and Mes <= '" & MesX1 & "' and Ano <= '" & Cmb_ano_ate & "'"
Pesquisa1 = "PIVOT (Sum(Valor) for Mes In (" & Par1 & "))"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT ID_PC, Revisao, " & Par1 & " From (Select ID_PC, Mes, Ano, Revisao, Valor from Usuarios_Setor_Previsao Where ID_CC = " & TBCarteira!ID_CC & " and Revisao = " & TBCarteira!Revisao & " and " & Pesquisa & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
ProcFiltrar1PCOrcado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1PCOrcado()
On Error GoTo tratar_erro

If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        ProcCriarResumidoPCOrcado
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarResumidoPCOrcado()
On Error GoTo tratar_erro

Permitido = True
qt = MesX
Qtd = MesX1
Do While qt <= Qtd
    ProcEnviaDadosResumidoPCOrcado
    qt = qt + 1
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumidoPCOrcado()
On Error GoTo tratar_erro

DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
DataTexto = "01/" & qt & "/" & Cmb_ano_de

'Verifica descrição da conta contábil
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Codigo, Txt_descricao, Nivel from tbl_familia where int_codfamilia = " & TBAbrir!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Select Case TBFI!Nivel
        Case 1: CODIGO = Left(TBFI!CODIGO, 1)
        Case 2: CODIGO = Left(TBFI!CODIGO, 4)
        Case 3: CODIGO = Left(TBFI!CODIGO, 7)
        Case 4: CODIGO = Left(TBFI!CODIGO, 10)
        Case 5: CODIGO = Left(TBFI!CODIGO, 13)
        Case 6: CODIGO = Left(TBFI!CODIGO, 16)
        Case 7: CODIGO = Left(TBFI!CODIGO, 19)
        Case 8: CODIGO = Left(TBFI!CODIGO, 22)
    End Select
    Familiatext = CODIGO & " - " & TBFI!Txt_descricao
End If
TBFI.Close

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
TBProdutividade!Nota = 2
TBProdutividade!Fase = TBCarteira!ID_CC
TBProdutividade!Turno = TBAbrir!ID_PC
Select Case qt
    Case 1: TotalPCOrcado = IIf(IsNull(TBAbrir![1]), 0, Format(TBAbrir![1], "###,##0.00"))
    Case 2: TotalPCOrcado = IIf(IsNull(TBAbrir![2]), 0, Format(TBAbrir![2], "###,##0.00"))
    Case 3: TotalPCOrcado = IIf(IsNull(TBAbrir![3]), 0, Format(TBAbrir![3], "###,##0.00"))
    Case 4: TotalPCOrcado = IIf(IsNull(TBAbrir![4]), 0, Format(TBAbrir![4], "###,##0.00"))
    Case 5: TotalPCOrcado = IIf(IsNull(TBAbrir![5]), 0, Format(TBAbrir![5], "###,##0.00"))
    Case 6: TotalPCOrcado = IIf(IsNull(TBAbrir![6]), 0, Format(TBAbrir![6], "###,##0.00"))
    Case 7: TotalPCOrcado = IIf(IsNull(TBAbrir![7]), 0, Format(TBAbrir![7], "###,##0.00"))
    Case 8: TotalPCOrcado = IIf(IsNull(TBAbrir![8]), 0, Format(TBAbrir![8], "###,##0.00"))
    Case 9: TotalPCOrcado = IIf(IsNull(TBAbrir![9]), 0, Format(TBAbrir![9], "###,##0.00"))
    Case 10: TotalPCOrcado = IIf(IsNull(TBAbrir![10]), 0, Format(TBAbrir![10], "###,##0.00"))
    Case 11: TotalPCOrcado = IIf(IsNull(TBAbrir![11]), 0, Format(TBAbrir![11], "###,##0.00"))
    Case 12: TotalPCOrcado = IIf(IsNull(TBAbrir![12]), 0, Format(TBAbrir![12], "###,##0.00"))
End Select
TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
TBProdutividade!qtdeOK = IIf(IsNull(TotalPCOrcado), 0, Format(TotalPCOrcado, "###,##0.00"))
TBProdutividade!Ordem = qt

TBProdutividade!maquina = Familiatext

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelasPCReal()
On Error GoTo tratar_erro

qt = FunVerificaMes(Cmb_mes_de)
Qtd = FunVerificaMes(Cmb_mes_ate)
MesX = qt
MesX1 = Qtd

Par1 = ""
Permitido = False

Do While qt <= Qtd
    If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
    Permitido = True
    qt = qt + 1
Loop
Pesquisa = "Mes >= '" & MesX & "' and Ano >= '" & Cmb_ano_de & "' and Mes <= '" & MesX1 & "' and Ano <= '" & Cmb_ano_ate & "'"
Pesquisa1 = "PIVOT (Sum(Valor) for Mes In (" & Par1 & "))"

'Estoque
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT Operacao, ID_estoque, ID_financeiro, ID_PC, Codigo, txt_descricao, Nivel, " & Par1 & " From (Select Operacao, ID_estoque, ID_financeiro, ID_PC, Codigo, txt_descricao, Nivel, Mes, Ano, Valor from Centro_de_custo_real_esoque_PC_resumido Where ID_CC = " & TBGravar!ID_CC & " and " & Pesquisa & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
ProcFiltrar1PCReal

'Financeiro
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT Operacao, ID_estoque, ID_financeiro, ID_PC, Codigo, txt_descricao, Nivel, " & Par1 & " From (Select Operacao, ID_estoque, ID_financeiro, ID_PC, Codigo, txt_descricao, Nivel, Mes, Ano, Valor from Centro_de_custo_real_financeiro_PC_resumido Where ID_CC = " & TBGravar!ID_CC & " and ID_financeiro is not null and " & Pesquisa & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
ProcFiltrar1PCReal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1PCReal()
On Error GoTo tratar_erro

If TBAbrir.EOF = False Then
    Permitido = True
    Do While TBAbrir.EOF = False
        ProcCriarResumidoPCReal
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarResumidoPCReal()
On Error GoTo tratar_erro

Permitido = True
qt = MesX
Qtd = MesX1
Do While qt <= Qtd
    Qtde = 0
    If IsNull(TBAbrir!ID_financeiro) = False And TBAbrir!ID_financeiro <> "" Then
        'Verifica valor total da conta
        Qtde = 0
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select LogSit, dbl_valorpagto, ValorPago from tbl_ContasPagar  where IDintconta = " & IIf(IsNull(TBAbrir!ID_financeiro), 0, TBAbrir!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            If TBFI!Logsit = "S" Then Qtde = TBFI!ValorPago Else Qtde = TBFI!dbl_valorpagto
            
            'Verifica valor em percentual por CC
            If Qtde > 0 Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Valor from CC_realizado where ID_CC = " & TBGravar!ID_CC & " and CC_realizado.ID_financeiro = " & IIf(IsNull(TBAbrir!ID_financeiro), 0, TBAbrir!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Qtde = (TBFI!valor / Qtde) * 100
                End If
            End If
        End If
        TBFI.Close
    End If
    ProcEnviaDadosResumidoPCReal
    qt = qt + 1
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumidoPCReal()
On Error GoTo tratar_erro

DataFiltro = "Execucaoprev = '" & qt & "/" & Cmb_ano_de & "'"
DataTexto = "01/" & qt & "/" & Cmb_ano_de

Select Case TBAbrir!Nivel
    Case 1: CODIGO = Left(TBAbrir!CODIGO, 1)
    Case 2: CODIGO = Left(TBAbrir!CODIGO, 4)
    Case 3: CODIGO = Left(TBAbrir!CODIGO, 7)
    Case 4: CODIGO = Left(TBAbrir!CODIGO, 10)
    Case 5: CODIGO = Left(TBAbrir!CODIGO, 13)
    Case 6: CODIGO = Left(TBAbrir!CODIGO, 16)
    Case 7: CODIGO = Left(TBAbrir!CODIGO, 19)
    Case 8: CODIGO = Left(TBAbrir!CODIGO, 22)
End Select
Familiatext = CODIGO & " - " & TBAbrir!Txt_descricao

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Fase = " & TBGravar!ID_CC & " and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")
TBProdutividade!Nota = 2
TBProdutividade!Fase = TBGravar!ID_CC
TBProdutividade!Turno = TBAbrir!ID_PC
Select Case qt
    Case 1: TotalReal = IIf(IsNull(TBAbrir![1]), 0, Format(TBAbrir![1], "###,##0.00"))
    Case 2: TotalReal = IIf(IsNull(TBAbrir![2]), 0, Format(TBAbrir![2], "###,##0.00"))
    Case 3: TotalReal = IIf(IsNull(TBAbrir![3]), 0, Format(TBAbrir![3], "###,##0.00"))
    Case 4: TotalReal = IIf(IsNull(TBAbrir![4]), 0, Format(TBAbrir![4], "###,##0.00"))
    Case 5: TotalReal = IIf(IsNull(TBAbrir![5]), 0, Format(TBAbrir![5], "###,##0.00"))
    Case 6: TotalReal = IIf(IsNull(TBAbrir![6]), 0, Format(TBAbrir![6], "###,##0.00"))
    Case 7: TotalReal = IIf(IsNull(TBAbrir![7]), 0, Format(TBAbrir![7], "###,##0.00"))
    Case 8: TotalReal = IIf(IsNull(TBAbrir![8]), 0, Format(TBAbrir![8], "###,##0.00"))
    Case 9: TotalReal = IIf(IsNull(TBAbrir![9]), 0, Format(TBAbrir![9], "###,##0.00"))
    Case 10: TotalReal = IIf(IsNull(TBAbrir![10]), 0, Format(TBAbrir![10], "###,##0.00"))
    Case 11: TotalReal = IIf(IsNull(TBAbrir![11]), 0, Format(TBAbrir![11], "###,##0.00"))
    Case 12: TotalReal = IIf(IsNull(TBAbrir![12]), 0, Format(TBAbrir![12], "###,##0.00"))
End Select
TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
If IsNull(TBAbrir!ID_estoque) = False And TBAbrir!ID_estoque <> "" Then
    If TBAbrir!Operacao = "Crédito" Then Valor2 = TBProdutividade!qtdeNC - TotalReal Else Valor2 = TBProdutividade!qtdeNC + TotalReal
Else
    Valor2 = TBProdutividade!qtdeNC + Format((TotalReal * Qtde) / 100, "###,##0.00")
End If
TBProdutividade!qtdeNC = Valor2

Valor1 = TBProdutividade!qtdeOK 'Orçado

'Variação
Valor_Cofins_Prod = Valor1 - Valor2
TBProdutividade!Terceiros = Valor_Cofins_Prod

'Percentual
If Valor1 = 0 Then
    Valor_total = -100
ElseIf Valor1 <> 0 And Valor_Cofins_Prod <> 0 Then
        Valor_total = (Valor_Cofins_Prod / Valor1) * 100
    Else
        Valor_total = 0
End If
TBProdutividade!impostos = Valor_total

TBProdutividade!Ordem = qt
TBProdutividade!maquina = Familiatext

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_resumido_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_resumido.ListItems.Count = 0 Then Exit Sub

'Resumido
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PR.* from Producao_Relatorios PR INNER JOIN Usuarios_Setor US ON US.ID = PR.Fase where US.ID = " & Lista_resumido.SelectedItem & " and PR.Responsavel = '" & pubUsuario & "' and PR.Modulo = '" & Formulario & "' and PR.Nota = 2 order by PR.Maquina", Conexao, adOpenKeyset, adLockOptimistic
ProcCarregaListaResumidoPC

qt = FunVerificaMes(Cmb_mes_de)

'Detalhado
'Orçado
Lista_det_orcado.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select USP.*, F.Codigo, F.txt_descricao from (Usuarios_setor_previsao USP INNER JOIN tbl_familia F ON USP.ID_PC = F.int_codfamilia) INNER JOIN Usuarios_Setor US ON US.ID = USP.ID_CC where US.ID = " & Lista_resumido.SelectedItem & " and USP.Mes >= " & qt & " and USP.Ano >= " & Cmb_ano_de & " and USP.Mes <= " & Qtd & " and USP.Ano <= " & Cmb_ano_ate & " and USP.Revisao = " & Txt_rev_prev & " order by USP.Ano desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista_det_orcado.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            If IsNull(TBLISTA!Mes) = False And TBLISTA!Mes <> "" Then .Item(.Count).SubItems(3) = FunVerificaNumeroMes(TBLISTA!Mes)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ano), "", TBLISTA!Ano)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

'Real
Conexao.Execute "DELETE from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Nota = '3'"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir!Texto1 = Lista_resumido.SelectedItem
    TBAbrir.Update
End If
TBAbrir.Close

Lista_det_real.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CCR.* from CC_realizado CCR INNER JOIN Usuarios_Setor US ON US.ID = CCR.ID_CC where US.ID = " & Lista_resumido.SelectedItem & " and Month(CCR.Data) >= " & qt & " and Year(CCR.Data) >= " & Cmb_ano_de & " and Month(CCR.Data) <= " & Qtd & " and Year(CCR.Data) <= " & Cmb_ano_ate & " and (CCR.Bloqueado IS NULL or CCR.Bloqueado = 0) order by CCR.Data desc, CCR.ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        
        'Grava dados detalhado (realizado) na tabela
        Set TBProdutividade = CreateObject("adodb.recordset")
        TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
        TBProdutividade.AddNew
        TBProdutividade!Responsavel = pubUsuario
        TBProdutividade!Modulo = Formulario
        TBProdutividade!Execucaoprev = Lista_resumido.SelectedItem 'Centro de custo
        TBProdutividade!Nota = 3
        TBProdutividade!Data = IIf(IsNull(TBLISTA!Data), Null, TBLISTA!Data) 'Data
        TBProdutividade!Totalhsprev = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel) 'Responsável
        'Módulo
        If IsNull(TBLISTA!ID_estoque) = False And TBLISTA!ID_estoque <> 0 Then
            TBProdutividade!maquina = "Estoque"
        Else
            TBProdutividade!maquina = "Financeiro"
        End If
        
        'Verifica se existe o PC realizado no previsto
        Permitido = True
        If IsNull(TBLISTA!ID_PC) = False And TBLISTA!ID_PC <> "" And TBLISTA!ID_PC <> "0" Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select USP.ID from (Usuarios_setor_previsao USP INNER JOIN tbl_familia F ON USP.ID_PC = F.int_codfamilia) INNER JOIN Usuarios_Setor US ON US.ID = USP.ID_CC where US.ID = " & Lista_resumido.SelectedItem & " and USP.Mes = " & Month(TBProdutividade!Data) & " and USP.Ano = " & Year(TBProdutividade!Data) & " and USP.Revisao = " & Txt_rev_prev & " and USP.ID_PC = " & TBLISTA!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                Permitido = False
            End If
            TBAbrir.Close
        End If
        
        If IsNull(TBLISTA!ID_estoque) = False And TBLISTA!ID_estoque <> 0 Then
            With Lista_det_real.ListItems.Add(, , TBLISTA!ID)
                .SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
                If IsNull(TBLISTA!ID_origem) = False And TBLISTA!ID_origem <> 0 Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select US.Setor from Usuarios_Setor US INNER JOIN CC_realizado CCR ON US.Id = CCR.ID_CC where CCR.Id = " & TBLISTA!ID_origem, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .SubItems(3) = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
                        TBProdutividade!Totalhsutil = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor) 'CC origem
                    End If
                End If
                
                .SubItems(4) = "Estoque"
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Entrada, Documento, Lote from Estoque_movimentacao where Idoperacao = " & TBLISTA!ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If TBAbrir!Entrada > 0 Then DocumentoRef = "Ped. " & TBAbrir!LOTE Else DocumentoRef = TBAbrir!Documento
                End If
                .SubItems(5) = DocumentoRef
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Codigo, txt_descricao from tbl_familia where int_codfamilia = " & TBLISTA!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .SubItems(6) = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
                    TBProdutividade!Data1 = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO) 'Código contábil
                    .SubItems(7) = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
                    TBProdutividade!Data2 = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao) 'Descrição contábil
                End If
                
                Select Case TBLISTA!Operacao
                    Case "Crédito": ValorTexto = IIf(TBLISTA!valor > 0, "-", "") & IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                    Case "Débito": ValorTexto = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                End Select
                .SubItems(8) = ValorTexto
                TBProdutividade!OS = ValorTexto 'Valor
                
                If Permitido = False Then
                    .ForeColor = vbRed
                    .ListSubItems(1).ForeColor = vbRed
                    .ListSubItems(2).ForeColor = vbRed
                    .ListSubItems(3).ForeColor = vbRed
                    .ListSubItems(4).ForeColor = vbRed
                    .ListSubItems(5).ForeColor = vbRed
                    .ListSubItems(6).ForeColor = vbRed
                    .ListSubItems(7).ForeColor = vbRed
                    .ListSubItems(8).ForeColor = vbRed
                End If
            End With
        Else
            'Verifica perecentual do centro de custo na conta
            Contador2 = 0
            Qtde = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select CP.LogSit, CP.dbl_valorpagto, CP.ValorPago, CCR.Valor from (tbl_ContasPagar CP INNER JOIN CC_realizado CCR ON CP.IDintconta = CCR.ID_financeiro) INNER JOIN Usuarios_Setor US ON US.ID = CCR.ID_CC where US.ID = " & Lista_resumido.SelectedItem & " and CCR.ID_financeiro = " & IIf(IsNull(TBLISTA!ID_financeiro), 0, TBLISTA!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!Logsit = "S" Then Qtde = TBAbrir!ValorPago Else Qtde = TBAbrir!dbl_valorpagto
                If Qtde > 0 Then Qtde = (TBAbrir!valor / Qtde) * 100
            End If
            
            'Verifica número do documento e se a conta já foi paga/recebida
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select txt_NDocumento, Logsit from " & IIf(TBLISTA!Operacao = "Crédito", "tbl_Contas_Receber", "tbl_ContasPagar") & " where Idintconta = " & IIf(IsNull(TBLISTA!ID_financeiro), 0, TBLISTA!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!Logsit = "N" Then DocumentoRef = "Doc. " & TBAbrir!txt_ndocumento & " | N" Else DocumentoRef = "Doc. " & TBAbrir!txt_ndocumento & " | S"
            End If
            
            'Verifica contas contábeis na conta e dilui o valor por centro de custo
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select F.int_codfamilia, F.Codigo, F.txt_descricao, FF.Valor from tbl_familia F INNER JOIN Familia_financeiro FF ON F.int_codfamilia = FF.ID_PC where FF.IDConta = " & IIf(IsNull(TBLISTA!ID_financeiro), 0, TBLISTA!ID_financeiro) & " and FF.TipoConta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Contador2 = TBAbrir.RecordCount
                Qtde = Format(Qtde / Contador2, "###.##0.00")
                
                Do While TBAbrir.EOF = False
                    With Lista_det_real.ListItems.Add(, , TBLISTA!ID)
                        .SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                        .SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
                        
                        If IsNull(TBLISTA!ID_origem) = False And TBLISTA!ID_origem <> 0 Then
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select US.Setor from Usuarios_Setor US INNER JOIN CC_realizado CCR ON US.Id = CCR.ID_CC where CCR.Id = " & TBLISTA!ID_origem, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = False Then
                                .SubItems(3) = IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
                                TBProdutividade!Totalhsutil = IIf(IsNull(TBFI!Setor), "", TBFI!Setor) 'CC origem
                            End If
                            TBFI.Close
                        End If
                        
                        .SubItems(4) = "Financeiro"
                        .SubItems(5) = DocumentoRef
                        .SubItems(6) = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
                        TBProdutividade!Data1 = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO) 'Código contábil
                        .SubItems(7) = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
                        TBProdutividade!Data2 = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao) 'Descrição contábil
                        
                        Valor3 = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                        Valor_Cofins_Prod = IIf(IsNull(TBAbrir!valor), "", Format((TBAbrir!valor * Qtde) / 100, "###,##0.00"))
                        
                        Select Case TBLISTA!Operacao
                            Case "Crédito": ValorTexto = IIf(TBAbrir!valor > 0, "-", "") & IIf(IsNull(TBAbrir!valor), "", Format((TBAbrir!valor * Qtde) / 100, "###,##0.00"))
                            Case "Débito": ValorTexto = IIf(IsNull(TBAbrir!valor), "", Format((TBAbrir!valor * Qtde) / 100, "###,##0.00"))
                        End Select
                        .SubItems(8) = ValorTexto
                        TBProdutividade!OS = ValorTexto 'Valor
                        
                        'Verifica se existe o PC realizado no previsto
                        If IsNull(TBAbrir!int_codfamilia) = False And TBAbrir!int_codfamilia <> "" And TBAbrir!int_codfamilia <> "0" Then
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select USP.ID from (Usuarios_setor_previsao USP INNER JOIN tbl_familia F ON USP.ID_PC = F.int_codfamilia) INNER JOIN Usuarios_Setor US ON US.ID = USP.ID_CC where US.ID = " & Lista_resumido.SelectedItem & " and USP.Mes = " & Month(TBProdutividade!Data) & " and USP.Ano = " & Year(TBProdutividade!Data) & " and USP.Revisao = " & Txt_rev_prev & " and USP.ID_PC = " & TBAbrir!int_codfamilia, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = True Then
                                .ForeColor = vbRed
                                .ListSubItems(1).ForeColor = vbRed
                                .ListSubItems(2).ForeColor = vbRed
                                .ListSubItems(3).ForeColor = vbRed
                                .ListSubItems(4).ForeColor = vbRed
                                .ListSubItems(5).ForeColor = vbRed
                                .ListSubItems(6).ForeColor = vbRed
                                .ListSubItems(7).ForeColor = vbRed
                                .ListSubItems(8).ForeColor = vbRed
                            End If
                            TBFI.Close
                        End If
                    End With
                    TBAbrir.MoveNext
                Loop
            Else
                With Lista_det_real.ListItems.Add(, , TBLISTA!ID)
                    .SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                    .SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
                    
                    If IsNull(TBLISTA!ID_origem) = False And TBLISTA!ID_origem <> 0 Then
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select US.Setor from Usuarios_Setor US INNER JOIN CC_realizado CCR ON US.Id = CCR.ID_CC where CCR.Id = " & TBLISTA!ID_origem, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            .SubItems(3) = IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
                            TBProdutividade!Totalhsutil = IIf(IsNull(TBFI!Setor), "", TBFI!Setor) 'CC origem
                        End If
                        TBFI.Close
                    End If
                    
                    .SubItems(4) = "Financeiro"
                    .SubItems(5) = DocumentoRef
                    
                    Select Case TBLISTA!Operacao
                        Case "Crédito": ValorTexto = IIf(TBLISTA!valor > 0, "-", "") & IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                        Case "Débito": ValorTexto = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                    End Select
                    .SubItems(8) = ValorTexto
                    TBProdutividade!OS = ValorTexto 'Valor
                    
                    If Permitido = False Then
                        .ForeColor = vbRed
                        .ListSubItems(1).ForeColor = vbRed
                        .ListSubItems(2).ForeColor = vbRed
                        .ListSubItems(3).ForeColor = vbRed
                        .ListSubItems(4).ForeColor = vbRed
                        .ListSubItems(5).ForeColor = vbRed
                        .ListSubItems(6).ForeColor = vbRed
                        .ListSubItems(7).ForeColor = vbRed
                        .ListSubItems(8).ForeColor = vbRed
                    End If
                End With
            End If
        End If
        
        TBProdutividade!DescEvento = DocumentoRef 'Referência
        
        TBProdutividade.Update
        TBProdutividade.Close
        
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
    
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaColunasResumidoPC()
On Error GoTo tratar_erro

Lista_res_PC.ColumnHeaders.Clear
contador = 1

With Lista_res_PC.ColumnHeaders
    .Add
    .Item(contador).Text = "ID_PC"
    .Item(contador).Width = 0
    contador = 2
    .Add
    .Item(contador).Text = "Conta contábil"
    .Item(contador).Width = 3500
    
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    Do While qt <= Qtd
        .Add
        contador = contador + 1
        .Item(contador).Text = qt & "/" & Cmb_ano_de
        .Item(contador).Width = 3800
        .Item(contador).Alignment = lvwColumnRight
        qt = qt + 1
    Loop
    
    .Add
    contador = contador + 1
    .Item(contador).Text = "Vlr. total orçado  |  Real  |  Var.  |  Perc."
    .Item(contador).Width = 3800
    .Item(contador).Alignment = lvwColumnRight
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaResumidoPC()
On Error GoTo tratar_erro

Familiatext = ""
Contador1 = 1
Posicao = 0
Lista_res_PC.ListItems.Clear
If TBLISTA.EOF = False Then
    
    ProcCriaColunasResumidoPC
    
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If TBLISTA!maquina <> "" Then
            With Lista_res_PC.ListItems
                Contador1 = 2
                Do While Lista_resumido.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                    Contador1 = Contador1 + 1
                Loop
Pula:
                If TBLISTA!maquina <> Familiatext Then
                    .Add , , TBLISTA!Turno
                    .Item(.Count).SubItems(1) = TBLISTA!maquina
                    Posicao = Posicao + 1
                End If
                                
                valor = IIf(IsNull(TBLISTA!qtdeOK), 0, TBLISTA!qtdeOK) 'Orçado
                Valor1 = IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC) 'Real
                Valor2 = IIf(IsNull(TBLISTA!Terceiros), 0, TBLISTA!Terceiros) 'Variação
                Valor3 = IIf(IsNull(TBLISTA!impostos), 0, TBLISTA!impostos) 'Percentual
                
                Valor_Cofins_Prod = IIf(IsNull(TBLISTA!Qtdetotalprod), 0, TBLISTA!Qtdetotalprod) 'Total orçado
                Valor_Cofins_Serv = IIf(IsNull(TBLISTA!OS), 0, TBLISTA!OS) 'Total real
                Valor_CSLL_Prod = IIf(IsNull(TBLISTA!Lucro), 0, TBLISTA!Lucro) 'Variação total
                Valor_CSLL_Serv = IIf(IsNull(TBLISTA!material), 0, TBLISTA!material) 'Percentual
                
                .Item(.Count).SubItems(Contador1) = Format(valor, "###,##0.00") & "  |  " & Format(Valor1, "###,##0.00") & "  |  " & Format(Valor2, "###,##0.00") & "  |  " & Format(Valor3, "###,##0.00") & "%"
            
                'Carrega valor total
                Contador1 = 1
                Do While Lista_resumido.ColumnHeaders(Contador1 + 1).Text <> "Vlr. total orçado  |  Real  |  Var.  |  Perc."
                    Contador1 = Contador1 + 1
                Loop
                
                .Item(.Count).SubItems(Contador1) = Format(Valor_Cofins_Prod, "###,##0.00") & "  |  " & Format(Valor_Cofins_Serv, "###,##0.00") & "  |  " & Format(Valor_CSLL_Prod, "###,##0.00") & "  |  " & Format(Valor_CSLL_Serv, "###,##0.00") & "%"
            End With
        End If
        Familiatext = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
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

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
If Opt_comparativo.Value = True Then
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

ProcLimpaListaeCampos
If Opt_individual.Value = True Then cmbTexto.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_rev_prev_Change()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
If Txt_rev_prev <> "" Then
    VerifNumero = Txt_rev_prev
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_rev_prev = ""
        Txt_rev_prev.SetFocus
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
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 3: ProcJustificar
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
