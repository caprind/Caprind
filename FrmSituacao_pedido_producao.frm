VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSituacao_pedido_producao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Situação da produção"
   ClientHeight    =   10035
   ClientLeft      =   1725
   ClientTop       =   10170
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSituacao_pedido_producao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   9045
      Left            =   60
      TabIndex        =   96
      Top             =   990
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   15954
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   12632064
      ForeColorSel    =   16777215
      BackColorBkg    =   14737632
      GridColor       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4515
      Left            =   60
      TabIndex        =   8
      Top             =   2820
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7964
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   12632064
      ForeColorSel    =   16777215
      BackColorBkg    =   14737632
      GridColor       =   0
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.CheckBox Chk_servico 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Servico"
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   1080
      Width           =   915
   End
   Begin VB.CheckBox Chk_atualizar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Atualizar a cada: "
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
      Height          =   210
      Left            =   12480
      TabIndex        =   20
      Top             =   1050
      Width           =   1545
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   510
      Top             =   0
   End
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
   Begin VB.CheckBox chkperiodo 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   12780
      TabIndex        =   4
      Top             =   1350
      Width           =   195
   End
   Begin VB.CheckBox chkQuebra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quebrar linhas"
      Height          =   195
      Left            =   10620
      TabIndex        =   19
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar no período      "
      Enabled         =   0   'False
      Height          =   1425
      Left            =   11370
      TabIndex        =   80
      Top             =   1380
      Width           =   2295
      Begin VB.ComboBox cmbTipoData 
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
         Height          =   330
         ItemData        =   "FrmSituacao_pedido_producao.frx":000C
         Left            =   720
         List            =   "FrmSituacao_pedido_producao.frx":001C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Data."
         Top             =   270
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker txtinicio 
         Height          =   315
         Left            =   720
         TabIndex        =   6
         ToolTipText     =   "Data início para pesquisa."
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   198574083
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txtfinal 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         ToolTipText     =   "Data final para pesquisa."
         Top             =   990
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   198574081
         CurrentDate     =   39057
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   83
         Top             =   1050
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   330
         TabIndex        =   82
         Top             =   690
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   81
         Top             =   345
         Width           =   450
      End
   End
   Begin VB.CheckBox Chk_OS_atraso 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OS c/ apont. em atraso"
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
      Height          =   210
      Left            =   8250
      TabIndex        =   18
      Top             =   1080
      Width           =   2145
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   77
      Top             =   9750
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
   Begin VB.CheckBox Chk_componente 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Componente"
      Height          =   195
      Left            =   2805
      TabIndex        =   15
      Top             =   1080
      Width           =   1245
   End
   Begin VB.CheckBox Chk_produto_final 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produto final"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   1245
   End
   Begin VB.CheckBox Chk_subconjunto 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subconjunto"
      Height          =   195
      Left            =   1530
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkOrdem 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exibir ordens concluídas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   210
      Left            =   5445
      TabIndex        =   17
      Top             =   1080
      Width           =   2565
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1425
      Left            =   55
      TabIndex        =   52
      Top             =   1380
      Width           =   11265
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   6300
         TabIndex        =   95
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
            TabIndex        =   11
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
            TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   12
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.ComboBox cmbOrdenar 
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
         ItemData        =   "FrmSituacao_pedido_producao.frx":0048
         Left            =   9060
         List            =   "FrmSituacao_pedido_producao.frx":0055
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Ordenar."
         Top             =   970
         Width           =   2025
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
         ItemData        =   "FrmSituacao_pedido_producao.frx":0076
         Left            =   180
         List            =   "FrmSituacao_pedido_producao.frx":00A4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   6015
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
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   970
         Width           =   8865
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
         Left            =   180
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Familia."
         Top             =   970
         Width           =   8865
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por"
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
         Left            =   9562
         TabIndex        =   84
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   2767
         TabIndex        =   54
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3877
         TabIndex        =   53
         Top             =   780
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1425
      Left            =   13710
      TabIndex        =   51
      Top             =   1380
      Width           =   1545
      Begin VB.CheckBox chkCancelada 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkProduzindo 
         BackColor       =   &H0000FFFF&
         Caption         =   "Produzindo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   500
         Width           =   1215
      End
      Begin VB.CheckBox chkConcluida 
         BackColor       =   &H0000FF00&
         Caption         =   "Concluída"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   790
         Width           =   1215
      End
      Begin VB.CheckBox chkProduzir 
         BackColor       =   &H000000FF&
         Caption         =   "A produzir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   210
         Width           =   1215
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   79
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
      RightColor1     =   15195350
      RightColor2     =   16315633
      ShowEndPanel    =   0   'False
      ShowGripper     =   0   'False
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
      ButtonKey4      =   "9"
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
      ButtonKey5      =   "10"
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
      ButtonKey6      =   "11"
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
      Begin DrawSuite2022.USButton btnAlterarLista 
         Height          =   465
         Left            =   13350
         TabIndex        =   97
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
         Caption         =   "Alterar lista"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   16777215
         BorderColorDisabled=   16777215
         GradientColor1  =   128
         GradientColor2  =   192
         GradientColor3  =   255
         GradientColor4  =   12632319
         GradientColorDisabled1=   128
         GradientColorDisabled2=   192
         GradientColorDisabled3=   255
         GradientColorDisabled4=   12632319
         GradientColorOver1=   128
         GradientColorOver2=   192
         GradientColorOver3=   255
         GradientColorOver4=   12632319
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5040
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "FrmSituacao_pedido_producao.frx":016C
         Count           =   1
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   12330
      TabIndex        =   85
      Top             =   930
      Width           =   2775
      Begin MSComCtl2.DTPicker Txt_tempo_atualizacao 
         Height          =   315
         Left            =   1710
         TabIndex        =   21
         ToolTipText     =   "Tempo para atualização automática."
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   198639618
         CurrentDate     =   39055
      End
   End
   Begin VB.Frame Frame30 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do produto/serviço"
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
      Height          =   885
      Left            =   2790
      TabIndex        =   57
      Top             =   7380
      Width           =   12465
      Begin VB.CommandButton cmd_Visualizar_arquivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1980
         Picture         =   "FrmSituacao_pedido_producao.frx":2F60
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Visualizar arquivo."
         Top             =   420
         Width           =   315
      End
      Begin VB.TextBox txtreferencia 
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
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Código de referência."
         Top             =   420
         Width           =   1755
      End
      Begin VB.TextBox txtProduto 
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
         Left            =   4590
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   420
         Width           =   7680
      End
      Begin VB.TextBox txtRevitem 
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
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Revisão."
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox txtDesenho 
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
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8085
         TabIndex        =   94
         Top             =   210
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. de referência"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3022
         TabIndex        =   93
         Top             =   210
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Rev."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2415
         TabIndex        =   92
         Top             =   210
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   810
         TabIndex        =   58
         Top             =   210
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da ordem"
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
      Height          =   2355
      Left            =   55
      TabIndex        =   55
      Top             =   7380
      Width           =   2715
      Begin VB.TextBox txtOP 
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   98
         TabStop         =   0   'False
         ToolTipText     =   "Numero da ordem de produção"
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   1920
         Width           =   2460
      End
      Begin VB.TextBox txtconcluida 
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
         Left            =   1395
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Data de conclusão."
         Top             =   913
         Width           =   1215
      End
      Begin VB.TextBox CmbStatus 
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   1416
         Width           =   2490
      End
      Begin VB.TextBox mskprazofina 
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
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Prazo final."
         Top             =   410
         Width           =   915
      End
      Begin VB.TextBox txtquantidade 
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
         Left            =   810
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peças para produzir."
         Top             =   410
         Width           =   870
      End
      Begin VB.TextBox txtData 
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Data de emissão."
         Top             =   913
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   255
         TabIndex        =   99
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   893
         TabIndex        =   74
         Top             =   1740
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1133
         TabIndex        =   63
         Top             =   1230
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Conclusão"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1635
         TabIndex        =   62
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   465
         TabIndex        =   61
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1950
         TabIndex        =   60
         Top             =   210
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   56
         Top             =   210
         Width           =   510
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empenhos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1475
      Left            =   2790
      TabIndex        =   59
      Top             =   8260
      Width           =   12465
      Begin MSComctlLib.ListView Lista_pedidos 
         Height          =   1115
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   1958
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
            Object.Tag             =   "N"
            Text            =   "Cód. cart."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Ped. int./SPR"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   2297
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Cod. ref."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   2914
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Qtde. emp."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Qtde. ent."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Saldo"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Object.Tag             =   "D"
            Text            =   "Pr. final"
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Tag             =   "T"
            Text            =   "Ped. cliente"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Tag             =   "T"
            Text            =   "N. item"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame FrameOS 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da OS"
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
      Height          =   2355
      Left            =   55
      TabIndex        =   64
      Top             =   7380
      Visible         =   0   'False
      Width           =   15195
      Begin VB.CheckBox ChkProcesso_controlado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Processo contr."
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13575
         TabIndex        =   47
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtexecucao 
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
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Tempo de execução previsto."
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox txtpreparacao 
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Tempo de preparação previsto."
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox txtPcHora 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "1"
         ToolTipText     =   "Total de peças por tempo de execução prevista."
         Top             =   1050
         Width           =   1155
      End
      Begin VB.TextBox TxtA3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4125
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   1095
         Width           =   1095
      End
      Begin VB.TextBox Cmb_prazo_os 
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
         Height          =   315
         Left            =   12750
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Prazo final."
         Top             =   390
         Width           =   1155
      End
      Begin VB.TextBox cmbOSMaquina 
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
         Height          =   315
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Posto de trabalho."
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txtOSLote 
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
         Height          =   315
         Left            =   11655
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   390
         Width           =   1085
      End
      Begin VB.TextBox txtOSFase 
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
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Fase."
         Top             =   390
         Width           =   655
      End
      Begin VB.TextBox txtdescmaquina 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do posto de trabalho."
         Top             =   390
         Width           =   9315
      End
      Begin VB.TextBox txtStatus_OS 
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
         Height          =   315
         Left            =   5550
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   1050
         Width           =   6450
      End
      Begin VB.CheckBox chkRetrabalho 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Retrabalho"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13575
         TabIndex        =   48
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkControlada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Controlada"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12105
         TabIndex        =   46
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtTempoTotal 
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
         Height          =   315
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total utilizado pela OS."
         Top             =   390
         Width           =   1080
      End
      Begin VB.CheckBox chkPchora 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pçs x T. exec.?"
         Enabled         =   0   'False
         Height          =   210
         Left            =   12105
         TabIndex        =   45
         Top             =   960
         Width           =   1425
      End
      Begin DrawSuite2022.USButton Cmd_abrir_apontamento 
         Height          =   525
         Left            =   13110
         TabIndex        =   50
         ToolTipText     =   "Visualizar o(s) apontamento(s) da OS."
         Top             =   1650
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   926
         Caption         =   "Visualizar o(s) apontamento(s) da OS"
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
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin RichTextLib.RichTextBox txtDescricao 
         Height          =   525
         Left            =   0
         TabIndex        =   49
         ToolTipText     =   "Instruções de trabalho."
         Top             =   1650
         Width           =   12885
         _ExtentX        =   22728
         _ExtentY        =   926
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"FrmSituacao_pedido_producao.frx":3522
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Posto de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   945
         TabIndex        =   91
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11777
         TabIndex        =   90
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do posto de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5865
         TabIndex        =   89
         Top             =   180
         Width           =   2235
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Execução x peça"
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
         Left            =   3975
         TabIndex        =   88
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pç(s) x exec."
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
         Left            =   2632
         TabIndex        =   87
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Execução"
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
         Left            =   1395
         TabIndex        =   86
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo final"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   12952
         TabIndex        =   76
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo total"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   14033
         TabIndex        =   75
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   8543
         TabIndex        =   69
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Preparação"
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
         Left            =   225
         TabIndex        =   68
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " /                               ="
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   2400
         TabIndex        =   67
         Top             =   1110
         Width           =   1620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Instruções de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5805
         TabIndex        =   65
         Top             =   1450
         Width           =   1635
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   335
         TabIndex        =   66
         Top             =   180
         Width           =   345
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
      TabIndex        =   78
      Top             =   9780
      Width           =   3315
   End
End
Attribute VB_Name = "FrmSituacao_pedido_producao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Situacao_Pedido_Localizar As String 'OK
Dim FormulaRel_Situacao_Pedido As String 'OK
Dim Ultima_atualizacao As Date 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

If Vendas = True Then FunAbrirVideoWeb ("http://www.youtube.com/watch?v=MASkwmvd21k&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=35&feature=plcp") Else FunAbrirVideoWeb ("http://www.youtube.com/watch?v=S5Q9hhu80qc&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=26&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnAlterarLista_Click()
On Error GoTo tratar_erro

If MSFlexGrid1.Visible = True Then
    MSFlexGrid1.Visible = False
    MSFlexGrid2.Visible = True
    
Else
    MSFlexGrid1.Visible = True
    MSFlexGrid2.Visible = False
    
End If

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_atualizar_Click()
On Error GoTo tratar_erro

If Chk_atualizar.Value = 1 Then
    Frame6.Enabled = False
    Ultima_atualizacao = 0
    Timer1.Enabled = True
Else
    Frame6.Enabled = True
    Timer1.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_OS_atraso_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS
ProcCarregaComboData

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_produto_final_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_componente_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_servico_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_subconjunto_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkOrdem_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS
With cmbTipoData
    .Clear
    .AddItem "Emissão"
    .AddItem "Prazo final"
    .AddItem "Venda"
    If chkOrdem.Value = 0 Then .AddItem "Conclusão"
    .Text = "Prazo final"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkperiodo_Click()
On Error GoTo tratar_erro

If chkPeriodo.Value = 1 Then
    'chkOrdem.Enabled = True
    Frame4.Enabled = True
    ProcCarregaComboData
    cmbTipoData.SetFocus
Else
'    With chkOrdem
'        .Enabled = False
'        .Value = 1
'    End With
    Frame4.Enabled = False
    txtinicio.Value = Date
    txtFinal.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkQuebra_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTipoData_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_abrir_apontamento_Click()
On Error GoTo tratar_erro

If MSFlexGrid1.rows = 0 Then Exit Sub
If MSFlexGrid1.Col = 0 Then
    USMsgBox ("Informe a OS antes de visualizar o(s) apontamento(s)."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
PCP_Ordem = False
With frmSituacao_Producao
    .ProcAbrir
    .Show
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If MSFlexGrid1.rows = 0 Then Exit Sub
NomeRel = "Pcp_situacao_producao.rpt"
ProcImprimirRel FormulaRel_Situacao_Pedido, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txtdesenho = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtdesenho & "' and imagem IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then ProcAbrirArquivo TBProduto!imagem
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Activate()
On Error GoTo tratar_erro

If Screen.ActiveForm.Name = Me.Name Then
    If Chk_atualizar.Value = 1 Then Timer1.Enabled = True
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
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

txtQuantidade = ""
mskprazofina = ""
txtData.Text = ""
cmbStatus = ""
txtconcluida = ""
txtdesenho = ""
txtRevitem = ""
txtreferencia = ""
txtProduto = ""
txtmaquina = ""
txtrevprod = ""
txtnomelista = ""
txtLista = ""
txtRev = ""
Txt_ID_cliente = ""
txtCliente = ""
txtResponsavel = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCamposOS()
On Error GoTo tratar_erro

txtOSFase = ""
cmbOSMaquina = ""
txtdescmaquina = ""
txtOSLote = ""
Cmb_prazo_os = ""
TxtTempoTotal = ""
txtpreparacao = ""
txtexecucao = ""
txtPcHora = 1
TxtA3 = "00:00:00"
txtStatus_OS = ""
chkPchora.Value = 0
chkControlada.Value = 0
ChkProcesso_controlado.Value = 0
chkRetrabalho.Value = 0
txtdescricao = ""

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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 6, True

If Vendas = True Then Caption = "Vendas - Situação da produção"
ProcLimparCampos
cmbfiltrarpor = "Ordem"
cmbTipoData = "Prazo final"
cmbOrdenar = "Ordem"
txtTexto.Visible = True
cmbfamilia.Visible = False
txtinicio.Value = Date
txtFinal.Value = Date
Txt_tempo_atualizacao = "00:01:00"
Timer1.Enabled = False

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparLista()
On Error GoTo tratar_erro

With MSFlexGrid1
    .rows = 0
    .Cols = 0
    .Refresh
End With
Lista_pedidos.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS
ProcCarregaComboData

With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Posto de trabalho" Or cmbfiltrarpor = "Setor" Or cmbfiltrarpor = "Prioridade" Or cmbfiltrarpor = "Reposição" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True' or compras = 'False'", True
            ElseIf cmbfiltrarpor = "Prioridade" Then
                    .AddItem "Urgente"
                    .AddItem "Normal"
                ElseIf cmbfiltrarpor = "Reposição" Then
                        .AddItem "Sim"
                        .AddItem "Não"
                    Else
                        If cmbfiltrarpor = "Posto de trabalho" Then NomeCampo = "Maquina" Else NomeCampo = "Setor"
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select " & NomeCampo & " from CadMaquinas where " & NomeCampo & " is not null Group by " & NomeCampo, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .AddItem ""
                            Do While TBAbrir.EOF = False
                                If NomeCampo = "Maquina" Then
                                    If TBAbrir!maquina <> "" Then .AddItem TBAbrir!maquina
                                Else
                                    If TBAbrir!Setor <> "" Then .AddItem TBAbrir!Setor
                                End If
                                TBAbrir.MoveNext
                            Loop
                        End If
        End If
    Else
        If cmbfiltrarpor = "Ordem" And IsNumeric(txtTexto) = False Then txtTexto = ""
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboData()
On Error GoTo tratar_erro

TipoDataAntigo = cmbTipoData
With cmbTipoData
    .Clear
    .AddItem "Prazo final"
    If Chk_OS_atraso.Value = 1 Then
        .Text = "Prazo final"
    Else
        If cmbfiltrarpor = "Posto de trabalho" Or cmbfiltrarpor = "Setor" Then
            .AddItem "Apontamento"
            .Text = "Apontamento"
        Else
            .AddItem "Conclusão"
            .AddItem "Emissão"
            .AddItem "Venda"
            If TipoDataAntigo <> "" And (TipoDataAntigo = "Conclusão" Or TipoDataAntigo = "Emissão" Or TipoDataAntigo = "Prazo final" Or TipoDataAntigo = "Venda") Then .Text = TipoDataAntigo Else .Text = "Prazo final"
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

ProcLimparCampos
Lista_pedidos.ListItems.Clear
ProcLimparCamposOS

Inicio = Time

Tipo = ""
TipoRel = ""
Concluida = ""
ConcluidaRel = ""
FormulaRel_Situacao_PedidoSubReport = ""

If Chk_componente.Value = 1 Then
    OrdemFab = "P.tipo = 'F'"
    OrdemFabRel = "{producao.tipo} = 'F'"
Else
    OrdemFab = ""
    OrdemFabRel = ""
End If
If Chk_subconjunto.Value = 1 Then
    If OrdemFab = "" Then
        OrdemMon = "P.tipo = 'M'"
        OrdemMonRel = "{producao.tipo} = 'M'"
    Else
        OrdemMon = "or P.tipo = 'M'"
        OrdemMonRel = "or {producao.tipo} = 'M'"
    End If
Else
    OrdemMon = ""
    OrdemMonRel = ""
End If
If Chk_produto_final.Value = 1 Then
    If OrdemFab = "" And OrdemMon = "" Then
        OrdemExp = "P.tipo = 'E'"
        OrdemExpRel = "{producao.tipo} = 'E'"
    Else
        OrdemExp = "or P.tipo = 'E'"
        OrdemExpRel = "or {producao.tipo} = 'E'"
    End If
Else
    OrdemExp = ""
    OrdemExpRel = ""
End If
If Chk_servico.Value = 1 Then
    If OrdemFab = "" And OrdemMon = "" And OrdemExp = "" Then
        OrdemServ = "P.tipo = 'S'"
        OrdemServRel = "{producao.tipo} = 'S'"
    Else
        OrdemServ = "or P.tipo = 'S'"
        OrdemServRel = "or {producao.tipo} = 'S'"
    End If
Else
    OrdemServ = ""
    OrdemServRel = ""
End If

If OrdemFab <> "" Or OrdemMon <> "" Or OrdemExp <> "" Or OrdemServ <> "" Then
    Tipo = "(" & OrdemFab & OrdemMon & OrdemExp & OrdemServ & ")"
    TipoRel = OrdemFabRel & OrdemMonRel & OrdemExpRel & OrdemServRel
Else
    Tipo = "P.tipo IS NOT NULL"
    TipoRel = "NOT(ISNULL({producao.tipo}))"
End If

If chkOrdem.Value = 0 Then
    Concluida = " and P.Status <> 'Concluída' "
    ConcluidaRel = " and {Producao.Status} <> 'Concluída' "
End If

If cmbOrdenar <> "" Then
    If cmbOrdenar = "Código interno" Then
        If cmbTipoData = "Apontamento" Then Ordenar = "P.desenho, P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "P.desenho, P.ordem, OS.IDproducao"
    ElseIf cmbOrdenar = "Ordem" Then
            If cmbTipoData = "Apontamento" Then Ordenar = "P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "P.ordem, OS.IDproducao"
        Else
            Select Case cmbTipoData
                Case "Venda": If cmbTipoData = "Apontamento" Then Ordenar = "VC.Datavendas, P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "VC.Datavendas, P.ordem, OS.IDproducao"
                Case "Emissão": If cmbTipoData = "Apontamento" Then Ordenar = "P.Data, P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "P.Data, P.ordem, OS.IDproducao"
                Case "Prazo final":
                    If Chk_OS_atraso.Value = 1 Then
                        If cmbTipoData = "Apontamento" Then Ordenar = "OS.Prazofinalinicio, P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "OS.Prazofinalinicio, P.ordem, OS.IDproducao"
                    Else
                        If cmbTipoData = "Apontamento" Then Ordenar = "P.Prazoentrega, P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "P.Prazoentrega, P.ordem, OS.IDproducao"
                    End If
                Case "Conclusão": If cmbTipoData = "Apontamento" Then Ordenar = "P.Dataentrega, P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "P.Dataentrega, P.ordem, OS.IDproducao"
                Case "Apontamento": If cmbTipoData = "Apontamento" Then Ordenar = "PF.data, P.ordem, OSMU.OS, OSMU.Maquina" Else Ordenar = "PF.data, P.ordem, OS.IDproducao"
            End Select
    End If
End If

If cmbTipoData = "Apontamento" Then
    SQLInicio = "Select P.*, OSMU.Maquina, CM.Setor FROM (((Producao P INNER JOIN Ordemservico_maq_utilizadas OSMU ON P.Ordem = OSMU.Ordem) INNER JOIN Projproduto PROD ON P.Desenho = PROD.Desenho) INNER JOIN ProducaoFases PF ON PF.OS = OSMU.OS) INNER JOIN CadMaquinas CM ON CM.Maquina = OSMU.Maquina"
Else
    If cmbfiltrarpor = "Pedido interno" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Pedido cliente" Or cmbfiltrarpor = "Código do produto" Or cmbfiltrarpor = "Descrição do produto" Or cmbTipoData = "Venda" Then
        SQLInicio = "Select P.*, OS.Idproducao, OS.Maquina, CM.Setor FROM (((((Producao P INNER JOIN Producao_pedidos PP ON P.Ordem = PP.Ordem) INNER JOIN Vendas_carteira VC ON PP.IDcarteira = VC.codigo) INNER JOIN Vendas_Proposta VP ON VP.Cotacao = VC.Cotacao) INNER JOIN Projproduto PROD ON P.Desenho = PROD.Desenho) INNER JOIN Ordemservico OS ON OS.Ordem = P.Ordem) INNER JOIN CadMaquinas CM ON CM.Maquina = OS.Maquina"
    Else
        SQLInicio = "Select P.*, OS.Idproducao, OS.Maquina, CM.Setor FROM ((Producao P INNER JOIN Projproduto PROD ON P.Desenho = PROD.Desenho) INNER JOIN Ordemservico OS ON OS.Ordem = P.Ordem) INNER JOIN CadMaquinas CM ON CM.Maquina = OS.Maquina"
    End If
End If

DataFiltro = ""
DataFiltroRel = ""
If chkPeriodo.Value = 1 Then
    Select Case cmbTipoData
        Case "Venda": DataTexto = "VC.Datavendas"
        Case "Emissão": DataTexto = "P.Data"
        Case "Prazo final": If Chk_OS_atraso.Value = 1 Then DataTexto = "OS.Prazofinalinicio" Else DataTexto = "P.Prazoentrega"
        Case "Conclusão": DataTexto = "P.Dataentrega"
        Case "Apontamento": DataTexto = "PF.data"
    End Select
    DataFiltro = " and " & DataTexto & " Between '" & txtinicio.Value & "' And '" & txtFinal.Value & "'"
    Select Case Left(DataTexto, 2)
        Case "VC": DataTextoRel = Replace(DataTexto, "VC.", "Vendas_Carteira.")
        Case "P.": DataTextoRel = Replace(DataTexto, "P.", "Producao.")
        Case "OS": DataTextoRel = Replace(DataTexto, "OS.", "OrdemServico.")
        Case "PF": DataTextoRel = Replace(DataTexto, "PF.", "ProducaoFases.")
    End Select
    DataFiltroRel = " and {" & DataTextoRel & "} >= Date(" & Year(txtinicio.Value) & "," & Month(txtinicio.Value) & "," & Day(txtinicio.Value) & ") and {" & DataTextoRel & "} <= Date(" & _
                                Year(txtFinal.Value) & "," & Month(txtFinal.Value) & "," & Day(txtFinal.Value) & ")"
End If

If Chk_OS_atraso.Value = 1 Then
    FamiliaAntiga = " and OS.Pronto = 'NÃO'"
    OSAtrasoRel = " and {OrdemServico.Pronto} = 'NÃO'"
Else
    FamiliaAntiga = ""
    OSAtrasoRel = " and Not(IsNull({producao.tipo}))"
End If

TextoFiltroPadrao = Tipo & Concluida & DataFiltro & FamiliaAntiga & " and P.Status <> 'Cancelada' order by " & Ordenar
TextoFiltroPadraoRel = TipoRel & ConcluidaRel & DataFiltroRel & OSAtrasoRel & " and {Producao.Status} <> 'Cancelada'"

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Família" Then
        StrSql_Situacao_Pedido_Localizar = SQLInicio & " where PROD.Classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
        FormulaRel_Situacao_Pedido = "{projproduto.classe} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
    ElseIf cmbfiltrarpor = "Ordem" Then
            StrSql_Situacao_Pedido_Localizar = SQLInicio & " where P.ordem = " & txtTexto.Text & " and " & TextoFiltroPadrao
            FormulaRel_Situacao_Pedido = "{Producao.ordem} = " & txtTexto & " and " & TextoFiltroPadraoRel
        ElseIf cmbfiltrarpor = "Posto de trabalho" Then
                If cmbTipoData = "Apontamento" Then TabelaFiltro = "OSMU." Else TabelaFiltro = "OS."
                TabelaFiltroRel = IIf(Left(TabelaFiltro, 3) = "OS.", Replace(TabelaFiltro, "OS.", "Ordemservico."), Replace(TabelaFiltro, "OSMU.", "Ordemservico_maq_utilizadas."))
                StrSql_Situacao_Pedido_Localizar = SQLInicio & " where " & TabelaFiltro & "Maquina = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                FormulaRel_Situacao_Pedido = "{" & TabelaFiltroRel & "Maquina} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                
                If Chk_OS_atraso.Value = 1 Then OSAtrasoSubRel = "and {" & TabelaFiltro & "Pronto} = 'NÃO'" Else OSAtrasoSubRel = "and {producao.tipo} <> 'Null'"
                FormulaRel_Situacao_PedidoSubReport = "{" & TabelaFiltroRel & "Ordem} = {?Pm-Producao.Ordem} and {" & TabelaFiltroRel & "Maquina} = '" & cmbfamilia & "' and " & DataFiltroRel & " " & OSAtrasoSubRel
            ElseIf cmbfiltrarpor = "Setor" Then
                    If cmbTipoData = "Apontamento" Then
                        TabelaFiltroRel = "Ordemservico_maq_utilizadas"
                        TabelaFiltroRel1 = "CadMaquinas"
                    Else
                        TabelaFiltroRel = "Ordemservico"
                        TabelaFiltroRel1 = "CadMaquinas_1"
                    End If
                    StrSql_Situacao_Pedido_Localizar = SQLInicio & " where CM.Setor = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                    FormulaRel_Situacao_Pedido = "{CadMaquinas.Setor} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                    
                    If Chk_OS_atraso.Value = 1 Then OSAtrasoSubRel = "and {" & TabelaFiltroRel & ".Pronto} = 'NÃO'" Else OSAtrasoSubRel = "and {producao.tipo} <> 'Null'"
                    FormulaRel_Situacao_PedidoSubReport = "{" & TabelaFiltroRel & ".Ordem} = {?Pm-Producao.Ordem} and {" & TabelaFiltroRel1 & ".Setor} = '" & cmbfamilia & "' and " & DataFiltroRel & " " & OSAtrasoSubRel
                ElseIf cmbfiltrarpor = "Prioridade" Then
                        If cmbfamilia = "Urgente" Then TabelaFiltro = "True" Else TabelaFiltro = "False"
                        StrSql_Situacao_Pedido_Localizar = SQLInicio & " where P.IMPREQ = '" & TabelaFiltro & "' and " & TextoFiltroPadrao
                        FormulaRel_Situacao_Pedido = "{Producao.IMPREQ} = " & TabelaFiltro & " and " & TextoFiltroPadraoRel
                    ElseIf cmbfiltrarpor = "Reposição" Then
                            If cmbfamilia = "Sim" Then TabelaFiltro = "True" Else TabelaFiltro = "False"
                            StrSql_Situacao_Pedido_Localizar = SQLInicio & " where P.Reposicao = '" & TabelaFiltro & "' and " & TextoFiltroPadrao
                            FormulaRel_Situacao_Pedido = "{Producao.Reposicao} = " & TabelaFiltro & " and " & TextoFiltroPadraoRel
                        Else
                            Select Case cmbfiltrarpor
                                Case "Código interno": TextoFiltro = "P.desenho"
                                Case "Código do produto": TextoFiltro = "VC.desenho"
                                Case "Código de referência": TextoFiltro = "P.n_referencia"
                                Case "Descrição": TextoFiltro = "P.produto"
                                Case "Descrição do produto": TextoFiltro = "VC.Descricao_tecnica"
                                Case "Pedido interno": TextoFiltro = "VP.Ncotacao"
                                Case "Cliente": TextoFiltro = "VP.Cliente"
                                Case "Pedido cliente": TextoFiltro = "VC.PCCliente"
                            End Select
                            Select Case Left(DataTexto, 2)
                                Case "VC": TextoFiltroRel = Replace(TextoFiltro, "VC.", "Vendas_Carteira.")
                                Case "P.": TextoFiltroRel = Replace(TextoFiltro, "P.", "Producao.")
                                Case "VP": TextoFiltroRel = Replace(TextoFiltro, "VP.", "Vendas_proposta.")
                            End Select
                            StrSql_Situacao_Pedido_Localizar = SQLInicio & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                            FormulaRel_Situacao_Pedido = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
    End If
Else
    StrSql_Situacao_Pedido_Localizar = SQLInicio & " where " & TextoFiltroPadrao
    FormulaRel_Situacao_Pedido = TextoFiltroPadraoRel
    If Chk_OS_atraso.Value = 1 Then FormulaRel_Situacao_PedidoSubReport = "{Ordemservico.Ordem} = {?Pm-Producao.Ordem} and " & DataFiltroRel & " and {Ordemservico.Pronto} = 'NÃO'"
End If

'Debug.print StrSql_Situacao_Pedido_Localizar

If MSFlexGrid1.Visible = True Then
    ProcCarregaGridSitProd MSFlexGrid1, StrSql_Situacao_Pedido_Localizar, PBLista, IIf(chkQuebra.Value = 1, True, False), cmbTipoData
Else
    ProcCarregaGridSitProd MSFlexGrid2, StrSql_Situacao_Pedido_Localizar, PBLista, IIf(chkQuebra.Value = 1, True, False), cmbTipoData
End If

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub Lista_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_pedidos, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MSFlexGrid1_DblClick()
On Error GoTo tratar_erro

Formulario = "PCP/Gerenciamento de ordem"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With MSFlexGrid1
    If .Col = 0 Then
        If .Text = "" Then
            ProcLimparCampos
            Exit Sub
        End If
'        Texto = ""
'        Numero = 0
'        TamanhoTexto = Len(.Text)
'        Texto1 = Mid(.Text, 8, TamanhoTexto)
'        Numero1 = Len(Texto1)
'        Do While Numero1 <> 0
'            If Texto = "-" Then GoTo Pula
'            Texto = Left(Texto1, (Numero + 1))
'            Texto = Right(Texto, Len(Texto) - Numero)
'            Numero = Numero + 1
'            Numero1 = Numero1 - 1
'        Loop
Pula:
        'tEXTO2 = Left(Texto1, (Numero - 2))
        Ordem = txtOP
        frmprod.Show
        frmprod.ProcLimpar True
        frmprod.ProcCarregaOrdem
        frmprod.ProcLimpaCamposAP
        frmprod.Proclimpaevento
        frmprod.ProcCarregaAPOS
        frmprod.SSTab1.Tab = 1
    Else
        If .Text <> "" Then
            Texto = ""
            Numero = 0
            TamanhoTexto = Len(.Text)
            Texto1 = Mid(.Text, 5, TamanhoTexto)
            Numero1 = Len(Texto1)
            Do While Numero1 <> 0
                If Texto = "-" Then GoTo Pula1
                Texto = Left(Texto1, (Numero + 1))
                Texto = Right(Texto, Len(Texto) - Numero)
                Numero = Numero + 1
                Numero1 = Numero1 - 1
            Loop
Pula1:
            OS = Left(Texto1, (Numero - 2))
            Set TBAliquota = CreateObject("adodb.recordset")
            TBAliquota.Open "Select * from ordemservico where IDProducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
            If TBAliquota.EOF = False Then
                Ordem = TBAliquota!Ordem
                frmprod.Show
                frmprod.ProcLimpar True
                frmprod.ProcCarregaOrdem
                frmprod.ProcLimpaCamposAP
                frmprod.Proclimpaevento
                frmprod.ProcCarregaAPOS
                frmprod.SSTab1.Tab = 4
                frmprod.cmbAPOS = OS
            End If
            TBAliquota.Close
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MSFlexGrid1_SelChange()
On Error GoTo tratar_erro

With MSFlexGrid1
    If .Col = 0 Then
        If .Text = "" Then
            ProcLimparCampos
            Exit Sub
        End If
        Texto = ""
        Numero = 0
        TamanhoTexto = Len(.Text)
        'Texto1 = Mid(.Text, 8, TamanhoTexto)
       Ordem = .Text
Pula:
        Ordem = Left(Ordem, 8)
        Ordem = Right(Ordem, Len(Ordem) - 4)
        txtOP.Text = Ordem
        
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from producao where Ordem = " & Int(Ordem), Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            mskprazofina = Format(TBOrdem!PrazoEntrega, "dd/mm/yy")
            txtQuantidade = IIf(IsNull(TBOrdem!Quant), "", TBOrdem!Quant)
            txtdesenho = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
            txtRevitem = IIf(IsNull(TBOrdem!Revitem), "", TBOrdem!Revitem)
            txtProduto = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
            txtData = Format(TBOrdem!Data, "dd/mm/yy")
            cmbStatus = IIf(IsNull(TBOrdem!status), "", TBOrdem!status)
            txtResponsavel = IIf(IsNull(TBOrdem!Responsavel), "", TBOrdem!Responsavel)
            Txt_ID_cliente = IIf(IsNull(TBOrdem!IDCliente), "", TBOrdem!IDCliente)
            txtCliente = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
            txtreferencia = IIf(IsNull(TBOrdem!N_referencia), "", TBOrdem!N_referencia)
            If TBOrdem!pronta = "SIM" Then txtconcluida = Format(TBOrdem!DataEntrega, "dd/mm/yy") Else txtconcluida = ""
            If TBOrdem!Consignacao = True Then
                txtCliente = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
            Else
                txtCliente = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
            End If
            ProcCarregaListaPedidos
        End If
        TBOrdem.Close
        Frame5.Visible = True
        Frame30.Visible = True
        Frame14.Visible = True
        FrameOS.Visible = False
    Else
        If .Text = "" Then
            ProcLimparCamposOS
            Exit Sub
        End If
        Texto = ""
        Numero = 0
        TamanhoTexto = Len(.Text)
        Texto1 = Mid(.Text, 5, TamanhoTexto)
        Numero1 = Len(Texto1)
        Do While Numero1 <> 0
            If Texto = "-" Then GoTo Pula1
            Texto = Left(Texto1, (Numero + 1))
            Texto = Right(Texto, Len(Texto) - Numero)
            Numero = Numero + 1
            Numero1 = Numero1 - 1
        Loop
Pula1:
        OS = Left(Texto1, (Numero - 2))
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select * from ordemservico where IDProducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            txtOSFase = TBFases!Fase
            If IsNull(TBFases!maquina) = False And TBFases!maquina <> "" Then
                cmbOSMaquina = TBFases!maquina
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from CadMaquinas where Maquina = '" & TBFases!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    txtdescmaquina = IIf(IsNull(TBMaquinas!Descricao), "", TBMaquinas!Descricao)
                End If
                TBMaquinas.Close
            End If
            txtOSLote = IIf(IsNull(TBFases!quantidade), "", TBFases!quantidade)
            Cmb_prazo_os = IIf(IsNull(TBFases!PrazoFinal), "", TBFases!PrazoFinal)
            TxtTempoTotal = IIf(IsNull(TBFases!TempoTotalLote), "", TBFases!TempoTotalLote)
            txtpreparacao = IIf(IsNull(TBFases!Preparacao), "", TBFases!Preparacao)
            txtexecucao = IIf(IsNull(TBFases!Execucao), "", TBFases!Execucao)
            txtPcHora = IIf(IsNull(TBFases!pc_te), 0, TBFases!pc_te)
            TxtA3 = FunCalculaSegPC(IIf(IsNull(TBFases!Execucao), 0, TBFases!Execucao), txtPcHora)
            TxtA3 = FormataTempo(TxtA3)
            txtStatus_OS = IIf(IsNull(TBFases!status), "", TBFases!status)
            If TBFases!pecahora = True Then chkPchora.Value = 1 Else chkPchora.Value = 0
            If TBFases!OSControlada = True Then chkControlada.Value = 1 Else chkControlada.Value = 0
            If TBFases!Processo_controlado = True Then ChkProcesso_controlado.Value = 1 Else ChkProcesso_controlado.Value = 0
            If TBFases!Retrabalho = True Then chkRetrabalho.Value = 1 Else chkRetrabalho.Value = 0
            txtdescricao.TextRTF = IIf(IsNull(TBFases!descfase), "", TBFases!descfase)
        End If
        TBFases.Close
        Frame5.Visible = False
        Frame30.Visible = False
        Frame14.Visible = False
        FrameOS.Visible = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaPedidos()
On Error GoTo tratar_erro

Lista_pedidos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VC.*, PP.ID, PP.OrdemEmpenho, PP.Qtde_empenho, PP.Qtde_entrada FROM vendas_carteira VC INNER JOIN Producao_pedidos PP on VC.codigo = PP.IDCarteira where PP.Ordem = " & TBOrdem!Ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_pedidos.ListItems.Add(, , TBLISTA!ID)
            .SubItems(1) = TBLISTA!CODIGO
            
            Set TBCFOP = CreateObject("adodb.recordset")
            If IsNull(TBLISTA!ID_solicitacao) = True Or TBLISTA!ID_solicitacao = 0 Then
                TBCFOP.Open "Select Ncotacao, Revisao, cliente FROM vendas_proposta where cotacao = " & TBLISTA!Cotacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then
                    .SubItems(2) = IIf(IsNull(TBCFOP!Ncotacao), "", TBCFOP!Ncotacao)
                    .SubItems(3) = IIf(IsNull(TBCFOP!Revisao), "", TBCFOP!Revisao)
                    .SubItems(4) = IIf(IsNull(TBCFOP!Cliente), "", TBCFOP!Cliente)
                End If
            Else
                TBCFOP.Open "Select Requisicaotexto FROM Outros_SolicitacaoPCP where ID = " & TBLISTA!ID_solicitacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then .SubItems(2) = IIf(IsNull(TBCFOP!Requisicaotexto), "", TBCFOP!Requisicaotexto)
            End If
            TBCFOP.Close
            
            .SubItems(5) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .SubItems(6) = IIf(IsNull(TBLISTA!Rev_codinterno), "", TBLISTA!Rev_codinterno)
            .SubItems(7) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .SubItems(8) = IIf(IsNull(TBLISTA!descricao_tecnica), "", TBLISTA!descricao_tecnica)
            .SubItems(9) = IIf(IsNull(TBLISTA!OrdemEmpenho), "", TBLISTA!OrdemEmpenho)
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .SubItems(10) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_entrada), 0, TBLISTA!Qtde_entrada)
            .SubItems(11) = Valor1
            .SubItems(12) = valor - Valor1
            .SubItems(13) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .SubItems(14) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
            .SubItems(15) = IIf(IsNull(TBLISTA!PCCliente), "", TBLISTA!PCCliente)
            .SubItems(16) = IIf(IsNull(TBLISTA!N_item), "", TBLISTA!N_item)
            .SubItems(17) = IIf(IsNull(TBLISTA!Liberacao), "", TBLISTA!Liberacao)
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub





Private Sub Optfim_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

If Screen.ActiveForm.Name = "FrmSituacao_pedido_producao" Then
    Dataini = Txt_tempo_atualizacao
    If Ultima_atualizacao + Dataini <= Format(Now, "hh:mm:ss") Then
        ProcFiltrar
        Ultima_atualizacao = Time
    End If
Else
    Timer1.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtfinal_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtinicio_Click()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ProcLimparLista
ProcLimparCampos
ProcLimparCamposOS
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
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


