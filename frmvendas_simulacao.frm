VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmvendas_simulacao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Simulação"
   ClientHeight    =   10035
   ClientLeft      =   735
   ClientTop       =   1665
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmvendas_simulacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Index           =   0
      Left            =   60
      TabIndex        =   24
      Top             =   990
      Width           =   1575
      Begin VB.OptionButton Opt_servico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serviço"
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
         Left            =   180
         TabIndex        =   15
         ToolTipText     =   "Jurídica"
         Top             =   1080
         Width           =   1305
      End
      Begin VB.OptionButton Opt_compras 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Matéria-prima"
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
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "Jurídica"
         Top             =   870
         Width           =   1305
      End
      Begin VB.OptionButton Opt_fabricacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Componente"
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
         Left            =   180
         TabIndex        =   13
         ToolTipText     =   "Jurídica"
         Top             =   660
         Width           =   1305
      End
      Begin VB.OptionButton Opt_montagem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Subconjunto"
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
         Left            =   180
         TabIndex        =   12
         ToolTipText     =   "Jurídica"
         Top             =   450
         Width           =   1305
      End
      Begin VB.OptionButton Opt_vendas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto final"
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
         Left            =   180
         TabIndex        =   11
         ToolTipText     =   "Jurídica"
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar nível"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   1650
      TabIndex        =   25
      Top             =   990
      Width           =   1695
      Begin VB.CheckBox chkEstrutura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Toda estrutura"
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
         TabIndex        =   40
         Top             =   990
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.OptionButton Opt_baixo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "para baixo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   17
         ToolTipText     =   "Jurídica"
         Top             =   660
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton Opt_cima 
         BackColor       =   &H00E0E0E0&
         Caption         =   "para cima"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   16
         ToolTipText     =   "Jurídica"
         Top             =   390
         Width           =   1125
      End
   End
   Begin VB.Frame Frame8 
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
      Height          =   375
      Left            =   3060
      TabIndex        =   32
      Top             =   990
      Width           =   2775
      Begin VB.ComboBox cmbVersao 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmvendas_simulacao.frx":0442
         Left            =   1950
         List            =   "frmvendas_simulacao.frx":0494
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Versão."
         Top             =   30
         Width           =   795
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa por versão :"
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
         Left            =   60
         TabIndex        =   35
         Top             =   30
         Width           =   1800
      End
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   60
      TabIndex        =   26
      Top             =   9210
      Width           =   15165
      Begin VB.TextBox txtValorMaxCusto 
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
         Left            =   13560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de custo."
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox txtValorMinCusto 
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
         Left            =   10680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de custo."
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox Txt_valor_total_custo 
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
         Left            =   12120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de custo."
         Top             =   360
         Width           =   1425
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   330
         Width           =   10365
         _ExtentX        =   18283
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. max. custo"
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
         Left            =   13650
         TabIndex        =   39
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. min. custo"
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
         Left            =   10800
         TabIndex        =   38
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total custo"
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
         Left            =   12210
         TabIndex        =   28
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   55
      TabIndex        =   20
      Top             =   2370
      Width           =   15165
      Begin VB.TextBox Txt_qtde_vendida 
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
         Left            =   12600
         TabIndex        =   3
         ToolTipText     =   "Quantidade vendida."
         Top             =   390
         Width           =   825
      End
      Begin VB.CheckBox Chk_valor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar valor"
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
         TabIndex        =   10
         Top             =   450
         Width           =   1485
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
         Height          =   510
         Index           =   19
         Left            =   3630
         TabIndex        =   33
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton optFim 
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
            TabIndex        =   8
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton optInicio 
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
            TabIndex        =   6
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optMeio 
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
            TabIndex        =   7
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
            TabIndex        =   9
            Top             =   180
            Width           =   705
         End
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
         Left            =   8520
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   4065
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
         ItemData        =   "frmvendas_simulacao.frx":04E6
         Left            =   180
         List            =   "frmvendas_simulacao.frx":04FF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3375
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
         Left            =   8520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Familia."
         Top             =   390
         Width           =   4065
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   14880
         Picture         =   "frmvendas_simulacao.frx":0574
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgFolder 
         Height          =   240
         Left            =   14610
         Picture         =   "frmvendas_simulacao.frx":0AFE
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qt. vend."
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
         Left            =   12660
         TabIndex        =   34
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label5 
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
         Left            =   1447
         TabIndex        =   22
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label4 
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
         Left            =   9817
         TabIndex        =   21
         Top             =   180
         Width           =   1470
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   23
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
      ButtonCaption3  =   "Similar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Visualizar estoque dos produtos similares (F7)"
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
      ButtonWidth3    =   38
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
      ButtonLeft4     =   133
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "4"
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
      ButtonLeft5     =   137
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "5"
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
      ButtonLeft6     =   175
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonKey7      =   "7"
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   203
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11430
         Top             =   60
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmvendas_simulacao.frx":1088
         Count           =   1
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   5955
      Left            =   60
      TabIndex        =   4
      Top             =   3240
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   10504
      Cols            =   2
      DefaultFontSize =   6.75
      GridColor       =   12632256
      ReadOnly        =   -1  'True
      Rows            =   2
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "                                         "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   3360
      TabIndex        =   29
      Top             =   990
      Width           =   11865
      Begin VB.TextBox Txt_ID_desc_versao 
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
         TabIndex        =   31
         Text            =   "0"
         ToolTipText     =   "ID descrição da versão."
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox Txt_descricao_versao 
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
         Height          =   825
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Descrição da versão."
         Top             =   390
         Width           =   11475
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição da versão"
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
         Left            =   5190
         TabIndex        =   30
         Top             =   180
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmvendas_simulacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Vendas_Simulacao As String 'OK
Dim VersaoEstrutura_Simulacao As String 'OK

'GridEstrutura
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String, valor As String, DataValidacao As String, RespValidacao As String
Public IDProduto As Long

Sub ProcAjuda()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDadosProduto()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Descricao from projproduto where codproduto = " & IDProduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Desenho = TBProduto!Desenho
    DT = TBProduto!Descricao
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_valor_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
Frame6.Enabled = False
VersaoEstrutura_Simulacao = cmbVersao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Simulação"
Direitos

'With Chk_valor
'    .Enabled = True
'    Set TBAcessos = CreateObject("adodb.recordset")
'    TBAcessos.Open "Select IDAcesso from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = '"Vendas/Simulação"/Visualizar valor de custo'", Conexao, adOpenKeyset, adLockOptimistic
'    If TBAcessos.EOF = True Then
'        .Value = 0
'        Chk_valor.Enabled = False
'    End If
'    TBAcessos.Close
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_RowColChange(ByVal Row As Long, ByVal Col As Long)
On Error GoTo tratar_erro

If Grid1.rows = 1 Then Exit Sub
IDProduto = Grid1.Cell(Row, 3).Text
VersaoEstrutura = Grid1.Cell(Row, 7).Text
Qtd_Prog = Grid1.Cell(Row, 8).Text

'Carrega descrição da versão
ProcLimpaCamposDescVersao
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Projconjunto_desc_versao where Codproduto = " & IDProduto & " and Versao = '" & VersaoEstrutura & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_desc_versao = TBAbrir!ID
    Txt_descricao_versao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt_compras_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
If Opt_compras.Value = True Then
    Opt_baixo.Enabled = False
    Opt_cima.Enabled = True
    Opt_cima.Value = True
    If Opt_fabricacao.Value = True Then ProcCarregaComboFamilia cmbfamilia, "Compras = 'True'", False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt_fabricacao_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
ProcLiberaNivel
If Opt_fabricacao.Value = True Then ProcCarregaComboFamilia cmbfamilia, "Familia is not null", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub opt_montagem_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
ProcLiberaNivel
If Opt_montagem.Value = True Then ProcCarregaComboFamilia cmbfamilia, "Familia is not null", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_vendas_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
ProcLiberaNivel
If Opt_vendas.Value = True Then ProcCarregaComboFamilia cmbfamilia, "vendas = 'True'", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaNivel()
On Error GoTo tratar_erro

Opt_cima.Enabled = True
Opt_baixo.Enabled = True
Opt_baixo.Value = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
Dim SubTipoProduto As String
Dim CamposFiltro As String
Dim INNERJOINTEXTO1 As String, INNERJOINTEXTO  As String, TextoFiltroPadrao As String, TextoFiltro As String, TextoFiltro1 As String
Dim TextoFiltroVersao  As String

Familiatext = ""
FamiliaAntiga = ""
Pesquisa = ""

Acao = "filtrar"
If txtTexto = "" And cmbfamilia = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    If txtTexto.Visible = True Then txtTexto.SetFocus Else cmbfamilia.SetFocus
    Exit Sub
End If
quantidade = IIf(Txt_qtde_vendida = "", 0, Txt_qtde_vendida)
If quantidade = 0 Then
    NomeCampo = "a quantidade vendida"
    ProcVerificaAcao
    Txt_qtde_vendida.SetFocus
    Exit Sub
End If

ProcExcluirDadosProducaoRelatorios

If Opt_vendas.Value = True Then SubTipoProduto = "P.subtipoitem = 1"
If Opt_montagem.Value = True Then SubTipoProduto = "P.subtipoitem = 2"
If Opt_fabricacao.Value = True Then SubTipoProduto = "P.subtipoitem = 3"
If Opt_compras.Value = True Then SubTipoProduto = "P.subtipoitem = 0"
If Opt_servico.Value = True Then SubTipoProduto = "P.subtipoitem = 5"

CamposFiltro = "P.Desenho, P.Codproduto, P.Producao, P.Descricao, P.Unidade, P.PCusto, P.SubTipoItem"
If Opt_cima.Value = True Then
    CamposFiltro = CamposFiltro & ", PC.Versao"
    INNERJOINTEXTO1 = "INNER JOIN Projconjunto PC ON PC.Desenho = P.Desenho"
    If cmbfiltrarpor = "Descrição da versão" Then
        If cmbVersao <> "" Then TextoFiltroVersao = " and PC.Versao_desenho = '" & cmbVersao & "'" Else TextoFiltroVersao = ""
    Else
        TextoFiltroVersao = " and PC.Versao_desenho = '" & cmbVersao & "'"
    End If
Else
    If cmbfiltrarpor = "Descrição da versão" Then
        CamposFiltro = CamposFiltro & ", PC.Versao"
        INNERJOINTEXTO1 = "LEFT JOIN Projconjunto PC ON PC.Codproduto = P.Codproduto"
        If cmbVersao <> "" Then TextoFiltroVersao = " and PC.Versao = '" & cmbVersao & "'" Else TextoFiltroVersao = ""
    Else
        INNERJOINTEXTO1 = ""
        TextoFiltroVersao = ""
    End If
End If
INNERJOINTEXTO = "Select " & CamposFiltro & " from Projproduto P " & INNERJOINTEXTO1 & " LEFT JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto " & IIf(cmbfiltrarpor = "Descrição da versão", "and PCDV.Versao = PC.Versao", "")
TextoFiltroPadrao = SubTipoProduto & " and P.bloqueado <> 'True'" & TextoFiltroVersao & " order by P.desenho " & IIf(cmbfiltrarpor = "Descrição da versão", ", PC.Versao", "")

TextoFiltro1 = ""
If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Família" Then
        StrSql_Vendas_Simulacao = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor
            Case "Código interno": TextoFiltro = "P.desenho"
            Case "Código de referência": TextoFiltro = "IA.N_referencia"
            Case "Descrição": TextoFiltro = "P.descricao"
            Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
            Case "Descrição da versão":
                TextoFiltro = "PCDV.Descricao"
                TextiFiltro1 = " and PC.Codigo IS NOT NULL"
        End Select
        StrSql_Vendas_Simulacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltro1 & " and " & TextoFiltroPadrao
    End If
End If

Erase arrNodes 'Zera o array
If Opt_cima.Value = True Then
    ''ReDim arrNodes(20000)
    ProcVerifNivelAcima
Else
    ''ReDim arrNodes(2000)
End If

ProcCarregaLista
IDProduto = 0

ProcExcluirDadosProducaoRelatoriosTotal
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Responsavel, Modulo, Texto, QtdePrevista, QtdeProduzida, QtdeNC from Producao_Relatorios_total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    TBLISTA!Responsavel = pubUsuario
    TBLISTA!Modulo = "Vendas/Simulação"
    If Chk_valor.Value = 1 Then TBLISTA!Texto = "S" Else TBLISTA!Texto = "N"
    TBLISTA!QtdePrevista = Txt_qtde_vendida
    TBLISTA!QtdeProduzida = qtdeliberada
    TBLISTA!qtdeNC = IIf(qtdeliberada - quantidade < 0, (qtdeliberada - quantidade) * -1, 0)
    TBLISTA.Update
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifNivelAcima()
On Error GoTo tratar_erro

Desenho = ""
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Engenharia_Estrutura, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = True
    Contador = 0
    Do While Not TBLISTA.EOF
        If Desenho <> TBLISTA!Desenho Then
            DesenhoProduto = TBLISTA!Desenho
            ProcNivel2EstruturaAcima frmproj_produto_estrutura, cmbVersao, IIf(chkEstrutura.Value = 1, True, False)
        End If
        Desenho = TBLISTA!Desenho
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    If Familiatext <> "" Then
        CamposFiltro = "P.Desenho, P.Codproduto, P.Producao, P.Descricao, P.PCusto, P.SubTipoItem, PC.Versao"
        INNERJOINTEXTO = "Select " & CamposFiltro & " from (Projproduto P LEFT JOIN Projconjunto PC ON PC.Desenho = P.Desenho) LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto and PCDV.Versao = PC.Versao"
        StrSql_Engenharia_Estrutura = INNERJOINTEXTO & " where " & FamiliaAntiga & Familiatext & Pesquisa & " and (SubTipoItem = 0 or SubTipoItem = 1 or SubTipoItem = 2 or SubTipoItem = 3) order by desenho"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Call m_Tree.Nodes.Clear
Grid1.rows = 1

m_Row = 1
m_Col = 1
Desenho = ""
Familiatext = ""
ValorPago = 0
ValorTotalPagar = 0
ValorTotalPago = 0
Txt_valor_total_custo = "0,00"
txtValorMaxCusto = "0,00"
txtValorMinCusto = "0,00"
Contador1 = -1
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Vendas_Simulacao, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While Not TBLISTA.EOF
        Permitido = False
        If cmbfiltrarpor = "Descrição da versão" Then
            Tipo = IIf(IsNull(TBLISTA!versao), IIf(cmbVersao = "", "A", cmbVersao), TBLISTA!versao)
            If Desenho <> TBLISTA!Desenho Or Desenho = TBLISTA!Desenho And (IsNull(TBLISTA!versao) Or Tipo <> Familiatext) Then Permitido = True
        Else
            Tipo = cmbVersao
            If Desenho <> TBLISTA!Desenho Then Permitido = True
        End If
        If Permitido = True Then
            CodRef = ""
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBLISTA!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                CodRef = TBFI!N_referencia
            End If
            TBFI.Close
            
'            PartNumber = ""
'            If IsNull(TBLISTA!ID_partnumber_fabricante) = False Then
'                Set TBComponente = CreateObject("adodb.recordset")
'                TBComponente.Open "Select Part_number from Projproduto_fabricante where ID = " & TBLISTA!ID_partnumber_fabricante, Conexao, adOpenKeyset, adLockOptimistic
'                If TBComponente.EOF = False Then PartNumber = TBComponente!Part_number
'                TBComponente.Close
'            End If
                                    
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select Responsavel, Modulo, maquina, Qtdeprev, Eficiencia, Terceiros from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Responsavel = pubUsuario
            TBGravar!Modulo = "Vendas/Simulação"
            TBGravar!maquina = TBLISTA!Desenho
            
            qtdeliberada = FunVerificaQtdeEstoque(TBLISTA!Desenho, 0, "") 'Verifica quantidade no estoque com material consignado
            If Chk_valor.Value = 1 Then
                'Verifica custo do estoque
                StrSql = "Select MAX(vlrUnit) as ValorMaximo, MIN(vlrUnit) as ValorMinimo, AVG(vlrUnit) as ValorMedio from Estoque_Movimentacao where Desenho = '" & TBLISTA!Desenho & "' and Operacao LIKE 'ENTRADA%'"
                
                'Debug.print StrSql
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                
                If TBAbrir.EOF = False Then
                    Valor_Cofins_Prod = Format(IIf(IsNull(TBAbrir!ValorMaximo), 0, TBAbrir!ValorMaximo), "###,##0.00000000")
                    Valor_Cofins_Serv = Format(IIf(IsNull(TBAbrir!ValorMinimo), 0, TBAbrir!ValorMinimo), "###,##0.00000000")
                    CTMedioEst = Format(IIf(IsNull(TBAbrir!ValorMedio), 0, TBAbrir!ValorMedio), "###,##0.00000000")
                    valor = CTMedioEst ' * TBAbrir!Saldo
                End If
                TBAbrir.Close

                
'                Call FunVerificaQtdeEstoque(TBLISTA!Desenho, 0, "and Consignacao = 'False'") 'Verifica valor no estoque sem material consignado
'
'                If CTMedioEst <> 0 Then
'                    valor = Format(CTMedioEst * quantidade, "###,##0.00000000") 'Medio custo
'                    If Chk_valor = 1 Then
'                        Set TBComponente = CreateObject("adodb.recordset")
'                        TBComponente.Open "Select Max(Valor_total / estoque_real) as MAX, Min(Valor_total / estoque_real) as Min from Qtde_estoque_produto where desenho = '" & TBLISTA!Desenho & "' and Consignacao = 'False' and Valor_total > 0", Conexao, adOpenKeyset, adLockOptimistic
'                        If TBComponente.EOF = False Then
'                            Valor_Cofins_Prod = Format(IIf(IsNull(TBComponente!Max), 0, TBComponente!Max) * quantidade, "###,##0.00000000") 'Max custo
'                            Valor_Cofins_Serv = Format(IIf(IsNull(TBComponente!Min), 0, TBComponente!Min) * quantidade, "###,##0.00000000") 'Min custo
'                        End If
'                        TBComponente.Close
'                    End If
'                Else
'                    'Verifica custo de compras
'                    valor = FunVerificaVlrUltCompra(TBLISTA!Desenho) 'Verifica valor unitário da última compra
'                    If valor <> 0 Then
'                        valor = Format(valor * quantidade, "###,##0.00000000")
'                        If Chk_valor = True Then
'                            Set TBComponente = CreateObject("adodb.recordset")
'                            TBComponente.Open "Select MAX(CPL.preco_unitario * ISNULL(CC.Valor_moeda, 1)) AS max, MIN(CPL.preco_unitario * ISNULL(CC.Valor_moeda, 1)) AS Min from dbo.Compras_pedido_lista AS CPL LEFT OUTER JOIN dbo.Compras_comercial AS CC ON CPL.IDPedido = CC.IdPedido where CPL.Desenho = '" & TBLISTA!Desenho & "' and CPL.IDpedido <> 0 AND CPL.preco_unitario > 0", Conexao, adOpenKeyset, adLockOptimistic
'                            If TBComponente.EOF = False Then
'                                Valor_Cofins_Prod = Format(IIf(IsNull(TBComponente!Max), 0, TBComponente!Max) * quantidade, "###,##0.00000000") 'Max custo
'                                Valor_Cofins_Serv = Format(IIf(IsNull(TBComponente!Min), 0, TBComponente!Min) * quantidade, "###,##0.00000000") 'Min custo
'                            End If
'                            TBComponente.Close
'                        End If
'                    Else
'                        'Verifica custo do item
'                        valor = Format(IIf(IsNull(TBLISTA!PCusto), 0, TBLISTA!PCusto) * quantidade, "###,##0.00000000")
'                        If Chk_valor = 1 Then
'                            Valor_Cofins_Prod = valor 'Max custo
'                            Valor_Cofins_Serv = valor 'Min custo
'                        End If
'                    End If
'                End If
                
                ValorPago = ValorPago + valor 'Valor total custo
                TBGravar!QtdePrev = Format(valor, "###,##0.00000000")
                If Chk_valor = 1 Then
                    ValorTotalPagar = ValorTotalPagar + Valor_Cofins_Prod 'Valor total Max custo
                    ValorTotalPago = ValorTotalPago + Valor_Cofins_Serv 'Valor total Min custo
                    TBGravar!Eficiencia = Format(Valor_Cofins_Prod, "###,##0.00000000")
                    TBGravar!Terceiros = Format(Valor_Cofins_Serv, "###,##0.00000000")
                End If
            Else
                'Mensagem = ""
                valor = 0
                Valor_Cofins_Prod = 0
                Valor_Cofins_Serv = 0
                TBGravar!QtdePrev = 0
                TBGravar!Eficiencia = 0
                TBGravar!Terceiros = 0
            End If
            
            TBGravar.Update
            TBGravar.Close
            
            DataValidacao = ""
            RespValidacao = ""
            If TBLISTA!SubTipoItem <> 0 Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from Projconjunto_desc_versao where codproduto = " & TBLISTA!Codproduto & " and Versao = '" & Tipo & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    DataValidacao = IIf(IsNull(TBFI!DtValidacao), "", TBFI!DtValidacao)
                    RespValidacao = IIf(IsNull(TBFI!RespValidacao), "", TBFI!RespValidacao)
                End If
            End If
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 0
            arrNodes(Contador1).Text = TBLISTA!Desenho & vbTab & "" & vbTab & TBLISTA!Codproduto & vbTab & CodRef & vbTab & TBLISTA!Descricao & vbTab & TBLISTA!Unidade & vbTab & Tipo & vbTab & Format(quantidade, "###,##0.0000") & vbTab & Format(qtdeliberada, "###,##0.0000") & vbTab & Format(IIf(qtdeliberada - quantidade < 0, (qtdeliberada - quantidade) * -1, 0), "###,##0.0000") & vbTab & Format(Valor_Cofins_Serv, "###,##0.00000000") & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(Valor_Cofins_Prod, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao & IIf(qtdeliberada - quantidade < 0, 1, 0)
            
            Codproduto = TBLISTA!Codproduto
            
            ProcNivel2Estrutura frmvendas_simulacao, Tipo, IIf(Chk_valor.Value = 0, False, True), False, True, True
        End If
        Desenho = TBLISTA!Desenho
        If cmbfiltrarpor = "Descrição da versão" Then Familiatext = Tipo
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    
    With Grid1
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 18
        .rows = m_Row
        .Cell(0, 1).Text = "Cód. interno"
        .Cell(0, 2).Text = "Pos."
        .Cell(0, 3).Text = "ID"
        .Cell(0, 4).Text = "Cód. de ref."
        .Cell(0, 5).Text = "Descrição"
        .Cell(0, 6).Text = "Un."
        .Cell(0, 7).Text = "Ver."
        .Cell(0, 8).Text = "Qtde."
        .Cell(0, 9).Text = "Qtde. est."
        .Cell(0, 10).Text = "Necessidade"
        .Cell(0, 11).Text = "Vlr. min. custo"
        .Cell(0, 12).Text = "Vlr. custo"
        .Cell(0, 13).Text = "Vlr. max. custo"
        .Cell(0, 14).Text = "Dt. validação"
        .Cell(0, 15).Text = "Resp. validação"
        .Cell(0, 16).Text = "Part number"
        .Cell(0, 17).Text = "Observações"
        .Range(0, 1, 0, 17).Alignment = cellCenterCenter
        .Column(1).Width = 200
        .Column(2).Width = 30
        .Column(3).Width = 0
        .Column(4).Width = 80
        .Column(5).Width = 300
        .Column(6).Width = 40
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Width = 40
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Width = 100
        .Column(8).Alignment = cellRightCenter
        .Column(9).Width = 100
        .Column(9).Alignment = cellRightCenter
        .Column(10).Width = 100
        .Column(10).Alignment = cellRightCenter
        .Column(11).Width = 100
        .Column(11).Alignment = cellRightCenter
        .Column(12).Width = 100
        .Column(12).Alignment = cellRightCenter
        .Column(13).Width = 100
        .Column(13).Alignment = cellRightCenter
        .Column(14).Width = 120
        .Column(15).Width = 100
        .Column(16).Width = 150
        .Column(17).Width = 400
        
        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem Left(arrNodes(0).Text, Len(arrNodes(0).Text) - 1)
        If Right(arrNodes(0).Text, 1) = 1 Then
            .Cell(1, 1).BackColor = Yellow
            .Cell(1, 2).BackColor = Yellow
            .Cell(1, 3).BackColor = Yellow
            .Cell(1, 4).BackColor = Yellow
            .Cell(1, 5).BackColor = Yellow
            .Cell(1, 6).BackColor = Yellow
            .Cell(1, 7).BackColor = Yellow
            .Cell(1, 8).BackColor = Yellow
            .Cell(1, 9).BackColor = Yellow
            .Cell(1, 10).BackColor = Yellow
            .Cell(1, 11).BackColor = Yellow
            .Cell(1, 12).BackColor = Yellow
            .Cell(1, 13).BackColor = Yellow
            .Cell(1, 14).BackColor = Yellow
            .Cell(1, 15).BackColor = Yellow
            .Cell(1, 16).BackColor = Yellow
            .Cell(1, 17).BackColor = Yellow
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

            .AddItem Left(arrNodes(intIndex).Text, Len(arrNodes(intIndex).Text) - 1)
            If Right(arrNodes(intIndex).Text, 1) = 1 Then
                .Cell(intIndex + 1, 1).BackColor = Yellow
                .Cell(intIndex + 1, 2).BackColor = Yellow
                .Cell(intIndex + 1, 3).BackColor = Yellow
                .Cell(intIndex + 1, 4).BackColor = Yellow
                .Cell(intIndex + 1, 5).BackColor = Yellow
                .Cell(intIndex + 1, 6).BackColor = Yellow
                .Cell(intIndex + 1, 7).BackColor = Yellow
                .Cell(intIndex + 1, 8).BackColor = Yellow
                .Cell(intIndex + 1, 9).BackColor = Yellow
                .Cell(intIndex + 1, 10).BackColor = Yellow
                .Cell(intIndex + 1, 11).BackColor = Yellow
                .Cell(intIndex + 1, 12).BackColor = Yellow
                .Cell(intIndex + 1, 13).BackColor = Yellow
                .Cell(intIndex + 1, 14).BackColor = Yellow
                .Cell(intIndex + 1, 15).BackColor = Yellow
                .Cell(intIndex + 1, 16).BackColor = Yellow
                .Cell(intIndex + 1, 17).BackColor = Yellow
            End If
        Next
        
        .AutoRedraw = True
        .Refresh
    End With
Else
    ProcLimpaCamposDescVersao
    Frame6.Enabled = False
End If
Txt_valor_total_custo = Format(ValorPago, "###,##0.00")
txtValorMaxCusto = Format(ValorTotalPagar, "###,##0.00")
txtValorMinCusto = Format(ValorTotalPago, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If IDProduto = 0 Then
    USMsgBox ("Informe o registro antes de visualizar a impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

ProcPuxaDadosProduto

NomeRel = "Vendas_simulacao.rpt"
ProcImprimirRel "{projproduto.Desenho} = '" & Desenho & "' and {Projconjunto.Versao} = '" & VersaoEstrutura_Simulacao & "' and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "' and {Producao_Relatorios.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "'", "{Projconjunto.codproduto} = {?Pm-projproduto.codproduto} and {Projconjunto.Versao} = '" & VersaoEstrutura_Simulacao & "' and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "' and {Producao_Relatorios.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "'"

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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 7, True

Formulario = "Vendas/Simulação"
Direitos
ProcCarregaComboFamilia cmbfamilia, "vendas = 'True'", False
Codproduto = 0

ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, 0, "Produtos/Serviços", "E", False
If Permitido = False Then cmbfiltrarpor = "Código interno"

txtTexto.Visible = True
cmbfamilia.Visible = False
ProcCarregaVersao

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_baixo_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
Optinicio.Enabled = True
Optmeio.Enabled = True
Optfim.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_cima_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
chkEstrutura.Visible = True
optIgual.Value = True
Optinicio.Enabled = False
Optmeio.Enabled = False
Optfim.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
chkEstrutura.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_vendida_Change()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao

If Txt_qtde_vendida <> "" Then
    VerifNumero = tTxt_qtde_vendida
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_vendida = ""
        Txt_qtde_vendida.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_vendida_LostFocus()
On Error GoTo tratar_erro

Txt_qtde_vendida = Format(Txt_qtde_vendida, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparGrid()
On Error GoTo tratar_erro

Grid1.rows = 1
Txt_valor_total_custo = "0,00"
txtValorMinCusto = "0,00"
txtValorMaxCusto = "0,00"

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

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcLimparGrid
ProcLimpaCamposDescVersao
ProcCarregaVersao
If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposDescVersao()
On Error GoTo tratar_erro

Txt_ID_desc_versao = 0
Txt_descricao_versao = ""

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
    Case 3: ProcSimilar
    Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaVersao()
On Error GoTo tratar_erro

With cmbVersao
    .Clear
    If cmbfiltrarpor = "Descrição da versão" Then .AddItem ""
    .AddItem "A"
    .AddItem "B"
    .AddItem "C"
    .AddItem "D"
    .AddItem "E"
    .AddItem "F"
    .AddItem "G"
    .AddItem "H"
    .AddItem "I"
    .AddItem "J"
    .AddItem "K"
    .AddItem "L"
    .AddItem "M"
    .AddItem "N"
    .AddItem "O"
    .AddItem "P"
    .AddItem "Q"
    .AddItem "R"
    .AddItem "S"
    .AddItem "T"
    .AddItem "U"
    .AddItem "V"
    .AddItem "W"
    .AddItem "X"
    .AddItem "Y"
    .AddItem "Z"
    .Text = "A"
End With
    
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

If CheckEditStatus() Then Exit Sub
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

Private Sub Grid1_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
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

Private Function CheckEditStatus() As Boolean
On Error GoTo tratar_erro
Dim hWnd As Long
Dim strClassName As String
Dim intPos As Integer

strClassName = Space(256)
hWnd = GetFocus()
Call GetClassName(hWnd, strClassName, 256)
intPos = InStr(1, strClassName, Chr(0))
strClassName = Left(strClassName, intPos - 1)
If strClassName = "ThunderRT6TextBox" Then CheckEditStatus = True    'Editing Else    CheckEditStatus = False

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcSimilar()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID_similar, Desenho from Projproduto where Codproduto = " & IDProduto & " and ID_similar IS NOT NULL and ID_similar <> 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    IDlista = IIf(IsNull(TBAbrir!ID_similar), 0, TBAbrir!ID_similar)
    With frmvendas_simulacao_similar
        .Caption = .Caption & " - Cód. interno: " & TBAbrir!Desenho
        .Show 1
    End With
Else
    USMsgBox ("Não existe nenhum produto similar a este."), vbInformation, "CAPRIND v5.0"
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
