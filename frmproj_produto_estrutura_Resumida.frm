VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmproj_produto_estrutura_Resumida 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Estrutura"
   ClientHeight    =   10035
   ClientLeft      =   735
   ClientTop       =   525
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmproj_produto_estrutura_Resumida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1035
      Left            =   60
      TabIndex        =   19
      Top             =   990
      Width           =   2775
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
         Left            =   1350
         TabIndex        =   10
         ToolTipText     =   "Jurídica"
         Top             =   540
         Width           =   825
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
         Left            =   1350
         TabIndex        =   9
         ToolTipText     =   "Jurídica"
         Top             =   300
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
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Jurídica"
         Top             =   780
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
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Jurídica"
         Top             =   540
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
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Jurídica"
         Top             =   300
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1035
      Left            =   2850
      TabIndex        =   20
      Top             =   990
      Width           =   1605
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
         Left            =   120
         TabIndex        =   28
         Top             =   780
         Visible         =   0   'False
         Width           =   1395
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
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Jurídica"
         Top             =   510
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
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Jurídica"
         Top             =   270
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
      Height          =   345
      Left            =   4590
      TabIndex        =   26
      Top             =   1020
      Width           =   2385
      Begin VB.ComboBox cmbVersao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmproj_produto_estrutura_Resumida.frx":0442
         Left            =   1710
         List            =   "frmproj_produto_estrutura_Resumida.frx":0494
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Versão."
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa por versão :"
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
         Left            =   90
         TabIndex        =   27
         Top             =   -30
         Width           =   1560
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   60
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
      TabIndex        =   21
      Top             =   9210
      Width           =   15165
      Begin VB.TextBox Txt_valor_total 
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
         Left            =   13170
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Valor total."
         Top             =   360
         Width           =   1770
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   390
         Width           =   12825
         _ExtentX        =   22622
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total"
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
         Left            =   13620
         TabIndex        =   23
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1035
      Left            =   8910
      TabIndex        =   16
      Top             =   990
      Width           =   6315
      Begin VB.OptionButton optIgual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Igual"
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
         Left            =   4740
         TabIndex        =   32
         Top             =   330
         Width           =   705
      End
      Begin VB.OptionButton Optmeio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Meio"
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
         Left            =   3480
         TabIndex        =   31
         Top             =   330
         Width           =   645
      End
      Begin VB.OptionButton Optinicio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Início"
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
         Left            =   2700
         TabIndex        =   30
         Top             =   330
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton Optfim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fim"
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
         Left            =   4170
         TabIndex        =   29
         Top             =   330
         Width           =   555
      End
      Begin VB.CheckBox Chk_valor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar valor"
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
         Left            =   4950
         TabIndex        =   5
         Top             =   30
         Width           =   1275
      End
      Begin VB.TextBox txtTexto 
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
         Left            =   2490
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmproj_produto_estrutura_Resumida.frx":04E6
         Left            =   180
         List            =   "frmproj_produto_estrutura_Resumida.frx":04FF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbFamilia 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   2490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   600
         Width           =   3615
      End
      Begin VB.Image imgFolder 
         Height          =   240
         Left            =   5520
         Picture         =   "frmproj_produto_estrutura_Resumida.frx":0574
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   5790
         Picture         =   "frmproj_produto_estrutura_Resumida.frx":0AFE
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   975
         TabIndex        =   17
         Top             =   390
         Width           =   705
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   1035
      Left            =   4470
      TabIndex        =   24
      Top             =   990
      Width           =   4425
      Begin VB.TextBox txtcodProduto 
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
         Left            =   3210
         TabIndex        =   33
         Text            =   "0"
         ToolTipText     =   "ID descrição da versão."
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txt_ID_desc_versao 
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
         Left            =   2460
         TabIndex        =   25
         Text            =   "0"
         ToolTipText     =   "ID descrição da versão."
         Top             =   0
         Visible         =   0   'False
         Width           =   735
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
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Descrição da versão."
         Top             =   390
         Width           =   3495
      End
      Begin DrawSuite2022.USButton Cmd_salvar_desc_versao 
         Height          =   585
         Left            =   3630
         TabIndex        =   15
         ToolTipText     =   "Salvar descrição da versão (F6)."
         Top             =   390
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1032
         DibPicture      =   "frmproj_produto_estrutura_Resumida.frx":1088
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
   End
   Begin FlexCell.Grid Grid1 
      Height          =   7155
      Left            =   60
      TabIndex        =   3
      Top             =   2040
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   12621
      BackColorBkg    =   16777215
      Cols            =   2
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   2
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   12
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
      ButtonCaption2  =   "Novo"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Novo (Insert)"
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
      ButtonWidth2    =   33
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Alterar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Alterar (F3)"
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
      ButtonLeft3     =   75
      ButtonTop3      =   2
      ButtonWidth3    =   41
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Excluir"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Excluir (F4)"
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
      ButtonWidth4    =   39
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Relatório"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Relatório (F5)"
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
      ButtonLeft5     =   159
      ButtonTop5      =   2
      ButtonWidth5    =   51
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Copiar"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Copiar (F7)"
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
      ButtonLeft6     =   212
      ButtonTop6      =   2
      ButtonWidth6    =   39
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Versão"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Criar versão (F8)"
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
      ButtonLeft7     =   253
      ButtonTop7      =   2
      ButtonWidth7    =   41
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Validação"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Validar/Cancelar validação (F9)"
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
      ButtonLeft8     =   296
      ButtonTop8      =   2
      ButtonWidth8    =   53
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonAlignment9=   2
      ButtonType9     =   1
      ButtonStyle9    =   -1
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   -1
      ButtonLeft9     =   351
      ButtonTop9      =   4
      ButtonWidth9    =   2
      ButtonHeight9   =   54
      ButtonCaption10 =   "Ajuda"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Ajuda (F1)"
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
      ButtonLeft10    =   355
      ButtonTop10     =   2
      ButtonWidth10   =   36
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Sair"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Sair (Esc)"
      ButtonKey11     =   "11"
      ButtonAlignment11=   2
      BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft11    =   393
      ButtonTop11     =   2
      ButtonWidth11   =   26
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      ButtonEnabled12 =   0   'False
      BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState12   =   5
      ButtonLeft12    =   421
      ButtonTop12     =   2
      ButtonWidth12   =   24
      ButtonHeight12  =   24
      ButtonUseMaskColor12=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11430
         Top             =   60
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmproj_produto_estrutura_Resumida.frx":181A
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmproj_produto_estrutura_Resumida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Estrutura As Boolean 'OK
Dim StrSql_Engenharia_Estrutura As String 'OK
Public VersaoEstrutura As String 'OK

'GridEstrutura
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String, ValorCusto As String, DataValidacao As String, RespValidacao As String
Public IDProduto As Long, IDestrutura As Long

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=noo6adXSNZM&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=33&feature=plcp")

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

Private Sub ProcPuxaDadosProdutoNovo()
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
VersaoEstrutura = cmbVersao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_desc_versao_Click()
On Error GoTo tratar_erro

If Frame6.Enabled = False Then Exit Sub
If FunVerificaRegistroValidado("Projconjunto_desc_versao", "ID = " & Txt_ID_desc_versao, "versão da estrutura", "deste registro", "cadastrar descrição da versão", False, True) = False Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Projconjunto_desc_versao where ID = " & Txt_ID_desc_versao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Codproduto = IDProduto
TBAbrir!versao = VersaoEstrutura
TBAbrir!Descricao = Txt_descricao_versao
TBAbrir.Update
Txt_ID_desc_versao = TBAbrir!ID
TBAbrir.Close
USMsgBox ("Descrição da versão cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Engenharia/Estrutura/Resumida"
Evento = "Cadastrar descrição da versão"
ID_documento = Txt_ID_desc_versao
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho from Projproduto where Codproduto = " & IDProduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Documento = "Cód. interno: " & TBProduto!Desenho & " - Versão: " & VersaoEstrutura
End If
TBProduto.Close
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Engenharia/Estrutura/Resumida"
Direitos

With Chk_valor
    .Enabled = True
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select IDAcesso from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Engenharia/Estrutura/Visualizar valor de custo'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then
        .Value = 0
        Chk_valor.Enabled = False
    End If
    TBAcessos.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_RowColChange(ByVal Row As Long, ByVal Col As Long)
On Error GoTo tratar_erro

If Novo_Estrutura = False Then
    IDestrutura = 0
    IDProduto = 0
End If

If Grid1.rows = 1 Then Exit Sub
Frame6.Enabled = False
With USToolBar1
    .ButtonState(8) = 0
    IDestrutura = IIf(Grid1.Cell(Row, 22).Text = "", 0, Grid1.Cell(Row, 22).Text)
    IDProduto = IIf(Grid1.Cell(Row, 4).Text = "", 0, Grid1.Cell(Row, 4).Text)
    
    If IDestrutura <> 0 Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Versao from Projconjunto where codigo = " & IDestrutura, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            VersaoEstrutura = TBAbrir!versao
        End If
        TBAbrir.Close
    Else
        IDestrutura = 0
        VersaoEstrutura = Grid1.Cell(Row, 10).Text
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Codproduto from Projproduto where Codproduto = " & IDProduto & " and Subtipoitem = 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then .ButtonState(8) = 5 Else Frame6.Enabled = True
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Desenho, Descricao from projproduto where codproduto = " & IDProduto & " and (Vendas = 'True' or Producao = 'True')", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ButtonState(2) = 0
    Else
        .ButtonState(2) = 5
    End If
    
    If IDestrutura <> 0 Then
        .ButtonState(3) = 0
        .ButtonState(4) = 0
    Else
        .ButtonState(3) = 5
        .ButtonState(4) = 5
    End If
    
    'Carrega descrição da versão
    ProcLimpaCamposDescVersao
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Projconjunto_desc_versao where Codproduto = " & IDProduto & " and Versao = '" & VersaoEstrutura & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_ID_desc_versao = TBAbrir!ID
        Txt_descricao_versao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
        txtcodproduto = IDProduto
    End If
    
    .Refresh
End With

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
quantidade = 1 'Obrigatório pois utiliza ao calcular a quantidade na estrutura

ProcExcluirDadosProducaoRelatorios
ProcLimpaCamposDescVersao

If Opt_vendas.Value = True Then SubTipoProduto = "P.subtipoitem = 1"
If Opt_montagem.Value = True Then SubTipoProduto = "P.subtipoitem = 2"
If Opt_fabricacao.Value = True Then SubTipoProduto = "P.subtipoitem = 3"
If Opt_compras.Value = True Then SubTipoProduto = "P.subtipoitem = 0"
If Opt_servico.Value = True Then SubTipoProduto = "P.subtipoitem = 5"

CamposFiltro = "P.Desenho, P.Codproduto, P.Producao, P.Descricao, P.Unidade, P.PCusto, P.SubTipoItem, P.Largura, P.Comprimento"
If Opt_cima.Value = True Then
    CamposFiltro = CamposFiltro & ", PC.Versao"
    INNERJOINTEXTO1 = "INNER JOIN Projconjunto PC ON PC.Desenho = P.Desenho"
    If cmbfiltrarpor = "Descrição da versão" Then
        If cmbVersao <> "" Then TextoFiltroVersao = " and PC.Versao = '" & cmbVersao & "'" Else TextoFiltroVersao = ""
    Else
        TextoFiltroVersao = " and PC.Versao = '" & cmbVersao & "'"
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
        StrSql_Engenharia_Estrutura = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor
            Case "Código interno": TextoFiltro = "P.desenho"
            Case "Código de referência": TextoFiltro = "IA.N_referencia"
            Case "Descrição": TextoFiltro = "P.descricao"
            Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
            Case "Descrição da versão":
                TextoFiltro = "PCDV.Descricao"
                TextoFiltro1 = " and PC.Codigo IS NOT NULL"
        End Select
        StrSql_Engenharia_Estrutura = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltro1 & " and " & TextoFiltroPadrao
    End If
End If

Erase arrNodes 'Zera o array

If Opt_cima.Value = True Then
   ' ReDim arrNodes(20000)
    ProcVerifNivelAcima
Else
  '  ReDim arrNodes(20000)
End If

'Debug.print StrSql_Engenharia_Estrutura


ProcCarregaLista
IDProduto = 0
IDestrutura = 0

ProcExcluirDadosProducaoRelatoriosTotal
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Responsavel, Modulo, Texto from Producao_Relatorios_total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    TBLISTA!Responsavel = pubUsuario
    TBLISTA!Modulo = "Engenharia/Estrutura/Resumida"
    If Chk_valor.Value = 1 Then TBLISTA!Texto = "S" Else TBLISTA!Texto = "N"
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
        CamposFiltro = "P.Desenho, P.Codproduto, P.Producao, P.Descricao, P.unidade, P.PCusto, P.SubTipoItem, PC.Versao, P.Largura, P.Comprimento"
        INNERJOINTEXTO = "Select " & CamposFiltro & " from (Projproduto P INNER JOIN Projconjunto PC ON PC.CodProduto = P.Codproduto) LEFT JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto and PCDV.Versao = PC.Versao"
        StrSql_Engenharia_Estrutura = INNERJOINTEXTO & " where P.bloqueado <> 'True' and " & FamiliaAntiga & Familiatext & Pesquisa & " order by P.desenho"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If IDestrutura = 0 Then Exit Sub
If USMsgBox("Deseja realmente excluir este registro da estrutura?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Posicao, codproduto, Versao, Desenho from ProjConjunto where Codigo = " & IDestrutura, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        If FunVerificaRegistroValidado("Projconjunto_desc_versao", "Codproduto = " & TBProduto!Codproduto & " and Versao = '" & VersaoEstrutura & "'", "versão da estrutura", "o registro", "excluir", False, True) = False Then Exit Sub
        Id_Item = TBProduto!Codproduto
        Desenho = TBProduto!Desenho
        
        Conexao.Execute "Update ProjConjunto Set Posicao = Posicao - 1 where Posicao > " & TBProduto!Posicao & " and Posicao IS NOT NULL and codproduto = " & TBProduto!Codproduto & " and Versao = '" & TBProduto!versao & "'"
        TBProduto.Delete
    End If
    TBProduto.Close
    USMsgBox ("Registro excluído da estrutura com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Engenharia/Estrutura/Resumida"
    Evento = "Excluir"
    ID_documento = IDestrutura
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Desenho from Projproduto where codproduto = " & Id_Item, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Documento = "Cód. interno: " & TBProduto!Desenho
    End If
    Documento1 = "Cód. interno: " & Desenho
    TBProduto.Close
    ProcGravaEvento
    '==================================
    IDProduto = 0
    IDestrutura = 0
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlterar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If IDestrutura = 0 Then
    USMsgBox ("Informe o registro antes de alterar a estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from projconjunto where Codigo = " & IDestrutura, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
'Debug.print VersaoEstrutura & TBAbrir!Codproduto
    If FunVerificaRegistroValidado("Projconjunto_desc_versao", "Codproduto = " & TBAbrir!Codproduto & " and Versao = '" & VersaoEstrutura & "'", "versão da estrutura", "o registro", "alterar", False, True) = False Then Exit Sub
    PCP_Ordem = False
    Novo_Estrutura = False
    Formulario = "Engenharia/Estrutura/Resumida"
    With frmproj_EstruturaLocaliza_item
        .ProcPuxaDados
        .Show 1
    End With
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If IDProduto = 0 Then
    USMsgBox ("Informe o registro antes de criar nova estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcPuxaDadosProdutoNovo
If USMsgBox("Deseja agregar um novo registro abaixo deste: " & vbCrLf & "Código interno : " & Desenho & vbCrLf & "Descrição: " & DT & " ", vbYesNo) = vbYes Then
    If FunVerificaRegistroValidado("Projconjunto_desc_versao", "ID = " & Txt_ID_desc_versao, "versão da estrutura", "abaixo do código interno : " & Desenho, "agregar um novo registro", False, True) = False Then Exit Sub
    
    PCP_Ordem = False
    Novo_Estrutura = True
    'frmProj_Produto_Estrutura_TipoItem.Show 1
    frmproj_EstruturaLocaliza_item.Show 1
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
ValorTotal = 0
Txt_valor_total = "0,00"
Contador1 = -1
ValorPago = 0
Set TBLISTA = CreateObject("adodb.recordset")
'Debug.print StrSql_Engenharia_Estrutura

TBLISTA.Open StrSql_Engenharia_Estrutura, Conexao, adOpenKeyset, adLockOptimistic
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
            
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select Responsavel, Modulo, Maquina, Qtdeprev from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Responsavel = pubUsuario
            TBGravar!Modulo = "Engenharia/Estrutura/Resumida"
            TBGravar!maquina = TBLISTA!Desenho
            
            If Chk_valor.Value = 1 Then
                'Verifica custo médido do estoque sem material consignado
                Call FunVerificaQtdeEstoque(TBLISTA!Desenho, 0, "and Consignacao = 'False'")
                If CTMedioEst <> 0 Then
                    valor = Format(CTMedioEst, "###,##0.00000000")
                    TBGravar!QtdePrev = Format(CTMedioEst, "###,##0.00000000")
                Else
                    'Verifica valor unitário da última compra
                    valor = FunVerificaVlrUltCompra(TBLISTA!Desenho)
                    If valor <> 0 Then
                        TBGravar!QtdePrev = Format(valor, "###,##0.00000000")
                    Else
                        valor = Format(IIf(IsNull(TBLISTA!PCusto), 0, TBLISTA!PCusto), "###,##0.00000000")
                        TBGravar!QtdePrev = Format(IIf(IsNull(TBLISTA!PCusto), 0, TBLISTA!PCusto), "###,##0.00000000")
                    End If
                End If
            Else
                valor = 0
                TBGravar!QtdePrev = 0
            End If
            TBGravar.Update
            TBGravar.Close
            
            DataValidacao = ""
            RespValidacao = ""
            If TBLISTA!SubTipoItem = 1 Or TBLISTA!SubTipoItem = 2 Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from Projconjunto_desc_versao where codproduto = " & TBLISTA!Codproduto & " and Versao = '" & Tipo & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    DataValidacao = IIf(IsNull(TBFI!DtValidacao), "", TBFI!DtValidacao)
                    RespValidacao = IIf(IsNull(TBFI!RespValidacao), "", TBFI!RespValidacao)
                End If
            End If
            
            Contador1 = Contador1 + 1
            arrNodes(Contador1).Level = 0
            arrNodes(Contador1).Text = TBLISTA!Desenho & vbTab & "" & vbTab & "Principal" & vbTab & TBLISTA!Codproduto & vbTab & CodRef & vbTab & vbTab & TBLISTA!Descricao & vbTab & vbTab & TBLISTA!Unidade & vbTab & Tipo & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(TBLISTA!Largura, "###,##0.00") & vbTab & Format(TBLISTA!Comprimento, "###,##0.00") & vbTab & "" & vbTab & "" & vbTab & Format(valor, "###,##0.00000000") & vbTab & Format(DataValidacao, "dd/mm/yy") & vbTab & RespValidacao & vbTab & IDestrutura
            
            Codproduto = TBLISTA!Codproduto
            
            ProcNivel2Estrutura frmproj_produto_estrutura, Tipo, IIf(Chk_valor.Value = 0, False, True), False, True, False
            
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
        .Cols = 23
        .rows = m_Row
        .Cell(0, 1).Text = "Cód. interno"
        .Cell(0, 2).Text = "Pos."
        .Cell(0, 3).Text = "Tipo"
        
        .Cell(0, 4).Text = "ID"
        .Cell(0, 5).Text = "Cód. de ref."
        .Cell(0, 6).Text = "Part number"
        .Cell(0, 7).Text = "Descrição"
        .Cell(0, 8).Text = "Observações"
        .Cell(0, 9).Text = "Un."
        .Cell(0, 10).Text = "Ver."
        .Cell(0, 11).Text = "Vlr./un"
        .Cell(0, 12).Text = "Un/vlr."
        .Cell(0, 13).Text = "Dim/mm"
        .Cell(0, 14).Text = "Vlr./pç"
        .Cell(0, 15).Text = "Largura/mm"
        .Cell(0, 16).Text = "Comprimento/mm"
        .Cell(0, 17).Text = "Qtde."
        .Cell(0, 18).Text = "Total"
        .Cell(0, 19).Text = "Vlr. custo"
        .Cell(0, 20).Text = "Dt. validação"
        .Cell(0, 21).Text = "Resp. validação"
        .Cell(0, 22).Text = "ID estr."
        .Column(1).Width = 150 'Codigo
        .Column(2).Width = 25 'Posição
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Width = 70 'Tipo
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Width = 0 'ID
        
        .Column(5).Width = 70 'Cod Referencia
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Width = 90 'Part Number
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Width = 220 'Descricao
        .Column(7).Alignment = cellLeftCenter
        .Column(8).Width = 150 ' observacoes
        .Column(8).Alignment = cellCenterCenter
        .Column(9).Width = 30 'unidade
        .Column(9).Alignment = cellCenterCenter
        .Column(10).Width = 40 'Versao
        .Column(10).Alignment = cellCenterCenter
        .Column(11).Width = 0 'Valor Unitario
        .Column(11).Alignment = cellRightCenter
        .Column(12).Width = 0 'Unidade x valor
        .Column(12).Alignment = cellRightCenter
        .Column(13).Width = 0 'Dim mm
        .Column(13).Alignment = cellRightCenter
        .Column(14).Width = 0 'valor x peça
        .Column(14).Alignment = cellRightCenter
        .Column(15).Width = 0 'largura mm
        .Column(15).Alignment = cellRightCenter
        .Column(16).Width = 0 'comprimento mm
        .Column(16).Alignment = cellRightCenter
        .Column(17).Width = 60 'Quantidade
        .Column(17).Alignment = cellRightCenter
        .Column(18).Width = 0 'Total
        .Column(18).Width = 0 'Valor custo
        .Column(19).Width = 80 'Data validacao
        .Column(19).Alignment = cellCenterCenter
        .Column(20).Width = 80 'Responsavel validacao
        .Column(20).Alignment = cellCenterCenter
        .Column(21).Width = 80 'Responsavel validacao
        .Column(21).Alignment = cellCenterCenter
        .Column(22).Width = 0 'IDproduto
        .Column(22).Alignment = cellCenterCenter
        
        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem arrNodes(0).Text
        
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
            .AddItem arrNodes(intIndex).Text
        Next
        
        .AutoRedraw = True
        .Refresh
    End With
Else
    ProcLimpaCamposDescVersao
    Frame6.Enabled = False
End If
Txt_valor_total = Format(ValorPago, "###,##0.00")

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
NomeRel = "Engenharia_estrutura.rpt"
'ProcImprimirRel "{projproduto.Codproduto} = " & IDProduto & " and {Projconjunto.Versao} = '" & VersaoEstrutura & "' and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", "{Projconjunto.codproduto} = {?Pm-projproduto.codproduto} and {Projconjunto.Versao} = '" & VersaoEstrutura & "' and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
ProcImprimirRel "{projproduto.Codproduto} = " & IDProduto & " and {Projconjunto.Versao} = '" & VersaoEstrutura & "'", "{Projconjunto.codproduto} = {?Pm-projproduto.codproduto} and {Projconjunto.Versao} = '" & VersaoEstrutura & "'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If IDProduto = 0 Then
    USMsgBox ("Informe o registro antes de copiar a estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar essa estrutura?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where codproduto = " & IDProduto & " and Subtipoitem = 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Não é permitido copiar estrutura de matéria-prima."), vbExclamation, "CAPRIND v5.0"
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
    
    Sit_REG = 3
    frmprocessos_Novo.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarVersao()
On Error GoTo tratar_erro

Acao = "criar a versão"
If IDProduto = 0 Then
    NomeCampo = "o registro"
    ProcVerificaAcao
    Exit Sub
End If
If USMsgBox("Deseja realmente criar nova versão para essa estrutura?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Engenharia_Conjuntos = False
    frmproj_conjunto_criar_versao.Show 1
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
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcAlterar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF6: Cmd_salvar_desc_versao_Click
    Case vbKeyF7: ProcCopiar
    Case vbKeyF8: ProcCriarVersao
    Case vbKeyF9:
        If ProcVerifTemEstrutura = False Then Exit Sub
        Formulario = "Engenharia/Estrutura/Resumida"
        frmValidar.Show
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

ProcCarregaToolBar1 Me, 15195, 11, True

Formulario = "Engenharia/Estrutura/Resumida"
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
chkEstrutura.Visible = False
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
'With Grid1
'    .Rows = 2
'    .Cols = 2
'    .Cell(0, 1).Text = ""
'    .Cell(1, 1).Text = ""
'    .Refresh
'End With
Txt_valor_total = "0,00"

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
txtcodproduto = 0
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcNovo
    Case 3: ProcAlterar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcCopiar
    Case 7: ProcCriarVersao
    Case 8:
            If ProcVerifTemEstrutura = False Then Exit Sub
            Formulario = "Engenharia/Estrutura/Resumida"
            frmValidar.Show
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifTemEstrutura() As Boolean
On Error GoTo tratar_erro

If IDProduto = 0 Then
    USMsgBox ("Informe o registro antes de validar/cancelar validação."), vbExclamation, "CAPRIND v5.0"
    ProcVerifTemEstrutura = False
    Exit Function
End If


    
ProcVerifTemEstrutura = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Codproduto from Projproduto where Codproduto = " & IDProduto & " and Subtipoitem = 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é possivel validar estrutura para matéria-prima."), vbExclamation, "CAPRIND v5.0"
    ProcVerifTemEstrutura = False
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

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
