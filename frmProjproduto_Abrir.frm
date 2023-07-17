VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProjproduto_Abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Engenharia - Produtos e serviços - Localizar"
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9990
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjproduto_Abrir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   741
      DibPicture      =   "frmProjproduto_Abrir.frx":000C
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   39
      Top             =   4380
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   714
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3690
      TabIndex        =   37
      Top             =   1410
      Width           =   3135
      Begin VB.CheckBox chkEmbalagem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sem embalagem"
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
         TabIndex        =   19
         Top             =   250
         Width           =   1485
      End
      Begin VB.CheckBox chkInspecao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sem  inspeção"
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
         Left            =   1710
         TabIndex        =   21
         Top             =   250
         Width           =   1365
      End
      Begin VB.CheckBox chkGravacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sem gravação"
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
         TabIndex        =   20
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Validado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   8670
      TabIndex        =   36
      Top             =   1380
      Width           =   1275
      Begin VB.ComboBox Cmb_validado 
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
         ItemData        =   "frmProjproduto_Abrir.frx":365C
         Left            =   180
         List            =   "frmProjproduto_Abrir.frx":3669
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Validado."
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6840
      TabIndex        =   33
      Top             =   1410
      Width           =   1815
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "frmProjproduto_Abrir.frx":3679
         Left            =   180
         List            =   "frmProjproduto_Abrir.frx":3686
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Status."
         Top             =   270
         Width           =   1470
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   795
      Left            =   30
      TabIndex        =   32
      Top             =   1410
      Width           =   1275
      Begin VB.CheckBox chkServicos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serviços"
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
         Top             =   480
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkProdutos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produtos"
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
         Top             =   250
         Value           =   1  'Checked
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicação "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1320
      TabIndex        =   31
      Top             =   1410
      Width           =   2355
      Begin VB.CheckBox chkCompras 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compras"
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
         TabIndex        =   16
         Top             =   480
         Width           =   930
      End
      Begin VB.CheckBox Chk_qualidade 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Qualidade"
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
         Left            =   1230
         TabIndex        =   18
         Top             =   480
         Width           =   1050
      End
      Begin VB.CheckBox Chk_PCP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PCP"
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
         Left            =   1230
         TabIndex        =   17
         Top             =   250
         Width           =   600
      End
      Begin VB.CheckBox chkVendas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vendas"
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
         Top             =   250
         Width           =   825
      End
   End
   Begin VB.CheckBox optPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data do cadastro"
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
      Left            =   4110
      TabIndex        =   10
      Top             =   3990
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   30
      TabIndex        =   25
      Top             =   2190
      Width           =   9915
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4950
         TabIndex        =   38
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   3
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
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
            Height          =   255
            Left            =   1470
            TabIndex        =   2
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   4
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.CommandButton Cmd_excluir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9420
         Picture         =   "frmProjproduto_Abrir.frx":36A1
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir filtro para pesquisa (F4)."
         Top             =   1065
         Width           =   315
      End
      Begin VB.CommandButton Cmd_salvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9090
         Picture         =   "frmProjproduto_Abrir.frx":37DF
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salvar filtro para pesquisa (F3)."
         Top             =   1065
         Width           =   315
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmProjproduto_Abrir.frx":3832
         Left            =   180
         List            =   "frmProjproduto_Abrir.frx":3863
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4725
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
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1065
         Width           =   8895
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1065
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.Label Label1 
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
         Left            =   3892
         TabIndex        =   27
         Top             =   840
         Width           =   1470
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
         Left            =   2107
         TabIndex        =   26
         Top             =   180
         Width           =   870
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   30
      Top             =   420
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor1  =   16777215
      GradientColor2  =   14737632
      GradientColorDown1=   10802943
      GradientColorDown2=   7979263
      GradientColorDownRight1=   10802943
      GradientColorDownRight2=   7979263
      GradientColorOver1=   14417407
      GradientColorOver2=   12317439
      GradientColorOverRight1=   14417407
      GradientColorOverRight2=   12317439
      IsStrech        =   -1  'True
      RightColor1     =   14737632
      RightColor2     =   16777215
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
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonUseMaskColor2=   0   'False
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3330
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProjproduto_Abrir.frx":392D
         Count           =   1
      End
   End
   Begin VB.Frame FrameData 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   30
      TabIndex        =   28
      Top             =   3705
      Width           =   9915
      Begin VB.ComboBox Cmb_ordenar 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmProjproduto_Abrir.frx":5B15
         Left            =   1350
         List            =   "frmProjproduto_Abrir.frx":5B1F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Ordenar por."
         Top             =   225
         Width           =   2265
      End
      Begin MSComCtl2.DTPicker Msk_final 
         Height          =   315
         Left            =   8250
         TabIndex        =   12
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   225
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
         Format          =   176488449
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker Msk_inicio 
         Height          =   315
         Left            =   6360
         TabIndex        =   11
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   225
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
         Format          =   176488449
         CurrentDate     =   39057
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por:"
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
         TabIndex        =   35
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label2 
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
         Left            =   6015
         TabIndex        =   34
         Top             =   270
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
         Left            =   7845
         TabIndex        =   29
         Top             =   270
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2115
      Left            =   60
      TabIndex        =   24
      Top             =   4815
      Visible         =   0   'False
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   3731
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Filtrar por"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Local da frase"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Texto para pesquisa"
         Object.Width           =   11509
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "IDTexto"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmProjproduto_Abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Label1.Visible = True
cmbfamilia.Visible = False
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Grupo do cliente" Or cmbfiltrarpor = "Aplicação" Then
    txtTexto.Visible = False
    With cmbfamilia
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            If Engenharia_Produtos = True Then ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
            If Compras_Produtos = True Then ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
            If Vendas_Produtos = True Then ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", True
        ElseIf cmbfiltrarpor = "Grupo" Then
                If Engenharia_Produtos = True Then ProcCarregaComboGrupoFamilia cmbfamilia, "Grupo <> 'Null'", True
                If Compras_Produtos = True Then ProcCarregaComboGrupoFamilia cmbfamilia, "Grupo <> 'Null' and compras = 'True'", True
                If Vendas_Produtos = True Then ProcCarregaComboGrupoFamilia cmbfamilia, "Grupo <> 'Null' and vendas = 'True'", True
            ElseIf cmbfiltrarpor = "Cliente" Then
                    ProcCarregaComboCliForn cmbfamilia, True
                ElseIf cmbfiltrarpor = "Grupo do cliente" Then
                        Set TBFamilia = CreateObject("adodb.recordset")
                        TBFamilia.Open "Select ID, Texto from Clientes_grupos where Texto <> 'Null' group by ID, Texto", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFamilia.EOF = False Then
                            Do While TBFamilia.EOF = False
                                .AddItem TBFamilia!Texto
                                .ItemData(.NewIndex) = TBFamilia!ID
                                TBFamilia.MoveNext
                            Loop
                        End If
                        TBFamilia.Close
                    Else
                        ProcCarregaComboCliForn cmbfamilia, False
        End If
    End With
Else
    txtTexto.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With Msk_final
    If FunVerificaDataFinal(Msk_inicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

ModuloFiltro = ""
ModuloFiltroRel = ""
If Compras_Produtos = True Then
    ModuloFiltro = " and P.Compras = 'True'"
    ModuloFiltroRel = " and {Projproduto.Compras} = True"
ElseIf Vendas_Produtos = True Then
        ModuloFiltro = " and P.Vendas = 'True'"
        ModuloFiltroRel = " and {Projproduto.Vendas} = True"
End If
If chkProdutos.Value = 1 Then
    TipoProduto = "P.Tipo <> 'S'"
    TipoProdutoRel = "{projproduto.Tipo} <> 'S'"
Else
    TipoProduto = "P.Tipo = 'S'"
    TipoProdutoRel = "{projproduto.Tipo} = 'S'"
End If
If chkServicos.Value = 1 Then
    TipoProduto = TipoProduto & " or P.Tipo = 'S'"
    TipoProdutoRel = TipoProdutoRel & " or {projproduto.Tipo} = 'S'"
End If

AplicacaoFiltro = ""
AplicacaoFiltroRel = ""
If Not (chkVendas.Value = 1 And chkCompras.Value = 1 And Chk_PCP.Value = 1 And Chk_qualidade.Value = 1 Or chkVendas.Value = 0 And chkCompras.Value = 0 And Chk_PCP.Value = 0 And Chk_qualidade.Value = 0) Then
    If chkVendas.Value = 1 Then
        AplicacaoFiltro = "P.Vendas = 'True'"
        AplicacaoFiltroRel = "{Projproduto.vendas} = True"
    End If
    If chkCompras.Value = 1 Then
        If AplicacaoFiltro <> "" Then
            AplicacaoFiltro = AplicacaoFiltro & " or P.Compras = 'True'"
            AplicacaoFiltroRel = AplicacaoFiltroRel & " or {Projproduto.Compras} = True"
        Else
            AplicacaoFiltro = "P.Compras = 'True'"
            AplicacaoFiltroRel = "{Projproduto.Compras} = True"
        End If
    End If
    If Chk_PCP.Value = 1 Then
        If AplicacaoFiltro <> "" Then
            AplicacaoFiltro = AplicacaoFiltro & " or P.Producao = 'True'"
            AplicacaoFiltroRel = AplicacaoFiltroRel & " or {Projproduto.Producao} = True"
        Else
            AplicacaoFiltro = "P.Producao = 'True'"
            AplicacaoFiltroRel = "{Projproduto.Producao} = True"
        End If
    End If
    If Chk_qualidade.Value = 1 Then
        If AplicacaoFiltro <> "" Then
            AplicacaoFiltro = AplicacaoFiltro & " or P.Qualidade = 'True'"
            AplicacaoFiltroRel = AplicacaoFiltroRel & " or {Projproduto.Qualidade} = True"
        Else
            AplicacaoFiltro = "P.Qualidade = 'True'"
            AplicacaoFiltroRel = "{Projproduto.Qualidade} = True"
        End If
    End If
    AplicacaoFiltro = " and (" & AplicacaoFiltro & ")"
    AplicacaoFiltroRel = " and (" & AplicacaoFiltroRel & ")"
End If

AplicacaoFiltro2 = ""
AplicacaoFiltroRel2 = ""
If chkEmbalagem.Value = 1 Then
    AplicacaoFiltro2 = "(P.Embalagem IS NULL or P.Embalagem = N'')"
    AplicacaoFiltroRel2 = "{Projproduto.Embalagem} = NULL"
End If
If chkGravacao.Value = 1 Then
    If AplicacaoFiltro2 <> "" Then
        AplicacaoFiltro2 = AplicacaoFiltro2 & " or (P.Gravacao IS NULL or P.Gravacao = N'')"
        AplicacaoFiltroRel2 = AplicacaoFiltroRel2 & " and {Projproduto.Gravacao} = NULL"
    Else
        AplicacaoFiltro2 = "(P.Gravacao IS NULL or P.Gravacao = N'')"
        AplicacaoFiltroRel2 = "{Projproduto.Gravacao} = NULL"
    End If
End If
If chkInspecao.Value = 1 Then
    If AplicacaoFiltro2 <> "" Then
        AplicacaoFiltro2 = AplicacaoFiltro2 & " or (P.Inspecao IS NULL or P.Inspecao = N'')"
        AplicacaoFiltroRel2 = AplicacaoFiltroRel2 & " or {Projproduto.Inspecao} = NULL"
    Else
        AplicacaoFiltro2 = "(P.Inspecao IS NULL or P.Inspecao = N'')"
        AplicacaoFiltroRel2 = "{Projproduto.Inspecao} = NULL"
    End If
End If
If AplicacaoFiltro2 <> "" Then
    AplicacaoFiltro2 = " and (" & AplicacaoFiltro2 & ")"
    AplicacaoFiltroRel2 = " and (" & AplicacaoFiltroRel2 & ")"
End If
StatusFiltro = ""
StatusFiltroRel = ""
If cmbStatus <> "" Then
    If cmbStatus = "Liberado" Then
        StatusFiltro = " and P.bloqueado = 'False'"
        StatusFiltroRel = " and {Projproduto.bloqueado} = False"
    Else
        StatusFiltro = " and P.bloqueado = 'True'"
        StatusFiltroRel = " and {Projproduto.bloqueado} = True"
    End If
End If

ValidFiltro = ""
ValidFiltroRel = ""
If Cmb_validado <> "" Then
    If Cmb_validado = "Sim" Then
        ValidFiltro = " and P.DtValidacao IS NOT NULL"
        ValidFiltroRel = " and NOT(ISNULL({Projproduto.DtValidacao}))"
    Else
        ValidFiltro = " and P.DtValidacao IS NULL"
        ValidFiltroRel = " and ISNULL({Projproduto.DtValidacao}) = True"
    End If
End If

If optPeriodo.Value = 1 Then
    DataFiltro = "(P.data) Between '" & Msk_inicio.Value & "' And '" & Msk_final.Value & "'"
    DataFiltroRel = "{Projproduto.data} >= Date(" & Year(Msk_inicio.Value) & "," & Month(Msk_inicio.Value) & "," & Day(Msk_inicio.Value) & ") and {Projproduto.data} <= Date(" & _
                        Year(Msk_final.Value) & "," & Month(Msk_final.Value) & "," & Day(Msk_final.Value) & ")"
Else
    DataFiltro = "P.desenho <> 'Null'"
    DataFiltroRel = "{Projproduto.desenho} <> 'Null'"
End If
If Cmb_ordenar = "Código interno" Then Ordenar = "P.Desenho" Else Ordenar = "P.Descricao"

CamposFiltro = "P.codProduto, P.Desenho, P.Classe, P.Descricao, P.RevDesenho, P.ID_CF, P.PCusto, P.PConsumo, P.PRevenda, P.Data, P.Espessura, P.Largura, P.Comprimento, P.Dureza, P.DtValidacao, P.DtValidacaoConj, DtValidacaoPlano"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (((((Projproduto P LEFT JOIN item_aplicacoes IA ON P.codproduto = IA.codproduto) LEFT JOIN Projproduto_clientes PC ON PC.codproduto = P.codproduto) LEFT JOIN Projproduto_fornecedor PF ON PF.codproduto = P.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto) LEFT JOIN Projfamilia PFA ON PFA.Familia = P.Classe) LEFT JOIN Clientes C ON C.IDcliente = PC.IDcliente"
TextoFiltroPadrao = DataFiltro & StatusFiltro & ValidFiltro & ModuloFiltro & AplicacaoFiltro & AplicacaoFiltro2 & " and (" & TipoProduto & ") group by " & CamposFiltro & " order by " & Ordenar
'TextoFiltroPadraoRel = DataFiltroRel & StatusFiltroRel & ValidFiltroRel & ModuloFiltroRel & AplicacaoFiltroRel & AplicacaoFiltroRel2 & " and (" & TipoProdutoRel & ")and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
TextoFiltroPadraoRel = DataFiltroRel & StatusFiltroRel & ValidFiltroRel & ModuloFiltroRel & AplicacaoFiltroRel & AplicacaoFiltroRel2 & " and (" & TipoProdutoRel & ")"
    
If Lista.ListItems.Count = 0 Then
    With frmproj_produto
        .Lista.ListItems.Clear
        If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
            If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Or cmbfiltrarpor = "Fornecedor" Then
                Select Case cmbfiltrarpor
                    Case "Cliente":
                        TextoFiltro = "PC.IDCliente"
                        TextoFiltroRel = "Projproduto_clientes.IDCliente"
                    Case "Grupo do cliente":
                        TextoFiltro = "C.IDGrupo"
                        TextoFiltroRel = "Clientes.IDGrupo"
                    Case "Fornecedor":
                        TextoFiltro = "PF.IDfornecedor"
                        TextoFiltroRel = "Projproduto_fornecedor.IDfornecedor"
                End Select
                .Sql_Produto = INNERJOINTEXTO & " where " & TextoFiltro & " = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
                .FormulaRel_Produto = "{" & TextoFiltroRel & "} = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadraoRel
            ElseIf cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
                    If cmbfiltrarpor = "Família" Then
                        TextoFiltro = "P.Classe"
                        TextoFiltroRel = "Projproduto.Classe"
                    Else
                        TextoFiltro = "PFA.Grupo"
                        TextoFiltroRel = "Projfamilia.Grupo"
                    End If
                    .Sql_Produto = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                    .FormulaRel_Produto = "{" & TextoFiltroRel & "} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                ElseIf cmbfiltrarpor = "Comprimento" Or cmbfiltrarpor = "Largura" Or cmbfiltrarpor = "Espessura" Then
                        Select Case cmbfiltrarpor
                            Case "Comprimento": TextoFiltro = "P.Comprimento"
                            Case "Largura": TextoFiltro = "P.Largura"
                            Case "Espessura": TextoFiltro = "P.Espessura"
                        End Select
                        valor = txtTexto
                        NovoValor = Replace(valor, ",", ".")
                        .Sql_Produto = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
                        .FormulaRel_Produto = "{Projproduto." & TextoFiltro & "} = " & NovoValor & " and " & TextoFiltroPadraoRel
                    Else
                        Select Case cmbfiltrarpor
                            Case "Código interno": TextoFiltro = "P.desenho"
                            Case "Código de referência": TextoFiltro = "IA.N_referencia"
                            Case "Número do desenho": TextoFiltro = "IA.desenho"
                            Case "Descrição": TextoFiltro = "P.descricao"
                            Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                            Case "Dureza": TextoFiltro = "P.Dureza"
                            Case "Part number": TextoFiltro = "PFAB.Part_number"
                        End Select
                        .Sql_Produto = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                        .FormulaRel_Produto = "{" & IIf(Left(TextoFiltro, 2) = "IA", Replace(TextoFiltro, "IA.", "item_aplicacoes."), IIf(Left(TextoFiltro, 2) = "P.", Replace(TextoFiltro, "P.", "Projproduto."), Replace(TextoFiltro, "PFAB.", "Projproduto_fabricante."))) & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
            End If
        Else
            .Sql_Produto = INNERJOINTEXTO & " where " & TextoFiltroPadrao
            .FormulaRel_Produto = TextoFiltroPadraoRel
        End If
        .ProcAtualizalista (1)
    End With
Else
    TextoFiltroLista = ""
    TextoFiltroListaRel = ""
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems(InitFor).ListSubItems(1) = "Cliente" Or .ListItems(InitFor).ListSubItems(1) = "Grupo do cliente" Or .ListItems(InitFor).ListSubItems(1) = "Fornecedor" Then
                Select Case .ListItems(InitFor).ListSubItems(1)
                    Case "Cliente":
                        TextoFiltro = "PC.IDCliente"
                        TextoFiltroRel = "Projproduto_clientes.IDCliente"
                    Case "Grupo do cliente":
                        TextoFiltro = "C.IDGrupo"
                        TextoFiltroRel = "Projproduto_clientes.IDGrupo"
                    Case "Fornecedor":
                        TextoFiltro = "PF.IDfornecedor"
                        TextoFiltroRel = "Projproduto_fornecedor.IDfornecedor"
                End Select
                If TextoFiltroLista = "" Then
                    TextoFiltroLista = INNERJOINTEXTO & " where " & TextoFiltro & " = " & .ListItems(InitFor).ListSubItems(4)
                    TextoFiltroListaRel = "{" & TextoFiltroRel & "} = " & .ListItems(InitFor).ListSubItems(4)
                Else
                    TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & " = " & .ListItems(InitFor).ListSubItems(4)
                    TextoFiltroListaRel = TextoFiltroListaRel & "and {" & TextoFiltroRel & "} = " & .ListItems(InitFor).ListSubItems(4)
                End If
            ElseIf .ListItems(InitFor).ListSubItems(1) = "Família" Or .ListItems(InitFor).ListSubItems(1) = "Grupo" Then
                    If .ListItems(InitFor).ListSubItems(1) = "Família" Then
                        TextoFiltro = "P.Classe"
                        TextoFiltroRel = "Projproduto.Classe"
                    Else
                        TextoFiltro = "PF.Grupo"
                        TextoFiltroRel = "Projfamilia.Grupo"
                    End If
                    If TextoFiltroLista = "" Then
                        TextoFiltroLista = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                        TextoFiltroListaRel = "{" & TextoFiltroRel & "} = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                    Else
                        TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & " = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                        TextoFiltroListaRel = TextoFiltroListaRel & " and {" & TextoFiltroRel & "} = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                    End If
                ElseIf .ListItems(InitFor).ListSubItems(1) = "Comprimento" Or .ListItems(InitFor).ListSubItems(1) = "Largura" Or .ListItems(InitFor).ListSubItems(1) = "Espessura" Then
                        Select Case .ListItems(InitFor).ListSubItems(1)
                            Case "Comprimento": TextoFiltro = "P.Comprimento"
                            Case "Largura": TextoFiltro = "P.Largura"
                            Case "Espessura": TextoFiltro = "P.Espessura"
                        End Select
                        valor = .ListItems(InitFor).ListSubItems(3)
                        NovoValor = Replace(valor, ",", ".")
                        If TextoFiltroLista = "" Then
                            TextoFiltroLista = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor
                            TextoFiltroListaRel = "{Projproduto." & TextoFiltro & "} = " & NovoValor
                        Else
                            TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & " = " & NovoValor
                            TextoFiltroListaRel = TextoFiltroListaRel & " and {Projproduto." & TextoFiltro & "} = " & NovoValor
                        End If
                    Else
                        Select Case .ListItems(InitFor).ListSubItems(1)
                            Case "Código interno": TextoFiltro = "P.desenho"
                            Case "Código de referência": TextoFiltro = "IA.N_referencia"
                            Case "Número do desenho": TextoFiltro = "IA.desenho"
                            Case "Descrição": TextoFiltro = "P.descricao"
                            Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                            Case "Dureza": TextoFiltro = "P.Dureza"
                            Case "Part number": TextoFiltro = "PFAB.Part_number"
                        End Select
                        If TextoFiltroLista = "" Then
                            TextoFiltroLista = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3))
                            TextoFiltroListaRel = "{" & IIf(Left(TextoFiltro, 2) = "IA", Replace(TextoFiltro, "IA.", "item_aplicacoes."), IIf(Left(TextoFiltro, 2) = "P.", Replace(TextoFiltro, "P.", "Projproduto."), Replace(TextoFiltro, "PFAB.", "Projproduto_fabricante."))) & "}" & FunVerifTipoFiltroIMFListaRel(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3))
                        Else
                            TextoFiltroLista = TextoFiltroLista & " and " & TextoFiltro & FunVerifTipoFiltroIMFLista(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3))
                            TextoFiltroListaRel = TextoFiltroListaRel & " and {" & IIf(Left(TextoFiltro, 2) = "IA", Replace(TextoFiltro, "IA.", "item_aplicacoes."), IIf(Left(TextoFiltro, 2) = "P.", Replace(TextoFiltro, "P.", "Projproduto."), Replace(TextoFiltro, "PFAB.", "Projproduto_fabricante."))) & "}" & FunVerifTipoFiltroIMFListaRel(.ListItems(InitFor).ListSubItems(2), .ListItems(InitFor).ListSubItems(3))
                        End If
            End If
        Next InitFor
    End With
    With frmproj_produto
        .Lista.ListItems.Clear
        .Sql_Produto = TextoFiltroLista & " and " & TextoFiltroPadrao
        .FormulaRel_Produto = TextoFiltroListaRel & " and " & TextoFiltroPadraoRel
        .ProcAtualizalista (1)
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_excluir_Click()
On Error GoTo tratar_erro

Permitido = False
Inicio:
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If USMsgBox("Deseja realmente excluir este(s) filtro(s) para pesquisa?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                End If
                Permitido = True
                .ListItems.Remove (InitFor)
                GoTo Inicio
            End If
        Next InitFor
    End With
    If Permitido = False Then
        USMsgBox ("Informe o(s) filtro(s) para pesquisa antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Filtro(s) para pesquisa excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
        If Lista.ListItems.Count = 0 Then
            Lista.Visible = False
            Me.Height = 4395
        End If
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_Click()
On Error GoTo tratar_erro

If txtTexto.Visible = True And txtTexto = "" Or cmbfamilia.Visible = True And cmbfamilia = "" Then
    USMsgBox ("Informe o texto para pesquisa antes de adicionar o filtro na lista."), vbExclamation, "CAPRIND v5.0"
    If txtTexto.Visible = True Then txtTexto.SetFocus Else cmbfamilia.SetFocus
    Exit Sub
End If

With Lista.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = cmbfiltrarpor.Text
    If Optinicio.Value = True Then .Item(.Count).SubItems(2) = "Início"
    If Optmeio.Value = True Then .Item(.Count).SubItems(2) = "Meio"
    If Optfim.Value = True Then .Item(.Count).SubItems(2) = "Fim"
    If optIgual.Value = True Then .Item(.Count).SubItems(2) = "Igual"
    If txtTexto.Visible = True Then
        .Item(.Count).SubItems(3) = txtTexto
    Else
        .Item(.Count).SubItems(3) = cmbfamilia.Text
        .Item(.Count).SubItems(4) = cmbfamilia.ItemData(cmbfamilia.ListIndex)
    End If
End With
Lista.Visible = True
Me.Height = 6555

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: Cmd_salvar_Click
    Case vbKeyF4: Cmd_excluir_Click
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15480, 5, True
If Compras_Produtos = True Then
    Caption = "Compras - Produtos e serviços - Localizar"
    Familiatext = "C"
    With chkCompras
        .Value = 1
        .Enabled = False
    End With
ElseIf Vendas_Produtos = True Then
        Caption = "Vendas - Produtos e serviços - Localizar"
        Familiatext = "V"
        With chkVendas
            .Value = 1
            .Enabled = False
    End With
    Else
        Familiatext = "E"
End If
cmbStatus = "Liberado"
ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, 0, "Produtos/Serviços", Familiatext, False
If Permitido = False Then cmbfiltrarpor = "Código interno"

Cmb_ordenar = "Código interno"
Msk_final.Value = Date
Msk_inicio.Value = Date

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
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
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

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

If optPeriodo.Value = 1 Then
    FrameData.Enabled = True
    Msk_inicio.SetFocus
Else
    FrameData.Enabled = False
    Msk_inicio.Value = Date
    Msk_final.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
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
