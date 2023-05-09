VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Prod_Serv_DI 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Importação"
   ClientHeight    =   10425
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   12135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   820
      DibPicture      =   "frmFaturamento_Prod_Serv_DI.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmFaturamento_Prod_Serv_DI.frx":7180
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   102
      Top             =   10020
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7020
      Top             =   630
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFaturamento_Prod_Serv_DI.frx":749A
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   120
      TabIndex        =   61
      Top             =   840
      Width           =   11925
      _ExtentX        =   21034
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
      ButtonCaption1  =   "Novo"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Novo (Insert)"
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
      ButtonWidth1    =   33
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Salvar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Salvar (F3)"
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
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   38
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Excluir"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Excluir (F4)"
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
      ButtonLeft3     =   77
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
      ButtonLeft4     =   118
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
      ButtonLeft5     =   122
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
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonKey7      =   "7"
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9510
      Left            =   60
      TabIndex        =   53
      Top             =   480
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   16775
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados da DI"
      TabPicture(0)   =   "frmFaturamento_Prod_Serv_DI.frx":A872
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label31"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label32"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label33"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label34"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label35"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "PBLista"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Lista"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtTotalacessorias"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtTotalICMS"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtTotalBCIcms"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtTotalCofins"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtTotalPIS"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtTotalFrete"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Despesas"
      TabPicture(1)   =   "frmFaturamento_Prod_Serv_DI.frx":A88E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListaDespesas"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).ControlCount=   3
      Begin VB.TextBox TxtTotalFrete 
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
         Left            =   6285
         TabIndex        =   110
         ToolTipText     =   "Valor do frete internacional."
         Top             =   9030
         Width           =   1285
      End
      Begin VB.TextBox TxtTotalPIS 
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
         Left            =   1740
         TabIndex        =   109
         ToolTipText     =   "Valor do PIS."
         Top             =   9030
         Width           =   885
      End
      Begin VB.TextBox TxtTotalCofins 
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
         Left            =   2640
         TabIndex        =   108
         ToolTipText     =   "Valor do COFINS."
         Top             =   9030
         Width           =   975
      End
      Begin VB.TextBox TxtTotalBCIcms 
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
         Left            =   5040
         TabIndex        =   107
         ToolTipText     =   "Valor da base de cálculo do ICMS."
         Top             =   9030
         Width           =   1215
      End
      Begin VB.TextBox TxtTotalICMS 
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
         Left            =   7590
         TabIndex        =   106
         ToolTipText     =   "Valor do ICMS."
         Top             =   9030
         Width           =   1155
      End
      Begin VB.TextBox TxtTotalacessorias 
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
         Left            =   3630
         TabIndex        =   105
         ToolTipText     =   "Valor do COFINS."
         Top             =   9030
         Width           =   1395
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   60
         TabIndex        =   90
         Top             =   1260
         Width           =   11925
         Begin VB.CheckBox Chk_salvar_valores_produto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Salvar valores por produto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   9180
            TabIndex        =   0
            Top             =   390
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.TextBox Txt_responsavel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   92
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   345
            Width           =   4245
         End
         Begin VB.TextBox Txt_data 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            MaxLength       =   25
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   345
            Width           =   1125
         End
         Begin VB.ComboBox cmbUF_desembaraco 
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
            Height          =   330
            ItemData        =   "frmFaturamento_Prod_Serv_DI.frx":A8AA
            Left            =   5520
            List            =   "frmFaturamento_Prod_Serv_DI.frx":A902
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "UF."
            Top             =   915
            Width           =   690
         End
         Begin VB.TextBox txtCodigo_fabricante 
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
            Left            =   9630
            MaxLength       =   60
            TabIndex        =   7
            ToolTipText     =   "Código fabricante estrangeiro."
            Top             =   915
            Width           =   2085
         End
         Begin VB.TextBox txtCodigo_exportador 
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
            Left            =   7530
            MaxLength       =   60
            TabIndex        =   6
            ToolTipText     =   "Código do exportador."
            Top             =   915
            Width           =   2085
         End
         Begin VB.TextBox txtLocal_desembaraco 
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
            Left            =   180
            MaxLength       =   60
            TabIndex        =   3
            ToolTipText     =   "Local de desembaraço."
            Top             =   915
            Width           =   5320
         End
         Begin VB.TextBox txtDocumento_importacao 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
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
            Left            =   5580
            MaxLength       =   10
            TabIndex        =   1
            ToolTipText     =   "Número do documento de importação."
            Top             =   345
            Width           =   2205
         End
         Begin MSComCtl2.DTPicker Cmb_data_registro 
            Height          =   315
            Left            =   7800
            TabIndex        =   2
            ToolTipText     =   "Data de registro."
            Top             =   345
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
            Format          =   196149251
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker Cmb_data_desembaraco 
            Height          =   315
            Left            =   6210
            TabIndex        =   5
            ToolTipText     =   "Data de desembaraço."
            Top             =   915
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
            Format          =   196149251
            CurrentDate     =   39057
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
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
            Left            =   2985
            TabIndex        =   101
            Top             =   150
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
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
            Left            =   540
            TabIndex        =   100
            Top             =   150
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. fabricante estrangeiro"
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
            Left            =   9675
            TabIndex        =   99
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código do exportador"
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
            TabIndex        =   98
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. desembaraço"
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
            TabIndex        =   97
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
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
            Index           =   0
            Left            =   5775
            TabIndex        =   96
            Top             =   720
            Width           =   195
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local de desembaraço"
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
            Left            =   1995
            TabIndex        =   95
            Top             =   720
            Width           =   1590
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. de registro"
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
            Left            =   7950
            TabIndex        =   94
            Top             =   150
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DI (nº do docto. de import.)"
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
            Index           =   2
            Left            =   5715
            TabIndex        =   93
            Top             =   150
            Width           =   2010
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   55
         TabIndex        =   62
         Top             =   2610
         Width           =   11925
         Begin VB.TextBox txtAcessorias 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   22
            ToolTipText     =   "Valor do COFINS."
            Top             =   960
            Width           =   975
         End
         Begin DrawSuite2022.USLabel USLabel1 
            Height          =   195
            Left            =   5790
            TabIndex        =   117
            Top             =   2730
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   344
            Caption         =   "Obs: Despesas acessórias = II + Desp aduan. + Siscomex + Pis + Cofins"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            NoHTMLCaption   =   "Obs: Despesas acessórias = II + Desp aduan. + Siscomex + Pis + Cofins"
         End
         Begin VB.TextBox txt_NCM 
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
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   9
            ToolTipText     =   "Número da adição."
            Top             =   360
            Width           =   1380
         End
         Begin VB.ComboBox Cmb_ID_prod 
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
            ItemData        =   "frmFaturamento_Prod_Serv_DI.frx":A974
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_DI.frx":A976
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "ID do produto."
            Top             =   360
            Width           =   1155
         End
         Begin VB.CheckBox Chk_soma_II 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Somar valor do II no total das despesas acessórias"
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
            Height          =   210
            Left            =   180
            TabIndex        =   31
            Top             =   1920
            Value           =   1  'Checked
            Width           =   5445
         End
         Begin VB.CheckBox Chk_soma_aduaneiras 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Somar valor das despesas aduaneiras no total das despesas acessórias"
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
            Height          =   210
            Left            =   180
            TabIndex        =   32
            Top             =   2130
            Value           =   1  'Checked
            Width           =   6465
         End
         Begin VB.CheckBox Chk_recalcula_IPI 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Diluir valor do IPI por produto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   7230
            TabIndex        =   36
            Top             =   1950
            Width           =   2475
         End
         Begin VB.ComboBox Cmb_forma_importacao 
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
            ItemData        =   "frmFaturamento_Prod_Serv_DI.frx":A978
            Left            =   5160
            List            =   "frmFaturamento_Prod_Serv_DI.frx":A985
            Style           =   2  'Dropdown List
            TabIndex        =   30
            ToolTipText     =   "Formar de importação."
            Top             =   1530
            Width           =   6555
         End
         Begin VB.TextBox Txt_valor_AFRMM 
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
            Left            =   3510
            TabIndex        =   29
            ToolTipText     =   "Valor da AFRMM - Adicional ao Frete para Renovação da Marinha Mercante."
            Top             =   1530
            Width           =   1635
         End
         Begin VB.ComboBox Cmb_via_transporte 
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
            ItemData        =   "frmFaturamento_Prod_Serv_DI.frx":A9EB
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_DI.frx":AA0D
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Via de transporte internacional."
            Top             =   1530
            Width           =   3315
         End
         Begin VB.CheckBox Chk_soma_Cofins 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Somar valor do COFINS no total das despesas acessórias"
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
            Height          =   210
            Left            =   180
            TabIndex        =   35
            Top             =   2790
            Value           =   1  'Checked
            Width           =   5445
         End
         Begin VB.CheckBox Chk_soma_PIS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Somar valor do PIS no total das despesas acessórias"
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
            Height          =   210
            Left            =   180
            TabIndex        =   34
            Top             =   2565
            Value           =   1  'Checked
            Width           =   5445
         End
         Begin VB.CheckBox Chk_soma_siscomex 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Somar valor do siscomex no total das despesas acessórias"
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
            Height          =   210
            Left            =   180
            TabIndex        =   33
            Top             =   2355
            Value           =   1  'Checked
            Width           =   5445
         End
         Begin VB.TextBox Txt_vlr_IOF 
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
            Left            =   10740
            TabIndex        =   27
            ToolTipText     =   "Valor do imposto sobre operações financeiras."
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtNumero_adicao 
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
            Left            =   2730
            MaxLength       =   3
            TabIndex        =   10
            ToolTipText     =   "Número da adição."
            Top             =   360
            Width           =   1290
         End
         Begin VB.TextBox txtNumero_sequencial 
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
            Left            =   4035
            MaxLength       =   3
            TabIndex        =   11
            ToolTipText     =   "Número sequencial."
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox Txt_vlr_ICMS 
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
            Left            =   9750
            TabIndex        =   26
            ToolTipText     =   "Valor do ICMS."
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Txt_vlr_BC_ICMS_fator 
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
            Left            =   8370
            TabIndex        =   25
            ToolTipText     =   "Valor da base de cálculo do ICMS sem fator."
            Top             =   960
            Width           =   1365
         End
         Begin VB.TextBox Txt_vlr_BC_ICMS 
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
            Left            =   7350
            TabIndex        =   24
            ToolTipText     =   "Valor da base de cálculo do ICMS."
            Top             =   960
            Width           =   1005
         End
         Begin VB.TextBox Txt_vlr_COFINS 
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
            Left            =   4290
            TabIndex        =   21
            ToolTipText     =   "Valor do COFINS."
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Txt_vlr_PIS 
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
            Left            =   3390
            TabIndex        =   20
            ToolTipText     =   "Valor do PIS."
            Top             =   960
            Width           =   885
         End
         Begin VB.TextBox Txt_vlr_bc_PIS_COFINS 
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
            Left            =   6270
            TabIndex        =   23
            ToolTipText     =   "Valor da base de cálculo do PIS/COFINS."
            Top             =   960
            Width           =   1065
         End
         Begin VB.TextBox Txt_vlr_siscomex 
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
            Left            =   2340
            TabIndex        =   19
            ToolTipText     =   "Valor do siscomex."
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox Txt_vlr_desp_aduan 
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
            Left            =   1200
            TabIndex        =   18
            ToolTipText     =   "Valor das despesas aduaneiras."
            Top             =   960
            Width           =   1125
         End
         Begin VB.TextBox Txt_vlr_IPI 
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
            Left            =   10440
            TabIndex        =   16
            ToolTipText     =   "Valor do IPI."
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox Txt_vlr_II 
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
            Left            =   210
            TabIndex        =   17
            ToolTipText     =   "Valor do imposto de importação."
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Txt_vlr_seguro 
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
            Left            =   9150
            TabIndex        =   15
            ToolTipText     =   "Valor do seguro."
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox Txt_vlr_frete 
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
            Left            =   7848
            TabIndex        =   14
            ToolTipText     =   "Valor do frete internacional."
            Top             =   360
            Width           =   1285
         End
         Begin VB.TextBox Txt_vlr_fob 
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
            Left            =   6547
            TabIndex        =   13
            ToolTipText     =   "Valor do FOB."
            Top             =   360
            Width           =   1285
         End
         Begin VB.TextBox Txt_vlr_aduaneiro 
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
            Left            =   5246
            TabIndex        =   12
            ToolTipText     =   "Valor do aduaneiro."
            Top             =   360
            Width           =   1275
         End
         Begin VB.OptionButton Opt_vlr_BC_ICMS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Salvar valor da base do ICMS no total da base"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   7230
            TabIndex        =   37
            Top             =   2205
            Value           =   -1  'True
            Width           =   4485
         End
         Begin VB.OptionButton Opt_vlr_BC_ICMS_fator 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Salvar valor da base do ICMS s/ fator no total da base"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   420
            Left            =   7230
            TabIndex        =   38
            Top             =   2295
            Width           =   4485
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Acessórias"
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
            Left            =   5400
            TabIndex        =   104
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ID do prod."
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
            Index           =   2
            Left            =   345
            TabIndex        =   89
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de importação"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   7680
            TabIndex        =   88
            Top             =   1320
            Width           =   1515
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. AFRMM"
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
            Left            =   3915
            TabIndex        =   87
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Via de transporte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   1200
            TabIndex        =   86
            Top             =   1320
            Width           =   1260
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. IOF"
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
            TabIndex        =   81
            Top             =   780
            Width           =   555
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº da adição"
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
            Index           =   0
            Left            =   2925
            TabIndex        =   80
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº sequencial"
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
            Left            =   4155
            TabIndex        =   79
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. ICMS"
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
            Left            =   9930
            TabIndex        =   78
            Top             =   780
            Width           =   660
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BC ICMS s/ fator"
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
            Left            =   8460
            TabIndex        =   77
            Top             =   780
            Width           =   1200
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. BC ICMS"
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
            Left            =   7410
            TabIndex        =   76
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cofins     ="
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
            Left            =   4530
            TabIndex        =   75
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS      +"
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
            Left            =   3705
            TabIndex        =   74
            Top             =   780
            Width           =   630
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BC Pis Cofins"
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
            Left            =   6360
            TabIndex        =   73
            Top             =   780
            Width           =   930
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Siscomex  +"
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
            Left            =   2565
            TabIndex        =   72
            Top             =   780
            Width           =   870
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desp. aduan.  +"
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
            Left            =   1245
            TabIndex        =   71
            Top             =   780
            Width           =   1185
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. IPI"
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
            Left            =   10770
            TabIndex        =   70
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. seguro"
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
            Left            =   9390
            TabIndex        =   69
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. frete int."
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
            Left            =   8025
            TabIndex        =   68
            Top             =   180
            Width           =   930
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. FOB"
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
            Left            =   6900
            TabIndex        =   67
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "valor II   +"
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
            Index           =   0
            Left            =   405
            TabIndex        =   66
            Top             =   780
            Width           =   780
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. aduaneiro"
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
            Left            =   5385
            TabIndex        =   64
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NCM"
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
            Left            =   1860
            TabIndex        =   63
            Top             =   180
            Width           =   330
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74945
         TabIndex        =   58
         Top             =   9420
         Width           =   11925
         Begin VB.TextBox txtValorTotal 
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
            Left            =   10140
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Valor total."
            Top             =   180
            Width           =   1560
         End
         Begin DrawSuite2022.USProgressBar PBLista1 
            Height          =   255
            Left            =   180
            TabIndex        =   59
            Top             =   210
            Width           =   8775
            _ExtentX        =   15478
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
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor total :"
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
            Left            =   9090
            TabIndex        =   60
            Top             =   180
            Width           =   975
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74945
         TabIndex        =   54
         Top             =   1320
         Width           =   11925
         Begin VB.TextBox Txt_descricao_PC 
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
            Left            =   2070
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   990
            Width           =   8025
         End
         Begin VB.CommandButton Cmd_localizar_PC 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Left            =   10110
            Picture         =   "frmFaturamento_Prod_Serv_DI.frx":AAC3
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Localizar conta contábil."
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox Txt_ID_PC 
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
            MaxLength       =   255
            MouseIcon       =   "frmFaturamento_Prod_Serv_DI.frx":ABC5
            MousePointer    =   99  'Custom
            TabIndex        =   46
            Text            =   "0"
            ToolTipText     =   "ID PC."
            Top             =   990
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Txt_ID_fornecedor 
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
            Left            =   3990
            MaxLength       =   255
            TabIndex        =   42
            ToolTipText     =   "Código."
            Top             =   375
            Width           =   765
         End
         Begin VB.CommandButton Cmd_localizar_fornecedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Left            =   10110
            Picture         =   "frmFaturamento_Prod_Serv_DI.frx":AECF
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Localizar fornecedor."
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox Txt_valor 
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
            Left            =   10530
            TabIndex        =   50
            ToolTipText     =   "Valor despesa."
            Top             =   990
            Width           =   1200
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   2595
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   1185
         End
         Begin VB.TextBox Txt_codigo_PC 
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
            MaxLength       =   255
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Código."
            Top             =   990
            Width           =   1875
         End
         Begin VB.TextBox Txt_fornecedor 
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
            Left            =   4770
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   375
            Width           =   5325
         End
         Begin MSComCtl2.DTPicker Cmb_data_vencimento 
            Height          =   315
            Left            =   10530
            TabIndex        =   45
            ToolTipText     =   "Data de vencimento."
            Top             =   375
            Width           =   1200
            _ExtentX        =   2117
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. vencto."
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
            Left            =   10703
            TabIndex        =   85
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   5722
            TabIndex        =   84
            Top             =   780
            Width           =   720
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
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
            Index           =   5
            Left            =   855
            TabIndex        =   83
            Top             =   780
            Width           =   510
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
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
            Left            =   7020
            TabIndex        =   82
            Top             =   180
            Width           =   825
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            TabIndex        =   57
            Top             =   780
            Width           =   360
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
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
            Index           =   11
            Left            =   2220
            TabIndex        =   56
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
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
            Index           =   10
            Left            =   600
            TabIndex        =   55
            Top             =   180
            Width           =   345
         End
      End
      Begin MSComctlLib.ListView ListaDespesas 
         Height          =   6645
         Left            =   -74940
         TabIndex        =   51
         Top             =   2775
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   11721
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
         BorderStyle     =   1
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Fornecedor"
            Object.Width           =   5851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Dt. vcto."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Código"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   5851
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Pago"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   3105
         Left            =   60
         TabIndex        =   39
         Top             =   5700
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   5477
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   512
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "DI"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Dt. registro"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "NCM"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "N° adição"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "N° sequencial"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "ID_do Produto"
            Object.Width           =   2295
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   60
         TabIndex        =   65
         Top             =   8250
         Width           =   11925
         _ExtentX        =   21034
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
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total frete"
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
         Left            =   6465
         TabIndex        =   116
         Top             =   8850
         Width           =   765
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total PIS"
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
         Left            =   1875
         TabIndex        =   115
         Top             =   8850
         Width           =   645
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cofins"
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
         Left            =   2700
         TabIndex        =   114
         Top             =   8850
         Width           =   855
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total BC ICMS"
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
         Left            =   5100
         TabIndex        =   113
         Top             =   8850
         Width           =   1020
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS"
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
         Left            =   7770
         TabIndex        =   112
         Top             =   8850
         Width           =   960
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Acessórias"
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
         Left            =   3750
         TabIndex        =   111
         Top             =   8850
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_DI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_DI  As Boolean
Dim Novo_DI1  As Boolean

Private Sub ProcCalculaAcessorias()
On Error GoTo tratar_erro
Dim vlrII As Double
Dim vlrDespAduan As Double
Dim vlrSiscomex As Double

vlrII = IIf(Txt_vlr_II <> "", Txt_vlr_II, 0)

vlrDespAduan = IIf(Txt_vlr_desp_aduan <> "", Txt_vlr_desp_aduan, 0)

vlrSiscomex = IIf(Txt_vlr_siscomex <> "", Txt_vlr_siscomex, 0)

VlrPIS = IIf(Txt_vlr_PIS <> "", Txt_vlr_PIS, 0)

VlrCofins = IIf(Txt_vlr_COFINS <> "", Txt_vlr_COFINS, 0)


txtAcessorias.Text = vlrII + vlrDespAduan + vlrSiscomex + VlrPIS + VlrCofins
txtAcessorias.Text = Format(txtAcessorias, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub ProcSomaTotais()
On Error GoTo tratar_erro

StrSql = "Select Sum(Valor_bc_ICMS) as TotalBaseICMS,SUM(Valor_ICMS) as TotalICMS, SUM(Valor_PIS) as TotalPIS, SUM(Valor_Cofins) as TotalCofins, SUM(valor_frete_int) as TotalFrete, SUM(Valor_imposto_importacao+Valor_despesas+Valor_sixcomex+Valor_PIS+Valor_cofins) as TotalAcessorias from tbl_Detalhes_Nota_NFe where ID_nota = '26766'"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
TxtTotalPIS = Format(TBAbrir!TotalPIS, "###,##0.00")
TxtTotalCofins = Format(TBAbrir!TotalCofins, "###,##0.00")
TxtTotalacessorias = Format(TBAbrir!TotalAcessorias, "###,##0.00")
TxtTotalBCIcms = Format(TBAbrir!TotalBaseICMS, "###,##0.00")
TxtTotalFrete = Format(TBAbrir!TotalFrete, "###,##0.00")
txtTotalICMS = Format(TBAbrir!TotalICMS, "###,##0.00")
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
Dim valor As Double

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmFaturamento_Prod_Serv
    If .txtNFiscal = "" Then
        If FunVerifValidacaoRegistro("salvar", .txtDtValidacao, "ordem de faturamento", "os dados de importação", False) = False Then Exit Sub
    Else
        If FunVerifValidacaoRegistro("salvar", .txtDtValidacao, "nota fiscal", "os dados de importação", False) = False Then Exit Sub
    End If
End With
Acao = "salvar"
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtDocumento_importacao = "" Then
    NomeCampo = "o número da DI"
    ProcVerificaAcao
    txtDocumento_importacao.SetFocus
    Exit Sub
End If
If txtLocal_desembaraco = "" Then
    NomeCampo = "o local de desembaraço"
    ProcVerificaAcao
    txtLocal_desembaraco.SetFocus
    Exit Sub
End If
If cmbUF_desembaraco = "" Then
    NomeCampo = "a UF"
    ProcVerificaAcao
    cmbUF_desembaraco.SetFocus
    Exit Sub
End If
If txtCodigo_exportador = "" Then
    NomeCampo = "o código do exportador"
    ProcVerificaAcao
    txtCodigo_exportador.SetFocus
    Exit Sub
End If
If txtCodigo_fabricante = "" Then
    NomeCampo = "o código do fabricante estrangeiro"
    ProcVerificaAcao
    txtCodigo_fabricante.SetFocus
    Exit Sub
End If
If txt_NCM = "" Then
    NomeCampo = "o NCM"
    ProcVerificaAcao
    txt_NCM.SetFocus
    Exit Sub
End If
If txtNumero_adicao = "" Then
    NomeCampo = "o número da adição"
    ProcVerificaAcao
    txtNumero_adicao.SetFocus
    Exit Sub
End If
valor = IIf(txtNumero_sequencial = "", 0, txtNumero_sequencial)
If valor <= 0 Then
    NomeCampo = "o número sequencial"
    ProcVerificaAcao
    txtNumero_sequencial.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_aduaneiro = "", 0, Txt_vlr_aduaneiro)
If Txt_vlr_aduaneiro = "" Or valor < 0 Then
    NomeCampo = "o valor aduaneiro"
    ProcVerificaAcao
    Txt_vlr_aduaneiro.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_fob = "", 0, Txt_vlr_fob)
If Txt_vlr_fob = "" Or valor < 0 Then
    NomeCampo = "o valor do fob"
    ProcVerificaAcao
    Txt_vlr_fob.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_frete = "", 0, Txt_vlr_frete)
If Txt_vlr_frete = "" Or valor < 0 Then
    NomeCampo = "o valor do frete"
    ProcVerificaAcao
    Txt_vlr_frete.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_seguro = "", 0, Txt_vlr_seguro)
If Txt_vlr_seguro = "" Or valor < 0 Then
    NomeCampo = "o valor do seguro"
    ProcVerificaAcao
    Txt_vlr_seguro.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_II = "", 0, Txt_vlr_II)
If Txt_vlr_II = "" Or valor < 0 Then
    NomeCampo = "o valor do imposto de importação"
    ProcVerificaAcao
    Txt_vlr_II.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_IPI = "", 0, Txt_vlr_IPI)
If Txt_vlr_IPI = "" Or valor < 0 Then
    NomeCampo = "o valor do IPI"
    ProcVerificaAcao
    Txt_vlr_IPI.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_desp_aduan = "", 0, Txt_vlr_desp_aduan)
If Txt_vlr_desp_aduan = "" Or valor < 0 Then
    NomeCampo = "o valor das despesas aduaneiras"
    ProcVerificaAcao
    Txt_vlr_desp_aduan.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_siscomex = "", 0, Txt_vlr_siscomex)
If Txt_vlr_siscomex = "" Or valor < 0 Then
    NomeCampo = "o valor do siscomex"
    ProcVerificaAcao
    Txt_vlr_siscomex.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_bc_PIS_COFINS = "", 0, Txt_vlr_bc_PIS_COFINS)
If Txt_vlr_bc_PIS_COFINS = "" Or valor < 0 Then
    NomeCampo = "o valor da base de cálculo do PIS/COFINS"
    ProcVerificaAcao
    Txt_vlr_bc_PIS_COFINS.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_PIS = "", 0, Txt_vlr_PIS)
If Txt_vlr_PIS = "" Or valor < 0 Then
    NomeCampo = "o valor do PIS"
    ProcVerificaAcao
    Txt_vlr_PIS.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_COFINS = "", 0, Txt_vlr_COFINS)
If Txt_vlr_COFINS = "" Or valor < 0 Then
    NomeCampo = "o valor do COFINS"
    ProcVerificaAcao
    Txt_vlr_COFINS.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_BC_ICMS = "", 0, Txt_vlr_BC_ICMS)
If Txt_vlr_BC_ICMS = "" Or valor < 0 Then
    NomeCampo = "o valor da base de cálculo do ICMS"
    ProcVerificaAcao
    Txt_vlr_BC_ICMS.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_BC_ICMS_fator = "", 0, Txt_vlr_BC_ICMS_fator)
If Txt_vlr_BC_ICMS_fator = "" Or valor < 0 Then
    NomeCampo = "o valor da base de cálculo do ICMS sem fator"
    ProcVerificaAcao
    Txt_vlr_BC_ICMS_fator.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_ICMS = "", 0, Txt_vlr_ICMS)
If Txt_vlr_ICMS = "" Or valor < 0 Then
    NomeCampo = "o valor do ICMS"
    ProcVerificaAcao
    Txt_vlr_ICMS.SetFocus
    Exit Sub
End If
valor = IIf(Txt_vlr_IOF = "", 0, Txt_vlr_IOF)
If Txt_vlr_IOF = "" Or valor < 0 Then
    NomeCampo = "o valor do IOF"
    ProcVerificaAcao
    Txt_vlr_IOF.SetFocus
    Exit Sub
End If
If Cmb_via_transporte = "" Then
    NomeCampo = "a via de transporte"
    ProcVerificaAcao
    Cmb_via_transporte.SetFocus
    Exit Sub
End If
If Left(Cmb_via_transporte, 1) = 1 Then
    valor = IIf(Txt_valor_AFRMM = "", 0, Txt_valor_AFRMM)
    If Txt_valor_AFRMM = "" Or valor < 0 Then
        NomeCampo = "o Valor da AFRMM"
        ProcVerificaAcao
        Txt_valor_AFRMM.SetFocus
        Exit Sub
    End If
End If
If Cmb_forma_importacao = "" Then
    NomeCampo = "a forma de importação"
    ProcVerificaAcao
    Cmb_forma_importacao.SetFocus
    Exit Sub
End If

With frmFaturamento_Prod_Serv
    Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set Alterar = 'False' where ID = " & .txtId

    TextoFiltro = ""
    If Chk_salvar_valores_produto.Value = 1 Then
        valor = Txt_vlr_aduaneiro.Text 'Txt_vlr_fob - Txt_vlr_frete
        NovoValor = Replace(valor, ",", ".")
        TextoFiltro = " and NFP.dbl_ValorTotal = " & NovoValor

'        'Verifica se existem dois produtos com o mesmo valor total e pede para escolher um deles
'        Set TBAbrir = CreateObject("adodb.recordset")
'        TBAbrir.Open "Select NFP.Int_codigo from tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.ID_nota = " & .txtId & " and CF.IDIntClasse = '" & Cmb_NCM & "' and NFP.Retorno = 'False' and NFP.Remessa = 'False'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
'        If TBAbrir.EOF = False Then
'            If Cmb_ID_prod = "" Then
'                If TBAbrir.RecordCount > 1 Then
'                    USMsgBox ("Existe mais de um produto com o valor total de " & Format(valor, "###,##0.00") & ", favor informar o ID do produto."), vbInformation, "CAPRIND v5.0"
'                    With Cmb_ID_prod
'                        .Clear
'                        .Locked = False
'                        .TabStop = True
'                        .SetFocus
'                        Do While TBAbrir.EOF = False
'                            .AddItem TBAbrir!Int_codigo
'                            TBAbrir.MoveNext
'                        Loop
'                    End With
'                    TBAbrir.Close
'                    Exit Sub
'                Else
'                    With Cmb_ID_prod
'                        .Clear
'                        .AddItem TBAbrir!Int_codigo
'                        .Text = TBAbrir!Int_codigo
'                    End With
'                End If
'            End If
'        Else
'            USMsgBox ("Não é permitido utilizar a opção de salvar valores por produto, pois não foi encontrado nenhum produto na nota fiscal com o valor total de " & Format(valor, "###,##0.00") & "."), vbInformation, "CAPRIND v5.0"
'            TBAbrir.Close
'            Exit Sub
'        End If
'        TBAbrir.Close

        TextoFiltro = " and NFP.Int_codigo = " & Cmb_ID_prod
'    End If
'
'    'Dilui os valores digitados por produto com a NCM informada
'    'Soma valor total dos produtos
'    ValorTotal = 0
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select Sum(NFP.dbl_ValorTotal) as Valortotal from tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.ID_nota = " & .txtId & " and CF.IDIntClasse = '" & Cmb_NCM & "' and NFP.Retorno = 'False' and NFP.Remessa = 'False'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        ValorTotal = IIf(IsNull(TBAbrir!ValorTotal), 0, TBAbrir!ValorTotal)
'    End If
'
'    Set TBAbrir = CreateObject("adodb.recordset")
'    TBAbrir.Open "Select NFP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.ID_nota = " & .txtId & " and CF.IDIntClasse = '" & Cmb_NCM & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
'    If TBAbrir.EOF = False Then
'        'Do While TBAbrir.EOF = False
'            If ValorTotal <> 0 Then
'                Qtde = (IIf(IsNull(TBAbrir!dbl_ValorTotal), 0, TBAbrir!dbl_ValorTotal) * 100) / ValorTotal
'            Else
'                If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) = 2 Then Qtde = TBAbrir.RecordCount Else Qtde = 0
'            End If

            Set TBGravar = CreateObject("adodb.recordset")
            StrSql = "Select * from tbl_Detalhes_Nota_NFe where ID = " & IDlista
            'Debug.print StrSql
            
            TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                Novo_DI = True
                Evento = "Nova DI"
                'Id_Item = frmFaturamento_Prod_Serv.txtId
            Else
                Novo_DI = False
                Evento = "Alterar DI"
            End If
            '==================================
            Modulo = Formulario & "/Importação"
            'ID_documento = TBAbrir!Int_codigo
            .ProcVerificaTipoNF False
            If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
            Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
            Documento1 = "DI: " & txtDocumento_importacao & " - NCM: " & txt_NCM & " - Adição: " & txtNumero_adicao & " - Sequencial: " & txtNumero_sequencial
            ProcGravaEvento
            '==================================
            
            TBGravar!ID_nota = .txtId
            If Chk_salvar_valores_produto.Value = 1 Then TBGravar!Salvar_valores_produto = True Else TBGravar!Salvar_valores_produto = False
            TBGravar!Id_Item = Cmb_ID_prod
            If Txt_data = "" Then TBGravar!Data = Date Else TBGravar!Data = Txt_data
            If Txt_responsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = Txt_responsavel
            TBGravar!Documento_importacao = IIf(txtDocumento_importacao = "", Null, txtDocumento_importacao)
            TBGravar!Data_registro = Cmb_data_registro
            TBGravar!Local_desembaraco = IIf(txtLocal_desembaraco = "", Null, txtLocal_desembaraco)
            TBGravar!UF_desembaraco = IIf(cmbUF_desembaraco = "", Null, cmbUF_desembaraco)
            TBGravar!Data_desembaraco = Cmb_data_desembaraco
            TBGravar!Codigo_exportador = IIf(txtCodigo_exportador = "", Null, txtCodigo_exportador)
            TBGravar!Codigo_fabricante = IIf(txtCodigo_fabricante = "", Null, txtCodigo_fabricante)
            TBGravar!NCM = txt_NCM
            TBGravar!Numero_adicao = IIf(txtNumero_adicao = "", Null, txtNumero_adicao)
            TBGravar!Numero_sequencial = IIf(txtNumero_sequencial = "", Null, txtNumero_sequencial)
            If Chk_soma_II.Value = 1 Then TBGravar!Soma_II = True Else TBGravar!Soma_II = False
            If Chk_soma_aduaneiras.Value = 1 Then TBGravar!Soma_aduaneiras = True Else TBGravar!Soma_aduaneiras = False
            If Chk_soma_siscomex.Value = 1 Then TBGravar!Soma_siscomex = True Else TBGravar!Soma_siscomex = False
            If Chk_soma_PIS.Value = 1 Then TBGravar!Soma_PIS = True Else TBGravar!Soma_PIS = False
            If Chk_soma_Cofins.Value = 1 Then TBGravar!Soma_Cofins = True Else TBGravar!Soma_Cofins = False
            If Chk_recalcula_IPI.Value = 1 Then TBGravar!Recalcula_IPI = True Else TBGravar!Recalcula_IPI = False
            If Opt_vlr_BC_ICMS.Value = True Then TBGravar!Opt_valor_bc_ICMS = True Else TBGravar!Opt_valor_bc_ICMS = False
            
                                        
            TBGravar!Via_transp = Left(Cmb_via_transporte, 1)
            TBGravar!Valor_AFRMM = IIf(Txt_valor_AFRMM = "", Null, Txt_valor_AFRMM)
            TBGravar!Forma_imp = Left(Cmb_forma_importacao, 1)
                                        
            procSalvaValores
            TBGravar.Update
           
            TBGravar.Close
            
'
'            TBAbrir.MoveNext
'        Loop
    End If
   ' TBAbrir.Close
    
    If Novo_DI = True Then USMsgBox ("Nova DI cadastrada com sucesso."), vbInformation, "CAPRIND v5.0" Else USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        
    'Verifica totais por NCM
    VLFRETE = 0
    VLSEGURO = 0
    valor = 0
    Valor1 = 0
    Valor2 = 0
    ValorPagar = 0
    ValorPago = 0
    Valor3 = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select NFPNF.* from (tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF) INNER JOIN tbl_Detalhes_Nota_NFe NFPNF ON NFPNF.ID_item = NFP.Int_codigo where NFP.ID_nota = " & .txtId & " and CF.IDIntClasse = '" & txt_NCM & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            VLFRETE = VLFRETE + IIf(IsNull(TBAbrir!Valor_frete_int), 0, TBAbrir!Valor_frete_int)
            VLSEGURO = VLSEGURO + IIf(IsNull(TBAbrir!Valor_seguro), 0, TBAbrir!Valor_seguro)
            
            If TBAbrir!Soma_II = True Then valor = valor + TBAbrir!Valor_imposto_importacao
            If TBAbrir!Soma_aduaneiras = True Then valor = valor + TBAbrir!Valor_despesas
            If TBAbrir!Soma_siscomex = True Then valor = valor + TBAbrir!Valor_sixcomex
            If TBAbrir!Soma_PIS = True Then valor = valor + TBAbrir!valor_pis
            If TBAbrir!Soma_Cofins = True Then valor = valor + TBAbrir!valor_cofins
            Valor1 = Valor1 + TBAbrir!Valor_IPI
            Valor2 = Valor2 + TBAbrir!valor_pis
            ValorPagar = ValorPagar + TBAbrir!valor_cofins
            ValorPago = ValorPago + (TBAbrir!Valor_bc_ICMS + TBAbrir!Valor_bc_ICMS_fator)
            Valor3 = Valor3 + TBAbrir!Valor_ICMS
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
                               'Frete   Seguro    D. Ac  IPI     PIS     Cofins      BC ICMS    VLR. ICMS
    ProcCalculaValoresProdutos VLFRETE, VLSEGURO, valor, Valor1, Valor2, ValorPagar, ValorPago, Valor3
    
    ProcCarregaLista
    If Novo_DI = False Then
        If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
            Lista.SelectedItem = Lista.ListItems(CodigoLista)
            Lista.SetFocus
        End If
    End If
    
    '.ProcCarregaLista
    
1:
    Novo_DI = False
End With

'With Cmb_ID_prod
'    .Locked = True
'    .TabStop = False
'End With

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvarDespesas()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If Frame5.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Txt_fornecedor = "" Then
    NomeCampo = "o fornecedor"
    ProcVerificaAcao
    Cmd_localizar_fornecedor_Click
    Exit Sub
End If
If Txt_ID_PC = 0 Then
    NomeCampo = "a conta contábil"
    ProcVerificaAcao
    Cmd_localizar_PC_Click
    Exit Sub
End If
If Txt_valor = "" Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    Txt_valor.SetFocus
    Exit Sub
End If

With frmFaturamento_Prod_Serv
    If Novo_DI1 = True Then
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select tbl_contaspagar.* from tbl_contaspagar INNER JOIN Familia_financeiro ON tbl_contaspagar.IDIntconta = Familia_financeiro.IDconta where tbl_contaspagar.ID_nota = " & .txtId & " and tbl_contaspagar.int_codforn = " & Txt_ID_fornecedor & " and tbl_contaspagar.dt_Pagamento = '" & Cmb_data_vencimento & "' and tbl_contaspagar.Despesas_NF = 'True' and Familia_financeiro.ID_PC = " & Txt_ID_PC, Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            USMsgBox ("Esta conta contábil está sendo utilizada nesta nota fiscal para este fornecedor."), vbExclamation, "CAPRIND v5.0"
            TBContas.Close
            Exit Sub
        End If
    End If

    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_contaspagar where int_codforn = " & Txt_ID_fornecedor & " and dt_Pagamento = '" & Cmb_data_vencimento & "' and Despesas_NF = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        With frmFaturamento_Prod_Serv
            If .txtNFiscal = "" Then
                If FunVerifValidacaoRegistro("alterar", .txtDtValidacao, "ordem de faturamento", "essa despesa", False) = False Then Exit Sub
            Else
                If FunVerifValidacaoRegistro("alterar", .txtDtValidacao, "nota fiscal", "essa despesa", False) = False Then Exit Sub
            End If
        End With
        
        If TBContas!Logsit = "S" Then
            USMsgBox ("Não é permitido alterar esta despesa, pois a mesma já foi baixada."), vbExclamation, "CAPRIND v5.0"
            TBContas.Close
            Exit Sub
        End If
        Evento = "Alterar despesa"
    Else
        TBContas.AddNew
        TBContas!Despesas_NF = True
        TBContas!Antecipacao = False
        TBContas!Devolucao = False
        TBContas!Parcial = False
        TBContas!impresso = False
        TBContas!Bloqueado = False
        TBContas!Logsit = "N"
        Evento = "Nova despesa"
    End If
    TBContas!ID_nota = .txtId
    TBContas!Data_transacao = txtData
    TBContas!Dt_emissao = Date
    TBContas!dt_Pagamento = Cmb_data_vencimento.Value
    TBContas!dbl_valorpagto = Txt_valor
    'TBContas!Banco = cmbbanco
    'TBContas!FormaBaixa = cmb_forma
    'TBContas!txt_observacoes = txtObs.Text
    'TBContas!txt_pedido = txtNPedido.Text
    
    TBContas!Tipo = "FO"
    TBContas!int_codforn = Txt_ID_fornecedor
    TBContas!Txt_fornecedor = Txt_fornecedor
    
    TBContas!txt_ndocumento = IIf(.txtNFiscal = "", Null, .txtNFiscal)
    'TBContas!Class_conta = cmbtipo_conta.Text
    TBContas!Responsavel = txtResponsavel
    TBContas!txt_Parcela = "01/01"
    TBContas!status = "TÍTULO EM ABERTO"
    TBContas!ID_empresa = .txtIDEmpresa.Text
    
    TBContas.Update
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = True Then TBFluxo.AddNew
    TBFluxo!IDintconta = TBContas!IDintconta
    TBFluxo!Operacao = "À Debitar"
    TBFluxo!Data = Cmb_data_vencimento
    TBFluxo!valor = Txt_valor
    TBFluxo!Descricao = Txt_fornecedor
    TBFluxo!status = "N"
    TBFluxo!int_NotaFiscal = IIf(.txtNFiscal = "", Null, .txtNFiscal)
    TBFluxo!Instituicao = Null
    TBFluxo!Bloqueado = False
    TBFluxo!ID_empresa = .txtIDEmpresa.Text
    TBFluxo.Update
    Conexao.Execute "UPDATE tbl_contaspagar set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBContas!IDintconta
    TBFluxo.Close
    
    'Conta contábil
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from familia_financeiro where IDconta = " & TBContas!IDintconta & " and ID_PC = " & Txt_ID_PC & " and TipoConta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then TBGravar.AddNew
    TBGravar!ID_PC = Txt_ID_PC
    TBGravar!IDConta = TBContas!IDintconta
    TBGravar!valor = Txt_valor
    TBGravar!TipoConta = "P"
    TBGravar.Update
    TBGravar.Close
    
    'Verifica valor total da conta, de acordo com as contas contábeis
    valor = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(Valor) as Valor from familia_financeiro where IDconta = " & TBContas!IDintconta & " and TipoConta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
    End If
    TBAbrir.Close
    NovoValor = Replace(valor, ",", ".")
    
    Conexao.Execute "UPDATE tbl_contaspagar set dbl_valorpagto = " & NovoValor & " where IdIntConta = " & TBContas!IDintconta
    Conexao.Execute "UPDATE tbl_Fluxo_de_caixa set Valor = " & NovoValor & " where IDFluxo = " & TBContas!IDFluxo
    
    '==================================
    Modulo = Formulario & "/Importação"
    ID_documento = TBContas!IDintconta
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
    Documento1 = "Fornecedor: " & Txt_fornecedor & " - Dt. de vencimento: " & Format(Cmb_data_vencimento, "dd/mm/yy") & " - Código contábil: " & Txt_codigo_PC & " - Descrição: " & Txt_descricao_PC
    ProcGravaEvento
    '==================================
    TBContas.Close
    
    ProcCarregaListaDespesas
    ProcAtualizarEstoque
    If Novo_DI1 = True Then
        USMsgBox ("Nova despesa cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Nova despesa"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar despesa"
        If CodigoLista1 <> 0 And ListaDespesas.ListItems.Count <> 0 Then
            ListaDespesas.SelectedItem = ListaDespesas.ListItems(CodigoLista1)
            ListaDespesas.SetFocus
        End If
    End If

1:
End With
    Novo_DI1 = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmFaturamento_Prod_Serv
    If .txtNFiscal = "" Then
        If FunVerifValidacaoRegistro("criar novos", .txtDtValidacao, "ordem de faturamento", "dados de importação", False) = False Then Exit Sub
    Else
        If FunVerifValidacaoRegistro("criar novos", .txtDtValidacao, "nota fiscal", "dados de importação", False) = False Then Exit Sub
    End If
End With
ProcLimpaCampos
If txtCodigo_exportador = "" Then
    If USMsgBox("Deseja utilizar o ID do destinatário no código do Código do exportador?", vbYesNo, "CAPRIND v5.0") = vbYes Then txtCodigo_exportador = frmFaturamento_Prod_Serv.txtIDcliente
End If
Frame1.Enabled = True
Novo_DI = True
If txtDocumento_importacao = "" Then txtDocumento_importacao.SetFocus Else txt_NCM.SetFocus
IDlista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcNovoDespesas()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmFaturamento_Prod_Serv
    If .txtNFiscal = "" Then
        If FunVerifValidacaoRegistro("criar nova", .txtDtValidacao, "ordem de faturamento", "despesa", False) = False Then Exit Sub
    Else
        If FunVerifValidacaoRegistro("criar nova", .txtDtValidacao, "nota fiscal", "despesa", False) = False Then Exit Sub
    End If
End With
ProcLimpaCamposDespesas
Frame5.Enabled = True
Novo_DI1 = True
Cmd_localizar_fornecedor_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_data = Format(Date, "dd/mm/yy")
Txt_responsavel = pubUsuario
txt_NCM.Text = ""
'Cmb_ID_prod.Clear
txtNumero_adicao = ""
txtNumero_sequencial = ""
Txt_vlr_aduaneiro = ""
Txt_vlr_fob = ""
Txt_vlr_frete = ""
Txt_vlr_seguro = ""
Txt_vlr_II = ""
'Chk_recalcula_IPI.Value = 0
Txt_vlr_IPI = ""
Txt_vlr_desp_aduan = ""
'Chk_soma_II.Value = 0
'Chk_soma_aduaneiras.Value = 0
'Chk_soma_siscomex.Value = 0
Txt_vlr_siscomex = ""
Txt_vlr_bc_PIS_COFINS = ""
'Chk_soma_PIS.Value = 0
Txt_vlr_PIS = ""
'Chk_soma_Cofins.Value = 0
Txt_vlr_COFINS = ""
Opt_vlr_BC_ICMS.Value = True
Txt_vlr_BC_ICMS = ""
Opt_vlr_BC_ICMS_fator.Value = False
Txt_vlr_BC_ICMS_fator = ""
Txt_vlr_ICMS = ""
Txt_vlr_IOF = ""
Novo_DI = False
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLimpaCamposDespesas()
On Error GoTo tratar_erro

txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""
Txt_valor = ""
Novo_DI1 = False
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub




Private Sub Cmb_ID_prod_Change()
On Error GoTo tratar_erro

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select CF.IDIntClasse from tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.int_codigo = '" & Cmb_ID_prod & "' and NFP.Retorno = 'False' and NFP.Remessa = 'False' Group by CF.IDIntClasse", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
            txt_NCM.Text = TBAbrir!IDIntClasse
    End If
    TBAbrir.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_ID_prod_Click()
On Error GoTo tratar_erro

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select CF.IDIntClasse from tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.int_codigo = '" & Cmb_ID_prod & "' and NFP.Retorno = 'False' and NFP.Remessa = 'False' Group by CF.IDIntClasse", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
            txt_NCM.Text = TBAbrir!IDIntClasse
    End If
    TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_via_transporte_Click()
On Error GoTo tratar_erro

With Txt_valor_AFRMM
    If Left(Cmb_via_transporte, 1) <> 1 Then
        .Text = ""
        .Locked = True
        .TabStop = False
    Else
        .Locked = False
        .TabStop = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_localizar_fornecedor_Click()
On Error GoTo tratar_erro

Sit_REG = 5
ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False
FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro
  
Plano_contas_produtos = False
Plano_contas_familias = False
Plano_centro_de_custo = False
Plano_instituicao = False
Plano_opcoesgerais = False
Plano_Faturamento = True
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
Plano_PCP = False
Sit_REG = 2
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            'Case vbKeyF1: ProcAjuda
           ' Case 13: SendKeys "{TAB}"
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoDespesas
            Case vbKeyF3: ProcSalvarDespesas
            Case vbKeyF4: ProcExcluirDespesas
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11925, 7, True

If Formulario = "Faturamento/Nota fiscal/Própria" Then
     Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Importação"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Importação"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            Caption = "Estoque - Ordem de faturamento - Importação"
        Else
            Caption = "Estoque - Nota fiscal - Importação"
End If
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Cmb_data_registro.Value = Date
Cmb_data_desembaraco.Value = Date
Cmb_data_vencimento.Value = Date


With Cmb_ID_prod
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Int_codigo from tbl_Detalhes_Nota where ID_nota = " & frmFaturamento_Prod_Serv.txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Int_codigo
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close

End With

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcPuxaDados()
On Error GoTo tratar_erro

Txt_data = IIf(IsNull(TBAbrir!Data), Format(Date, "dd/mm/yy"), Format(TBAbrir!Data, "dd/mm/yy"))
Txt_responsavel = IIf(IsNull(TBAbrir!Responsavel), pubUsuario, TBAbrir!Responsavel)
If IsNull(TBAbrir!NCM) = False And TBAbrir!NCM <> "" Then txt_NCM.Text = TBAbrir!NCM

'If TBAbrir!Salvar_valores_produto = True Then
'    Chk_salvar_valores_produto.Value = 1
'    With Cmb_ID_prod
'        .Clear
'        .AddItem TBAbrir!Id_Item
'        .Text = TBAbrir!Id_Item
'        .Locked = True
'        .TabStop = False
'    End With
'Else
'    Chk_salvar_valores_produto.Value = 0
'End If

txtDocumento_importacao = IIf(IsNull(TBAbrir!Documento_importacao), "", TBAbrir!Documento_importacao)
Cmb_data_registro.Value = IIf(IsNull(TBAbrir!Data_registro), Format(Date, "dd/mm/yy"), Format(TBAbrir!Data_registro, "dd/mm/yy"))
txtLocal_desembaraco = IIf(IsNull(TBAbrir!Local_desembaraco), "", TBAbrir!Local_desembaraco)
If IsNull(TBAbrir!UF_desembaraco) = False And TBAbrir!UF_desembaraco <> "" Then cmbUF_desembaraco = TBAbrir!UF_desembaraco
Cmb_data_desembaraco.Value = IIf(IsNull(TBAbrir!Data_desembaraco), Format(Date, "dd/mm/yy"), Format(TBAbrir!Data_desembaraco, "dd/mm/yy"))
txtCodigo_exportador = IIf(IsNull(TBAbrir!Codigo_exportador), "", TBAbrir!Codigo_exportador)
txtCodigo_fabricante = IIf(IsNull(TBAbrir!Codigo_fabricante), "", TBAbrir!Codigo_fabricante)
txtNumero_adicao = IIf(IsNull(TBAbrir!Numero_adicao), "", TBAbrir!Numero_adicao)
txtNumero_sequencial = IIf(IsNull(TBAbrir!Numero_sequencial), "", TBAbrir!Numero_sequencial)
Txt_vlr_aduaneiro = IIf(IsNull(TBAbrir!valor), "0,00", Format(TBAbrir!valor, "###,##0.0000"))
Txt_vlr_fob = IIf(IsNull(TBAbrir!Valor1), "0,00", Format(TBAbrir!Valor1, "###,##0.0000"))
Txt_vlr_frete = IIf(IsNull(TBAbrir!Valor2), "0,00", Format(TBAbrir!Valor2, "###,##0.0000"))
Txt_vlr_seguro = IIf(IsNull(TBAbrir!ValorPagar), "0,00", Format(TBAbrir!ValorPagar, "###,##0.0000"))
Txt_vlr_II = IIf(IsNull(TBAbrir!ValorPago), "0,00", Format(TBAbrir!ValorPago, "###,##0.0000"))
Txt_vlr_desp_aduan = IIf(IsNull(TBAbrir!ValorPagoParcial), "0,00", Format(TBAbrir!ValorPagoParcial, "###,##0.00"))
If TBAbrir!Recalcula_IPI = True Then Chk_recalcula_IPI.Value = 1 Else Chk_recalcula_IPI.Value = 0
Txt_vlr_IPI = IIf(IsNull(TBAbrir!Valor_IPI), "0,00", Format(TBAbrir!Valor_IPI, "###,##0.0000"))
If TBAbrir!Soma_II = True Then Chk_soma_II.Value = 1 Else Chk_soma_II.Value = 0
If TBAbrir!Soma_aduaneiras = True Then Chk_soma_aduaneiras.Value = 1 Else Chk_soma_aduaneiras.Value = 0
If TBAbrir!Soma_siscomex = True Then Chk_soma_siscomex.Value = 1 Else Chk_soma_siscomex.Value = 0
Txt_vlr_siscomex = IIf(IsNull(TBAbrir!Valorparcela), "0,00", Format(TBAbrir!Valorparcela, "###,##0.0000"))
Txt_vlr_bc_PIS_COFINS = IIf(IsNull(TBAbrir!Valor_total), "0,00", Format(TBAbrir!Valor_total, "###,##0.0000"))
If TBAbrir!Soma_PIS = True Then Chk_soma_PIS.Value = 1 Else Chk_soma_PIS.Value = 0
Txt_vlr_PIS = IIf(IsNull(TBAbrir!Valor_PIS_Prod), "0,00", Format(TBAbrir!Valor_PIS_Prod, "###,##0.0000"))
If TBAbrir!Soma_Cofins = True Then Chk_soma_Cofins.Value = 1 Else Chk_soma_Cofins.Value = 0
Txt_vlr_COFINS = IIf(IsNull(TBAbrir!Valor_Cofins_Prod), "0,00", Format(TBAbrir!Valor_Cofins_Prod, "###,##0.0000"))
If TBAbrir!Opt_valor_bc_ICMS = True Then Opt_vlr_BC_ICMS.Value = True Else Opt_vlr_BC_ICMS_fator.Value = True
Txt_vlr_BC_ICMS = IIf(IsNull(TBAbrir!ValorConta), "0,00", Format(TBAbrir!ValorConta, "###,##0.0000"))
Txt_vlr_BC_ICMS_fator = IIf(IsNull(TBAbrir!Valor_DAS), "", Format(TBAbrir!Valor_DAS, "###,##0.0000"))
Txt_vlr_ICMS = IIf(IsNull(TBAbrir!ValorICMS), "0,00", Format(TBAbrir!ValorICMS, "###,##0.0000"))
Txt_vlr_IOF = IIf(IsNull(TBAbrir!Valor_ICMS_SN), "0,00", Format(TBAbrir!Valor_ICMS_SN, "###,##0.0000"))

With Cmb_via_transporte
    Select Case TBAbrir!Via_transp
        Case "1": .Text = "1 - Marítima"
        Case "2": .Text = "2 - Fluvial"
        Case "3": .Text = "3 - Lacustre"
        Case "4": .Text = "4 - Aérea"
        Case "5": .Text = "5 - Postal"
        Case "6": .Text = "6 - Ferroviária"
        Case "7": .Text = "7 - Rodoviária"
        Case "8": .Text = "8 - Conduto - Rede Transmissão"
        Case "9": .Text = "9 - Meios Próprios"
        Case "10": .Text = "10 - Entrada ou Saída ficta"
    End Select
End With
Txt_valor_AFRMM = IIf(IsNull(TBAbrir!Valor_AFRMM), "", Format(TBAbrir!Valor_AFRMM, "###,##0.0000"))
With Cmb_forma_importacao
    Select Case TBAbrir!Forma_imp
        Case "1": .Text = "1 - Importação por conta própria"
        Case "2": .Text = "2 - Importação por conta e ordem"
        Case "3": .Text = "3 - Importação por encomenda"
    End Select
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcPuxaDadosDespesas()
On Error GoTo tratar_erro

txtData = IIf(IsNull(TBAbrir!Data_transacao), Format(Date, "dd/mm/yy"), Format(TBAbrir!Data_transacao, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), pubUsuario, TBAbrir!Responsavel)
Txt_ID_fornecedor = IIf(IsNull(TBAbrir!int_codforn), "", TBAbrir!int_codforn)
Txt_fornecedor = IIf(IsNull(TBAbrir!Txt_fornecedor), "", TBAbrir!Txt_fornecedor)
Cmb_data_vencimento = IIf(IsNull(TBAbrir!dt_Pagamento), "", TBAbrir!dt_Pagamento)
Txt_ID_PC = IIf(IsNull(TBAbrir!int_codfamilia), "", TBAbrir!int_codfamilia)
Txt_codigo_PC = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
Txt_descricao_PC = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
Txt_valor = IIf(IsNull(TBAbrir!valor), "", Format(TBAbrir!valor, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.txtId, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                
                ProcVerificaRegistroUtilizadoSemMsg "tbl_ContasPagar", "ID_Nota = " & frmFaturamento_Prod_Serv.txtId & " and Despesas_NF = 'True'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
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
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If frmFaturamento_Prod_Serv.txtNFiscal = "" Then Familiatext = "ordem de faturamento" Else Familiatext = "nota fiscal"
                If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.txtId, Familiatext, "dados de importação", "excluir estes", False, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            Mensagem = "Não é permitido excluir esses dados da DI, pois já exitem despesas cadastradas no módulo"
            ProcVerificaRegistroUtilizado "tbl_ContasPagar", "ID_Nota = " & frmFaturamento_Prod_Serv.txtId & " and Despesas_NF = 'True'", "Faturamento/Nota fiscal/Importação/Despesas"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
CamposFiltro = "Salvar_valores_produto, Data, Responsavel, Documento_importacao, Data_registro, Local_desembaraco, UF_desembaraco, Data_desembaraco, Codigo_exportador, Codigo_fabricante, NCM, Numero_adicao, Numero_sequencial, Opt_valor_bc_ICMS, Soma_aduaneiras, Soma_siscomex, Soma_PIS, Soma_Cofins, Soma_II, Sum(Valor_BC_importacao) as Valor, Sum(Valor_FOB) as Valor1, Sum(Valor_frete_int) as Valor2, Sum(Valor_seguro) as ValorPagar, Sum(Valor_imposto_importacao) as ValorPago, Sum(Valor_despesas) as ValorPagoParcial, Sum(Valor_sixcomex) as Valorparcela, Sum(Valor_bc_PIS_Cofins) as Valor_Total, Sum(Valor_PIS) as Valor_PIS_Prod, Sum(Valor_Cofins) as Valor_Cofins_Prod, Sum(Valor_bc_ICMS) as ValorConta, Sum(Valor_bc_ICMS_fator) as Valor_DAS, Sum(Valor_ICMS) as ValorICMS, Sum(Valor_imposto_OperacoesFinanceiras) as Valor_ICMS_SN, Sum(Valor_IPI) as Valor_IPI, Via_transp, Valor_AFRMM, Forma_imp, Recalcula_IPI"
CamposGrupo = " Salvar_valores_produto, Data, Responsavel, Documento_importacao, Data_registro, Local_desembaraco, UF_desembaraco, Data_desembaraco, Codigo_exportador, Codigo_fabricante,  NCM, Numero_adicao, Numero_sequencial, Opt_valor_bc_ICMS, Soma_aduaneiras, Soma_siscomex, Soma_PIS, Soma_Cofins, Soma_II, Via_transp, Valor_AFRMM, Forma_imp, Recalcula_IPI"
If Chk_salvar_valores_produto.Value = 1 Then
    CamposFiltro = CamposFiltro & ", ID_item"
    CamposGrupo = CamposGrupo & ", ID_item"
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select " & CamposFiltro & " from tbl_Detalhes_Nota_NFe where id_nota = " & frmFaturamento_Prod_Serv.txtId & " and NCM = '" & Lista.SelectedItem.ListSubItems(3) & "' and Numero_adicao = " & Lista.SelectedItem.ListSubItems(4) & " and Numero_sequencial = " & Lista.SelectedItem.ListSubItems(5) & " Group by " & CamposGrupo, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    Frame1.Enabled = True
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close
IDlista = Lista.SelectedItem
Cmb_ID_prod = Lista.SelectedItem.ListSubItems.Item(6)
ProcCalculaAcessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaDespesas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaDespesas
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.txtId, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                
                If .ListItems.Item(InitFor).ListSubItems(6) = "S" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaDespesas, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaDespesas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaDespesas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If frmFaturamento_Prod_Serv.txtNFiscal = "" Then Familiatext = "ordem de faturamento" Else Familiatext = "nota fiscal"
                If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & frmFaturamento_Prod_Serv.txtId, Familiatext, "despesa", "excluir esta", False, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            If .ListItems.Item(InitFor).ListSubItems(6) = "S" Then
                USMsgBox ("Não é permitido excluir esta despesa, pois a mesma já está baixada"), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaDespesas_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaDespesas.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select tbl_ContasPagar.Data_transacao, tbl_ContasPagar.responsavel, tbl_ContasPagar.int_codforn, tbl_ContasPagar.txt_Fornecedor, tbl_ContasPagar.dt_Pagamento, tbl_familia.int_codfamilia, tbl_familia.Codigo, tbl_familia.txt_descricao, Familia_financeiro.Valor from (tbl_ContasPagar INNER JOIN Familia_financeiro ON tbl_ContasPagar.IDintconta = Familia_financeiro.IDConta) INNER JOIN tbl_familia ON tbl_familia.int_codfamilia = Familia_financeiro.ID_PC where tbl_ContasPagar.IDintconta = " & ListaDespesas.SelectedItem & " and Familia_financeiro.TipoConta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposDespesas
    ProcPuxaDadosDespesas
    Frame5.Enabled = True
    CodigoLista1 = ListaDespesas.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If

Select Case SSTab1.Tab
    Case 0:
        If Lista.Visible = True Then Lista.SetFocus
        ProcCarregaLista
    Case 1:
        ListaDespesas.SetFocus
        ProcCarregaListaDespesas
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_ID_fornecedor_Change()
On Error GoTo tratar_erro

If Txt_ID_fornecedor <> "" Then
    VerifNumero = Txt_ID_fornecedor
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ID_fornecedor = ""
        Txt_ID_fornecedor.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_ID_fornecedor_LostFocus()
On Error GoTo tratar_erro

Txt_fornecedor = ""
If Txt_ID_fornecedor = "" Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & Txt_ID_fornecedor, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then Txt_fornecedor = TBFornecedor!Nome_Razao
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_aduaneiro_Change()
On Error GoTo tratar_erro

If Txt_vlr_aduaneiro.Text <> "" Then
    VerifNumero = Txt_vlr_aduaneiro.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_aduaneiro.Text = ""
        Txt_vlr_aduaneiro.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_aduaneiro_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_aduaneiro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_aduaneiro_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_aduaneiro = Format(Txt_vlr_aduaneiro, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_BC_ICMS_Change()
On Error GoTo tratar_erro

If Txt_vlr_BC_ICMS.Text <> "" Then
    VerifNumero = Txt_vlr_BC_ICMS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_BC_ICMS.Text = ""
        Txt_vlr_BC_ICMS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_BC_ICMS_fator_Change()
On Error GoTo tratar_erro

If Txt_vlr_BC_ICMS_fator.Text <> "" Then
    VerifNumero = Txt_vlr_BC_ICMS_fator.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_BC_ICMS_fator.Text = ""
        Txt_vlr_BC_ICMS_fator.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_BC_ICMS_fator_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_BC_ICMS_fator

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_BC_ICMS_fator_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_BC_ICMS_fator = Format(Txt_vlr_BC_ICMS_fator, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_BC_ICMS_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_BC_ICMS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_BC_ICMS_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_BC_ICMS = Format(Txt_vlr_BC_ICMS, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_bc_PIS_COFINS_Change()
On Error GoTo tratar_erro

If Txt_vlr_bc_PIS_COFINS.Text <> "" Then
    VerifNumero = Txt_vlr_bc_PIS_COFINS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_bc_PIS_COFINS.Text = ""
        Txt_vlr_bc_PIS_COFINS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_bc_PIS_COFINS_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_bc_PIS_COFINS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_bc_PIS_COFINS_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_bc_PIS_COFINS = Format(Txt_vlr_bc_PIS_COFINS, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_COFINS_Change()
On Error GoTo tratar_erro

If Txt_vlr_COFINS.Text <> "" Then
    VerifNumero = Txt_vlr_COFINS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_COFINS.Text = ""
        Txt_vlr_COFINS.SetFocus
        Exit Sub
    End If
End If

ProcCalculaAcessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_COFINS_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_COFINS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_COFINS_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_COFINS = Format(Txt_vlr_COFINS, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_desp_aduan_Change()
On Error GoTo tratar_erro

If Txt_vlr_desp_aduan.Text <> "" Then
    VerifNumero = Txt_vlr_desp_aduan.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_desp_aduan.Text = ""
        Txt_vlr_desp_aduan.SetFocus
        Exit Sub
    End If
End If

ProcCalculaAcessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_desp_aduan_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_desp_aduan

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_desp_aduan_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_desp_aduan = Format(Txt_vlr_desp_aduan, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_fob_Change()
On Error GoTo tratar_erro

If Txt_vlr_fob.Text <> "" Then
    VerifNumero = Txt_vlr_fob.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_fob.Text = ""
        Txt_vlr_fob.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_fob_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_fob

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_fob_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_fob = Format(Txt_vlr_fob, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_frete_Change()
On Error GoTo tratar_erro

If Txt_vlr_frete.Text <> "" Then
    VerifNumero = Txt_vlr_frete.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_frete.Text = ""
        Txt_vlr_frete.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_frete_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_frete

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_frete_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_frete = Format(Txt_vlr_frete, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_ICMS_Change()
On Error GoTo tratar_erro

If Txt_vlr_ICMS.Text <> "" Then
    VerifNumero = Txt_vlr_ICMS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_ICMS.Text = ""
        Txt_vlr_ICMS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_ICMS_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_ICMS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_ICMS_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_ICMS = Format(Txt_vlr_ICMS, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_II_Change()
On Error GoTo tratar_erro

If Txt_vlr_II.Text <> "" Then
    VerifNumero = Txt_vlr_II.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_II.Text = ""
        Txt_vlr_II.SetFocus
        Exit Sub
    End If
End If

ProcCalculaAcessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_II_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_II

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_II_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_II = Format(Txt_vlr_II, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_IOF_Change()
On Error GoTo tratar_erro

If Txt_vlr_IOF.Text <> "" Then
    VerifNumero = Txt_vlr_IOF.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_IOF.Text = ""
        Txt_vlr_IOF.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_IOF_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_IOF

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_IOF_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_IOF = Format(Txt_vlr_IOF, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_IPI_Change()
On Error GoTo tratar_erro

If Txt_vlr_IPI.Text <> "" Then
    VerifNumero = Txt_vlr_IPI.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_IPI.Text = ""
        Txt_vlr_IPI.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_IPI_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_IPI = Format(Txt_vlr_IPI, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub Txt_vlr_PIS_Change()
On Error GoTo tratar_erro

If Txt_vlr_PIS.Text <> "" Then
    VerifNumero = Txt_vlr_PIS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_PIS.Text = ""
        Txt_vlr_PIS.SetFocus
        Exit Sub
    End If
End If

ProcCalculaAcessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_PIS_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_PIS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_PIS_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_PIS = Format(Txt_vlr_PIS, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_seguro_Change()
On Error GoTo tratar_erro

If Txt_vlr_seguro.Text <> "" Then
    VerifNumero = Txt_vlr_seguro.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_seguro.Text = ""
        Txt_vlr_seguro.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_seguro_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_seguro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_seguro_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_seguro = Format(Txt_vlr_seguro, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_siscomex_Change()
On Error GoTo tratar_erro

If Txt_vlr_siscomex.Text <> "" Then
    VerifNumero = Txt_vlr_siscomex.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_vlr_siscomex.Text = ""
        Txt_vlr_siscomex.SetFocus
        Exit Sub
    End If
End If

ProcCalculaAcessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_siscomex_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_vlr_siscomex

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_vlr_siscomex_LostFocus()
On Error GoTo tratar_erro

Txt_vlr_siscomex = Format(Txt_vlr_siscomex, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_Valor_Change()
On Error GoTo tratar_erro

If Txt_valor.Text <> "" Then
    VerifNumero = Txt_valor.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor.Text = ""
        Txt_valor.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_valor_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_valor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_Valor_LostFocus()
On Error GoTo tratar_erro

Txt_valor = Format(Txt_valor, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub





Private Sub txtNumero_adicao_Change()
On Error GoTo tratar_erro

If txtNumero_adicao.Text <> "" Then
    VerifNumero = txtNumero_adicao.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_adicao.Text = ""
        txtNumero_adicao.SetFocus
        Exit Sub
    End If
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtNumero_sequencial_Change()
On Error GoTo tratar_erro

If txtNumero_sequencial.Text <> "" Then
    VerifNumero = txtNumero_sequencial.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_sequencial.Text = ""
        txtNumero_sequencial.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case ButtonIndex
            Case 1: ProcNovo
            Case 2: ProcSalvar
            Case 3: ProcExcluir
            'Case 5: ProcAjuda
            Case 6: ProcSair
        End Select
    Case 1:
        Select Case ButtonIndex
            Case 1: ProcNovoDespesas
            Case 2: ProcSalvarDespesas
            Case 3: ProcExcluirDespesas
            'Case 5: ProcAjuda
            Case 6: ProcSair
        End Select
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_DI = True Then
    If USMsgBox("A DI ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_DI = True Then Exit Sub Else GoTo Sair
    End If
End If
If Novo_DI1 = True Then
    If USMsgBox("A despesa ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarDespesas
        If Novo_DI1 = True Then Exit Sub Else GoTo Sair
    End If
End If

Sair:
    Novo_DI = False
    Novo_DI1 = False
    
    If Lista.ListItems.Count = 0 Then TextoFiltro = "Alterar = 'False'" Else TextoFiltro = "Alterar = 'True'"
    Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set " & TextoFiltro & " where ID = " & frmFaturamento_Prod_Serv.txtId
  '  frmFaturamento_Prod_Serv.ProcCarregaLista
    
    Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
valor = 0
Valorparcela = 0
ValorTotalPago = 0
ValorPendente = 0
Valor2 = 0
Valor_Produto = 0
Valor_PIS_Prod = 0
Valor_Cofins_Prod = 0
ValorParcial = 0
ValorTotalPagar = 0
ValorNC = 0
ValorIPI = 0

'With Chk_salvar_valores_produto
'    .Enabled = True
'    .Value = 0
'End With

Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select id,ID_item, Salvar_valores_produto, Documento_importacao, Data_registro, NCM, Numero_adicao, Numero_sequencial, Opt_valor_bc_ICMS, Soma_II, Soma_aduaneiras, Soma_siscomex, Soma_PIS, Soma_Cofins, Valor_frete_int as Valor, Valor_seguro as Valorparcela, Valor_bc_ICMS as ValorTotalPago, Valor_bc_ICMS_fator as ValorPendente, Valor_ICMS as Valor2, Valor_IPI as ValorIPI, Valor_imposto_importacao as Valor_Produto, Valor_PIS as Valor_PIS_Prod, Valor_Cofins as Valor_Cofins_Prod, Valor_despesas as ValorParcial, Valor_sixcomex as ValorTotalPagar from tbl_Detalhes_Nota_NFe where id_nota = " & frmFaturamento_Prod_Serv.txtId & " order by Numero_adicao, Numero_sequencial"

'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Documento_importacao), "", TBLISTA!Documento_importacao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data_registro), "", Format(TBLISTA!Data_registro, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!NCM), "", TBLISTA!NCM)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Numero_adicao), "", TBLISTA!Numero_adicao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Numero_sequencial), "", TBLISTA!Numero_sequencial)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Id_Item), "", TBLISTA!Id_Item)
            
            valor = Format(valor + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor), "###,##0.00")
            Valorparcela = Format(Valorparcela + IIf(IsNull(TBLISTA!Valorparcela), 0, TBLISTA!Valorparcela), "###,##0.00")
            If TBLISTA!Opt_valor_bc_ICMS = True Then
                ValorTotalPago = Format(ValorTotalPago + IIf(IsNull(TBLISTA!ValorTotalPago), 0, TBLISTA!ValorTotalPago), "###,##0.00")
            Else
                ValorTotalPago = Format(ValorTotalPago + IIf(IsNull(TBLISTA!ValorPendente), 0, TBLISTA!ValorPendente), "###,##0.00")
            End If
            Valor2 = Format(Valor2 + IIf(IsNull(TBLISTA!Valor2), 0, TBLISTA!Valor2), "###,##0.00")
            ValorIPI = Format(ValorIPI + IIf(IsNull(TBLISTA!ValorIPI), 0, TBLISTA!ValorIPI), "###,##0.00")
            Valor_Produto = Format(Valor_Produto + IIf(IsNull(TBLISTA!Valor_Produto), 0, TBLISTA!Valor_Produto), "###,##0.00")
            Valor_PIS_Prod = Format(Valor_PIS_Prod + IIf(IsNull(TBLISTA!Valor_PIS_Prod), 0, TBLISTA!Valor_PIS_Prod), "###,##0.00")
            Valor_Cofins_Prod = Format(Valor_Cofins_Prod + IIf(IsNull(TBLISTA!Valor_Cofins_Prod), 0, TBLISTA!Valor_Cofins_Prod), "###,##0.00")
            ValorTotalPagar = Format(ValorTotalPagar + IIf(IsNull(TBLISTA!ValorTotalPagar), 0, TBLISTA!ValorTotalPagar), "###,##0.00")
            
            If TBLISTA!Soma_II = True Then ValorNC = Format(ValorNC + IIf(IsNull(TBLISTA!Valor_Produto), 0, TBLISTA!Valor_Produto), "###,##0.00")
            If TBLISTA!Soma_aduaneiras = True Then ValorNC = Format(ValorNC + IIf(IsNull(TBLISTA!ValorParcial), 0, TBLISTA!ValorParcial), "###,##0.00")
            If TBLISTA!Soma_siscomex = True Then ValorNC = Format(ValorNC + IIf(IsNull(TBLISTA!ValorTotalPagar), 0, TBLISTA!ValorTotalPagar), "###,##0.00")
            If TBLISTA!Soma_PIS = True Then ValorNC = Format(ValorNC + IIf(IsNull(TBLISTA!Valor_PIS_Prod), 0, TBLISTA!Valor_PIS_Prod), "###,##0.00")
            If TBLISTA!Soma_Cofins = True Then ValorNC = Format(ValorNC + IIf(IsNull(TBLISTA!Valor_Cofins_Prod), 0, TBLISTA!Valor_Cofins_Prod), "###,##0.00")
            
'            With Chk_salvar_valores_produto
'                If TBLISTA!Salvar_valores_produto = True Then .Value = 1 Else .Value = 0
'                .Enabled = False
'            End With
            
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    
    NovoValor = Replace(valor, ",", ".") 'Frete
    NovoValor1 = Replace(Valorparcela, ",", ".") 'Seguro
    NovoValor2 = Replace(ValorNC, ",", ".") 'Outras despesas
    NovoValor3 = Replace(ValorTotalPago, ",", ".") 'Base ICMS
    NovoValor4 = Replace(Valor2, ",", ".") 'ICMS
    NovoValor5 = Replace(Valor_Produto, ",", ".") 'Imposto Importação
    NovoValor6 = Replace(Valor_PIS_Prod, ",", ".") 'PIS
    NovoValor7 = Replace(Valor_Cofins_Prod, ",", ".") 'Cofins
    NovoValor9 = Replace(ValorIPI, ",", ".") 'IPI
   
   With frmFaturamento_Prod_Serv
        .txt_VlrFrete = Format(valor, "###,##0.00")
        .txt_vlrSeguro = Format(Valorparcela, "###,##0.00")
        .txt_OutrasDisp = Format(ValorNC, "###,##0.00")
        .txt_BaseICMS = Format(ValorTotalPago, "###,##0.00")
        .txt_vlrICMS = Format(Valor2, "###,##0.00")
        .txt_TotalIPI = Format(ValorIPI, "###,##0.00")
        Valor1 = .txt_ValorNota
        NovoValor8 = Replace(Valor1, ",", ".") 'Total da NF
        Conexao.Execute "Update tbl_Totais_Nota Set dbl_Valor_Frete = " & NovoValor & ", dbl_Valor_Seguro = " & NovoValor1 & ", dbl_Desp_Adicionais = " & NovoValor2 & ", dbl_Base_ICMS = " & NovoValor3 & ", dbl_Valor_ICMS = " & NovoValor4 & ", Valor_total_II = " & NovoValor5 & ", Total_PIS_prod = " & NovoValor6 & ", Total_Cofins_prod = " & NovoValor7 & ", dbl_Valor_Total_Nota = " & NovoValor8 & ", dbl_Valor_Total_IPI = " & NovoValor9 & " where ID_nota = " & .txtId
    End With
End If
TBLISTA.Close

If Chk_salvar_valores_produto.Value = 0 Then
    ProcCorrigeValoresProd valor, "NFP.Valor_frete"
    ProcCorrigeValoresProd Valorparcela, "NFP.Valor_seguro"
    ProcCorrigeValoresProd ValorNC, "NFP.Valor_acessorias"
    ProcCorrigeValoresProd ValorTotalPago, "CSTICMS.Valor_BC"
    ProcCorrigeValoresProd Valor2, "CSTICMS.Valor_ICMS"
    ProcCorrigeValoresProd ValorIPI, "NFP.dbl_valoripi"
    ProcCorrigeValoresProd Valor_PIS_Prod, "NFP.Total_PIS_prod"
    ProcCorrigeValoresProd Valor_Cofins_Prod, "NFP.Total_Cofins_prod"
End If

ProcSomaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCorrigeValoresProd(ValorTotal As Double, NomeCampo As String)
On Error GoTo tratar_erro

Contador = 1
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select " & NomeCampo & " as Valor2 from tbl_Detalhes_Nota NFP INNER JOIN tbl_Detalhes_Nota_CST_ICMS CSTICMS ON CSTICMS.ID_item = NFP.Int_codigo where NFP.id_nota = " & frmFaturamento_Prod_Serv.txtId & " and " & IIf(Left(NomeCampo, 3) = "CST", "NFP.int_ICMS", NomeCampo) & " > 0 and NFP.Remessa = 'False' and NFP.Retorno = 'False' order by NFP.int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Do While TBProduto.EOF = False
        If ValorTotal < 0 Then
            If (ValorTotal * -1) <= TBProduto!Valor2 Then
                TBProduto!Valor2 = TBProduto!Valor2 - (ValorTotal * -1)
                ValorTotal = 0
            Else
                ValorTotal = ValorTotal + TBProduto!Valor2
                TBProduto!Valor2 = 0
            End If
            TBProduto.Update
            GoTo Proximo
        ElseIf ValorTotal = 0 Then
                TBProduto!Valor2 = 0
                TBProduto.Update
            Else
                If Contador = TBProduto.RecordCount Then
                    TBProduto!Valor2 = Format(ValorTotal, "###,##0.00")
                    TBProduto.Update
                End If
        End If
        ValorTotal = ValorTotal - TBProduto!Valor2
Proximo:
        Contador = Contador + 1
        TBProduto.MoveNext
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCalculaValoresProdutos(Frete1 As Double, Seguro1 As Double, Acessorias1 As Double, ValorIPI1 As Double, Valor_PIS_Prod1 As Double, Valor_Cofins_Prod1 As Double, Valor_BCICMS1 As Double, Valor_ICMS1 As Double)
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    TextoFiltro = ""
    If Chk_salvar_valores_produto.Value = 1 Then
        ValorNC = Txt_vlr_fob
        NovoValor = Replace(ValorNC, ",", ".")
        'TextoFiltro = " and NFP.dbl_ValorTotal = " & NovoValor
        TextoFiltro = " and NFP.Int_codigo = " & Cmb_ID_prod
    End If
    
    'Valor Total de produtos
    ValorTotal = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(NFP.dbl_ValorTotal) as Valortotal from tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.ID_nota = " & .txtId & " and CF.IDIntClasse = '" & txt_NCM & "' and Retorno = 'False' and Remessa = 'False'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ValorTotal = IIf(IsNull(TBAbrir!ValorTotal), 0, TBAbrir!ValorTotal)
    End If
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select NFP.Int_codigo, NFP.dbl_ValorTotal, NFP.Valor_frete, NFP.Valor_seguro, NFP.Valor_acessorias, NFP.int_IPI, NFP.dbl_valoripi, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, CF.IDIntClasse from tbl_Detalhes_Nota NFP INNER JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.id_nota = " & .txtId.Text & " and CF.IDIntClasse = '" & txt_NCM & "' and NFP.Remessa = 'False' And NFP.Retorno = 'False' " & TextoFiltro & " order by NFP.int_codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Do While TBProduto.EOF = False
            Frete = 0
            Seguro = 0
            Acessorias = 0
            ValorIPI = 0
            Valor_PIS_Prod = 0
            Valor_Cofins_Prod = 0
            BC = 0
            ValorICMS = 0
        
            'IPI
            If Chk_salvar_valores_produto.Value = 1 Then
                Frete = Frete1
                Seguro = Seguro1
                Acessorias = Acessorias1
                TBProduto!dbl_valoripi = Format(ValorIPI1, "###,##0.00")
                Valor_PIS_Prod = Valor_PIS_Prod1
                Valor_Cofins_Prod = Valor_Cofins_Prod1
                BC = Valor_BCICMS1
                ValorICMS = ValorICMS1
            Else
                'Verifica valores para somar na base de calculo
                VltUnit = IIf(IsNull(TBProduto!dbl_ValorTotal), 0, TBProduto!dbl_ValorTotal)
                If ValorTotal <> 0 Then Qtd = (VltUnit * 100) / ValorTotal
                    
                'Frete
                If Frete1 <> 0 Then Frete = Format((Frete1 * Qtd) / 100, "###,##0.00")
                'Seguro
                If Seguro1 <> 0 Then Seguro = Format((Seguro1 * Qtd) / 100, "###,##0.00")
                'Acessorias
                If Acessorias1 <> 0 Then Acessorias = Format((Acessorias1 * Qtd) / 100, "###,##0.00")
                
                If TBProduto!int_IPI <> 0 Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select ID from tbl_Detalhes_Nota_NFe where ID_item = " & TBProduto!Int_codigo & " and Recalcula_IPI = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        If ValorIPI1 <> 0 Then ValorIPI = Format((ValorIPI1 * Qtd) / 100, "###,##0.00")
                        TBProduto!dbl_valoripi = Format(ValorIPI, "###,##0.00")
                    Else
                        TBProduto!dbl_valoripi = Format((VltUnit * TBProduto!int_IPI) / 100, "###,##0.00")
                    End If
                    TBAbrir.Close
                End If
                
                'PIS
                If Valor_PIS_Prod1 <> 0 Then Valor_PIS_Prod = Format((Valor_PIS_Prod1 * Qtd) / 100, "###,##0.00")
                'Cofins
                If Valor_Cofins_Prod1 <> 0 Then Valor_Cofins_Prod = Format((Valor_Cofins_Prod1 * Qtd) / 100, "###,##0.00")
                'BC ICMS
                If Valor_BCICMS1 <> 0 Then BC = Format((Valor_BCICMS1 * Qtd) / 100, "###,##0.00")
                'ICMS
                If Valor_ICMS1 <> 0 Then ValorICMS = Format((Valor_ICMS1 * Qtd) / 100, "###,##0.00")
            End If
            
            TBProduto!Valor_frete = Format(Frete, "###,##0.00")
            TBProduto!Valor_seguro = Format(Seguro, "###,##0.00")
            TBProduto!Valor_acessorias = Format(Acessorias, "###,##0.00")
            TBProduto!Total_PIS_prod = Format(Valor_PIS_Prod, "###,##0.00")
            TBProduto!Total_Cofins_prod = Format(Valor_Cofins_Prod, "###,##0.00")

            TBProduto.Update
            
            NovoValor = Replace(BC, ",", ".")
            NovoValor1 = Replace(ValorICMS, ",", ".")
            'Conexao.Execute "Update tbl_Detalhes_Nota_CST_ICMS Set Valor_BC = " & NovoValor & ", Valor_ICMS = " & NovoValor1 & " where ID_item = " & TBProduto!Int_codigo
            
            TBProduto.MoveNext
        Loop
    End If
    TBProduto.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaListaDespesas()
On Error GoTo tratar_erro

ListaDespesas.ListItems.Clear
Qtde = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select tbl_ContasPagar.IDintconta, tbl_ContasPagar.txt_Fornecedor, tbl_ContasPagar.dt_Pagamento, tbl_ContasPagar.LogSit, tbl_familia.Codigo, tbl_familia.txt_descricao, Familia_financeiro.Valor from (tbl_ContasPagar INNER JOIN Familia_financeiro ON tbl_ContasPagar.IDintconta = Familia_financeiro.IDConta) INNER JOIN tbl_familia ON tbl_familia.int_codfamilia = Familia_financeiro.ID_PC where tbl_ContasPagar.ID_nota = " & frmFaturamento_Prod_Serv.txtId & " and tbl_ContasPagar.Despesas_NF = 'True' order by tbl_ContasPagar.IDintconta", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaDespesas.ListItems
            .Add , , TBLISTA!IDintconta
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Txt_fornecedor), "", TBLISTA!Txt_fornecedor)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!dt_Pagamento), "", Format(TBLISTA!dt_Pagamento, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Logsit), "", TBLISTA!Logsit)
            Qtde = Qtde + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close
txtValorTotal = Format(Qtde, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) DI('s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set Alterar = 'False' where ID = " & frmFaturamento_Prod_Serv.txtId
            End If
            Permitido = True
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Detalhes_Nota_NFe where ID_nota = " & frmFaturamento_Prod_Serv.txtId & " and NCM = '" & .ListItems(InitFor).ListSubItems(3) & "' and Numero_adicao = " & .ListItems(InitFor).ListSubItems(4) & " and Numero_sequencial = " & .ListItems(InitFor).ListSubItems(5), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    '==================================
                    Modulo = Formulario & "/Importação"
                    Evento = "Excluir DI"
                    ID_documento = TBAbrir!Id_Item
                    With frmFaturamento_Prod_Serv
                        .ProcVerificaTipoNF False
                        If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
                        Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
                        Documento1 = "DI: " & TBAbrir!Documento_importacao & " - NCM: " & TBAbrir!NCM & " - Adição: " & TBAbrir!Numero_adicao & " - Sequencial: " & TBAbrir!Numero_sequencial
                    End With
                    ProcGravaEvento
                    '==================================
                    Conexao.Execute "DELETE from tbl_Detalhes_Nota_NFe where ID = " & TBAbrir!ID
                    
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) DI('s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("DI('s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
    
    frmFaturamento_Prod_Serv.ProcCarregaLista
    
    txtDocumento_importacao = ""
    Cmb_data_registro.Value = Date
    txtLocal_desembaraco = ""
    cmbUF_desembaraco.ListIndex = -1
    Cmb_data_desembaraco.Value = Date
    txtCodigo_exportador = ""
    txtCodigo_fabricante = ""
    Cmb_via_transporte.ListIndex = -1
    Txt_valor_AFRMM = ""
    Cmb_forma_importacao.ListIndex = -1
    ProcLimpaCampos
    
    Frame1.Enabled = False
    Novo_DI = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcExcluirDespesas()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaDespesas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) despesa(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select tbl_ContasPagar.IDIntconta, tbl_ContasPagar.IDFluxo, Familia_financeiro.ID from (tbl_ContasPagar INNER JOIN Familia_financeiro ON tbl_ContasPagar.IDintconta = Familia_financeiro.IDConta) INNER JOIN tbl_familia ON tbl_familia.int_codfamilia = Familia_financeiro.ID_PC where tbl_ContasPagar.IDintconta = " & .ListItems(InitFor) & " and tbl_familia.Codigo = '" & .ListItems(InitFor).ListSubItems(3) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Conexao.Execute "DELETE from Familia_financeiro where ID = " & TBFI!ID
                
                '==================================
                Modulo = Formulario & "/Importação"
                Evento = "Excluir despesa"
                ID_documento = .ListItems(InitFor)
                frmFaturamento_Prod_Serv.ProcVerificaTipoNF False
                If frmFaturamento_Prod_Serv.txtNFiscal = "" Then NomeCampo = "N° ordem: " & frmFaturamento_Prod_Serv.txtId Else NomeCampo = "N° nota: " & frmFaturamento_Prod_Serv.txtNFiscal
                Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & frmFaturamento_Prod_Serv.txtSerie
                Documento1 = "Fornecedor: " & .ListItems(InitFor).ListSubItems(1) & " - Dt. de vencimento: " & Format(.ListItems(InitFor).ListSubItems(2), "dd/mm/yy") & " - Código contábil: " & .ListItems(InitFor).ListSubItems(3) & " - Descrição: " & .ListItems(InitFor).ListSubItems(4)
                ProcGravaEvento
                '==================================
                
                'Verifica valor total da conta, de acordo com as contas contábeis
                valor = 0
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(Valor) as Valor from familia_financeiro where IDconta = " & TBFI!IDintconta & " and TipoConta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                End If
                TBAbrir.Close
                NovoValor = Replace(valor, ",", ".")
                
                If valor <= 0 Then
                    Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBFI!IDFluxo
                    Conexao.Execute "DELETE from tbl_contaspagar where IdIntConta = " & TBFI!IDintconta
                Else
                    Conexao.Execute "UPDATE tbl_contaspagar set dbl_valorpagto = " & NovoValor & " where IdIntConta = " & TBFI!IDintconta
                    Conexao.Execute "UPDATE tbl_Fluxo_de_caixa set Valor = " & NovoValor & " where IDFluxo = " & TBFI!IDFluxo
                End If
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) despesa(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Despesa(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    
    Txt_ID_fornecedor = ""
    Txt_fornecedor = ""
    
    ProcCarregaListaDespesas
    ProcLimpaCamposDespesas
    ProcAtualizarEstoque
    Frame5.Enabled = False
    Novo_DI1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizarEstoque()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
    'Valor Total de produtos
    VlttTotal = 0
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "select SUM(dbl_ValorTotal) as ValorNF from tbl_Detalhes_Nota where ID_Nota = " & .txtId & " and Retorno = 'False' and Remessa = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        VlttTotal = IIf(IsNull(TBProduto!ValorNF), 0, TBProduto!ValorNF)
    End If
    
    VltUnit = 0
    Qtd = 0
    valor = 0
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select tbl_Detalhes_Nota.*, tbl_Detalhes_Nota_NFe.Valor_ICMS, projproduto.Estoque from (tbl_Detalhes_Nota INNER JOIN projproduto ON tbl_Detalhes_Nota.int_Cod_Produto = Projproduto.Desenho) INNER JOIN tbl_Detalhes_Nota_NFe ON tbl_Detalhes_Nota_NFe.ID_item = tbl_Detalhes_Nota.Int_codigo where tbl_Detalhes_Nota.ID_Nota = " & frmFaturamento_Prod_Serv.txtId & " and tbl_Detalhes_Nota.Retorno = 'False' and tbl_Detalhes_Nota.Remessa = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Do While TBProduto.EOF = False
            'Verifica o percentual do produto sobre o valor total
            Qtde = IIf(IsNull(TBProduto!dbl_ValorTotal), 0, TBProduto!dbl_ValorTotal)
            If VlttTotal <> 0 Then VltUnit = (Qtde * 100) / VlttTotal
            
            Valor1 = Format(IIf(IsNull(TBProduto!Valor_ICMS), 0, TBProduto!Valor_ICMS) / IIf(IsNull(TBProduto!int_Qtd), 0, TBProduto!int_Qtd), "###,##0.00")
            
            valor = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(dbl_valorpagto) as Valor from tbl_ContasPagar where ID_nota = " & .txtId & " and Despesas_NF = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            TBAbrir.Close
            
            'Verifica o valor total das despesas de acordo com o pecentual do produto
            vlrTotalProd = 0
            If valor <> 0 Then vlrTotalProd = (valor * VltUnit) / 100
            
            'Atualiza valor de compra do produto
            Valor2 = Format((IIf(IsNull(TBProduto!dbl_ValorUnitario), "0", TBProduto!dbl_ValorUnitario) + vlrTotalProd) - Valor1, "###,##0.0000000000")
            NovoValor = Replace(Valor2, ",", ".")
            Conexao.Execute "Update projproduto Set PCusto = " & NovoValor & " where Desenho = '" & TBProduto!int_Cod_Produto & "'"
            
            'Atualiza valor do estoque
            If TBProduto!Estoque = True Then
                Qtd = 0
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from estoque_movimentacao where IDLista_recebimento = " & IIf(IsNull(TBProduto!CODIGO), 0, TBProduto!CODIGO) & " and Desenho = '" & IIf(IsNull(TBProduto!int_Cod_Produto), 0, TBProduto!int_Cod_Produto) & "' and Entrada is not null and Entrada <> 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    Do While TBFIltro.EOF = False
                        
                        Qtd = vlrTotalProd / TBFIltro!Entrada
                        TBFIltro!VlrUnit = Format((IIf(IsNull(TBProduto!dbl_ValorUnitario), "0", TBProduto!dbl_ValorUnitario) + Qtd) - Valor1, "###,##0.0000000000")
                        TBFIltro!vlrTotal = Format(TBFIltro!VlrUnit * TBFIltro!Entrada, "###,##0.00")
                        
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from estoque_controle where IdEstoque = " & TBFIltro!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            TBFI!valor_unitario = TBFIltro!VlrUnit
                            TBFI!Valor_total = Format(TBFI!valor_unitario * TBFI!estoque_real, "###,##0.00")
                            TBFI.Update
                            TBFI.MoveNext
                        End If
                        TBFI.Close
                        
                        TBFIltro.Update
                        TBFIltro.MoveNext
                    Loop
                End If
                TBFIltro.Close
            End If
            
            TBProduto.MoveNext
        Loop
    End If
    TBProduto.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Txt_valor_AFRMM_Change()
On Error GoTo tratar_erro

If Txt_valor_AFRMM <> "" Then
    VerifNumero = Txt_valor_AFRMM
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_AFRMM = ""
        Txt_valor_AFRMM.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procSalvaValores()
On Error GoTo tratar_erro

If Left(frmFaturamento_Prod_Serv.cmbFinalidade_emissao, 1) = 2 Then
    valor = Txt_vlr_aduaneiro
    TBGravar!Valor_BC_importacao = valor '/ Qtde
    
    valor = Txt_vlr_fob
    TBGravar!Valor_FOB = valor ' / Qtde
    
    valor = Txt_vlr_frete
    TBGravar!Valor_frete_int = valor '/ Qtde
    
    valor = Txt_vlr_seguro
    TBGravar!Valor_seguro = valor ' / Qtde
    
    valor = Txt_vlr_II
    TBGravar!Valor_imposto_importacao = valor ' / Qtde
    
    valor = Txt_vlr_desp_aduan
    TBGravar!Valor_despesas = valor ' / Qtde
    
    valor = Txt_vlr_siscomex
    TBGravar!Valor_sixcomex = valor ' / Qtde
    
    valor = Txt_vlr_bc_PIS_COFINS
    TBGravar!Valor_bc_PIS_Cofins = valor ' / Qtde
    
    valor = Txt_vlr_PIS
    TBGravar!valor_pis = valor ' / Qtde
    
    Valor1 = Txt_vlr_COFINS
    TBGravar!valor_cofins = valor ' / Qtde
    
    valor = Txt_vlr_BC_ICMS
    TBGravar!Valor_bc_ICMS = valor ' / Qtde
    
    valor = Txt_vlr_BC_ICMS_fator
    TBGravar!Valor_bc_ICMS_fator = valor ' / Qtde
    
    valor = Txt_vlr_ICMS
    TBGravar!Valor_ICMS = valor ' / Qtde
    'TBGravar!Valor_ICMS = (TBGravar!Valor_bc_ICMS '* TBAbrir!int_ICMS) / 100
    
    valor = Txt_vlr_IOF
    TBGravar!Valor_imposto_OperacoesFinanceiras = valor '/ Qtde
    
    valor = Txt_vlr_IPI
    TBGravar!Valor_IPI = valor '/ Qtde
Else
    valor = Txt_vlr_aduaneiro
    TBGravar!Valor_BC_importacao = valor ''* 100
    
    valor = Txt_vlr_fob
    TBGravar!Valor_FOB = valor ''* 100
    
    valor = Txt_vlr_frete
    TBGravar!Valor_frete_int = valor ''* 100
    
    valor = Txt_vlr_seguro
    TBGravar!Valor_seguro = valor ''* 100
    
    valor = Txt_vlr_II
    TBGravar!Valor_imposto_importacao = valor '* 100
    
    valor = Txt_vlr_desp_aduan
    TBGravar!Valor_despesas = valor '* 100
    
    valor = Txt_vlr_siscomex
    TBGravar!Valor_sixcomex = valor '* 100
    
    valor = Txt_vlr_bc_PIS_COFINS
    TBGravar!Valor_bc_PIS_Cofins = valor '* 100
    
    valor = Txt_vlr_PIS
    TBGravar!valor_pis = valor '* 100
    
    Valor1 = Txt_vlr_COFINS
    TBGravar!valor_cofins = Valor1 '* 100
    
    valor = Txt_vlr_BC_ICMS
    TBGravar!Valor_bc_ICMS = valor '* 100
    
    valor = Txt_vlr_BC_ICMS_fator
    TBGravar!Valor_bc_ICMS_fator = valor '* 100
    
   Set TBAliquota = CreateObject("adodb.recordset")
   TBAliquota.Open "Select * from tbl_Detalhes_nota where Int_codigo = " & Cmb_ID_prod, Conexao, adOpenKeyset, adLockOptimistic
   If TBAliquota.EOF = False Then
    TBGravar!Valor_ICMS = (TBGravar!Valor_bc_ICMS * TBAliquota!int_ICMS) / 100
    TBAliquota!Total_PIS_prod = Txt_vlr_PIS.Text
    TBAliquota!Total_Cofins_prod = Txt_vlr_COFINS
    TBAliquota!Valor_frete = Txt_vlr_frete
    TBAliquota!Valor_acessorias = txtAcessorias.Text
    
    TBAliquota.Update
   End If
   TBAliquota.Close
   
    valor = Txt_vlr_IOF
    TBGravar!Valor_imposto_OperacoesFinanceiras = valor '* 100
    
    valor = Txt_vlr_IPI
    TBGravar!Valor_IPI = valor '* 100
'    TBGravar.Update

End If

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & Cmb_ID_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
TBAliquota!Valor_ICMS = Txt_vlr_ICMS.Text 'TBGravar!Valor_ICMS
TBAliquota!Valor_BC = Txt_vlr_BC_ICMS.Text 'TBGravar!Valor_bc_ICMS
TBAliquota.Update
End If

TBAliquota.Close

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_PIS where ID_item = " & Cmb_ID_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
'TBAliquota!Valor_PIS = Txt_vlr_ICMS.Text 'TBGravar!Valor_ICMS
TBAliquota!Valor_BC = Txt_vlr_bc_PIS_COFINS.Text 'TBGravar!Valor_bc_ICMS
TBAliquota.Update
End If

TBAliquota.Close

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from tbl_Detalhes_Nota_CST_COFINS where ID_item = " & Cmb_ID_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
'TBAliquota!Valor_ICMS = Txt_vlr_ICMS.Text 'TBGravar!Valor_ICMS
TBAliquota!Valor_BC = Txt_vlr_bc_PIS_COFINS.Text 'TBGravar!Valor_bc_ICMS
TBAliquota.Update
End If

'TBAliquota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
