VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmestoque_item_relat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque | Relatórios"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6390
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
   ScaleHeight     =   4965
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   38
      Top             =   4560
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   767
      DibPicture      =   "frmestoque_item_relat.frx":0000
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
      Icon            =   "frmestoque_item_relat.frx":1C95
   End
   Begin VB.ComboBox Cmb_empresa 
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
      ItemData        =   "frmestoque_item_relat.frx":1FAF
      Left            =   1380
      List            =   "frmestoque_item_relat.frx":1FB1
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   600
      Width           =   4725
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   210
      TabIndex        =   20
      Top             =   2055
      Width           =   5895
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   180
         TabIndex        =   36
         Top             =   210
         Width           =   2925
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
            Height          =   240
            Left            =   2100
            TabIndex        =   50
            Top             =   240
            Width           =   705
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
            Height          =   240
            Left            =   1500
            TabIndex        =   49
            Top             =   240
            Width           =   555
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
            Height          =   240
            Left            =   810
            TabIndex        =   48
            Top             =   240
            Width           =   645
         End
         Begin VB.OptionButton OptInicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Inicio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
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
         Height          =   345
         Left            =   180
         MouseIcon       =   "frmestoque_item_relat.frx":1FB3
         TabIndex        =   18
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1020
         Width           =   5505
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
         ItemData        =   "frmestoque_item_relat.frx":22BD
         Left            =   3390
         List            =   "frmestoque_item_relat.frx":22D0
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2325
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
         ItemData        =   "frmestoque_item_relat.frx":2317
         Left            =   180
         List            =   "frmestoque_item_relat.frx":2319
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Familia."
         Top             =   1020
         Width           =   5505
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
         Left            =   4200
         TabIndex        =   22
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   1305
         TabIndex        =   21
         Top             =   810
         Width           =   3240
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para busca do relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   210
      TabIndex        =   23
      Top             =   1050
      Width           =   3225
      Begin DrawSuite2022.USOptionButton optSaldo 
         Height          =   285
         Left            =   1620
         TabIndex        =   39
         Top             =   600
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         Caption         =   "Saldo (detalhado)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin VB.CheckBox Chk_igual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sem saldo em estoque"
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
         Left            =   3780
         TabIndex        =   2
         Top             =   1020
         Width           =   1965
      End
      Begin VB.CheckBox Chk_maior 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Com saldo em estoque"
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
         Left            =   3780
         TabIndex        =   1
         Top             =   780
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin DrawSuite2022.USOptionButton optSaldo_resumido 
         Height          =   285
         Left            =   1620
         TabIndex        =   40
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         Caption         =   "Saldo (resumido)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton Opt_saldo_diario 
         Height          =   285
         Left            =   210
         TabIndex        =   41
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         Caption         =   "Saldo (diário)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton optMovimentacao 
         Height          =   285
         Left            =   210
         TabIndex        =   42
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Caption         =   "Movimentação"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton Opt_semi_acabado_SPED 
         Height          =   285
         Left            =   1890
         TabIndex        =   43
         Top             =   1530
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   503
         Caption         =   "Semi-acabado (SPED)"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton Opt_semi_acabado 
         Height          =   285
         Left            =   1890
         TabIndex        =   44
         Top             =   1290
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         Caption         =   "Semi-acabado"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USOptionButton Opt_obsolescencia 
         Height          =   285
         Left            =   1890
         TabIndex        =   45
         Top             =   1770
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         Caption         =   "Obsolescência"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
      End
   End
   Begin VB.Frame frameMovimentacao 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5040
      TabIndex        =   27
      Top             =   2070
      Width           =   1035
      Begin VB.CheckBox chkEntrada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrada"
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
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   810
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.CheckBox chkSaida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saida"
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
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Value           =   1  'Checked
         Width           =   705
      End
   End
   Begin DrawSuite2022.USButton btnImprimir 
      Height          =   825
      Left            =   210
      TabIndex        =   46
      Top             =   3570
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1455
      DibPicture      =   "frmestoque_item_relat.frx":231B
      Caption         =   "Visualizar impressão"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      PicAlign        =   7
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin VB.Frame FramePeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   3450
      TabIndex        =   28
      Top             =   1050
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ComboBox cmbPor 
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
         ItemData        =   "frmestoque_item_relat.frx":3FB0
         Left            =   630
         List            =   "frmestoque_item_relat.frx":3FB7
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Por."
         Top             =   1050
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         ToolTipText     =   "Data final."
         Top             =   570
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
         Format          =   197394433
         CurrentDate     =   39799
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         ToolTipText     =   "Data inicio."
         Top             =   210
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
         Format          =   197394433
         CurrentDate     =   39799
      End
      Begin VB.ComboBox Cmb_ano_de1 
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
         ItemData        =   "frmestoque_item_relat.frx":3FC0
         Left            =   630
         List            =   "frmestoque_item_relat.frx":3FC2
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Ano de."
         Top             =   1080
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_mes_de 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmestoque_item_relat.frx":3FC4
         Left            =   90
         List            =   "frmestoque_item_relat.frx":3FEC
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Mês de."
         Top             =   1140
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_de 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmestoque_item_relat.frx":402D
         Left            =   2250
         List            =   "frmestoque_item_relat.frx":402F
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Ano de."
         Top             =   1110
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox Cmb_mes_ate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmestoque_item_relat.frx":4031
         Left            =   0
         List            =   "frmestoque_item_relat.frx":4059
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Mês até."
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_ate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmestoque_item_relat.frx":409A
         Left            =   2370
         List            =   "frmestoque_item_relat.frx":409C
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Ano até."
         Top             =   1650
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox Cmb_ano_ate1 
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
         ItemData        =   "frmestoque_item_relat.frx":409E
         Left            =   630
         List            =   "frmestoque_item_relat.frx":40A0
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Ano até."
         Top             =   1590
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label9 
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
         Left            =   780
         TabIndex        =   51
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
         Left            =   720
         TabIndex        =   31
         Top             =   600
         Width           =   360
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
         Left            =   -300
         TabIndex        =   30
         Top             =   630
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
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
         Left            =   195
         TabIndex        =   29
         Top             =   1125
         Width           =   345
      End
   End
   Begin VB.Frame frameDia 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   3990
      TabIndex        =   24
      Top             =   1050
      Width           =   2115
      Begin MSComCtl2.DTPicker Txt_data 
         Height          =   315
         Left            =   630
         TabIndex        =   5
         ToolTipText     =   "Data."
         Top             =   540
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
         Format          =   197459969
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dia :"
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
         TabIndex        =   25
         Top             =   570
         Width           =   330
      End
   End
   Begin VB.Frame FrameObs 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   3990
      TabIndex        =   32
      Top             =   1050
      Visible         =   0   'False
      Width           =   2115
      Begin VB.TextBox Txt_dias_obs 
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
         Left            =   300
         MouseIcon       =   "frmestoque_item_relat.frx":40A2
         TabIndex        =   16
         ToolTipText     =   "Quantidade de dias."
         Top             =   1560
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker Txt_ate_obs 
         Height          =   315
         Left            =   300
         TabIndex        =   15
         ToolTipText     =   "A partir de."
         Top             =   465
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
         Format          =   197459969
         CurrentDate     =   39057
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dias"
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
         Left            =   1380
         TabIndex        =   35
         Top             =   1650
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Com mais de"
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
         Left            =   345
         TabIndex        =   34
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "A partir de"
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
         Left            =   540
         TabIndex        =   33
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa :"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   885
   End
End
Attribute VB_Name = "frmestoque_item_relat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormulaRel_Estoque_Movimentacao As String 'OK

Private Sub btnImprimir_Click()
On Error GoTo tratar_erro

If Opt_saldo_diario.Value = False And optMovimentacao.Value = False And optSaldo_resumido.Value = False And optSaldo.Value = False Then
  USMsgBox "Selecione uma das opções para filtrar o relatório", vbInformation, "CAPRIND v5.0"
  Exit Sub
End If

'Filtrar movimentação do estoque
If optMovimentacao.Value = True Then
  ProcMovimentacao
  Exit Sub
End If

'Filtrar saldo diario
If Opt_saldo_diario.Value = True Then
  procSaldoDiario
  Exit Sub
End If


ProcImprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSaldoDiario()
On Error GoTo tratar_erro

ProcGravarSaldoDiario
FormulaRel_Estoque_Movimentacao = ""
NomeRel = "Estoque_saldo_diario.rpt"
ProcImprimirRel FormulaRel_Estoque_Movimentacao, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub ProcMovimentacao()
On Error GoTo tratar_erro

ProcGravarMovimentacao
FormulaRel_Estoque_Movimentacao = ""
NomeRel = "Estoque_Movimentacao.rpt"
ProcImprimirRel FormulaRel_Estoque_Movimentacao, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

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

If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Local de armazenamento" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    If cmbfiltrarpor = "Família" Then
         StrSql = "select distinct(Familia) from Estoque_Controle_Saldo_RE"

         With cmbfamilia
             .Clear
             Set TBCarregarCombo = CreateObject("adodb.recordset")
             TBCarregarCombo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
             If TBCarregarCombo.EOF = False Then
                 If CarregarVazio = True Then .AddItem ""
                 Do While TBCarregarCombo.EOF = False
                     If IsNull(TBCarregarCombo!Familia) = False And TBCarregarCombo!Familia <> "" Then
                         .AddItem TBCarregarCombo!Familia
                     End If
                     TBCarregarCombo.MoveNext
                 Loop
             End If
         End With

        'ProcCarregaComboFamilia cmbfamilia, "Familia <> 'Null'", False
    ElseIf cmbfiltrarpor = "Grupo" Then
            ProcCarregaComboGrupoFamilia cmbfamilia, "Grupo <> 'Null'", False
        Else
            ProcCarregaComboLA cmbfamilia, False, True
    End If
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Opt_semi_acabado.Value = False And Opt_semi_acabado_SPED.Value = False And Chk_maior.Value = 0 And Chk_igual = 0 Then
    USMsgBox ("Informe uma das opções de filtro antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If optMovimentacao.Value = True And chkentrada.Value = 0 And chkSaida.Value = 0 Then
    USMsgBox ("Informe uma das opções de filtro da movimentação antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

IDlista = IIf(Txt_dias_obs = "", 0, Txt_dias_obs)

If Opt_obsolescencia.Value = True Then

    If IDlista <= 0 Then
        Acao = "visualizar impressão"
        NomeCampo = "a quantidade de dias"
        ProcVerificaAcao
        Txt_dias_obs.SetFocus
        Exit Sub
    End If
    
    ProcExcluirDadosProducaoRelatoriosTotal
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and QtdeOrdem = 8", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then TBGravar.AddNew
    TBGravar!QtdeOrdem = 8
    TBGravar!Data_final = Txt_ate_obs
    TBGravar!Responsavel = pubUsuario
    TBGravar!Modulo = Formulario
    TBGravar!QtdePrevista = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBGravar!QtdeProduzida = IDlista
    TBGravar.Update
    TBGravar.Close
End If

FormulaRel_Estoque_MovimentacaoSubReport = ""

If Chk_maior.Value = 1 And Chk_igual.Value = 0 Then
    If Opt_semi_acabado.Value = True Or Opt_semi_acabado_SPED.Value = True Then
        OpcaoFiltro = "{Estoque_produto_SA.Saldo} > 0"
'    ElseIf Opt_saldo_diario.Value = True Then
'            OpcaoFiltro = "{Estoque_diario.Qtde} > 0"
    Else
        If optSaldo_resumido.Value = True Then
            OpcaoFiltro = "{Estoque_controle_Saldo_RE.Saldo} > 0"
        Else
            OpcaoFiltro = "{Estoque_controle_Saldo_RE.Saldo} > 0"
        End If
    End If
ElseIf Chk_maior.Value = 0 And Chk_igual.Value = 1 Then
    If optSaldo_resumido.Value = True Then
        OpcaoFiltro = "{Qtde_estoque_produto.Estoque_real} <= 0"
    Else
        OpcaoFiltro = "{Estoque_controle_Saldo_RE.Saldo} <= 0"
    End If
Else
    If Opt_semi_acabado.Value = True Or Opt_semi_acabado_SPED.Value = True Then
        OpcaoFiltro = "{Estoque_produto_SA.Desenho} <> 'Null'"
'        ElseIf Opt_saldo_diario.Value = True Then
'                OpcaoFiltro = "{Estoque_diario.Desenho} <> 'Null'"
    Else
        OpcaoFiltro = "{projproduto.desenho} <> 'Null'"
    End If
End If

OpcaoFiltro_movimentacao = ""
OpcaoFiltro_movimentacao1 = ""
strdataIni = "Date(" & Format(msk_fltInicio.Value, "yyyy,mm,dd") & ")"
strDataFim = "Date(" & Format(msk_fltFim.Value, "yyyy,mm,dd") & ")"

'Datafim

If optMovimentacao.Value = True Then
    DataFiltro = "and {estoque_movimentacao.Data} >= " & strdataIni & " and {estoque_movimentacao.Data} <= " & strDataFim & " and ({Estoque_movimentacao.IDEstoque_recebimento} <> 0 and {Compras_pedido_lista.Remessa} = False or {Estoque_movimentacao.IDEstoque_recebimento} = 0)"
    'Debug.print DataFiltro
    
    If chkentrada.Value = 1 And chkSaida.Value = 0 Then
        OpcaoFiltro_movimentacao1 = " and {estoque_movimentacao.Entrada} <> 0"
    ElseIf chkentrada.Value = 0 And chkSaida.Value = 1 Then
            OpcaoFiltro_movimentacao1 = " and {estoque_movimentacao.Saida} <> 0"
    End If
    OpcaoFiltro_movimentacao = DataFiltro & OpcaoFiltro_movimentacao1 & " and {Estoque_movimentacao.Operacao} <> 'DEVOLUCAO_ALMOXARIFADO C/ PROB.'"
End If

ConsignacaoFiltro = ""
If Opt_semi_acabado_SPED.Value = True Then ConsignacaoFiltro = "and {Estoque_produto_SA.Consignacao} = False"

If txtTexto <> "" Or cmbfamilia <> "" Then
    If Opt_semi_acabado.Value = True Or Opt_semi_acabado_SPED.Value = True Then
        If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
            If cmbfiltrarpor = "Família" Then TextoFiltro = "{Estoque_produto_SA.classe}" Else TextoFiltro = "{Estoque_produto_SA.Grupo}"
            FormulaRel_Estoque_Movimentacao = TextoFiltro & " = '" & cmbfamilia & "' and {Estoque_produto_SA.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & OpcaoFiltro & " " & ConsignacaoFiltro
        Else
            Select Case cmbfiltrarpor
                Case "Código interno": TextoFiltro = "{Estoque_produto_SA.desenho}"
                Case "Descrição": TextoFiltro = "{Estoque_produto_SA.Descricao}"
                Case "Cliente": TextoFiltro = "{Estoque_produto_SA.NomeRazao}"
            End Select
            FormulaRel_Estoque_Movimentacao = TextoFiltro & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and {Estoque_produto_SA.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & OpcaoFiltro & " " & ConsignacaoFiltro
        End If
    ElseIf Opt_saldo_diario.Value = True Then
            If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
                Select Case cmbfiltrarpor
                    Case "Família": TextoFiltro = "{Estoque_diario.classe}"
                    Case "Grupo": TextoFiltro = "{Estoque_diario.Grupo}"
                End Select
                FormulaRel_Estoque_Movimentacao = TextoFiltro & " = '" & cmbfamilia & "' and {Estoque_diario.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
            Else
                If cmbfiltrarpor = "Código interno" Then TextoFiltro = "{Estoque_diario.desenho}" Else TextoFiltro = "{Estoque_diario.Descricao}"
                FormulaRel_Estoque_Movimentacao = TextoFiltro & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and {Estoque_diario.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
            End If
        ElseIf Opt_obsolescencia.Value = True Then
                TextoFiltroPadrao = "and {Estoque_obsolescencia.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {Estoque_obsolescencia.Dias_sem_movim} > " & IDlista & " and {Qtde_estoque_produto.Estoque_real} > 0"
                If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Local de armazenamento" Then
                    Select Case cmbfiltrarpor
                        Case "Família": TextoFiltro = "{Estoque_obsolescencia.classe}"
                        Case "Grupo": TextoFiltro = "{Estoque_obsolescencia.Grupo}"
                        Case "Local de armazenamento": TextoFiltro = "{Estoque_obsolescencia.local_armaz}"
                    End Select
                    FormulaRel_Estoque_Movimentacao = TextoFiltro & " = '" & cmbfamilia & "' " & TextoFiltroPadrao
                Else
                    If cmbfiltrarpor = "Código interno" Then TextoFiltro = "{Estoque_obsolescencia.desenho}" Else TextoFiltro = "{Estoque_obsolescencia.Descricao}"
                    FormulaRel_Estoque_Movimentacao = TextoFiltro & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " " & TextoFiltroPadrao
                End If
            Else
                If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Or cmbfiltrarpor = "Local de armazenamento" Then
                    Select Case cmbfiltrarpor
                        Case "Família": TextoFiltro = "{Estoque_Controle_Saldo_RE.Familia}"
                        Case "Grupo": TextoFiltro = "{Projfamilia.Grupo}"
                        Case "Local de armazenamento": TextoFiltro = "{Estoque_Controle_Saldo_RE.local_armaz}"
                    End Select
                    FormulaRel_Estoque_Movimentacao = TextoFiltro & " = '" & cmbfamilia & "' and " & OpcaoFiltro & " and {Empresa.Codigo} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {projproduto.Estoque} = True " & OpcaoFiltro_movimentacao
                Else
                    If cmbfiltrarpor = "Código interno" Then TextoFiltro = "{Estoque_Controle_Saldo_RE.Codigo}" Else TextoFiltro = "{Estoque_Controle_Saldo_RE.Descricao}"
                    FormulaRel_Estoque_Movimentacao = TextoFiltro & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & OpcaoFiltro & " and {Empresa.Codigo} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {projproduto.Estoque} = True " & OpcaoFiltro_movimentacao
                End If
    End If
Else
    If Opt_semi_acabado.Value = True Or Opt_semi_acabado_SPED.Value = True Then
        FormulaRel_Estoque_Movimentacao = "{Estoque_produto_SA.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & OpcaoFiltro & " " & ConsignacaoFiltro
    ElseIf Opt_saldo_diario.Value = True Then
            ProcGravarSaldoDiario
            FormulaRel_Estoque_Movimentacao = "{Estoque_consulta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
        ElseIf Opt_obsolescencia.Value = True Then
                FormulaRel_Estoque_Movimentacao = "{Estoque_obsolescencia.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {Estoque_obsolescencia.Dias_sem_movim} > " & IDlista & " and {Qtde_estoque_produto.Estoque_real} > 0"
            Else
                FormulaRel_Estoque_Movimentacao = OpcaoFiltro & " and {Empresa.Codigo} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {projproduto.Estoque} = True " & OpcaoFiltro_movimentacao
    End If
End If

' Saldo estoque detalhado
If optSaldo.Value = True Then
    NomeRel = "Estoque_saldo_detalhado.rpt"
End If

'Saldo estoque resumido
If optSaldo_resumido.Value = True Then
    NomeRel = "Estoque_saldo_resumido.rpt"
End If

If Opt_obsolescencia.Value = True Then
    NomeRel = "Estoque_obsolescencia.rpt"
    FormulaRel_Estoque_Movimentacao = FormulaRel_Estoque_Movimentacao & " And {Estoque_obsolescencia.Responsavel} = '" & pubUsuario & "'"
End If

If optMovimentacao.Value = True Then
    ProcGravarDataFiltroRel msk_fltInicio, msk_fltFim, True, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""
    NomeRel = "Estoque_movimentacao.rpt"
    FormulaRel_Estoque_Movimentacao = FormulaRel_Estoque_Movimentacao & " And {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'"
End If

If Opt_semi_acabado.Value = True Then
    NomeRel = "Estoque_semiacabado.rpt"
End If
If Opt_semi_acabado_SPED.Value = True Then
    NomeRel = "Estoque_semiacabado_SPED.rpt"
End If

'Debug.print FormulaRel_Estoque_Movimentacao
ProcImprimirRel FormulaRel_Estoque_Movimentacao, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPor_Click()
On Error GoTo tratar_erro

ProcMostrarEsconderCombosData

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarSaldoDiario()
On Error GoTo tratar_erro

Conexao.Execute "Delete from Estoque_Consulta"
Conexao.Execute "update Estoque_movimentacao set Estoque = PP.Estoque, Familia = PP.Classe, Unidade = PP.Unidade from Estoque_movimentacao EM Inner Join projproduto PP on EM.Desenho = PP.Desenho"
Conexao.Execute "update Estoque_movimentacao set Grupo = PF.Grupo from Estoque_movimentacao EM Inner Join projFamilia PF on EM.Familia = PF.Familia"
Conexao.Execute "update Estoque_movimentacao set ID_Empresa = EC.ID_empresa from Estoque_movimentacao EM Inner Join Estoque_controle EC on EM.IDEstoque = EC.IDestoque"
Conexao.Execute "update Estoque_Movimentacao set Local_Armaz = EC.Local_Armaz from Estoque_movimentacao EM inner Join Estoque_Controle EC on EC.IdEstoque = EM.IdEstoque"


If cmbfiltrarpor.Text = "Código interno" And txtTexto <> "" Then

If Optinicio.Value = True Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Codigo Like '" & txtTexto.Text & "%' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If

If Optmeio.Value = True Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Codigo Like '%" & txtTexto.Text & "%' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If

If Optfim.Value = True Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Codigo Like '%" & txtTexto.Text & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If

If optIgual.Value = True Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Codigo = '" & txtTexto.Text & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If
End If

If cmbfiltrarpor.Text = "Descrição" And txtTexto <> "" Then
    If Optinicio.Value = True Then
    StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Descricao Like '" & txtTexto.Text & "%' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
    End If
    
    If Optmeio.Value = True Then
    StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Descricao Like '%" & txtTexto.Text & "%' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
    End If
    
    If Optfim.Value = True Then
    StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Descricao Like '%" & txtTexto.Text & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
    End If
    
    If optIgual.Value = True Then
    StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Descricao = '" & txtTexto.Text & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
    End If
    
End If

If cmbfiltrarpor.Text = "Família" And cmbfamilia.Text <> "" Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Familia = '" & cmbfamilia.Text & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If

If cmbfiltrarpor.Text = "Grupo" And cmbfamilia.Text <> "" Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and grupo = '" & cmbfamilia.Text & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If

If cmbfiltrarpor.Text = "Local de armazenamento" And cmbfamilia.Text <> "" Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Local_Armaz = '" & cmbfamilia.Text & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If


If txtTexto = "" And cmbfamilia.Text = "" Then
StrSql = "Insert Into Estoque_consulta(ID_Empresa,Grupo, Familia, Desenho, Descricao, Unidade, VlrUnit, Saldo) Select ID_Empresa, Grupo, Familia,Codigo,Descricao,un,sum(valor_total)/sum(saldo) as vlrUnit, Sum(Saldo) as Saldo from Estoque_Controle_Saldo_RE where ID_empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' AND Estoque = '1' and data <= '" & Txt_data.Value & "' and Saldo > 0 group by  ID_Empresa,Codigo,Descricao,Un, Grupo, Familia"
End If



'Debug.print StrSql


Conexao.Execute Trim(StrSql)

StrSql = "Update Estoque_Consulta set DtInicio = '" & Txt_data.Value & "', Dtfinal = '" & Txt_data.Value & "'"

'Debug.print StrSql
Conexao.Execute StrSql

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarMovimentacaoOLD()
On Error GoTo tratar_erro

Conexao.Execute "Delete from Estoque_Consulta_Movimentacao"
Conexao.Execute "Delete from Estoque_Consulta_Movimentacao_Totais"

strdataIni = msk_fltInicio.Value
strDataFim = msk_fltFim.Value

If cmbfiltrarpor.Text = "Código interno" And txtTexto <> "" Then

    If Optinicio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote ,Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao, EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.Entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDOperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque,Desenho, Data, Lote"
    End If
    
    If Optmeio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '%" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDOperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
    End If
    
    If Optfim.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '%" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
    End If
    
    If optIgual.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
    End If

End If

If cmbfiltrarpor.Text = "Descrição" And txtTexto <> "" Then

    If Optinicio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
    End If
    
    If Optmeio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '%" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
    End If
    
    If Optfim.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '%" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit, EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
    End If
    
    If optIgual.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit, EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
    End If
    
End If
'Filtrar por Familia
If cmbfiltrarpor.Text = "Família" And cmbfamilia.Text <> "" Then
  StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Familia = '" & cmbfamilia.Text & "' And EM.Data >= '" & strdataIni & "' and EM.Data <= '" & strDataFim & "' Group by EM.Grupo, EM.IDEstoque, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
End If

'Filtrar por Grupo
If cmbfiltrarpor.Text = "Grupo" And cmbfamilia.Text <> "" Then
  StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Grupo = '" & cmbfamilia.Text & "' And EM.Data >= '" & strdataIni & "' and EM.Data <= '" & strDataFim & "' Group by EM.Grupo, EM.IDEstoque, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
End If

'Filtrar por local de armazenamento
If cmbfiltrarpor.Text = "Local de armazenamento" And cmbfamilia.Text <> "" Then

  StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And EM.Data >= '" & strdataIni & "' and EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data, Lote"
End If

'Filtrar tudo
If txtTexto = "" And cmbfamilia.Text = "" Then
  If Chk_maior = 1 And Chk_igual = 0 Then 'Com saldo em estoque
    StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao, EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Data >= '" & strdataIni & "' AND EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz, EM.Operacao ORDER BY EM.IDEstoque, Desenho, Data , Lote"
  '  StrSql = "Insert Into Estoque_consulta_Movimentacao(Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EM.local_armaz from Estoque_movimentacao EM Where EM.Data >= '" & strdataIni & "' AND EM.Data <= '" & strDataFim & "' Group by EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade, EM.VlrUnit,EM.IDoperacao, EM.local_armaz, EM.Operacao ORDER BY Desenho,EM.Data , Lote"
  
  End If
  
  If Chk_maior = 0 And Chk_igual = 1 Then 'Sem Saldo em estoque
  
  End If
  
  If Chk_maior = 0 And Chk_igual = 0 Then 'Com saldo e sem saldo
  
  End If
End If

'Debug.print StrSql

'Grava dados na tabela para o relatório
Conexao.Execute StrSql

'Atualiza Saldo nos itens

Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select * from Estoque_Consulta_Movimentacao order by ID_Consulta_Mov"
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = True Then Exit Sub


If TBAbrir.EOF = False Then
    Desenho = ""
    Loteitem = ""
    
    Do While TBAbrir.EOF = False
'====================================================================================================
        If TBAbrir!Desenho <> Desenho Or Loteitem <> TBAbrir!LOTE Then
        Desenho = TBAbrir!Desenho
        Loteitem = TBAbrir!LOTE
        
                Set TBSaldo = CreateObject("adodb.recordset")
                
                Select Case cmbfiltrarpor.Text
                    Case "Código interno":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.lote = '" & TBAbrir!LOTE & "'  and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho, EM.Lote"
                    Case "Descrição":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.lote = '" & TBAbrir!LOTE & "'  and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                    Case "Família":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.lote = '" & TBAbrir!LOTE & "'  and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                    Case "Grupo":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.lote = '" & TBAbrir!LOTE & "'  and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                    Case "Local de armazenamento":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And desenho = '" & TBAbrir!Desenho & "' and EM.lote = '" & TBAbrir!LOTE & "'  and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                End Select
                
                TBSaldo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                'Debug.print StrSql
                'Debug.print TBAbrir!Desenho
                
                If TBSaldo.EOF = False Then
                    Saldo = IIf(IsNull(TBSaldo!Saldo), 0, TBSaldo!Saldo)
                    Else
                    Saldo = 0
                End If
                
                Set TBSaldo = CreateObject("adodb.recordset")
                StrSql = "Select avg(vlrUnit) as ValorUnit from estoque_Consulta_Movimentacao where Desenho = '" & TBAbrir!Desenho & "' and Lote  = '" & TBAbrir!LOTE & "'"
                TBSaldo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                'Debug.print StrSql
                
                If TBSaldo.EOF = False Then
                    VlrUnit = IIf(IsNull(TBSaldo!ValorUnit), 0, TBSaldo!ValorUnit)
                    Else
                    VlrUnit = 0
                End If
                
        End If

'==========================================================
            TBAbrir!SaldoInicial = IIf(Saldo > 0, Saldo, 0)
            TBAbrir!Saldo = (TBAbrir!Entrada - TBAbrir!Saida) + Saldo
            TBAbrir!vlrTotal = VlrUnit * TBAbrir!Saldo
            Saldo = TBAbrir!Saldo
            TBAbrir.Update
    TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close
TBSaldo.Close

'Grava ID_empresa na tabela
StrSql = "Update Estoque_Consulta_Movimentacao set ID_Empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' , DtIni = '" & strdataIni & "', Dtfim = '" & strDataFim & "'"
Conexao.Execute StrSql


'Grava totais na tabela do item
StrSql = "Select Desenho from estoque_consulta_Movimentacao GROUP BY desenho"
Set TBSaldo = CreateObject("adodb.recordset")
TBSaldo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBSaldo.EOF = False Then
        Do While TBSaldo.EOF = False
        
            Select Case cmbfiltrarpor.Text
                Case "Código interno":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Descrição":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Família":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Grupo":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Local de armazenamento":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
            End Select
        
            'Debug.print StrSQLTotais
            Conexao.Execute StrSQLTotais
            TBSaldo.MoveNext
        Loop
    End If
TBSaldo.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarMovimentacao()
On Error GoTo tratar_erro

Conexao.Execute "Delete from Estoque_Consulta_Movimentacao"
Conexao.Execute "Delete from Estoque_Consulta_Movimentacao_Totais"

strdataIni = msk_fltInicio.Value
strDataFim = msk_fltFim.Value

If cmbfiltrarpor.Text = "Código interno" And txtTexto <> "" Then

    If Optinicio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote ,Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao, EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.Entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDOperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If
    
    If Optmeio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '%" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDOperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If
    
    If Optfim.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '%" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If
    
    If optIgual.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Desenho like '" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If

End If

If cmbfiltrarpor.Text = "Descrição" And txtTexto <> "" Then

    If Optinicio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If
    
    If Optmeio.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '%" & txtTexto.Text & "%' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If
    
    If Optfim.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '%" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit, EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If
    
    If optIgual.Value = True Then
      StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Descricao like '" & txtTexto.Text & "' And EM.Data >= '" & strdataIni & "' And EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit, EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
    End If
    
End If
'Filtrar por Familia
If cmbfiltrarpor.Text = "Família" And cmbfamilia.Text <> "" Then
  StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Familia = '" & cmbfamilia.Text & "' And EM.Data >= '" & strdataIni & "' and EM.Data <= '" & strDataFim & "' Group by EM.Grupo, EM.IDEstoque, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
End If

'Filtrar por Grupo
If cmbfiltrarpor.Text = "Grupo" And cmbfamilia.Text <> "" Then
  StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Grupo = '" & cmbfamilia.Text & "' And EM.Data >= '" & strdataIni & "' and EM.Data <= '" & strDataFim & "' Group by EM.Grupo, EM.IDEstoque, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
End If

'Filtrar por local de armazenamento
If cmbfiltrarpor.Text = "Local de armazenamento" And cmbfamilia.Text <> "" Then

  StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And EM.Data >= '" & strdataIni & "' and EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz,EM.Operacao ORDER BY Desenho, Data"
End If

'Filtrar tudo
If txtTexto = "" And cmbfamilia.Text = "" Then
  If Chk_maior = 1 And Chk_igual = 0 Then 'Com saldo em estoque
    StrSql = "Insert Into Estoque_consulta_Movimentacao(IDEstoque, Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, Certificado,Corrida, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao, EM.Lote, EC.Certificado, EC.Corrida, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EC.local_armaz from Estoque_movimentacao EM Inner join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EM.Data >= '" & strdataIni & "' AND EM.Data <= '" & strDataFim & "' Group by EM.IDEstoque, EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade,EC.Certificado,EC.corrida,EM.VlrUnit,EM.IDoperacao, EC.local_armaz, EM.Operacao ORDER BY Desenho, Data"
  '  StrSql = "Insert Into Estoque_consulta_Movimentacao(Grupo, Familia,Desenho, Unidade, Descricao,Data,Operacao,Lote, VlrUnit, Entrada, Saida, vlrTotal, Local_armaz) Select EM.Grupo, EM.Familia, EM.Desenho,EM.Unidade, EM.Descricao,Em.Data,EM.Operacao,EM.Lote, EM.VlrUnit, SUM(EM.Entrada) as Entrada, SUM(EM.Saida)as Saida, IIF(SUM(EM.entrada)>'0', SUM((EM.entrada) * EM.vlrUnit),SUM((EM.saida) * EM.vlrUnit)) as vlrTotal, EM.local_armaz from Estoque_movimentacao EM Where EM.Data >= '" & strdataIni & "' AND EM.Data <= '" & strDataFim & "' Group by EM.Grupo, EM.Familia, EM.Desenho, EM.Descricao,EM.Data,EM.Lote, EM.Unidade, EM.VlrUnit,EM.IDoperacao, EM.local_armaz, EM.Operacao ORDER BY Desenho,EM.Data , Lote"
  
  End If
  
  If Chk_maior = 0 And Chk_igual = 1 Then 'Sem Saldo em estoque
  
  End If
  
  If Chk_maior = 0 And Chk_igual = 0 Then 'Com saldo e sem saldo
  
  End If
End If

'Debug.print StrSql

'Grava dados na tabela para o relatório
Conexao.Execute StrSql

'Atualiza Saldo nos itens

Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select * from Estoque_Consulta_Movimentacao order by ID_Consulta_Mov"
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = True Then Exit Sub


If TBAbrir.EOF = False Then
    Desenho = ""
    Loteitem = ""
    
    Do While TBAbrir.EOF = False
'====================================================================================================
        If TBAbrir!Desenho <> Desenho Then
        Desenho = TBAbrir!Desenho
        Loteitem = TBAbrir!LOTE
        
                Set TBSaldo = CreateObject("adodb.recordset")
                
                Select Case cmbfiltrarpor.Text
                    Case "Código interno":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                    Case "Descrição":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                    Case "Família":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                    Case "Grupo":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBAbrir!Desenho & "' and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                    Case "Local de armazenamento":
                        StrSql = "SELECT sum(EM.entrada-EM.saida) as Saldo from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And desenho = '" & TBAbrir!Desenho & "' and EM.data <= '" & msk_fltInicio - 1 & "' GROUP BY Desenho"
                End Select
                
                TBSaldo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                'Debug.print StrSql
                'Debug.print TBAbrir!Desenho
                
                If TBSaldo.EOF = False Then
                    Saldo = IIf(IsNull(TBSaldo!Saldo), 0, TBSaldo!Saldo)
                    Else
                    Saldo = 0
                End If
                
                Set TBSaldo = CreateObject("adodb.recordset")
                StrSql = "Select avg(vlrUnit) as ValorUnit from estoque_Consulta_Movimentacao where Desenho = '" & TBAbrir!Desenho & "'"
                TBSaldo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                'Debug.print StrSql
                
                If TBSaldo.EOF = False Then
                    VlrUnit = IIf(IsNull(TBSaldo!ValorUnit), 0, TBSaldo!ValorUnit)
                    Else
                    VlrUnit = 0
                End If
                
        End If

'==========================================================
            TBAbrir!SaldoInicial = IIf(Saldo > 0, Saldo, 0)
            TBAbrir!Saldo = (TBAbrir!Entrada - TBAbrir!Saida) + Saldo
            If TBAbrir!Saldo > 0 Then
            TBAbrir!vlrTotal = VlrUnit * TBAbrir!Saldo
            Else
            TBAbrir!vlrTotal = 0
            End If
            
            Saldo = TBAbrir!Saldo
            TBAbrir.Update
    TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close
TBSaldo.Close

'Grava ID_empresa na tabela
StrSql = "Update Estoque_Consulta_Movimentacao set ID_Empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "' , DtIni = '" & strdataIni & "', Dtfim = '" & strDataFim & "'"
Conexao.Execute StrSql


'Grava totais na tabela do item
StrSql = "Select Desenho from estoque_consulta_Movimentacao GROUP BY desenho"
Set TBSaldo = CreateObject("adodb.recordset")
TBSaldo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBSaldo.EOF = False Then
        Do While TBSaldo.EOF = False
        
            Select Case cmbfiltrarpor.Text
                Case "Código interno":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Descrição":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Família":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Grupo":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
                Case "Local de armazenamento":
                    StrSQLTotais = "INSERT INTO Estoque_consulta_Movimentacao_Totais(Desenho,  vlrUnitMedio, TotalEntrada, TotalSaida,SaldoFinal) SELECT Desenho, AVG(vlrUnit) as VlrUnit, sum(EM.entrada) as Entrada,sum(EM.Saida) as Saida, (SELECT sum(EM.entrada-EM.saida) from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And desenho = '" & TBSaldo!Desenho & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho) as SaldoFinal from Estoque_movimentacao EM inner Join Estoque_Controle_Saldo_RE EC on EM.IdEstoque = EC.IdEstoque Where EC.local_armaz = '" & cmbfamilia.Text & "' And desenho = '" & TBSaldo!Desenho & "' and EM.data >= '" & msk_fltInicio & "' and EM.data <= '" & msk_fltFim & "' GROUP BY Desenho"
            End Select
        
            'Debug.print StrSQLTotais
            Conexao.Execute StrSQLTotais
            TBSaldo.MoveNext
        Loop
    End If
TBSaldo.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarEsconderCombosData()
On Error GoTo tratar_erro

If cmbPor = "Dia" Then
    msk_fltInicio.Visible = True
    msk_fltFim.Visible = True
    Cmb_mes_de.Visible = False
    Cmb_mes_ate.Visible = False
    Cmb_ano_de.Visible = False
    Cmb_ano_ate.Visible = False
    Cmb_ano_de1.Visible = False
    Cmb_ano_ate1.Visible = False
ElseIf cmbPor = "Mês" Then
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = True
        Cmb_mes_ate.Visible = True
        Cmb_ano_de.Visible = True
        Cmb_ano_ate.Visible = True
        Cmb_ano_de1.Visible = False
        Cmb_ano_ate1.Visible = False
    Else
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = False
        Cmb_mes_ate.Visible = False
        Cmb_ano_de.Visible = False
        Cmb_ano_ate.Visible = False
        Cmb_ano_de1.Visible = True
        Cmb_ano_ate1.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF5: ProcImprimir
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Código interno"
txtTexto.Visible = True
cmbfamilia.Visible = False
Txt_data.Value = Date
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
Txt_ate_obs.Value = Date
optMovimentacao.Value = True
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_obsolescencia_Click()
On Error GoTo tratar_erro

ProcCarregaComboFiltrarPor
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_semi_acabado_Click()
On Error GoTo tratar_erro

ProcCarregaComboFiltrarPor
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_semi_acabado_SPED_Click()
On Error GoTo tratar_erro

ProcCarregaComboFiltrarPor
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optMovimentacao_Click()
On Error GoTo tratar_erro

ProcCarregaComboFiltrarPor
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSaldo_Click()
On Error GoTo tratar_erro

ProcCarregaComboFiltrarPor
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_saldo_diario_Click()
On Error GoTo tratar_erro

ProcCarregaComboFiltrarPor
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSaldo_resumido_Click()
On Error GoTo tratar_erro

ProcCarregaComboFiltrarPor
ProcVerificaBloqueiaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboFiltrarPor()
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    .AddItem "Grupo"
    .AddItem "Local de armazenamento"
    If Opt_semi_acabado.Value = True Or Opt_semi_acabado_SPED.Value = True Then .AddItem "Cliente"
    If Opt_semi_acabado.Value = False And Opt_semi_acabado_SPED.Value = False And Opt_saldo_diario.Value = False Then .AddItem "Local de armazenamento"
    .Text = "Código interno"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaBloqueiaCampos()
On Error GoTo tratar_erro

With chkentrada
    .Value = 1
    .Enabled = False
End With
With chkSaida
    .Value = 1
    .Enabled = False
End With
With frameDia
    .Visible = True
    .Enabled = False
End With
With FramePeriodo
    .Visible = False
    .Enabled = False
End With
With FrameObs
    .Visible = False
    .Enabled = False
End With
With cmbPor
    .Text = "Dia"
    .Locked = False
    .TabStop = True
End With

If Opt_saldo_diario.Value = True Or Opt_semi_acabado.Value = True Or Opt_semi_acabado_SPED.Value = True Or Opt_obsolescencia.Value = True Then
    With Chk_maior
        .Value = 1
        If Opt_saldo_diario.Value = True Or Opt_obsolescencia.Value = True Then .Enabled = False Else .Enabled = True
    End With
    With Chk_igual
        .Value = 0
        .Enabled = False
    End With
    If Opt_saldo_diario.Value = True Then
        frameDia.Enabled = True
    ElseIf Opt_semi_acabado.Value = True Or Opt_semi_acabado_SPED.Value = True Then
            frameDia.Enabled = False
        Else
            frameDia.Visible = False
            FrameObs.Enabled = True
            FrameObs.Visible = True
    End If
Else
    With Chk_maior
        .Value = 1
        .Enabled = True
    End With
    With Chk_igual
        .Value = 0
        .Enabled = True
    End With
    frameDia.Enabled = False
    FrameObs.Visible = False
    If optMovimentacao.Value = True Then
        chkentrada.Enabled = True
        chkSaida.Enabled = True
        With FramePeriodo
            .Visible = True
            .Enabled = True
        End With
        frameDia.Visible = False
        With cmbPor
            .Clear
            .AddItem "Dia"
            .Text = "Dia"
            .Locked = True
            .TabStop = False
        End With
        ProcMostrarEsconderCombosData
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_dias_obs_Change()
On Error GoTo tratar_erro

If Txt_dias_obs <> "" Then
    VerifNumero = Txt_dias_obs
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_dias_obs = ""
        Txt_dias_obs.SetFocus
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

If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

