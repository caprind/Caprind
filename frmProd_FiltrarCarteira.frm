VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frmProd_FiltrarCarteira 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Ordem de faturamento | Filtrar carteira"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6510
   StartUpPosition =   3  'Padrão Windows
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar somente"
      Height          =   945
      Left            =   180
      TabIndex        =   24
      Top             =   3210
      Width           =   6045
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Empenhados"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Value           =   1  'Marcado
         Width           =   2055
      End
      Begin VB.CheckBox Chk_tem_estoque 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Com saldo em estoque"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Value           =   1  'Marcado
         Width           =   2265
      End
      Begin DrawSuite2014.USButton btnFiltrar 
         Height          =   645
         Left            =   4020
         TabIndex        =   27
         Top             =   180
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1138
         DibPicture      =   "frmProd_FiltrarCarteira.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Filtrar carteira"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         ForeColorDown   =   16777215
         ForeColorOver   =   16777215
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         Theme           =   4
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período"
      Height          =   945
      Left            =   180
      TabIndex        =   19
      Top             =   2250
      Width           =   2715
      Begin VB.CheckBox Chk_data 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dt. venda"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker msk_data 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         ToolTipText     =   "Data final para pesquisa."
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   890241025
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_data 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   22
         ToolTipText     =   "Data início para pesquisa."
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   890241027
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparente
         Caption         =   "à"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   25
         Left            =   1290
         TabIndex        =   23
         Top             =   600
         Width           =   225
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   945
      Left            =   2910
      TabIndex        =   10
      Top             =   2250
      Width           =   3315
      Begin VB.TextBox txtTexto 
         Alignment       =   2  'Centralizar
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
         Left            =   90
         TabIndex        =   16
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Width           =   3000
      End
      Begin VB.ComboBox cmbTexto 
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
         ItemData        =   "frmProd_FiltrarCarteira.frx":3650
         Left            =   90
         List            =   "frmProd_FiltrarCarteira.frx":3652
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Familia."
         Top             =   480
         Width           =   3015
      End
      Begin VB.OptionButton Optinicio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Início"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   210
         Width           =   645
      End
      Begin VB.OptionButton Optmeio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Meio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   870
         TabIndex        =   13
         Top             =   210
         Width           =   675
      End
      Begin VB.OptionButton Optfim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fim"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1590
         TabIndex        =   12
         Top             =   210
         Width           =   585
      End
      Begin VB.OptionButton optIgual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Igual"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   210
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Carregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   22
      Left            =   2910
      TabIndex        =   7
      Top             =   540
      Width           =   3315
      Begin VB.OptionButton optespecificacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição comercial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   18
         Top             =   480
         Width           =   1725
      End
      Begin VB.OptionButton Optdescricao 
         BackColor       =   &H00E0E0E0&
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
         Height          =   210
         Left            =   1170
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton Opt_servico_filtrar 
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
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   915
      End
      Begin VB.OptionButton Opt_produto_filtrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   26
      Left            =   180
      TabIndex        =   4
      Top             =   540
      Width           =   2715
      Begin VB.OptionButton Opt_filtrar_ped_int 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ped. interno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Opt_filtrar_ped_compra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ped. compra (remessa)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   5
         Top             =   450
         Width           =   2025
      End
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
      Height          =   960
      Index           =   25
      Left            =   180
      TabIndex        =   1
      Top             =   1290
      Width           =   6045
      Begin VB.ComboBox Cmb_empresa_filtro 
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
         ItemData        =   "frmProd_FiltrarCarteira.frx":3654
         Left            =   180
         List            =   "frmProd_FiltrarCarteira.frx":3656
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Empresa."
         Top             =   510
         Width           =   2625
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
         ItemData        =   "frmProd_FiltrarCarteira.frx":3658
         Left            =   2820
         List            =   "frmProd_FiltrarCarteira.frx":3671
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   510
         Width           =   3075
      End
   End
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   714
      DibPicture      =   "frmProd_FiltrarCarteira.frx":36DB
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProd_FiltrarCarteira.frx":A85B
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmProd_FiltrarCarteira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
