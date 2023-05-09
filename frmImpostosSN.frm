VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImpostosSN 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Impostos Simples nacional"
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   Icon            =   "frmImpostosSN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9090
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIDEmpresa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   750
      Width           =   525
   End
   Begin VB.TextBox txtempresa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   750
      Width           =   8115
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   24
      Top             =   6045
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   714
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tabela do simples nacional"
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
      Height          =   1005
      Left            =   150
      TabIndex        =   21
      Top             =   4860
      Width           =   5865
      Begin VB.ComboBox Cmb_tipo_TBSN 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmImpostosSN.frx":000C
         Left            =   180
         List            =   "frmImpostosSN.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Tabela do simples nacional."
         Top             =   390
         Width           =   5385
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Faturamento mensal 12 meses"
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
      Height          =   3135
      Left            =   180
      TabIndex        =   18
      Top             =   1590
      Width           =   2775
      Begin MSComctlLib.ListView Lista 
         Height          =   2775
         Left            =   60
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Mês"
            Object.Width           =   971
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Ano"
            Object.Width           =   1324
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4890
      Width           =   1665
   End
   Begin VB.TextBox txtValorDeduzir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3780
      Width           =   2955
   End
   Begin VB.TextBox txtAliqNom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4380
      Width           =   2955
   End
   Begin VB.TextBox txtFaixa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3180
      Width           =   2955
   End
   Begin VB.TextBox txtValorTotalFaturado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2610
      Width           =   2955
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aliquotas a serem aplicadas"
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
      Height          =   3165
      Left            =   6060
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
      Begin VB.TextBox txtISSQN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2820
         Width           =   585
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Left            =   300
         TabIndex        =   29
         Top             =   390
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% IRPJ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% IRPJ"
      End
      Begin VB.TextBox txtCSLL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   585
      End
      Begin VB.TextBox txtCOFINS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1065
         Width           =   585
      End
      Begin VB.TextBox txtPis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1425
         Width           =   585
      End
      Begin VB.TextBox txtCPP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1785
         Width           =   585
      End
      Begin VB.TextBox txtIPI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2130
         Width           =   585
      End
      Begin VB.TextBox txtICMS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2490
         Width           =   585
      End
      Begin VB.TextBox txtIRPJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   280
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   585
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   195
         Left            =   300
         TabIndex        =   30
         Top             =   740
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% CSLL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% CSLL"
      End
      Begin DrawSuite2022.USLabel USLabel3 
         Height          =   195
         Left            =   270
         TabIndex        =   31
         Top             =   1090
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% Cofins"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% Cofins"
      End
      Begin DrawSuite2022.USLabel USLabel4 
         Height          =   195
         Left            =   270
         TabIndex        =   32
         Top             =   1440
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% Pis/Pasep"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% Pis/Pasep"
      End
      Begin DrawSuite2022.USLabel USLabel5 
         Height          =   195
         Left            =   300
         TabIndex        =   33
         Top             =   1790
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% CPP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% CPP"
      End
      Begin DrawSuite2022.USLabel USLabel6 
         Height          =   195
         Left            =   300
         TabIndex        =   34
         Top             =   2140
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% IPI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% IPI"
      End
      Begin DrawSuite2022.USLabel USLabel7 
         Height          =   195
         Left            =   270
         TabIndex        =   35
         Top             =   2490
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% ICMS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% ICMS"
      End
      Begin DrawSuite2022.USLabel USLabel8 
         Height          =   195
         Left            =   270
         TabIndex        =   36
         Top             =   2820
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "% ISSQN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "% ISSQN"
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   714
      DibPicture      =   "frmImpostosSN.frx":0010
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
      Icon            =   "frmImpostosSN.frx":A133
   End
   Begin DrawSuite2022.USButton Cmd_salvar_tabelaSN 
      Height          =   1035
      Left            =   6060
      TabIndex        =   20
      Top             =   4830
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1826
      DibPicture      =   "frmImpostosSN.frx":A44D
      Caption         =   "Salvar tabela do Simples nacional"
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
      Theme           =   4
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Obs: Use F12 para recalcular faturamento."
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
      Height          =   255
      Left            =   210
      TabIndex        =   28
      Top             =   1320
      Width           =   3225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
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
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   26
      Top             =   540
      Width           =   7635
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valor a deduzir"
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
      Height          =   285
      Left            =   3570
      TabIndex        =   17
      Top             =   3600
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota nominal (% DAS)"
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
      Height          =   285
      Left            =   3570
      TabIndex        =   15
      Top             =   4170
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Faixa atual"
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
      Height          =   285
      Left            =   3570
      TabIndex        =   13
      Top             =   2970
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total faturado 12 meses"
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
      Height          =   285
      Index           =   0
      Left            =   3570
      TabIndex        =   4
      Top             =   2400
      Width           =   1875
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   1215
      Left            =   3660
      TabIndex        =   37
      Top             =   1110
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2143
      Image           =   "frmImpostosSN.frx":12E52
      Props           =   5
   End
End
Attribute VB_Name = "frmImpostosSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Contador = 0
valor = 0
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
'StrSql = "Select IFM.ID, IFM.ID_empresa, IFM.Ano, IFM.Mes,Sum(IFM.Valor-ISNULL(FDM.VlrTotalFat,0)) as TotalMes from Impostos_FaturamentoMes IFM LEFT OUTER JOIN Faturamento_12ultimos_meses_devolucao_mensal FDM on IFM.Ano = FDM.Ano and IFM.Mes = FDM.Mes and IFM.ID_empresa = FDM.ID_empresa where IFM.ID_empresa = " & IDEmp & "  group by IFM.ID_empresa,IFM.ID, IFM.Ano, IFM.mes order by Ano, Mes "
StrSql = "Select ID, ID_empresa, Ano, Mes,Valor as TotalMes from Impostos_FaturamentoMes  where ID_empresa = " & IDEmp & "  group by Valor, ID_empresa,ID, Ano, mes order by Ano, Mes "

'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = FunVerificaNumeroMes(IIf(IsNull(TBLISTA!Mes), "", TBLISTA!Mes))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ano), "", TBLISTA!Ano)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!TotalMes), "", Format(TBLISTA!TotalMes, "###,##0.00"))
            valor = valor + TBLISTA!TotalMes
            TBLISTA.MoveNext
            Contador = Contador + 1
        End With
    Loop
    Lista.ListItems.Add , , Contador + 1
    Lista.ListItems.Item(Contador + 1).SubItems(2) = "Total: "
    Lista.ListItems.Item(Contador + 1).SubItems(3) = Format(valor, "###,##0.00")
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_tabelaSN_Click()
On Error GoTo tratar_erro

Acao = "salvar"
If Cmb_tipo_TBSN = "" Then
    NomeCampo = "a tabela"
    ProcVerificaAcao
    Cmb_tipo_TBSN.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente alterar a tabela do simples nacional.", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Select Case Mid(Cmb_tipo_TBSN, 8, 3)
        Case "I -": TabelaSN = 1
        Case "II ": TabelaSN = 2
        Case "III": TabelaSN = 3
        Case "IV ": TabelaSN = 4
        Case "V ": TabelaSN = 5
    End Select
        With frmFaturamento_Prod_Serv
            If FunVerificaRegistroValidado("tbl_Dados_Nota_Fiscal", "ID = " & .txtId, IIf(.txtNFiscal = "", "ordem de faturamento", "nota fiscal"), "a tabela do simples nacional", "alterar", False, True) = False Then Exit Sub
            '.TabelaSN = TabelaSN
            Conexao.Execute "UPDATE tbl_Dados_Nota_Fiscal Set TabelaSN = " & TabelaSN & " where ID = " & .txtId
            .ProcCorrigeValorImpostosSN .txtId
            .ProcVerificaTipoNF False
            If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
            Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
        End With
    USMsgBox ("Tabela do simples nacional alterada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Alterar tabela do simples nacional"
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: If Cmd_salvar_tabelaSN.Enabled = True Then Cmd_salvar_tabelaSN_Click
    Case vbKeyEscape: Unload Me
    Case vbKeyF12: ProcAcertaFaturamento12Meses (IDEmp)
    ProcCarregaLista
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If FunVerifRegimeEmpresa(frmFaturamento_Prod_Serv.txtIDEmpresa.Text) = 1 Or FunVerifRegimeEmpresa(frmFaturamento_Prod_Serv.txtIDEmpresa.Text) = 4 Then
IDEmp = frmFaturamento_Prod_Serv.txtIDEmpresa.Text
txtIDEmpresa.Text = IDEmp
txtEmpresa.Text = frmFaturamento_Prod_Serv.txtEmpresa.Text

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & IDEmp & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        USMsgBox ("Não existe nenhuma tabela do simples nacional ativa."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    
txtValorTotalFaturado.Text = Format(FunVerifVlrTotalFat12UltMesesSomado(IDEmp), "###,##0.00")
ProcVerifImpostosSN (IDEmp), TBAbrir!Tabela, txtValorTotalFaturado.Text

txtIRPJ = Format(IRPJ_Prod, "###,##0.0000")
txtCSLL = Format(CSLL_Prod, "###,##0.0000")
txtCofins = Format(Cofins_Prod, "###,##0.0000")
txtPIS = Format(PIS_Prod, "###,##0.0000")
txtCPP = Format(CPP_Prod, "###,##0.0000")
txtIPI = Format(IPI_SN, "###,##0.0000")
txtICMS = Format(ICMS_SN, "###,##0.0000")
txtISSQN = Format(ISS_Serv, "###,##0.0000")
txtAliqNom = Format(DAS, "###,##0.0000")
txtValorDeduzir = Format(Valor2, "###,##0.0000")
txtFaixa = Var

TBAbrir.Close
End If
Valor2 = 0
Var = ""

ProcCarregaLista
'=====================================
With Cmb_tipo_TBSN
.Clear
'Carrega tabelas do simples cadastradas
Contador = 0
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & IDEmp & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
Do While TBFI.EOF = False
   Select Case TBFI!Tabela
       Case 1: .AddItem "Tabela I - Partilha do Simples Nacional – Comércio"
       Case 2: .AddItem "Tabela II - Partilha do Simples Nacional - Indústria"
       Case 3: .AddItem "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
       Case 4: .AddItem "Tabela IV - Partilha do Simples Nacional - Serviços"
       Case 5: .AddItem "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
   End Select
   
   TabelaSN = TBFI!Tabela
   Contador = Contador + 1
   TBFI.MoveNext
Loop
If Contador = 1 Then
   Select Case TabelaSN
       Case 1: .Text = "Tabela I - Partilha do Simples Nacional – Comércio"
       Case 2: .Text = "Tabela II - Partilha do Simples Nacional - Indústria"
       Case 3: .Text = "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
       Case 4: .Text = "Tabela IV - Partilha do Simples Nacional - Serviços"
       Case 5: .Text = "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
   End Select
   .Locked = True
   .TabStop = False
Else
   .Locked = False
   .TabStop = True
End If
Else
USMsgBox ("Não existe nenhuma tabela do simples nacional ativa, favor verificar."), vbExclamation, "CAPRIND v5.0"
TBFI.Close
Exit Sub
End If
TBFI.Close
End With
'End If
'=============================================================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
