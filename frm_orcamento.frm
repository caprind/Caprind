VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_orcamento 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Orçamento Caprind (Protótipo)"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da empresa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   60
      TabIndex        =   54
      Top             =   1080
      Width           =   5310
      Begin VB.TextBox txtRegime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   3810
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox txtEmpresa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   180
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtidregime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   3810
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtidempresa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   180
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   8
         Left            =   1665
         Top             =   270
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   344
         Caption         =   "Empresa"
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
         NoHTMLCaption   =   "Empresa"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   7
         Left            =   3900
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   344
         Caption         =   "Regime tributário"
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
         NoHTMLCaption   =   "Regime tributário"
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   960
      Left            =   5370
      TabIndex        =   37
      Top             =   1080
      Width           =   9855
      Begin VB.TextBox txtcodproduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   8940
         TabIndex        =   61
         ToolTipText     =   "Fase."
         Top             =   450
         Width           =   825
      End
      Begin VB.TextBox txtNCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   465
         Width           =   825
      End
      Begin DrawSuite2022.USButton btnItem 
         Height          =   285
         Left            =   1065
         TabIndex        =   52
         ToolTipText     =   "Localizar item"
         Top             =   465
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         DibPicture      =   "frm_orcamento.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   ""
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
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin VB.TextBox txtdescricao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   465
         Width           =   4890
      End
      Begin VB.TextBox txtunidade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   465
         Width           =   345
      End
      Begin VB.TextBox txtreferencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   465
         Width           =   1335
      End
      Begin VB.TextBox txtcodigoproduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   465
         Width           =   975
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   2
         Left            =   315
         Top             =   270
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Caption         =   "Código"
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
         NoHTMLCaption   =   "Código"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   3
         Left            =   1740
         Top             =   270
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   344
         Caption         =   "Referência"
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
         NoHTMLCaption   =   "Referência"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   4
         Left            =   2865
         Top             =   270
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   344
         Caption         =   "Un"
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
         NoHTMLCaption   =   "Un"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   5
         Left            =   6495
         Top             =   270
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   "Descrição"
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
         NoHTMLCaption   =   "Descrição"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   9
         Left            =   3435
         Top             =   270
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   344
         Caption         =   "NCM"
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
         NoHTMLCaption   =   "NCM"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote min."
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
         Left            =   9000
         TabIndex        =   62
         Top             =   270
         Width           =   660
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
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
      Height          =   8100
      Left            =   5370
      TabIndex        =   50
      Top             =   2010
      Width           =   9855
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   6900
         Left            =   225
         OleObjectBlob   =   "frm_orcamento.frx":3650
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   360
         Width           =   9330
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   60
      TabIndex        =   42
      Top             =   2010
      Width           =   5310
      Begin VB.TextBox txtresponsavel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   2340
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   315
         Width           =   2865
      End
      Begin VB.TextBox txtdata 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   1365
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   315
         Width           =   945
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   210
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   315
         Width           =   1125
      End
      Begin DrawSuite2022.USLabel USLabel21 
         Height          =   195
         Left            =   270
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         Caption         =   "N° Orçamento"
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
         NoHTMLCaption   =   "N° Orçamento"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   1
         Left            =   1680
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         Caption         =   "Data"
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
         NoHTMLCaption   =   "Data"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   6
         Left            =   3345
         Top             =   120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   344
         Caption         =   "Responsável"
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
         NoHTMLCaption   =   "Responsável"
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1050
         Width           =   345
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calcular..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   60
      TabIndex        =   32
      Top             =   9135
      Width           =   5295
      Begin VB.OptionButton optvalorsugerido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pelo valor de venda"
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
         Left            =   225
         TabIndex        =   11
         Top             =   585
         Width           =   1770
      End
      Begin VB.OptionButton optvalorcalculado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O valor de venda"
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
         Left            =   225
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.TextBox txtvv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2205
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   1590
      End
      Begin DrawSuite2022.USButton cmdCalcular_valores 
         Height          =   675
         Left            =   3900
         TabIndex        =   12
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1191
         DibPicture      =   "frm_orcamento.frx":6FF3
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Executar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total (+)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   29
      Top             =   7410
      Width           =   5295
      Begin VB.TextBox txtttcindireto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   195
         Width           =   1590
      End
      Begin VB.TextBox txtTotalPci 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   195
         Width           =   780
      End
      Begin DrawSuite2022.USLabel USLabel18 
         Height          =   195
         Left            =   4650
         Top             =   240
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total (+)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   28
      Top             =   4665
      Width           =   5295
      Begin VB.TextBox txtttcdireto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   225
         Width           =   1590
      End
      Begin VB.TextBox txtTotalPcd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   225
         Width           =   780
      End
      Begin DrawSuite2022.USLabel USLabel12 
         Height          =   195
         Left            =   4650
         Top             =   300
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resultados finais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   60
      TabIndex        =   24
      Top             =   8025
      Width           =   5295
      Begin VB.TextBox txtmarkup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   270
         Width           =   1590
      End
      Begin VB.TextBox txtct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   1590
      End
      Begin VB.TextBox txtTotalPcustos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   3825
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   630
         Width           =   780
      End
      Begin VB.TextBox txtpMargem 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   3825
         TabIndex        =   9
         Text            =   "20,00"
         Top             =   270
         Width           =   780
      End
      Begin DrawSuite2022.USLabel USLabel6 
         Height          =   195
         Left            =   960
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         Caption         =   "(+) Markup (10):"
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
         NoHTMLCaption   =   "(+) Markup (10):"
      End
      Begin DrawSuite2022.USLabel USLabel7 
         Height          =   195
         Left            =   840
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         Caption         =   "(=) Custo total:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "(=) Custo total:"
      End
      Begin DrawSuite2022.USLabel USLabel19 
         Height          =   195
         Left            =   4650
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel20 
         Height          =   195
         Left            =   4650
         Top             =   690
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custos indiretos "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   60
      TabIndex        =   20
      Top             =   5310
      Width           =   5295
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   0
         Left            =   840
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         Caption         =   "Administrativo (5):"
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
         NoHTMLCaption   =   "Administrativo (5):"
      End
      Begin VB.TextBox txtp9 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   3825
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   1680
         Width           =   780
      End
      Begin VB.TextBox txtv9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1590
      End
      Begin VB.TextBox txtp8 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   3825
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   1320
         Width           =   780
      End
      Begin VB.TextBox txtv8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1590
      End
      Begin VB.TextBox txtv7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   960
         Width           =   1590
      End
      Begin VB.TextBox txtv6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   1590
      End
      Begin VB.TextBox txtV5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2205
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1590
      End
      Begin VB.TextBox txtp7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   960
         Width           =   780
      End
      Begin VB.TextBox txtp6 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   3825
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   600
         Width           =   780
      End
      Begin VB.TextBox txtp5 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   3825
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   240
         Width           =   780
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   195
         Left            =   1110
         Top             =   660
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   344
         Caption         =   "Financeiro (6):"
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
         NoHTMLCaption   =   "Financeiro (6):"
      End
      Begin DrawSuite2022.USLabel USLabel4 
         Height          =   195
         Left            =   1470
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         Caption         =   "Frete (8):"
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
         NoHTMLCaption   =   "Frete (8):"
      End
      Begin DrawSuite2022.USLabel USLabel5 
         Height          =   195
         Left            =   1200
         Top             =   1710
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         Caption         =   "Comissão (9):"
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
         NoHTMLCaption   =   "Comissão (9):"
      End
      Begin DrawSuite2022.USLabel USLabel13 
         Height          =   195
         Left            =   4650
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel14 
         Height          =   195
         Left            =   4650
         Top             =   660
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel16 
         Height          =   195
         Left            =   4650
         Top             =   1380
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel17 
         Height          =   195
         Left            =   4650
         Top             =   1740
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USButton btnImpostos 
         Height          =   285
         Left            =   4620
         TabIndex        =   53
         Top             =   960
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USLabel USLabel3 
         Height          =   195
         Left            =   1200
         Top             =   1020
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   344
         Caption         =   "Impostos (7):"
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
         NoHTMLCaption   =   "Impostos (7):"
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custos diretos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   60
      TabIndex        =   15
      Top             =   2715
      Width           =   5295
      Begin DrawSuite2022.USButton btnMO 
         Height          =   285
         Left            =   510
         TabIndex        =   45
         Top             =   330
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         DibPicture      =   "frm_orcamento.frx":E173
         Alignment       =   2
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Mão de obra (1):"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin VB.TextBox txtv1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2205
         TabIndex        =   0
         Top             =   315
         Width           =   1590
      End
      Begin VB.TextBox txtp1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   315
         Width           =   780
      End
      Begin VB.TextBox txtv4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         TabIndex        =   3
         Top             =   1530
         Width           =   1590
      End
      Begin VB.TextBox txtv3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         TabIndex        =   2
         Top             =   1125
         Width           =   1590
      End
      Begin VB.TextBox txtv2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         TabIndex        =   1
         Top             =   720
         Width           =   1590
      End
      Begin VB.TextBox txtp4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1530
         Width           =   780
      End
      Begin VB.TextBox txtp3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1125
         Width           =   780
      End
      Begin VB.TextBox txtp2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   720
         Width           =   780
      End
      Begin DrawSuite2022.USLabel USLabel8 
         Height          =   195
         Left            =   4650
         Top             =   360
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel9 
         Height          =   195
         Left            =   4650
         Top             =   780
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel10 
         Height          =   195
         Left            =   4650
         Top             =   1170
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel11 
         Height          =   195
         Left            =   4650
         Top             =   1620
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USButton btnMteriais 
         Height          =   285
         Left            =   510
         TabIndex        =   46
         ToolTipText     =   "Cadastrar materiais"
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         DibPicture      =   "frm_orcamento.frx":FED0
         Alignment       =   2
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Materiais (2):"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton btnTerceiros 
         Height          =   285
         Left            =   510
         TabIndex        =   47
         Top             =   1140
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         DibPicture      =   "frm_orcamento.frx":11B7D
         Alignment       =   2
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Terceiros (3):"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton brnOutros 
         Height          =   285
         Left            =   510
         TabIndex        =   48
         Top             =   1530
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         DibPicture      =   "frm_orcamento.frx":133D1
         Alignment       =   2
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Outros (4):"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar2 
      Height          =   1065
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   1879
      ButtonCount     =   10
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
      ButtonCaption2  =   "Filtrar"
      ButtonEnabled2  =   0   'False
      ButtonToolTipText2=   "Filtrar registros"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   42
      ButtonHeight2   =   24
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Salvar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Salvar (F3)"
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
      ButtonLeft3     =   81
      ButtonTop3      =   2
      ButtonWidth3    =   38
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Relatório"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Relatório (F5)"
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
      ButtonLeft4     =   121
      ButtonTop4      =   2
      ButtonWidth4    =   51
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Anterior"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Registro anterior."
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
      ButtonLeft5     =   174
      ButtonTop5      =   2
      ButtonWidth5    =   47
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Próximo"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Próximo registro."
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
      ButtonLeft6     =   223
      ButtonTop6      =   2
      ButtonWidth6    =   46
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonAlignment7=   2
      ButtonType7     =   1
      ButtonStyle7    =   -1
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   -1
      ButtonLeft7     =   271
      ButtonTop7      =   4
      ButtonWidth7    =   2
      ButtonHeight7   =   60
      ButtonCaption8  =   "Ajuda"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Ajuda (F1)"
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
      ButtonLeft8     =   275
      ButtonTop8      =   2
      ButtonWidth8    =   36
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Sair"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Sair (Esc)"
      ButtonKey9      =   "9"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft9     =   313
      ButtonTop9      =   2
      ButtonWidth9    =   26
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonKey10     =   "10"
      ButtonAlignment10=   2
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState10   =   5
      ButtonLeft10    =   341
      ButtonTop10     =   2
      ButtonWidth10   =   24
      ButtonHeight10  =   24
      ButtonUseMaskColor10=   0   'False
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
         FormHeightDT    =   10755
         FormWidthDT     =   15480
         FormScaleHeightDT=   10290
         FormScaleWidthDT=   15360
         ResizeFormBackground=   -1  'True
         ResizePictureBoxContents=   -1  'True
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   10560
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_orcamento.frx":2DD7D
         Count           =   1
      End
   End
End
Attribute VB_Name = "frm_orcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public percentual1 As Currency
Public percentual2 As Currency
Public percentual3 As Currency
Public percentual4 As Currency
Public percentual5 As Currency
Public percentual6 As Currency
Public percentual7 As Currency
Public percentual8 As Currency
Public percentual9 As Currency
Public percentual10 As Currency


Public Valor1 As Currency
Public Valor2 As Currency
Public Valor3 As Currency
Public Valor4 As Currency
Public Valor5 As Currency
Public Valor6 As Currency
Public Valor7 As Currency
Public Valor8 As Currency
Public Valor9 As Currency
Public ValorMargem As Currency

Public SomaTotal As Currency
Public TotalVenda As Currency
Public TotalImposto As Currency
Public TotalCusto As Currency
Public TotalMargem As Currency
Public TotalPercentual As Currency

Public Sub ProcSomaTotalPercentualCustosDiretos()
On Error GoTo tratar_erro
'Usando

    percentual1 = IIf(txtp1.Text <> "", txtp1.Text, 0)
    percentual2 = IIf(txtp2.Text <> "", txtp2.Text, 0)
    percentual3 = IIf(txtp3.Text <> "", txtp3.Text, 0)
    percentual4 = IIf(txtp4.Text <> "", txtp4.Text, 0)


txtTotalPcd.Text = percentual1 + percentual2 + percentual3 + percentual4
txtTotalPcd.Text = Format(Round(txtTotalPcd, 2), "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcSomaTotalPercentualCustosIndiretos()
On Error GoTo tratar_erro
'Usando

    percentual5 = IIf(txtp5.Text <> "", txtp5.Text, 0)
    percentual6 = IIf(txtp6.Text <> "", txtp6.Text, 0)
    percentual7 = IIf(txtp7.Text <> "", txtp7.Text, 0)
    percentual8 = IIf(txtp8.Text <> "", txtp8.Text, 0)
    percentual9 = IIf(txtp9.Text <> "", txtp9.Text, 0)


txtTotalPci.Text = percentual5 + percentual6 + percentual7 + percentual8 + percentual9



txtTotalPci.Text = Format(Round(txtTotalPci, 2), "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcSomaCustosDiretos()
On Error GoTo tratar_erro
'Usando

Valor1 = IIf(txtv1 <> "", txtv1, 0)
Valor2 = IIf(txtv2 <> "", txtv2, 0)
Valor3 = IIf(txtv3 <> "", txtv3, 0)
Valor4 = IIf(txtv4 <> "", txtv4, 0)


txtttcdireto.Text = Valor1 + Valor2 + Valor3 + Valor4
txtttcdireto.Text = Format(Round(txtttcdireto, 2), "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcSomacustosIndiretos()
On Error GoTo tratar_erro
'Usando

txtttcindireto.Text = Valor5 + Valor6 + Valor7 + Valor8 + Valor9
txtttcindireto.Text = Format(Round(txtttcindireto, 2), "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcSomapercentuais()
On Error GoTo tratar_erro
'Usando

ProcSomaTotalPercentualCustosDiretos
ProcSomaTotalPercentualCustosIndiretos

    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCalculavalor1()
On Error GoTo tratar_erro
    
   ' ProcCarregaValores
   ' ProcCalculavalorvenda
   ' ProcCalculaImpostos
   ' ProcCalculaPercentual
   ' ProcRedesenhaGrafico
   ' ProcSomapercentuais

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcRedesenhaGrafico()
On Error GoTo tratar_erro
percentual10 = txtpMargem

    With MSChart1
     .ShowLegend = True
        For i = 1 To 10
            .Column = i
            valor = i
                Select Case i
                    Case 1: .Data = txtp1
                        .ColumnLabel = "Mão de obra " & Format(Round(txtp1, 2), "###,##0.00") & " %"
                    Case 2: .Data = txtp2
                        .ColumnLabel = "Material " & Format(Round(txtp2, 2), "###,##0.00") & " %"
                    Case 3: .Data = txtp3
                        .ColumnLabel = "Terceiro " & Format(Round(txtp3, 2), "###,##0.00") & " %"
                    Case 4: .Data = txtp4
                        .ColumnLabel = "Outros " & Format(Round(txtp4, 4), "###,##0.00") & " %"
                    Case 5: .Data = txtp5
                        .ColumnLabel = "Administrativo " & Format(Round(txtp5, 2), "###,##0.00") & " %"
                    Case 6: .Data = txtp6
                        .ColumnLabel = "Financeiro " & Format(Round(txtp6, 2), "###,##0.00") & " %"
                    Case 7: .Data = txtp7
                        .ColumnLabel = "Impostos " & Format(Round(txtp7, 2), "###,##0.00") & " %"
                    Case 8: .Data = txtp8
                        .ColumnLabel = "Frete " & Format(Round(txtp8, 2), "###,##0.00") & " %"
                    Case 9: .Data = txtp9
                        .ColumnLabel = "Comissão " & Format(Round(txtp9, 2), "###,##0.00") & " %"
                    Case 10: .Data = txtpMargem.Text
                        .ColumnLabel = "Markup " & Format(Round(txtpMargem, 2), "###,##0.00") & " %"
                End Select
        Next
        .Refresh
    End With
'frm_orcamento.Refresh

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCalculavalorvenda()
On Error GoTo tratar_erro

Dim ValorCustoDireto As Double
Dim pMargem As Double
Dim PCustosIndiretos As Double
Dim PCalculo As Double
Dim Imposto As Double
Dim ValorVenda As Double

'Carrega o valor total de custo na variavel
    ValorCustoDireto = txtttcdireto

'Carrega o valor percentual e da margem de lucro
    pMargem = txtpMargem
    PCustosIndiretos = IIf(txtTotalPci.Text <> "", txtTotalPci.Text, 0)

   PCalculo = pMargem + PCustosIndiretos

'Calcula o valor da margem de lucro
    ValorVenda = ValorCustoDireto / (1 - (PCalculo / 100))


'Carrega o valor de venda
    txtvv.Text = Format(Round(ValorVenda, 2), "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCalculaCustosindiretos()
On Error GoTo tratar_erro

    
    TotalVenda = txtvv

'Calcula os valores custos indiretos

    Valor5 = TotalVenda - (TotalVenda * (1 - (percentual5 / 100)))
    Valor6 = TotalVenda - (TotalVenda * (1 - (percentual6 / 100)))
    Valor7 = TotalVenda - (TotalVenda * (1 - (percentual7 / 100)))
    Valor8 = TotalVenda - (TotalVenda * (1 - (percentual8 / 100)))
    Valor9 = TotalVenda - (TotalVenda * (1 - (percentual9 / 100)))

'Carrega valores custos indiretos
    txtV5 = Format(Round(Valor5, 2), "###,##0.00")
    txtv6 = Format(Round(Valor6, 2), "###,##0.00")
    txtv7 = Format(Round(Valor7, 2), "###,##0.00")
    txtv8 = Format(Round(Valor8, 2), "###,##0.00")
    txtv9 = Format(Round(Valor9, 2), "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCalculaValorMargem()
On Error GoTo tratar_erro
Dim TotalCustoDireto As Currency
Dim TotalCustoIndireto As Currency
Dim TotalPcustoDireto As Currency
Dim TotalPcustoIndireto As Currency
Dim pMargem As Currency


'Carrega o valor do lucro
    TotalCustoDireto = txtttcdireto
    TotalCustoIndireto = txtttcindireto
    
    TotalCusto = TotalCustoDireto + TotalCustoIndireto
    TotalMargem = TotalVenda - TotalCusto
    txtmarkup = Format(Round(TotalMargem, 2), "###,##0.00")

'Carrega valor total de custos
    txtct = TotalCustoDireto + TotalCustoIndireto + TotalMargem
    txtct = Format(Round(txtct, 2), "###,##0.00")
    
'Carrega valor total de porcentagem
    TotalPcustoDireto = txtTotalPcd
    TotalPcustoIndireto = txtTotalPci
    pMargem = txtpMargem

    txtTotalPcustos = TotalPcustoDireto + TotalPcustoIndireto + pMargem
    
    If txtTotalPcustos.Text > 100 And Optcustos = True Then
        txtpMargem.Text = txtTotalPcustos.Text - 100
        txtTotalPcustos.Text = 100
    End If
    
    If txtTotalPcustos.Text < 100 And Optcustos = True Then
           txtpMargem.Text = 100 - (TotalPcustoDireto + TotalPcustoIndireto)
    End If
    
    
    txtTotalPcustos = Format(Round(txtTotalPcustos, 2), "###,##0.00")
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCalculaPercentual()
On Error GoTo tratar_erro

'Calcula os percentuais de calculo e carrega o valor de custo

'Calcula o valor dos percentuais
    If Valor1 > 0 Then percentual1 = (Valor1 / TotalVenda) * 100
    If Valor2 > 0 Then percentual2 = (Valor2 / TotalVenda) * 100
    If Valor3 > 0 Then percentual3 = (Valor3 / TotalVenda) * 100
    If Valor4 > 0 Then percentual4 = (Valor4 / TotalVenda) * 100
    

'Carrega percentuais
    If percentual1 > 0 Then txtp1 = Format(Round(percentual1, 2), "###,##0.00")
    If percentual2 > 0 Then txtp2 = Format(Round(percentual2, 2), "###,##0.00")
    If percentual3 > 0 Then txtp3 = Format(Round(percentual3, 2), "###,##0.00")
    If percentual4 > 0 Then txtp4 = Format(Round(percentual4, 2), "###,##0.00")


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub brnOutros_Click()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnImpostos_Click()
On Error GoTo tratar_erro

If txtidregime.Text = 1 Then
    frmImpostosSN.Show 1

Else
    frmImpostosLPLR.Show 1
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnItem_Click()
On Error GoTo tratar_erro

frm_Orcamento_item.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnMO_Click()
On Error GoTo tratar_erro

If txtID <> "" Then
frm_Orcamento_Processo.Show 1
Else
USMsgBox "Favor salvar o orçamento antes de iniciar o cadastro!", vbInformation, "CAPRIND V5.0"
Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnMteriais_Click()
On Error GoTo tratar_erro

If txtID <> "" Then
frm_Orcamento_Conjunto.Show 1
Else
USMsgBox "Favor salvar o orçamento antes de iniciar o cadastro!", vbInformation, "CAPRIND V5.0"
Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnTerceiros_Click()
On Error GoTo tratar_erro

If txtID <> "" Then
    frm_Orcamento_Terceiros.Show 1
Else
    USMsgBox "Favor salvar o orçamento antes de iniciar o cadastro!", vbInformation, "CAPRIND V5.0"
    Exit Sub
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdCalcular_valores_Click()
On Error GoTo tratar_erro

Contador = 4
Do While Contador > 0

    ProcSomapercentuais 'Soma percentuais custos indiretos + margem
    ProcSomaCustosDiretos 'Soma os custos diretos
    If optvalorcalculado.Value = True Then
    ProcCalculavalorvenda 'Calcula valor de venda
    End If
    ProcCalculaCustosindiretos 'Calcula valores dos custos indiretos em cima do valor de venda
    ProcCalculaPercentual 'Calcula percentuais custos diretos
    ProcSomaTotalPercentualCustosDiretos 'Soma total percentual custos diretos
    ProcSomacustosIndiretos 'Soma total custos indiretos
    ProcCalculaValorMargem 'Calcula valor margem de lucro
    ProcRedesenhaGrafico 'Redesenha gráfico de percentuais
Contador = Contador - 1
Loop



Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro

    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
'TEMACAPRIND = "Mxs2.skn"

'If TEMACAPRIND <> "" Then
'    frm_orcamento.Skin1.LoadSkin TEMACAPRIND
'    frm_orcamento.Skin1.ApplySkin Me.hWnd
'End If

ProcCarregaToolBar2 Me, 16495, 10, True

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
IDempresa = TBAliquota!CODIGO
txtidempresa.Text = TBAliquota!CODIGO
txtEmpresa.Text = TBAliquota!Empresa

    If TBAliquota!Simples = True Then
    txtRegime.Text = "Simples nacional"
    txtidregime.Text = 1
    End If
    
    If TBAliquota!Presumido = True Then
    txtRegime.Text = "Lucro presumido"
    txtidregime.Text = 2
    End If
    
    If TBAliquota!Real = True Then
    txtRegime.Text = "Lucro real"
    txtidregime.Text = 3
    End If
    
    If TBAliquota!Simples1 = True Then
    txtRegime.Text = "Simples|Exc. sub lim."
    txtidregime.Text = 4
    End If

End If
TBAliquota.Close
   
   
    With MSChart1
     .ShowLegend = True
        For i = 1 To 10
            .Column = i
            valor = i
                Select Case i
                    Case 1: .Data = txtp1
                        .ColumnLabel = "Mão de obra"
                    Case 2: .Data = txtp2
                        .ColumnLabel = "Material"
                    Case 3: .Data = txtp3
                        .ColumnLabel = "Terceiro"
                    Case 4: .Data = txtp4
                        .ColumnLabel = "Outros"
                    Case 5: .Data = txtp5
                        .ColumnLabel = "Administrativo"
                    Case 6: .Data = txtp6
                        .ColumnLabel = "Financeiro"
                    Case 7: .Data = txtp7
                        .ColumnLabel = "Impostos"
                    Case 8: .Data = txtp8
                        .ColumnLabel = "Frete"
                    Case 9: .Data = txtp9
                        .ColumnLabel = "comissão"
                    Case 10: .Data = txtpMargem
                        .ColumnLabel = "Markup"
                End Select
        Next
        .Refresh
    End With
txtv1.Text = Format(0, "###,##0.00")
txtv2.Text = Format(0, "###,##0.00")
txtv3.Text = Format(0, "###,##0.00")
txtv4.Text = Format(0, "###,##0.00")
txtp1.Text = Format(0, "###,##0.00")
txtp2.Text = Format(0, "###,##0.00")
txtp3.Text = Format(0, "###,##0.00")
txtp4.Text = Format(0, "###,##0.00")
txtttcdireto = Format(0, "###,##0.00")
txtttcindireto = Format(0, "###,##0.00")
txtTotalPcd = Format(0, "###,##0.00")
txtTotalPci = Format(0, "###,##0.00")
txtmarkup = Format(0, "###,##0.00")
txtTotalPcustos = Format(0, "###,##0.00")
txtV5 = Format(0, "###,##0.00")
txtv6 = Format(0, "###,##0.00")
txtv7 = Format(0, "###,##0.00")
txtv8 = Format(0, "###,##0.00")
txtv9 = Format(0, "###,##0.00")
txtvv = Format(0, "###,##0.00")
txtct = Format(0, "###,##0.00")


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



Private Sub ProcBuscaTotaisOrcamento()
On Error GoTo tratar_erro


Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select sum(vlrTotal) as v1 from Vendas_Orcamento_Fases where ID_Orcamento = '" & txtID & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtv1.Text = Format(TBAbrir!V1, "###,##0.00")
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

'ProcBuscaTotaisOrcamento

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtID_Change()
On Error GoTo tratar_erro

If IsNumeric(txtID) Then
 MSChart1.Visible = True
Else
 MSChart1.Visible = False
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp5_GotFocus()
On Error GoTo tratar_erro

txtp5.BackColor = &HC0FFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp5_LostFocus()
On Error GoTo tratar_erro

If txtp5.Text = "" Then txtp5.Text = 0
txtp5.Text = Format(Round(txtp5.Text, 2), "###,##0.00")
txtp5.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub txtp6_GotFocus()
On Error GoTo tratar_erro


txtp6.BackColor = &HC0FFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp6_LostFocus()
On Error GoTo tratar_erro


If txtp6.Text = "" Then txtp6.Text = 0
txtp6.Text = Format(Round(txtp6.Text, 2), "###,##0.00")
txtp6.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp7_GotFocus()
On Error GoTo tratar_erro


txtp7.BackColor = &HC0FFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp7_LostFocus()
On Error GoTo tratar_erro


If txtp7.Text = "" Then txtp7.Text = 0
txtp7.Text = Format(Round(txtp7.Text, 2), "###,##0.00")
txtp7.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp8_GotFocus()
On Error GoTo tratar_erro


txtp8.BackColor = &HC0FFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp8_LostFocus()
On Error GoTo tratar_erro


If txtp8.Text = "" Then txtp8.Text = 0
txtp8.Text = Format(Round(txtp8.Text, 2), "###,##0.00")
txtp8.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp9_GotFocus()
On Error GoTo tratar_erro


txtp9.BackColor = &HC0FFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtp9_LostFocus()
On Error GoTo tratar_erro


If txtp9.Text = "" Then txtp9.Text = 0
txtp9.Text = Format(Round(txtp9.Text, 2), "###,##0.00")
txtp9.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



Private Sub txtpMargem_GotFocus()
On Error GoTo tratar_erro


txtpMargem.BackColor = &HC0FFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtpMargem_LostFocus()
On Error GoTo tratar_erro


If txtpMargem.Text = "" Then txtpMargem.Text = 0
txtpMargem.Text = Format(Round(txtpMargem.Text, 2), "###,##0.00")
txtpMargem.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv1_GotFocus()
On Error GoTo tratar_erro


txtv1.BackColor = &HC0FFFF
'txtv1.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv1_LostFocus()
On Error GoTo tratar_erro


If txtv1.Text = "" Then txtv1.Text = 0
txtv1.Text = Format(Round(txtv1.Text, 2), "###,##0.00")
txtv1.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv2_GotFocus()
On Error GoTo tratar_erro

txtv2.BackColor = &HC0FFFF
'txtv2.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv2_LostFocus()
On Error GoTo tratar_erro

If txtv2.Text = "" Then txtv2.Text = 0
txtv2.Text = Format(Round(txtv2.Text, 2), "###,##0.00")
txtv2.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv3_GotFocus()
On Error GoTo tratar_erro

txtv3.BackColor = &HC0FFFF
'txtv3.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv3_LostFocus()
On Error GoTo tratar_erro

If txtv3.Text = "" Then txtv3.Text = 0
txtv3.Text = Format(Round(txtv3.Text, 2), "###,##0.00")
txtv3.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv4_GotFocus()
On Error GoTo tratar_erro

txtv4.BackColor = &HC0FFFF
'txtv4.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtv4_LostFocus()
On Error GoTo tratar_erro

If txtv4.Text = "" Then txtv4.Text = 0
txtv4.Text = Format(Round(txtv4.Text, 2), "###,##0.00")
txtv4.BackColor = &HFFFFFF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



Private Sub txtvv_GotFocus()
On Error GoTo tratar_erro

txtvv.BackColor = &HC0FFFF


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtvv_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
    'If KeyAscii = 13 Then
    '    ProcCalculaValores
    '    ProcRedesenhaGrafico
    'End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtvv_LostFocus()
On Error GoTo tratar_erro

txtvv.BackColor = &HFFFFFF
txtvv.Text = Format(Round(txtvv.Text, 2), "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub




Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
'    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
'    Case 7: ProcAtualizar
'    Case 10: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtCodigo = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Vendas_orcamento order by ID_Orcamento", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Codigo = " & txtCodigo & "")
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        IDlista = TBLISTA!ID_orcamento
        ProcCarregaorcamento
    Else
        USMsgBox ("Fim dos cadastros de orçamentos."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Desconto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtCodigo = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Vendas_orcamento order by ID_orcamento", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Codigo = " & txtCodigo & "")
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        IDlista = TBLISTA!ID_orcamento
        ProcCarregaorcamento
    Else
        USMsgBox ("Fim dos cadastros de orçamentos."), vbInformation, "CAPRIND v5.0"
    End If
End If

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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frm_Orcamento_abrir.Show 1
If IDlista <> 0 Then
ProcCarregaorcamento
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaorcamento()
On Error GoTo tratar_erro

If IDlista <> 0 Then

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_orcamento where ID_Orcamento = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
txtID.Text = TBLISTA!ID_orcamento
txtCodigo.Text = TBLISTA!CODIGO
txtData.Text = Format(TBLISTA!Data, "dd/mm/yyyy")
txtResponsavel.Text = TBLISTA!Responsavel
txtcodproduto.Text = IIf(IsNull(TBLISTA!Codproduto), 0, TBLISTA!Codproduto)
txtcodigoproduto.Text = TBLISTA!txtcodigoproduto
txtreferencia.Text = TBLISTA!txtreferencia
txtunidade.Text = TBLISTA!txtunidade
txtNCM.Text = IIf(IsNull(TBLISTA!NCM), "", TBLISTA!NCM)
txtdescricao.Text = IIf(IsNull(TBLISTA!txtdescricao), "", TBLISTA!txtdescricao)
txtLote.Text = Format(TBLISTA!Loteminimo, "###,##0.00")

txtv1.Text = Format(TBLISTA!V1, "###,##0.00")
txtp1.Text = Format(TBLISTA!p1, "###,##0.00")
txtv2.Text = Format(TBLISTA!V2, "###,##0.00")
txtp2.Text = Format(TBLISTA!p2, "###,##0.00")
txtv3.Text = Format(TBLISTA!V3, "###,##0.00")
txtp3.Text = Format(TBLISTA!p3, "###,##0.00")
txtv4.Text = Format(TBLISTA!V4, "###,##0.00")
txtp4.Text = Format(TBLISTA!p4, "###,##0.00")

txtttcdireto.Text = Format(TBLISTA!cd, "###,##0.00")
txtTotalPcd.Text = Format(TBLISTA!pcd, "###,##0.00")

txtV5.Text = Format(TBLISTA!V5, "###,##0.00")
txtp5.Text = Format(TBLISTA!p5, "###,##0.00")
txtv6.Text = Format(TBLISTA!V6, "###,##0.00")
txtp6.Text = Format(TBLISTA!p6, "###,##0.00")
txtv7.Text = Format(TBLISTA!V7, "###,##0.00")
txtp7.Text = Format(TBLISTA!p7, "###,##0.00")
txtv8.Text = Format(TBLISTA!V8, "###,##0.00")
txtp8.Text = Format(TBLISTA!p8, "###,##0.00")
txtv9.Text = Format(TBLISTA!V9, "###,##0.00")
txtp9.Text = Format(TBLISTA!p9, "###,##0.00")


txtttcindireto.Text = Format(TBLISTA!ci, "###,##0.00")
txtTotalPci.Text = Format(TBLISTA!pci, "###,##0.00")

txtmarkup.Text = Format(TBLISTA!mkp, "###,##0.00")
txtpMargem.Text = Format(TBLISTA!pmkp, "###,##0.00")

txtct.Text = Format(TBLISTA!CT, "###,##0.00")
txtTotalPcustos.Text = Format(TBLISTA!pct, "###,##0.00")

txtvv.Text = Format(TBLISTA!vv, "###,##0.00")

optvalorcalculado.Value = TBLISTA!Valorcalculado
optvalorsugerido.Value = TBLISTA!ValorSugerido


'USMsgBox "Dados carregados com sucesso!", vbInformation, "CAPRIND v5.0"
ProcRedesenhaGrafico 'Redesenha gráfico de percentuais

End If
TBLISTA.Close
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

txtID.Text = ""
txtv1.Text = "0,00"
txtv2.Text = "0,00"
txtv3.Text = "0,00"
txtv4.Text = "0,00"
txtttcdireto.Text = "0,00"

txtp5.Text = "0,00"
txtp6.Text = "0,00"
txtp7.Text = "0,00"
txtp8.Text = "0,00"
txtp9.Text = "0,00"
txtttcindireto.Text = "0,00"
txtTotalPci.Text = "0,00"
txtmarkup.Text = "0,00"

txtTotalPcustos.Text = "0,00"
txtct.Text = "0,00"

txtvv.Text = "0,00"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimparCampos
frm_Orcamento_item.Show 1

If txtcodproduto.Text <> "" Then
procCriaCodigo
Else
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub procCriaCodigo()
On Error GoTo tratar_erro
Dim CODIGO As String

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_orcamento where year(Data) = year('" & Date & "') order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
TBLISTA.MoveLast

CODIGO = TBLISTA!CODIGO
CODIGO = Left(CODIGO, 7)
CODIGO = Int(CODIGO)

CODIGO = Int(CODIGO) + 1
    Select Case Len(CODIGO)
        Case 1: CODIGO = "000000" & CODIGO
        Case 2: CODIGO = "00000" & CODIGO
        Case 3: CODIGO = "0000" & CODIGO
        Case 4: CODIGO = "000" & CODIGO
        Case 5: CODIGO = "00" & CODIGO
        Case 6: CODIGO = "0" & CODIGO
    End Select
    Ano = Right(Year(Date), 2)
CODIGO = CODIGO & "/" & Right(Year(Date), 2)
Else
    CODIGO = "0000001" & "/" & Right(Year(Date), 2)
End If
TBLISTA.Close
txtCodigo.Text = CODIGO
txtData.Text = Date
txtResponsavel.Text = pubUsuario

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
Dim CODIGO As String

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_orcamento where codigo = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
TBLISTA.AddNew
End If


TBLISTA!CODIGO = txtCodigo.Text
TBLISTA!Data = txtData.Text
TBLISTA!Responsavel = txtResponsavel.Text

TBLISTA!txtcodigoproduto = txtcodigoproduto.Text
TBLISTA!Codproduto = txtcodproduto.Text

TBLISTA!txtreferencia = txtreferencia.Text
TBLISTA!txtunidade = txtunidade.Text
TBLISTA!NCM = txtNCM.Text

TBLISTA!txtdescricao = txtdescricao.Text
TBLISTA!Loteminimo = IIf(txtLote.Text <> "", txtLote.Text, 0)

TBLISTA!V1 = txtv1.Text
TBLISTA!p1 = txtp1.Text
TBLISTA!V2 = txtv2.Text
TBLISTA!p2 = txtp2.Text
TBLISTA!V3 = txtv3.Text
TBLISTA!p3 = txtp3.Text
TBLISTA!V4 = txtv4.Text
TBLISTA!p4 = txtp4.Text

TBLISTA!cd = txtttcdireto.Text
TBLISTA!pcd = txtTotalPcd.Text

TBLISTA!V5 = txtV5.Text
TBLISTA!p5 = txtp5.Text
TBLISTA!V6 = txtv6.Text
TBLISTA!p6 = txtp6.Text
TBLISTA!V7 = txtv7.Text
TBLISTA!p7 = txtp7.Text
TBLISTA!V8 = txtv8.Text
TBLISTA!p8 = txtp8.Text
TBLISTA!V9 = txtv9.Text
TBLISTA!p9 = txtp9.Text


TBLISTA!ci = txtttcindireto.Text
TBLISTA!pci = txtTotalPci.Text

TBLISTA!mkp = txtmarkup.Text
TBLISTA!pmkp = txtpMargem.Text

TBLISTA!CT = txtct.Text
TBLISTA!pct = txtTotalPcustos.Text

TBLISTA!vv = txtvv.Text

TBLISTA!Valorcalculado = optvalorcalculado.Value
TBLISTA!ValorSugerido = optvalorsugerido.Value
TBLISTA.Update

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"

TBLISTA.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

