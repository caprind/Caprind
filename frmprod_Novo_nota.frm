VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmprod_Novo_Nota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Gerenciamento de ordem - Empenhar"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9900
   ClipControls    =   0   'False
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
   ScaleHeight     =   7035
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   58
      Top             =   6630
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   714
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por"
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   210
      TabIndex        =   54
      Top             =   600
      Width           =   9495
      Begin VB.ComboBox Cmb_RE 
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
         Left            =   1980
         TabIndex        =   60
         ToolTipText     =   "Número da rastreabilidade do estoque."
         Top             =   345
         Width           =   1275
      End
      Begin VB.ComboBox cmb_Lote 
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
         Left            =   4350
         Sorted          =   -1  'True
         TabIndex        =   59
         ToolTipText     =   "Numero do lote no estoque"
         Top             =   345
         Width           =   1215
      End
      Begin VB.OptionButton optLote 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lote"
         Height          =   255
         Left            =   3660
         TabIndex        =   56
         Top             =   390
         Width           =   705
      End
      Begin VB.OptionButton optRE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RE (Fifo)"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   900
         TabIndex        =   55
         Top             =   390
         Value           =   -1  'True
         Width           =   1335
      End
      Begin DrawSuite2022.USButton btnRE 
         Height          =   480
         Left            =   7680
         TabIndex        =   61
         ToolTipText     =   "Carregar RE"
         Top             =   210
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   847
         DibPicture      =   "frmprod_Novo_nota.frx":0000
         Caption         =   "Filtrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do item na requisição"
      Height          =   1095
      Left            =   210
      TabIndex        =   46
      Top             =   1470
      Width           =   9495
      Begin VB.TextBox txtIDEmpresa 
         Height          =   360
         Left            =   8370
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   150
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Data."
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox txtIDLista 
         Alignment       =   2  'Center
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   540
         Width           =   420
      End
      Begin VB.TextBox Txt_un 
         Alignment       =   2  'Center
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   540
         Width           =   400
      End
      Begin VB.TextBox txtDescricao 
         Alignment       =   2  'Center
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Local de armazenamento."
         Top             =   540
         Width           =   7170
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un"
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
         Left            =   1845
         TabIndex        =   51
         Top             =   330
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Index           =   0
         Left            =   5400
         TabIndex        =   50
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
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
         Left            =   690
         TabIndex        =   49
         Top             =   330
         Width           =   495
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   873
      DibPicture      =   "frmprod_Novo_nota.frx":3650
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
      Icon            =   "frmprod_Novo_nota.frx":D7FD
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados para empenho do item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   210
      TabIndex        =   32
      Top             =   5250
      Width           =   9510
      Begin DrawSuite2022.USButton btnEmpenhar 
         Height          =   825
         Left            =   7620
         TabIndex        =   45
         Top             =   210
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1455
         DibPicture      =   "frmprod_Novo_nota.frx":DB17
         Caption         =   "Empenhar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.TextBox Txt_qtde_req_PC 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1325
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peças requisitada."
         Top             =   1635
         Width           =   1125
      End
      Begin VB.TextBox txtQtde_aretirar_PC 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5920
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Necessidade de peças."
         Top             =   1635
         Width           =   1125
      End
      Begin VB.TextBox txtQtde_retirada_PC 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3620
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peças retirada."
         Top             =   1635
         Width           =   1140
      End
      Begin VB.TextBox txtQtde_Retirar_PC 
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
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peças a empenhar."
         Top             =   1635
         Width           =   1125
      End
      Begin VB.TextBox Txt_qtde_req 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade requisitada."
         Top             =   495
         Width           =   1515
      End
      Begin VB.TextBox txtQtde_aretirar 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Necessidade."
         Top             =   495
         Width           =   1515
      End
      Begin VB.TextBox txtQtde_retirada 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade retirada."
         Top             =   495
         Width           =   1515
      End
      Begin VB.TextBox txtQtde_Retirar 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5325
         TabIndex        =   19
         ToolTipText     =   "Quantidade a empenhar."
         Top             =   375
         Width           =   2205
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. req. PÇ"
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
         Left            =   1365
         TabIndex        =   43
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Necess. PÇ"
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
         Left            =   6075
         TabIndex        =   42
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retirado"
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
         Left            =   2265
         TabIndex        =   41
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qt. emp. PÇ"
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
         Left            =   8265
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisitado"
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
         Left            =   510
         TabIndex        =   36
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Necessidade"
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
         Left            =   3780
         TabIndex        =   35
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. ret. PÇ"
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
         Left            =   3690
         TabIndex        =   34
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. Empenhar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5715
         TabIndex        =   33
         Top             =   150
         Width           =   1455
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7860
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmprod_Novo_nota.frx":17CC4
      Count           =   1
   End
   Begin VB.Frame frameLote 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da RE | Lote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   210
      TabIndex        =   21
      Top             =   2670
      Width           =   9510
      Begin VB.TextBox Txt_est_disp_PC 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7845
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Estoque de peça disponível."
         Top             =   3375
         Width           =   1485
      End
      Begin VB.TextBox Txt_qtde_empenho_PC 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4790
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peça empenhada."
         Top             =   3375
         Width           =   1515
      End
      Begin VB.TextBox Txt_qtde_estoque_PC 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1713
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peça em estoque."
         Top             =   3375
         Width           =   1515
      End
      Begin VB.TextBox Txt_responsavel_cadastro 
         Alignment       =   2  'Center
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   525
         Width           =   3250
      End
      Begin VB.TextBox Txt_data_cadastro 
         Alignment       =   2  'Center
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   525
         Width           =   1005
      End
      Begin VB.TextBox Txt_est_disp 
         Alignment       =   2  'Center
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
         Left            =   7815
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Estoque disponível."
         Top             =   1725
         Width           =   1515
      End
      Begin VB.TextBox Txt_data 
         Alignment       =   2  'Center
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
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data."
         Top             =   525
         Width           =   1155
      End
      Begin VB.TextBox Txt_certificado 
         Alignment       =   2  'Center
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
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Certificado."
         Top             =   1725
         Width           =   2340
      End
      Begin VB.TextBox Txt_corrida 
         Alignment       =   2  'Center
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Corrida."
         Top             =   1725
         Width           =   2220
      End
      Begin VB.TextBox Txt_loca_armaz 
         Alignment       =   2  'Center
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Local de armazenamento."
         Top             =   525
         Width           =   3690
      End
      Begin VB.TextBox Txt_qtde_estoque 
         Alignment       =   2  'Center
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
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade em estoque."
         Top             =   1725
         Width           =   1485
      End
      Begin VB.TextBox Txt_qtde_empenho 
         Alignment       =   2  'Center
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
         Left            =   6270
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade empenhada."
         Top             =   1725
         Width           =   1525
      End
      Begin VB.TextBox txtFornecedor 
         Alignment       =   2  'Center
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cliente/Fornecedor."
         Top             =   1155
         Width           =   9135
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Est. disponível PÇ"
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
         Left            =   7935
         TabIndex        =   39
         Top             =   3180
         Width           =   1305
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. emp. PÇ"
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
         Left            =   4995
         TabIndex        =   38
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. estoque PÇ"
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
         Left            =   1815
         TabIndex        =   37
         Top             =   3180
         Width           =   1305
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
         Index           =   3
         Left            =   2370
         TabIndex        =   31
         Top             =   330
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
         Index           =   4
         Left            =   510
         TabIndex        =   30
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Est. disponível"
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
         Left            =   8040
         TabIndex        =   29
         Top             =   1530
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Left            =   4875
         TabIndex        =   28
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. estoque"
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
         Left            =   4980
         TabIndex        =   27
         Top             =   1530
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. empenhada"
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
         TabIndex        =   26
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local de armazenamento"
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
         Left            =   6600
         TabIndex        =   25
         Top             =   330
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente/Fornecedor"
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
         Left            =   3780
         TabIndex        =   24
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado"
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
         Left            =   3195
         TabIndex        =   23
         Top             =   1530
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corrida"
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
         Left            =   1035
         TabIndex        =   22
         Top             =   1530
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmprod_Novo_Nota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LOTE As Boolean
Dim RE As Boolean

Private Sub btnEmpenhar_Click()
On Error GoTo tratar_erro

If Cmb_RE.Text = "" Then
    MsgBox ("Selecione uma RE para empenhar!"), vbInformation + vbOKOnly
    Exit Sub
End If

ProcSalvar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnLote_Click()
On Error GoTo tratar_erro

If optRE.Value = True Then Exit Sub



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRE_Click()
On Error GoTo tratar_erro
ID_empresa = frmprod.Cmb_empresa.ItemData(frmprod.Cmb_empresa.ListIndex)

If optRE.Value = True Then
Dim Estoque_Disponivel As Double
Dim Estoque_Empenhado As Double
Dim Estoque_Retirar As Double

If Cmb_RE.Text = "" Then
USMsgBox "Escolha uma opção para filtrar", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

ProcLimpaCampos
With frmprod

Set TBAbrir = CreateObject("adodb.recordset")

TBAbrir.Open "select Sum(Entrada) - Sum(Saida) as Saldo from Estoque_movimentacao where Idestoque = '" & Cmb_RE.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
StrSql = "update estoque_controle set Estoque_real = '" & Replace(TBAbrir!Saldo, ",", ".") & "' from Estoque_controle where IDEstoque = '" & Cmb_RE.Text & "'"
'Debug.print StrSql
QTSaldo = TBAbrir!Saldo
Conexao.Execute (StrSql)
End If
TBAbrir.Close

    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from estoque_produtos where IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Txt_data = IIf(IsNull(TBEstoque!Data), "", Format(TBEstoque!Data, "dd/mm/yy"))
        'If optLote.Value = True Then
            cmb_Lote.Text = IIf(IsNull(TBEstoque!LOTE), "", TBEstoque!LOTE)
        'End If
        Txt_loca_armaz = IIf(IsNull(TBEstoque!local_armaz), "", TBEstoque!local_armaz)
        txt_Corrida = IIf(IsNull(TBEstoque!Corrida), "", TBEstoque!Corrida)
        txt_Certificado = IIf(IsNull(TBEstoque!Certificado), "", TBEstoque!Certificado)
        Txt_qtde_estoque = Format(TBEstoque!estoque_real, "###,##0.0000")
        Txt_qtde_estoque_PC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, Format(TBEstoque!estoque_real_PC, "###,##0.0000"))
        
        Set TBMaterial = CreateObject("adodb.recordset")
        TBMaterial.Open "Select Movimentar_estoque_pc from Empresa where Codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaterial.EOF = False And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
            With txtQtde_Retirar
                .Locked = True
                .TabStop = False
            End With
            With txtQtde_Retirar_PC
                .Locked = False
                .TabStop = True
            End With
        Else
            With txtQtde_Retirar
                .Locked = False
                .TabStop = True
            End With
            With txtQtde_Retirar_PC
                .Locked = True
                .TabStop = False
            End With
        End If
                
        Txt_qtde_empenho = IIf(IsNull(TBEstoque!Qtde_empenhada), 0, Format(TBEstoque!Qtde_empenhada, "###,##0.0000"))
        Txt_qtde_empenho_PC = IIf(IsNull(TBEstoque!Qtde_empenhada_PC), 0, Format(TBEstoque!Qtde_empenhada_PC, "###,##0.0000"))
        
        Estoque_Disponivel = IIf(IsNull(TBEstoque!Estoque_Disponivel), 0, Format(TBEstoque!Estoque_Disponivel, "###,##0.0000"))
        Estoque_Empenhado = IIf(IsNull(TBEstoque!Qtde_empenhada), 0, Format(TBEstoque!Qtde_empenhada, "###,##0.0000"))
        Estoque_Disponivel = Estoque_Disponivel - Estoque_Empenhado
        Estoque_Retirar = frmprod.ListaRequisicao.SelectedItem.ListSubItems.Item(4)
        Txt_est_disp = IIf(IsNull(TBEstoque!Estoque_Disponivel), 0, Format(TBEstoque!Estoque_Disponivel - TBEstoque!Qtde_empenhada, "###,##0.0000"))
        Txt_est_disp_PC = IIf(IsNull(TBEstoque!Estoque_disponivel_PC), 0, Format(TBEstoque!Estoque_disponivel_PC, "###,##0.0000"))
        
        If Estoque_Disponivel <= Estoque_Retirar Then
        txtQtde_Retirar.Text = Format(Estoque_Disponivel, "###,##0.0000")
        Else
        txtQtde_Retirar.Text = Format(txtQtde_aretirar, "###,##0.0000")
        End If
        
        
        If TBEstoque!Consignacao = True Then txtFornecedor = IIf(IsNull(TBEstoque!Cliente), "", TBEstoque!Cliente) Else txtFornecedor = IIf(IsNull(TBEstoque!Fornecedor), "", TBEstoque!Fornecedor)
    End If
    TBEstoque.Close
End With
Else
Cmb_RE.Clear
Set TBLISTA = CreateObject("adodb.recordset")

StrSql = "Select IDestoque, Estoque_disponivel, Qtde_empenhada,Lote from Estoque_produtos where Lote = '" & cmb_Lote & "' and ID_empresa = '" & ID_empresa & "' and Desenho = '" & txtCodigo.Text & "' and Liberado = 'SIM' AND Estoque_disponivel > 0 and Consignacao = 'False' order by Data, IDestoque"

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
   Contador = 0
   Do While TBLISTA.EOF = False
        
        If TBLISTA!Estoque_Disponivel - TBLISTA!Qtde_empenhada > 0 Then
            Cmb_RE.AddItem TBLISTA!IDEstoque
            Contador = Contador + 1
        End If
        
        TBLISTA.MoveNext
    Loop
    
    If Contador = 1 Then
        TBLISTA.MoveFirst
        Cmb_RE.Text = TBLISTA!IDEstoque
    End If
    
TBLISTA.Close


End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_lote_Click()
'On Error GoTo tratar_erro
'
'If optRE.Value = True Then Exit Sub
'
'Cmb_RE.Clear
'Set TBLISTA = CreateObject("adodb.recordset")
'TBLISTA.Open "SELECT * from Estoque_Produtos where Lote = '" & cmb_Lote & "' order by IDestoque", Conexao, adOpenKeyset, adLockOptimistic
'   contador = 0
'   Do While TBLISTA.EOF = False
'
'        If TBLISTA!Estoque_Disponivel - TBLISTA!Qtde_empenhada > 0 Then
'            Cmb_RE.AddItem TBLISTA!IDEstoque
'            contador = contador + 1
'        End If
'
'        TBLISTA.MoveNext
'    Loop
'
'    If contador = 1 Then
'        TBLISTA.MoveFirst
'        Cmb_RE.Text = TBLISTA!IDEstoque
'    End If
'
'TBLISTA.Close
'
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
End Sub

Private Sub Cmb_RE_Click()
On Error GoTo tratar_erro
Dim Estoque_Disponivel As Double
Dim Estoque_Empenhado As Double
Dim Estoque_Retirar As Double

ProcLimpaCampos
With frmprod

Set TBAbrir = CreateObject("adodb.recordset")

TBAbrir.Open "select Sum(Entrada) - Sum(Saida) as Saldo from Estoque_movimentacao where Idestoque = '" & Cmb_RE.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
StrSql = "update estoque_controle set Estoque_real = '" & Replace(TBAbrir!Saldo, ",", ".") & "' from Estoque_controle where IDEstoque = '" & Cmb_RE.Text & "'"
'Debug.print StrSql
QTSaldo = TBAbrir!Saldo
Conexao.Execute (StrSql)
End If
TBAbrir.Close

    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from estoque_produtos where IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Txt_data = IIf(IsNull(TBEstoque!Data), "", Format(TBEstoque!Data, "dd/mm/yy"))
        If optLote.Value = True Then
            cmb_Lote.Text = IIf(IsNull(TBEstoque!LOTE), "", TBEstoque!LOTE)
        End If
        Txt_loca_armaz = IIf(IsNull(TBEstoque!local_armaz), "", TBEstoque!local_armaz)
        txt_Corrida = IIf(IsNull(TBEstoque!Corrida), "", TBEstoque!Corrida)
        txt_Certificado = IIf(IsNull(TBEstoque!Certificado), "", TBEstoque!Certificado)
        Txt_qtde_estoque = Format(TBEstoque!estoque_real, "###,##0.0000")
        Txt_qtde_estoque_PC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, Format(TBEstoque!estoque_real_PC, "###,##0.0000"))

        Set TBMaterial = CreateObject("adodb.recordset")
        TBMaterial.Open "Select Movimentar_estoque_pc from Empresa where Codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaterial.EOF = False And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
            With txtQtde_Retirar
                .Locked = True
                .TabStop = False
            End With
            With txtQtde_Retirar_PC
                .Locked = False
                .TabStop = True
            End With
        Else
            With txtQtde_Retirar
                .Locked = False
                .TabStop = True
            End With
            With txtQtde_Retirar_PC
                .Locked = True
                .TabStop = False
            End With
        End If

        Txt_qtde_empenho = IIf(IsNull(TBEstoque!Qtde_empenhada), 0, Format(TBEstoque!Qtde_empenhada, "###,##0.0000"))
        Txt_qtde_empenho_PC = IIf(IsNull(TBEstoque!Qtde_empenhada_PC), 0, Format(TBEstoque!Qtde_empenhada_PC, "###,##0.0000"))

        Estoque_Disponivel = IIf(IsNull(TBEstoque!Estoque_Disponivel), 0, Format(TBEstoque!Estoque_Disponivel, "###,##0.0000"))
        Estoque_Empenhado = IIf(IsNull(TBEstoque!Qtde_empenhada), 0, Format(TBEstoque!Qtde_empenhada, "###,##0.0000"))
        Estoque_Disponivel = Estoque_Disponivel - Estoque_Empenhado
        Estoque_Retirar = frmprod.ListaRequisicao.SelectedItem.ListSubItems.Item(4)
        Txt_est_disp = IIf(IsNull(TBEstoque!Estoque_Disponivel), 0, Format(TBEstoque!Estoque_Disponivel - TBEstoque!Qtde_empenhada, "###,##0.0000"))
        Txt_est_disp_PC = IIf(IsNull(TBEstoque!Estoque_disponivel_PC), 0, Format(TBEstoque!Estoque_disponivel_PC, "###,##0.0000"))

        If Estoque_Disponivel <= Estoque_Retirar Then
        txtQtde_Retirar.Text = Format(Estoque_Disponivel, "###,##0.0000")
        Else
        txtQtde_Retirar.Text = Format(txtQtde_aretirar, "###,##0.0000")
        End If


        If TBEstoque!Consignacao = True Then txtFornecedor = IIf(IsNull(TBEstoque!Cliente), "", TBEstoque!Cliente) Else txtFornecedor = IIf(IsNull(TBEstoque!Fornecedor), "", TBEstoque!Fornecedor)
    End If
    TBEstoque.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

Acao = "adicionar o RE na ordem"
If Cmb_RE = "" Then
    NomeCampo = "o número da rastreabilidade do estoque"
    ProcVerificaAcao
    Cmb_RE.SetFocus
    Exit Sub
End If
If txtQtde_Retirar.Locked = False Then NomeCampo = "a quantidade a empenhar" Else NomeCampo = "a quantidade de peças a empenhar"
valor = IIf(txtQtde_Retirar = "", 0, txtQtde_Retirar)
If valor <= 0 Then
    ProcVerificaAcao
    If txtQtde_Retirar.Locked = False Then txtQtde_Retirar.SetFocus Else txtQtde_Retirar_PC.SetFocus
    Exit Sub
End If



With frmprod
    'Verifica se a qtde. requistada é menor que a qtde. do estoque - qtde empenhada
    If txtQtde_Retirar_PC = "" Or txtQtde_Retirar_PC = "0" Then
        Qtde = txtQtde_Retirar
        Requisitado = IIf(Txt_qtde_req = "", 0, Txt_qtde_req)
        quantestoque = IIf(Txt_qtde_estoque = "", 0, Txt_qtde_estoque)
        qtdeliberada = IIf(Txt_qtde_empenho = "", "0", Txt_qtde_empenho)
        
        FamiliaAntiga = Requisitado '.ListaRequisicao.SelectedItem.ListSubItems(4) 'Unidade
    Else
        Qtde = txtQtde_Retirar_PC
        Requisitado = IIf(Txt_qtde_req_PC = "", 0, Txt_qtde_req_PC)
        quantestoque = IIf(Txt_qtde_estoque_PC = "", 0, Txt_qtde_estoque_PC)
        qtdeliberada = IIf(Txt_qtde_empenho_PC = "", "0", Txt_qtde_empenho_PC)
        
        FamiliaAntiga = "PÇ"
    End If
    
    qtdeliberar = Format(quantestoque - qtdeliberada, "###,##0.0000")
    
    
    If Qtde > qtdeliberar Then
        'Verifica para quais ordens esta empenhado o RE
        Familiatext = ""
        Set TBMaterial = CreateObject("adodb.recordset")
        TBMaterial.Open "Select PNC.* from producaomaterial PM INNER JOIN Producao_NF_Consignada PNC ON PM.Ordem = PNC.Ordem and PM.Codigo = PNC.Codinterno where PM.Ordem <> " & .txtof & " and PM.codigo = '" & txtCodigo.Text & "' and PNC.IDestoque = " & Cmb_RE & " and (PM.Saida = 'NÃO' or PM.Saida = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaterial.EOF = False Then
            Do While TBMaterial.EOF = False
                If txtQtde_Retirar_PC = "" Or txtQtde_Retirar_PC = "0" Then
                    If Familiatext = "" Then Familiatext = TBMaterial!Ordem & " - " & Format(TBMaterial!quantidade - TBMaterial!Qtde_saida, "###,##0.0000") Else Familiatext = Familiatext & " | " & TBMaterial!Ordem & " - " & Format(TBMaterial!quantidade - TBMaterial!Qtde_saida, "###,##0.0000")
                Else
                    If Familiatext = "" Then Familiatext = TBMaterial!Ordem & " - " & Format(TBMaterial!Quantidade_PC - TBMaterial!Qtde_saida_PC, "###,##0.0000") Else Familiatext = Familiatext & " | " & TBMaterial!Ordem & " - " & Format(TBMaterial!Quantidade_PC - TBMaterial!Qtde_saida_PC, "###,##0.0000")
                End If
                TBMaterial.MoveNext
            Loop
        End If
        TBMaterial.Close
        
        If txtQtde_Retirar_PC = "" Or txtQtde_Retirar_PC = "0" Then
            USMsgBox ("Não é permitido empenhar este RE nesta ordem, pois a quantidade do material " & txtCodigo.Text & " disponível em estoque é menor que a quantidade a empenhar." & vbCrLf & "Em estoque: " & Format(quantestoque, "###,##0.0000") & " " & FamiliaAntiga & vbCrLf & "Empenhada: " & Format(qtdeliberada, "###,##0.0000") & " " & FamiliaAntiga & vbCrLf & "Disponível: " & Format(quantestoque - qtdeliberada, "###,##0.0000") & " " & FamiliaAntiga & vbCrLf & "Qtde. empenhar: " & Format(Qtde, "###,##0.0000") & " " & FamiliaAntiga & vbCrLf & "Lista de empenho: " & Familiatext), vbExclamation, "CAPRIND v5.0"
        Else
            USMsgBox ("Não é permitido empenhar este RE nesta ordem, pois a quantidade do material " & txtCodigo.Text & " disponível em estoque é menor que a quantidade a empenhar." & vbCrLf & "Em estoque: " & quantestoque & " " & FamiliaAntiga & vbCrLf & "Empenhada: " & qtdeliberada & " " & FamiliaAntiga & vbCrLf & "Disponível: " & quantestoque - qtdeliberada & " " & FamiliaAntiga & vbCrLf & "Qtde. empenhar: " & Qtde & " " & FamiliaAntiga & vbCrLf & "Lista de empenho: " & Familiatext), vbExclamation, "CAPRIND v5.0"
        End If
        Exit Sub
    End If
    
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from Producao_NF_Consignada where Ordem = " & .txtof & " and codinterno = '" & txtCodigo.Text & "' and IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        USMsgBox ("Este RE já foi empenhado nesta ordem."), vbExclamation, "CAPRIND v5.0"
        TBproducao.Close
        Exit Sub
    Else
        TBproducao.AddNew
        USMsgBox ("RE empenhado na ordem com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
    TBproducao!Data = Txt_data_cadastro
    TBproducao!Responsavel = Txt_responsavel_cadastro
    TBproducao!Ordem = .txtof
    TBproducao!Codinterno = txtCodigo 'txtcodigo.text
    TBproducao!IDEstoque = Cmb_RE
    TBproducao!quantidade = txtQtde_Retirar
    TBproducao!Quantidade_PC = IIf(txtQtde_Retirar_PC = "", 0, txtQtde_Retirar_PC)
    
    TBproducao.Update
    '==================================
    Modulo = "PCP/Gerenciamento de ordem"
    Evento = "Empenhar RE"
    ID_documento = TBproducao!ID
    Documento = "Ordem: " & .txtof.Text & " - Cód. interno: " & .txtdesenho
    Documento1 = "Cód. interno: " & txtCodigo & " - RE: " & Cmb_RE
    ProcGravaEvento
    '==================================
    TBproducao.Close
    
    valor = Txt_qtde_req
    NovoValor = Replace(valor, ",", ".")
    If Txt_qtde_req_PC = "" Then
        TextoFiltro = ""
    Else
        Valor1 = Txt_qtde_req_PC
        NovoValor1 = Replace(Valor1, ",", ".")
        TextoFiltro = ", Total_pc = " & NovoValor1
    End If
    Conexao.Execute "Update Producaomaterial Set Requisitado = " & NovoValor & " " & TextoFiltro & " where Idmateriaprima = " & .ListaRequisicao.SelectedItem

    
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select sum(quantidade) as Qtd, Sum(ISNULL(quantidade_PC, 0)) as Qt from Producao_NF_Consignada where Ordem = " & .txtof & " and Codinterno = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Qtd = IIf(IsNull(TBAbrir!Qtd), 0, TBAbrir!Qtd)
        qt = IIf(IsNull(TBAbrir!qt), "0", TBAbrir!qt)
    End If
    TBAbrir.Close
    Qtde = Txt_qtde_req 'frmprod.ListaRequisicao.SelectedItem.ListSubItems.Item(4)
    txtQtde_retirada = Format(Qtd, "###,##0.0000")
    txtQtde_retirada_PC = qt
    txtQtde_aretirar = Format(Qtde - Qtd, "###,##0.0000")
    txtQtde_aretirar_PC = Valor3 - qt
    txtQtde_Retirar = Format(Qtde - Qtd, "###,##0.0000")
    '.ProcCarregaListaRequisicao
    .ProcCarregaListaEmpenhos
    If txtQtde_Retirar = 0 Then
     Unload Me
    End If
    
End With
'Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmb_Lote.Locked = True

With frmprod
    
    Cmb_RE.Clear
    Do While TBLISTA.EOF = False
    If TBLISTA!Estoque_Disponivel - TBLISTA!Qtde_empenhada > 0 Then
        Cmb_RE.AddItem TBLISTA!IDEstoque
    End If
        TBLISTA.MoveNext
    Loop
    
    TBLISTA.MoveFirst
    
    Txt_data_cadastro = Format(Date, "dd/mm/yy")
    Txt_responsavel_cadastro = pubUsuario
    txtIDEmpresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
    Txt_cliente = .txtCliente
    Txt_qtde_req = .ListaRequisicao.SelectedItem.ListSubItems(4)
    Txt_un = .ListaRequisicao.SelectedItem.ListSubItems(3)
    TXTIDLista.Text = .ListaRequisicao.SelectedItem
    txtCodigo = .ListaRequisicao.SelectedItem.ListSubItems(1)
    txtdescricao = .ListaRequisicao.SelectedItem.ListSubItems(2)
    
    Qtde = Txt_qtde_req
    Valor3 = IIf(Txt_qtde_req_PC = "", 0, Txt_qtde_req_PC)
    
    Qtd = 0
    qt = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select sum(quantidade) as Qtd, Sum(ISNULL(quantidade_PC, 0)) as Qt from Producao_NF_Consignada where Ordem = " & .txtof & " and Codinterno = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Qtd = IIf(IsNull(TBAbrir!Qtd), 0, TBAbrir!Qtd)
        qt = IIf(IsNull(TBAbrir!qt), "0", TBAbrir!qt)
    End If
    TBAbrir.Close
    txtQtde_retirada = Format(Qtd, "###,##0.0000")
    txtQtde_retirada_PC = qt
    txtQtde_aretirar = Format(Qtde - Qtd, "###,##0.0000")
    txtQtde_aretirar_PC = Valor3 - qt
    txtQtde_Retirar = Format(Qtde - Qtd, "###,##0.0000")
End With

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

Private Sub optLote_Click()
On Error GoTo tratar_erro
    
    Cmb_RE.Clear
    cmb_Lote.Locked = False
    
    cmb_Lote.Clear
    
    txtQtde_retirada = Format(Qtd, "###,##0.0000")
    txtQtde_retirada_PC = qt
    txtQtde_aretirar = Format(Qtde - Qtd, "###,##0.0000")
    txtQtde_aretirar_PC = Valor3 - qt
    txtQtde_Retirar = Format(Qtde - Qtd, "###,##0.0000")
    
   
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Lote from Estoque_produtos where ID_empresa = " & txtIDEmpresa & " and Desenho = '" & txtCodigo & "' and Liberado = 'SIM' AND Estoque_disponivel > 0 " & Familiatext & " group by Lote", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBLISTA.EOF = False
        cmb_Lote.AddItem TBLISTA!LOTE
        TBLISTA.MoveNext
    Loop
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub optRE_Click()
On Error GoTo tratar_erro

    Cmb_RE.Locked = False
    cmb_Lote.Locked = True
    Cmb_RE.Clear
    
    txtQtde_retirada = Format(Qtd, "###,##0.0000")
    txtQtde_retirada_PC = qt
    txtQtde_aretirar = Format(Qtde - Qtd, "###,##0.0000")
    txtQtde_aretirar_PC = Valor3 - qt
    txtQtde_Retirar = Format(Qtde - Qtd, "###,##0.0000")
    
    
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select IDestoque, Estoque_disponivel, Qtde_empenhada, Lote from Estoque_produtos where ID_empresa = " & txtIDEmpresa & " and Desenho = '" & txtCodigo & "' and Liberado = 'SIM' AND Estoque_disponivel > 0 " & Familiatext & " order by Data, IDestoque", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBLISTA.EOF = False
        If TBLISTA!Estoque_Disponivel - TBLISTA!Qtde_empenhada > 0 Then
            Cmb_RE.AddItem TBLISTA!IDEstoque
        End If
        TBLISTA.MoveNext
    Loop
    TBLISTA.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub Txt_qtde_req_Change()
On Error GoTo tratar_erro

If Txt_qtde_req <> "" Then
    VerifNumero = Txt_qtde_req
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_req = ""
        Txt_qtde_req.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_req_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_req

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_req_LostFocus()
On Error GoTo tratar_erro

Txt_qtde_req = Format(Txt_qtde_req, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_data = ""
'cmb_Lote.Clear
Txt_loca_armaz = ""
txt_Corrida = ""
txt_Certificado = ""
Txt_qtde_estoque = ""
Txt_qtde_empenho = ""
Txt_est_disp = ""
txtFornecedor = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Retirar_Change()
On Error GoTo tratar_erro

If txtQtde_Retirar <> "" Then
    VerifNumero = txtQtde_Retirar
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_Retirar = ""
        txtQtde_Retirar.SetFocus
        Exit Sub
    End If
    If txtQtde_Retirar.Locked = False And txtQtde_aretirar_PC > 0 Then txtQtde_Retirar_PC = FunCalculaQtdePCKG(Txt_est_disp, Txt_est_disp_PC, txtQtde_Retirar, True)
Else
    txtQtde_Retirar_PC = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Retirar_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde_Retirar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Retirar_LostFocus()
On Error GoTo tratar_erro

txtQtde_Retirar = Format(txtQtde_Retirar, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Retirar_PC_Change()
On Error GoTo tratar_erro

If txtQtde_Retirar_PC <> "" Then
    VerifNumero = txtQtde_Retirar_PC
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_Retirar_PC = ""
        txtQtde_Retirar_PC.SetFocus
        Exit Sub
    End If
    If txtQtde_Retirar_PC.Locked = False Then txtQtde_Retirar = FunCalculaQtdePCKG(Txt_est_disp, Txt_est_disp_PC, txtQtde_Retirar_PC, False)
Else
    txtQtde_Retirar = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Retirar_PC_GotFocus()
On Error GoTo tratar_erro

FunGotFocus txtQtde_Retirar_PC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

