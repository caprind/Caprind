VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_Opcoes_Mrp 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP - Gerenciamento de ordem - Opções MRP"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   8025
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProd_Opcoes_Mrp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empenho"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   360
      TabIndex        =   64
      Top             =   6360
      Width           =   3915
      Begin DrawSuite2022.USCheckBox chkEmpenho 
         Height          =   285
         Left            =   180
         TabIndex        =   65
         ToolTipText     =   "Empenhar ordem de produção para o pedido interno"
         Top             =   390
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   503
         Caption         =   "Empenhar ordem de produção no PI?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   -1  'True
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   62
      Top             =   7530
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   714
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produzir lote(s) em"
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
      Height          =   855
      Left            =   5670
      TabIndex        =   59
      Top             =   5490
      Width           =   2025
      Begin VB.TextBox txtCopiarOP 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   120
         MaxLength       =   20
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "1"
         ToolTipText     =   "Quantidade de produtos a produzir"
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordem(ns)"
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
         Left            =   870
         TabIndex        =   61
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.Frame Frame_estoque 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações de estoque"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   300
      TabIndex        =   38
      Top             =   690
      Width           =   7395
      Begin VB.TextBox txtEstoque 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6150
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade produto disponível em estoque"
         Top             =   1050
         Width           =   1020
      End
      Begin VB.TextBox txtDescricao 
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
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   5385
      End
      Begin VB.TextBox TxtEstTotal 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total em estoque."
         Top             =   1050
         Width           =   915
      End
      Begin VB.TextBox TxtEstdisponivel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade disponível em estoque."
         Top             =   1050
         Width           =   885
      End
      Begin VB.TextBox TxtEstEmpenhado 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total empenhada."
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox TxtEstUN 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox TxtEstCodigo 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   1155
      End
      Begin VB.TextBox Txt_qtde_prod 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total produzindo."
         Top             =   1050
         Width           =   1005
      End
      Begin VB.TextBox Txt_qtde_emp_est 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total empenhada."
         Top             =   1050
         Width           =   945
      End
      Begin VB.TextBox Txt_qtde_disp_prod 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade disponível produzindo."
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   58
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   1920
         TabIndex        =   56
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "= Disponivel"
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
         Left            =   2010
         TabIndex        =   54
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   210
         TabIndex        =   53
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produzindo"
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
         Left            =   3285
         TabIndex        =   52
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "= Disponível"
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
         Left            =   5100
         TabIndex        =   51
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label1 
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
         Index           =   39
         Left            =   1470
         TabIndex        =   50
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em estoque"
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
         Index           =   40
         Left            =   180
         TabIndex        =   49
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Empenhado"
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
         Index           =   41
         Left            =   1050
         TabIndex        =   48
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Empenhado"
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
         Index           =   42
         Left            =   4140
         TabIndex        =   47
         Top             =   840
         Width           =   945
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Est. mínimo"
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
      Height          =   855
      Left            =   330
      TabIndex        =   36
      Top             =   5490
      Width           =   1125
      Begin VB.TextBox txtEstoqueMinimo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Quantidadeestoque mínimo"
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lote minimo"
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
      Height          =   855
      Left            =   2700
      TabIndex        =   34
      Top             =   5490
      Width           =   1575
      Begin VB.TextBox txtLoteminimo 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade lote mínimo a produzir"
         Top             =   315
         Width           =   1290
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProd_Opcoes_Mrp.frx":000C
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin DrawSuite2022.USButton cmdGerarmrp 
      Height          =   975
      Left            =   4290
      TabIndex        =   1
      ToolTipText     =   "Clique aqui para gerar as ordens de produção e requisições de materiais no estoque."
      Top             =   6360
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1720
      DibPicture      =   "frmProd_Opcoes_Mrp.frx":0028
      Caption         =   "Gerar MRP"
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
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   4
      ToolTipTitle    =   "CAPRIND V5.0"
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vendido"
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
      Height          =   855
      Left            =   1470
      TabIndex        =   29
      Top             =   5490
      Width           =   1215
      Begin VB.TextBox txtqtVendido 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade produtos vendidos"
         Top             =   315
         Width           =   1020
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Versão da estrutura"
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
      Left            =   3150
      TabIndex        =   28
      Top             =   4860
      Width           =   2025
      Begin VB.CheckBox Chk_versao_por_item_estrutura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por item"
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
         TabIndex        =   31
         Top             =   330
         Width           =   885
      End
      Begin VB.ComboBox cmbVersao_estrutura 
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
         ItemData        =   "frmProd_Opcoes_Mrp.frx":12C1
         Left            =   1020
         List            =   "frmProd_Opcoes_Mrp.frx":12C3
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Versão."
         Top             =   225
         Width           =   675
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estoque automático ?"
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
      Height          =   615
      Left            =   330
      TabIndex        =   27
      Top             =   4860
      Width           =   2775
      Begin VB.CheckBox Chk_entrar_estoque 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrada"
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
         Left            =   210
         TabIndex        =   10
         Top             =   300
         Width           =   885
      End
      Begin VB.CheckBox Chk_retirar_estoque 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saída"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções de impressão para tipo do relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4020
      TabIndex        =   26
      Top             =   2280
      Width           =   3675
      Begin VB.CheckBox Chk_etiqueta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gerar e imprimir etiqueta individual?"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1965
         Width           =   3075
      End
      Begin VB.CheckBox Chk_rm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Somente requisição de materiais?"
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
         TabIndex        =   18
         Top             =   1245
         Width           =   2835
      End
      Begin VB.CheckBox Chk_ordem_rm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordem com requisição (RM)?"
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
         TabIndex        =   15
         Top             =   510
         Width           =   2835
      End
      Begin VB.CheckBox Chk_ordem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Somente ordem?"
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
         TabIndex        =   14
         Top             =   270
         Width           =   2835
      End
      Begin VB.CheckBox Chk_visualizar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visualizar antes da impressão?"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2220
         Width           =   2835
      End
      Begin VB.CheckBox Chk_plano 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Somente plano(s) de inspeção?"
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
         TabIndex        =   19
         Top             =   1485
         Width           =   2835
      End
      Begin VB.CheckBox Chk_frequencia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Somente frequencia(s) de medição?"
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
         TabIndex        =   20
         Top             =   1725
         Width           =   2865
      End
      Begin VB.CheckBox Chk_ordem_manual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordem para apontamento manual?"
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
         TabIndex        =   17
         Top             =   1005
         Width           =   2835
      End
      Begin VB.CheckBox Chk_ordem_rm_resumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordem e requisição (resumido)?"
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
         TabIndex        =   16
         Top             =   750
         Width           =   2835
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Versão do processo"
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
      Left            =   5220
      TabIndex        =   25
      Top             =   4860
      Width           =   2475
      Begin VB.CheckBox Chk_versao_por_item 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por item"
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
         TabIndex        =   32
         Top             =   330
         Width           =   885
      End
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
         ItemData        =   "frmProd_Opcoes_Mrp.frx":12C5
         Left            =   1020
         List            =   "frmProd_Opcoes_Mrp.frx":12C7
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Versão."
         Top             =   225
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções emissão ordem"
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
      Height          =   2535
      Left            =   300
      TabIndex        =   23
      Top             =   2280
      Width           =   3645
      Begin VB.CheckBox chkRastreavel 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   63
         Top             =   150
         Width           =   525
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Left            =   330
         TabIndex        =   66
         Top             =   450
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   344
         Caption         =   "É urgente?"
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
         NoHTMLCaption   =   "É urgente?"
      End
      Begin VB.CheckBox optconsignacao 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   3
         Top             =   645
         Width           =   525
      End
      Begin VB.CheckBox chkEscopo 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   9
         Top             =   2190
         Width           =   525
      End
      Begin VB.CheckBox Chk_agrupar 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   8
         Top             =   1935
         Width           =   525
      End
      Begin VB.CheckBox chkValidada 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   7
         Top             =   1680
         Width           =   525
      End
      Begin VB.CheckBox Chk_empenhar_estoque 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   6
         Top             =   1410
         Width           =   525
      End
      Begin VB.CheckBox Opt_processo_controlado 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   5
         Top             =   1140
         Width           =   525
      End
      Begin VB.CheckBox OptOSControlada 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   4
         Top             =   885
         Width           =   525
      End
      Begin VB.CheckBox Chk_urgente 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2970
         TabIndex        =   2
         Top             =   390
         Width           =   525
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   195
         Left            =   330
         TabIndex        =   67
         Top             =   2220
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   344
         Caption         =   "Escopo na ordem?"
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
         NoHTMLCaption   =   "Escopo na ordem?"
      End
      Begin DrawSuite2022.USLabel USLabel3 
         Height          =   195
         Left            =   330
         TabIndex        =   68
         Top             =   1950
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   344
         Caption         =   "Agrupar itens na ordem?"
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
         NoHTMLCaption   =   "Agrupar itens na ordem?"
      End
      Begin DrawSuite2022.USLabel USLabel4 
         Height          =   195
         Left            =   330
         TabIndex        =   69
         Top             =   1710
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   344
         Caption         =   "Ordem validada?"
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
         NoHTMLCaption   =   "Ordem validada?"
      End
      Begin DrawSuite2022.USLabel USLabel5 
         Height          =   195
         Left            =   330
         TabIndex        =   70
         Top             =   1440
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   344
         Caption         =   "Considerar saldo no estoque?"
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
         NoHTMLCaption   =   "Considerar saldo no estoque?"
      End
      Begin DrawSuite2022.USLabel USLabel6 
         Height          =   195
         Left            =   330
         TabIndex        =   71
         Top             =   1170
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   344
         Caption         =   "Com processo controlado?"
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
         NoHTMLCaption   =   "Com processo controlado?"
      End
      Begin DrawSuite2022.USLabel USLabel7 
         Height          =   195
         Left            =   330
         TabIndex        =   72
         Top             =   930
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   344
         Caption         =   "A ordem é controlada?"
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
         NoHTMLCaption   =   "A ordem é controlada?"
      End
      Begin DrawSuite2022.USLabel USLabel8 
         Height          =   195
         Left            =   330
         TabIndex        =   73
         Top             =   690
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   344
         Caption         =   "Utiliza materia prima de terceiro?"
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
         NoHTMLCaption   =   "Utiliza materia prima de terceiro?"
      End
      Begin DrawSuite2022.USLabel USLabel9 
         Height          =   195
         Left            =   330
         TabIndex        =   74
         Top             =   210
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   344
         Caption         =   "Rastreabilidade individual"
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
         NoHTMLCaption   =   "Rastreabilidade individual"
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Qtde da OP"
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
      Height          =   855
      Left            =   4290
      TabIndex        =   24
      Top             =   5490
      Width           =   1365
      Begin VB.TextBox txtQuantidade 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   90
         MaxLength       =   20
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de produtos a produzir"
         Top             =   315
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmProd_Opcoes_Mrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_agrupar_Click()
On Error GoTo tratar_erro

With frmprod
    If Chk_agrupar.Value = 1 Then
        .Agrupar_ordem = True
    Else
        .Agrupar_ordem = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_empenhar_estoque_Click()
On Error GoTo tratar_erro

With frmprod
    If Chk_empenhar_estoque.Value = 1 Then
        .MRP_considerar_estoque = True
    Else
        .MRP_considerar_estoque = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_entrar_estoque_Click()
On Error GoTo tratar_erro

With frmprod
    If Chk_entrar_estoque.Value = 1 Then
        .EntrarEstoque = True
        .Chk_entrar_estoque.Value = 1
        .Caption = "Sim"
    Else
        .EntrarEstoque = False
        .Chk_entrar_estoque.Value = 0
        .Caption = "Não"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_etiqueta_Click()
On Error GoTo tratar_erro

With frmprod
    ProcHabDesabVisualizandoImpressao
        If Chk_etiqueta.Value = 1 Then
            .MRP_Imprimir_Etiq = True
        Else
            .MRP_Imprimir_Etiq = False
        End If

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_frequencia_Click()
On Error GoTo tratar_erro

With frmprod
    ProcHabDesabVisualizandoImpressao
    If Chk_frequencia.Value = 1 Then .MRP_Imprimir_Freq = True Else .MRP_Imprimir_Freq = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_Click()
On Error GoTo tratar_erro

With frmprod
    ProcHabDesabVisualizandoImpressao
    If Chk_ordem.Value = 1 Then .MRP_Imprimir_Ordem = True Else .MRP_Imprimir_Ordem = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_manual_Click()
On Error GoTo tratar_erro

With frmprod
ProcHabDesabVisualizandoImpressao
If Chk_ordem_manual.Value = 1 Then .MRP_Imprimir_Ordem_APM = True Else .MRP_Imprimir_Ordem_APM = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_rm_Click()
On Error GoTo tratar_erro

With frmprod
    ProcHabDesabVisualizandoImpressao
    If Chk_ordem_rm.Value = 1 Then .MRP_Imprimir_Ordem_RM = True Else .MRP_Imprimir_Ordem_RM = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_plano_Click()
On Error GoTo tratar_erro

With frmprod
    ProcHabDesabVisualizandoImpressao
    If Chk_plano.Value = 1 Then .MRP_Imprimir_Plano = True Else .MRP_Imprimir_Plano = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_rm_Click()
On Error GoTo tratar_erro

With frmprod
    ProcHabDesabVisualizandoImpressao
    If Chk_rm.Value = 1 Then .MRP_Imprimir_RM = True Else .MRP_Imprimir_RM = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_versao_por_item_Click()
On Error GoTo tratar_erro

With frmprod
    If Chk_versao_por_item.Value = 1 Then
        Frame4.Enabled = False
        cmbVersao.ListIndex = -1
        .Versao_por_item = True
    Else
        Frame4.Enabled = True
        cmbVersao = "A"
        .Versao_por_item = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_versao_por_item_estrutura_Click()
On Error GoTo tratar_erro

With frmprod
    If Chk_versao_por_item_estrutura.Value = 1 Then
        Frame5.Enabled = False
        cmbVersao_estrutura.ListIndex = -1
        .Versao_por_item_estrutura = True
    Else
        Frame5.Enabled = True
        cmbVersao_estrutura = "A"
        .Versao_por_item_estrutura = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_visualizar_Click()
On Error GoTo tratar_erro

If Chk_visualizar.Value = 1 Then Imprimir = False Else Imprimir = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_urgente_Click()
On Error GoTo tratar_erro

With frmprod
    If Chk_urgente.Value = 1 Then
        Urgencia = True
        .Chk_urgente.Value = 1
    Else
        Urgencia = False
        .Chk_urgente.Value = 0
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_retirar_estoque_Click()
On Error GoTo tratar_erro

With frmprod
    If Chk_retirar_estoque.Value = 1 Then
        .RetirarEstoque = True
        .Chk_retirar_estoque.Value = 1
        Chk_retirar_estoque.Caption = "Sim"
    Else
        .RetirarEstoque = False
        .Chk_retirar_estoque.Value = 0
        Chk_retirar_estoque.Caption = "Não"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ordem_rm_resumido_Click()
On Error GoTo tratar_erro

With frmprod
    ProcHabDesabVisualizandoImpressao
    If Chk_ordem_rm_resumido.Value = 1 Then .MRP_Imprimir_Ordem_RM_Res = True Else .MRP_Imprimir_Ordem_RM_Res = False
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

Private Sub ProcGerar()
On Error GoTo tratar_erro

With frmprod
    If Chk_etiqueta.Value = 1 Then
        If USMsgBox("Ordem com rastreabilidade individual, deseja imprimir as etiquetas?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem:
            Qtde_copias = InputBox("Favor informar a quantidade de cópias de etiqueta.", , 1)
            If Qtde_copias = "" Then Exit Sub
            
            If IsNumeric(Qtde_copias) = False Then
                USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            .MRP_Qtcopias_Etiq = Qtde_copias
            .MRP_Qtcopias_Etiq = FunArredondarPraCima(.MRP_Qtcopias_Etiq)
            If .MRP_Qtcopias_Etiq <= 0 Then
                USMsgBox ("So é permitido quantidade maior que 0."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
         Else
            .MRP_Imprimir_Etiq = False
        End If

    End If
    
    
    If .MRP_Prod = False Then 'Ordem manual
    .ProcGerarMrp
    Else
    ProcGeraMrpItem 'Ordem carteira
    End If
    
    If .MRPgerado = True Then
        USMsgBox ("MRP gerado com sucesso."), vbInformation, "CAPRIND v5.0"
        If .MRP_Prod = False Then
            '==================================
            Modulo = "PCP/Gerenciamento de ordem"
            Evento = "Gerar MRP"
            Documento = "Data de emissão: " & Date
            ProcGravaEvento
            '==================================
        Else
            For InitFor = 1 To .listaitens.ListItems.Count
                If .listaitens.ListItems.Item(InitFor).Checked = True Then
                    '==================================
                    Modulo = "PCP/Gerenciamento de ordem"
                    Evento = "Gerar MRP do(s) produto(s)/serviço(s)"
                    ID_documento = .listaitens.ListItems.Item(InitFor)
                    Documento = "Ped. interno: " & .listaitens.ListItems.Item(InitFor).ListSubItems(19) & " - Rev.: " & .listaitens.ListItems.Item(InitFor).ListSubItems(20) & " - Cód. interno: " & .listaitens.ListItems.Item(InitFor).ListSubItems(3)
                    Documento1 = ""
                    ProcGravaEvento
                    '==================================
                End If
            Next InitFor
        End If
        
        If .StrSql_Ordem_MRP <> "" And .MRPgerado = True Then
            .ProcAtualizalista_carteira
            .ProcCarregaDadosProdEstImaProc False
        End If
    Else
        Exit Sub
    End If
    Desenho = ""
    .DesenhoMRP = ""
    .MRPgerado = False
    .MRP_Prod = False
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGeraMrpItem()
On Error GoTo tratar_erro

If txtQuantidade <= 0 Then
USMsgBox "Não é permitido emitir ordem de produção para quantidade igual a zero, favor informar a quantidade a produzir", vbCritical, "CAPRIND v5.0"
txtQuantidade.SetFocus
Exit Sub
End If

Permitido1 = False
With frmprod
    .ProcLimpar True
    .Versao_estrutura = cmbVersao_estrutura
    .Versao_processo = cmbVersao
    
    QuantSolicitado = txtQuantidade
    
    Permitido = False
    For Init = 1 To .listaitens.ListItems.Count
        If .listaitens.ListItems.Item(Init).Checked = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Codigo from Empresa where Empresa = '" & .listaitens.ListItems.Item(Init).ListSubItems(2) & "' and Liberar_qtde_MRP = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Permitido = True
            End If
            TBAbrir.Close
            
            GoTo Prosseguir
        End If
    Next Init

Prosseguir:
    'Verifica se a quantidade de todos os produtos selecionados é menor que a quantidade a produzir
    If Permitido = False Then
        qtdeliberada = 0
        For Init = 1 To .listaitens.ListItems.Count
            If .listaitens.ListItems.Item(Init).Checked = True Then
                qtdeliberada = qtdeliberada + .listaitens.ListItems.Item(Init).ListSubItems(12)
            End If
        Next Init
        If QuantSolicitado > qtdeliberada Then
            If USMsgBox("Deseja realmente produzir com uma quantidade maior que a necessidade de estoque?.", vbYesNo, "CAPRINDV5.0") = vbNo Then
            txtQuantidade.SetFocus
            .MRPgerado = False
            Exit Sub
            End If
        End If
    End If
        
    'Verifica se está tudo de acordo
    For Init = 1 To .listaitens.ListItems.Count
        If .listaitens.ListItems.Item(Init).Checked = True Then
            Set TBCarteira = CreateObject("adodb.recordset")
            TBCarteira.Open "Select Desenho, Versao_estrutura, Versao_processo from Carteira_producao where Codigo = " & .listaitens.ListItems.Item(Init), Conexao, adOpenKeyset, adLockReadOnly
            If TBCarteira.EOF = False Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select SubTipoItem, desenho from projproduto where desenho = '" & TBCarteira!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                
                If TBAbrir.EOF = False Then
                    If IsNull(TBAbrir!SubTipoItem) = True Or TBAbrir!SubTipoItem = 0 Then
                        USMsgBox ("Informe o tipo do registro " & TBAbrir!Desenho & " no cadastro antes de gerar o MRP."), vbExclamation, "CAPRIND v5.0"
                        .MRPgerado = False
                        Exit Sub
                    End If
                End If
                TBAbrir.Close
                
                If IsNull(TBCarteira!Versao_estrutura) = False And TBCarteira!Versao_estrutura <> "" And TBCarteira!Versao_estrutura <> .Versao_estrutura Then
                    If USMsgBox("A versão selecionada da estrutura está diferente da versão informada pela engenharia para o produto " & TBCarteira!Desenho & ", deseja prosseguir mesmo assim?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                        .MRPgerado = False
                        Exit Sub
                    End If
                End If
                
                If IsNull(TBCarteira!Versao_processo) = False And TBCarteira!Versao_processo <> "" And TBCarteira!Versao_processo <> .Versao_processo Then
                    If USMsgBox("A versão selecionada do processo está diferente da versão informada pela engenharia para o produto " & TBCarteira!Desenho & ", deseja prosseguir mesmo assim?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                        .MRPgerado = False
                        Exit Sub
                    End If
                End If
            
            End If
        End If
    Next Init
    
    Contador2 = 0
    For Init = 1 To .listaitens.ListItems.Count
        If .listaitens.ListItems.Item(Init).Checked = True Then
            Set TBCarteira = CreateObject("adodb.recordset")
            TBCarteira.Open "Select VC.*, CP.ID_empresa, CP.IDcliente, CP.Necessidade from vendas_carteira VC INNER JOIN Carteira_producao CP ON CP.Codigo = VC.Codigo where VC.Codigo = " & .listaitens.ListItems.Item(Init), Conexao, adOpenKeyset, adLockReadOnly
            If TBCarteira.EOF = False Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Codproduto, Leadtime, SubTipoItem from projproduto where desenho = '" & TBCarteira!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Codproduto = TBAbrir!Codproduto
                    Leadtime = IIf(IsNull(TBAbrir!Leadtime), 0, TBAbrir!Leadtime)
                    SubTipoItem = TBAbrir!SubTipoItem
                Else
                    USMsgBox ("Não foi encontrado nenhum registro cadastrado com este código interno."), vbExclamation, "CAPRIND v5.0"
                    .MRPgerado = False
                    Exit Sub
                End If
                
                'Verifica se é utilizado material consignado no produto do pedido
                If optconsignacao.Value = 1 Then .Consignacao = True Else .Consignacao = False
                
                .MRPgerado = True
                Conexao.Execute "UPDATE vendas_carteira Set OE = 1, datapcp = '" & Date & "' where Codigo = " & TBCarteira!CODIGO
                
                If .MRP_considerar_estoque = True And TBCarteira!Cotacao <> 0 Then
                    ProcEmpenharProdEstoque TBCarteira!ID_empresa, TBCarteira!CODIGO, TBCarteira!Desenho, False, True, .listaitens.ListItems.Item(Init).ListSubItems(12)
                    If QuantSolicitado > 0 Then ProcEmpenharProdProduzindo TBCarteira!ID_empresa, TBCarteira!CODIGO, TBCarteira!Desenho, TBCarteira!PrazoFinal, False
                End If
                
                Contador2 = Contador2 + 1
                If QuantSolicitado > 0 And Contador2 = ContOrdem Then
                    Loteminimo = .ProcVerifQtdLoteMinimo(TBCarteira!Desenho, QuantSolicitado, TBCarteira!ID_empresa, False)
                    
                    If Loteminimo > QuantSolicitado Then
                    QuantSolicitado = Miminimo
                    End If
                    
                    Quant = .ProcVerifQtdLoteMinimo(TBCarteira!Desenho, QuantSolicitado, TBCarteira!ID_empresa, False)
                    
                    If Loteminimo > QuantSolicitado Then
                    QuantSolicitado = Quant
                    End If
                    
                    .ProcGerarOrdemProdutoVendido
                    .OrdemEmpenho1 = OF
                    .ProcNivel1 .OrdemEmpenho1
                End If
                Permitido1 = True
            End If
            TBCarteira.Close
        End If
    Next Init
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkEscopo_Click()
On Error GoTo tratar_erro

'With frmprod
'    If chkEscopo.Value = 1 Then
'    .Escopo = True
'    Else
'    .Escopo = False
'    End If
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkRastreavel_Click()
On Error GoTo tratar_erro

If chkRastreavel.Value = 1 Then
      Chk_etiqueta.Value = 1
      'Chk_etiqueta.Enabled = False
Else
      Chk_etiqueta.Value = 0
      'Chk_etiqueta.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkValidada_Click()
On Error GoTo tratar_erro

If chkValidada.Value = 1 Then
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "select IDUsuario from acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'PCP/Gerenciamento de ordem' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then
        USMsgBox ("Atenção usuário " & pubUsuario & ", você não tem autorização para validar ordem."), vbExclamation, "CAPRIND v5.0"
    Else
        frmprod.Validacao = True
    End If
    TBAcessos.Close
    
Else
    frmprod.Validacao = False
        chkValidada.Caption = "Não"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdGerarmrp_Click()
On Error GoTo tratar_erro

If txtCopiarOP = "" Then
    MsgBox ("Informe o numero de copias da OP!"), vbInformation + vbOKOnly
    txtCopiarOP.Text = "1"
    Exit Sub
End If

ProcGerar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcGerar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboVersãoEstrutura()
On Error GoTo tratar_erro

With cmbVersao_estrutura
    .Clear

        StrSql = "Select PCDV.Versao from Projproduto P INNER JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto where P.Desenho = '" & DesenhoProduto & "' and PCDV.DtValidacao IS NOT NULL group by PCDV.Versao"
        'Debug.print StrSql
       
        Set TBCarregarCombo = CreateObject("adodb.recordset")
        TBCarregarCombo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBCarregarCombo.EOF = False Then
            Do While TBCarregarCombo.EOF = False
                .AddItem TBCarregarCombo!versao
                TBCarregarCombo.MoveNext
            Loop
        TBCarregarCombo.Close
        Else
        .AddItem "A"
        End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosProdEstImaProc(BuscarEst As Boolean)
On Error GoTo tratar_erro
Dim QtdeEstoque As Double 'Quantidade em estoque
Dim QtdeProduzindo As Double 'Quantidade produzindo

Dim QtdeEmpenhoEstVenda As Double 'Quantidade empenhado em vendas
Dim QtdeEmpenhoEst As Double ' Quantidade empenhado em estoque
Dim QtdeEmpenhoProduzindo As Double ' Quantidade empenhado produzindo
Dim Estdisponivel As Double 'Estoque disponivel
Dim qtde_disp_prod As Double 'Produzindo disponivel

TxtEstUN = ""
TxtEstTotal = ""
TxtEstEmpenhado = ""
TxtEstdisponivel = ""
Txt_qtde_prod = ""
Txt_qtde_emp_est = ""
Txt_qtde_disp_prod = ""
'Txt_qtde_total_disp = ""
QtdeEmpenhoEst = 0

If TxtEstCodigo.Text = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Unidade, rastreavel from projproduto where desenho = '" & TxtEstCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TxtEstUN.Text = TBProduto!Unidade
    
    If TBProduto!rastreavel <> "" Then
        If TBProduto!rastreavel = 0 Then chkRastreavel.Value = 0 Else chkRastreavel.Value = 1
        If TBProduto!rastreavel = 0 Then frmprod.chkindividual.Value = 0 Else frmprod.chkindividual.Value = 1
    End If
    
End If
TBProduto.Close

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select ID_solicitacao, Cotacao, Desenho from vendas_carteira where Desenho = '" & TxtEstCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = False Then
'============================================================
'Verifica se o item é de uma solicitação de produção ou não
'============================================================
Set TBNivel1 = CreateObject("adodb.recordset")
If IsNull(TBCarteira!ID_solicitacao) = True Or TBCarteira!ID_solicitacao = 0 Then
  TBNivel1.Open "Select IDcliente, ID_empresa from vendas_proposta where cotacao = " & TBCarteira!Cotacao, Conexao, adOpenKeyset, adLockOptimistic
Else
  TBNivel1.Open "Select ID_empresa from Outros_SolicitacaoPCP where ID = " & TBCarteira!ID_solicitacao, Conexao, adOpenKeyset, adLockOptimistic
End If
If TBNivel1.EOF = False Then
'===================================================
'                     ESTOQUE                      =
'===================================================
' Verifica quantidade em estoque
'===================================================
QtdeEstoque = FunVerificaQtdeEstoque(TxtEstCodigo, TBNivel1!ID_empresa, "")
TxtEstTotal = Format(QtdeEstoque, "###,##0.0000")
'===================================================
' Verifica quantidade em estoque empenhado pra venda
'===================================================
QtdeEmpenhoEstVenda = FunVerificaQtdeEmpenhoEstVenda(TxtEstCodigo, TBNivel1!ID_empresa)
QtdeEmpenhoEst = FunVerificaQtdeEmpenhoEst(TxtEstCodigo, TBNivel1!ID_empresa)
TxtEstEmpenhado.Text = Format(QtdeEmpenhoEstVenda + QtdeEmpenhoEst, "###,##0.0000")
'===================================================
' Carrega quantidade em estoque disponível 1
'===================================================
Estdisponivel = QtdeEstoque - (QtdeEmpenhoEstVenda + QtdeEmpenhoEst)
TxtEstdisponivel.Text = Format(Estdisponivel, "###,##0.0000")
'===================================================
'                     PRODUÇÃO                     =
'===================================================
' Verifica quantidade produzindo
'===================================================
QtdeProduzindo = FunVerificaQtdeProduzindo(TxtEstCodigo, TBNivel1!ID_empresa)
Txt_qtde_prod = Format(QtdeProduzindo, "###,##0.0000")
'===================================================
' Verifica quantidade produzindo já empenhado
'===================================================
QtdeEmpenhoProduzindo = FunVerificaQtdeEmpenhoProduzindo(TxtEstCodigo, TBNivel1!ID_empresa)
Txt_qtde_emp_est = Format(QtdeEmpenhoProduzindo, "###,##0.0000")
'===================================================
' Verifica quantidade produzindo disponivel
'===================================================
qtde_disp_prod = QtdeProduzindo - QtdeEmpenhoProduzindo
Txt_qtde_disp_prod = Format(qtde_disp_prod, "###,##0.0000")
'===================================================
' Apresenta total em estoque
'===================================================
txtEstoque.Text = Format(Estdisponivel + qtde_disp_prod, "###,##0.0000")

End If
TBNivel1.Close
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim Qtvendido As Double
Dim Loteminimo As Double
Dim Desenho As String
Dim EstoqueMinimo As Double
Dim Produzir As Double

ProcCarregaComboVersao cmbVersao, False, False, False, False, ""
'ProcCarregaComboVersao cmbVersao_estrutura, False, False, False, False, ""



cmbVersao = "A"


With frmprod
    If .MRP_Prod = False Then
        If .MRP_utilizaMatConsPI = True Then
            optconsignacao.Value = 1
            .Consignacao = True
            .optconsignacao.Value = 1
        End If
        txtQuantidade = ""
        txtQuantidade.Locked = True
        txtQuantidade.TabStop = False
        Chk_agrupar.Enabled = True
    Else
        ContOrdem = 0
        quantidade = 0
        Permitido = False
        For InitFor = 1 To .listaitens.ListItems.Count
            If .listaitens.ListItems.Item(InitFor).Checked = True Then
                quantidade = quantidade + .listaitens.ListItems(InitFor).ListSubItems(12)
                Qtvendido = Qtvendido + .listaitens.ListItems(InitFor).ListSubItems(7)
                Desenho = .listaitens.ListItems.Item(InitFor).ListSubItems(3)
                
                'Verifica se é utilizado material consignado no produto do pedido
                If Permitido = False Then
                    Set TBCarteira = CreateObject("adodb.recordset")
                    TBCarteira.Open "Select Utiliza_mat_cons from vendas_carteira where Codigo = " & .listaitens.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                    If TBCarteira.EOF = False Then
                        If TBCarteira!Utiliza_mat_cons = True Then
                            optconsignacao.Value = 1
                            Permitido = True
                        End If
                    End If
                    TBCarteira.Close
                End If
                ContOrdem = ContOrdem + 1
            End If
            
            
        Next InitFor
        
Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select Estoque_minimo, Qtde_LoteMinimo, unidade,Descricao from projproduto where Desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockReadOnly
 If TBCarteira.EOF = False Then
  txtLoteminimo.Text = IIf(IsNull(TBCarteira!qtde_LoteMinimo) = False, TBCarteira!qtde_LoteMinimo, 0)
  txtEstoqueMinimo.Text = IIf(IsNull(TBCarteira!Estoque_minimo) = False, TBCarteira!Estoque_minimo, 0)
  txtLoteminimo.Text = Format(txtLoteminimo.Text, "###,##0.0000")
  txtEstoqueMinimo.Text = Format(txtEstoqueMinimo.Text, "###,##0.0000")
  TxtEstUN = TBCarteira!Unidade
  txtdescricao.Text = TBCarteira!Descricao
  Loteminimo = txtLoteminimo.Text
  EstoqueMinimo = txtEstoqueMinimo.Text
 End If
TBCarteira.Close
       
DesenhoProduto = Desenho

ProcCarregaComboVersãoEstrutura
cmbVersao_estrutura.ListIndex = 0

TxtEstCodigo.Text = Desenho
ProcCarregaDadosProdEstImaProc True


If quantidade < Loteminimo And Loteminimo > EstoqueMinimo Then
txtQuantidade = Format(Loteminimo, "###,##0.0000")
Else
txtQuantidade = Format(quantidade, "###,##0.0000")
End If

If quantidade < EstoqueMinimo And EstoqueMinimo > Loteminimo Then
txtQuantidade = Format(EstoqueMinimo, "###,##0.0000")
Else
txtQuantidade = Format(quantidade, "###,##0.0000")
End If

'    txtQuantidade = Format(IIf(quantidade < Loteminimo, Loteminimo, quantidade), "###,##0.0000")
    
    txtQuantidade.Locked = False
    txtqtVendido.Text = Format(Qtvendido, "###,##0.0000")
    txtQuantidade.TabStop = True
    Chk_agrupar.Enabled = False
End If

    Urgencia = False
    .Consignacao = optconsignacao.Value
    OSControlada = False
    Processo_controlado = False
    .EntrarEstoque = Chk_entrar_estoque.Value
    .RetirarEstoque = Chk_retirar_estoque.Value
    .MRP_considerar_estoque = Chk_empenhar_estoque.Value
    .Validacao = chkValidada.Value
    .Agrupar_ordem = Chk_agrupar.Value
    .Versao_por_item_estrutura = Chk_versao_por_item_estrutura.Value
    .Versao_por_item = Chk_versao_por_item.Value
    Tipo_Processo = False
    .MRP_Imprimir_Ordem = Chk_ordem.Value
    .MRP_Imprimir_Ordem_RM = Chk_ordem_rm.Value
    .MRP_Imprimir_Ordem_RM_Res = Chk_ordem_rm_resumido.Value
    .MRP_Imprimir_Ordem_APM = Chk_ordem_manual.Value
    .MRP_Imprimir_Plano = Chk_plano.Value
    .MRP_Imprimir_Freq = Chk_frequencia.Value
    Imprimir = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBuscaLoteMinino()
On Error GoTo tratar_erro

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select Qtde_LoteMinimo from projproduto where codproduto = " & frmprod.listaitens.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockReadOnly
  If TBCarteira.EOF = False Then
   txtLoteminimo.Text = TBCarteira!qtde_LoteMinimo
  End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
Private Sub Opt_processo_controlado_Click()
On Error GoTo tratar_erro

With frmprod
    If Opt_processo_controlado.Value = 1 Then
        Processo_controlado = True
        .Opt_processo_controlado.Value = 1
    Else
        Processo_controlado = False
        .Opt_processo_controlado.Value = 0
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optconsignacao_Click()
On Error GoTo tratar_erro

With frmprod
    If optconsignacao.Value = 1 Then
         .Consignacao = True
        .optconsignacao.Value = 1
    Else
        .Consignacao = False
        .optconsignacao.Value = 0
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptOSControlada_Click()
On Error GoTo tratar_erro

With frmprod
    If OptOSControlada.Value = 1 Then
        OSControlada = True
        .OptOSControlada.Value = 1
    Else
        OSControlada = False
        .OptOSControlada.Value = 0
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCopiarOP_Change()
On Error GoTo tratar_erro
Dim Qtvendido As Double
Dim QTCopias As Integer

Qtvendido = txtqtVendido.Text

If txtCopiarOP <> "" Then
    VerifNumero = txtCopiarOP
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCopiarOP = ""
        txtCopiarOP.SetFocus
        Exit Sub
    End If
End If

If txtCopiarOP = "0" Then
    MsgBox ("A quantidade de copias não pode ser inferior a 1!"), vbCritical + vbOKOnly
    txtCopiarOP.Text = "1"
    Exit Sub
End If

If txtCopiarOP.Text <> "" Then
QTCopias = txtCopiarOP.Text

txtQuantidade.Text = Qtvendido / QTCopias
Else
txtQuantidade.Text = Qtvendido
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

If txtQuantidade <> "" Then
    VerifNumero = txtQuantidade
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade = ""
        txtQuantidade.SetFocus
        Exit Sub
    End If
    QuantSolicitado = IIf(IsNumeric(txtQuantidade) = True, txtQuantidade, 0)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabDesabVisualizandoImpressao()
On Error GoTo tratar_erro

If Chk_ordem.Value = 0 And Chk_ordem_rm.Value = 0 And Chk_ordem_rm_resumido.Value = 0 And Chk_ordem_manual.Value = 0 And Chk_rm.Value = 0 And Chk_plano.Value = 0 And Chk_frequencia.Value = 0 And Chk_etiqueta.Value = 0 Then
    With Chk_visualizar
        .Value = 0
        .Enabled = False
    End With
Else
    Chk_visualizar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

