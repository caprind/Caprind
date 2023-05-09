VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_sucata 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND  v5.0 | Estoque | Movimentação | Gerar sucata/retalho"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   9750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações do estoque"
      Height          =   885
      Left            =   180
      TabIndex        =   48
      Top             =   4200
      Width           =   9345
      Begin VB.TextBox txtSaida 
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
         Left            =   3870
         MaxLength       =   50
         TabIndex        =   52
         ToolTipText     =   "Quantidade de itens a serem baixados do estoque na transformação por unidade de saida."
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox txtUnEntrada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6930
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         ToolTipText     =   "Unidade de estoque."
         Top             =   450
         Width           =   690
      End
      Begin VB.TextBox txtUnSaida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   50
         ToolTipText     =   "Unidade de estoque."
         Top             =   450
         Width           =   630
      End
      Begin VB.TextBox TxtEntrada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Left            =   5730
         MaxLength       =   50
         TabIndex        =   49
         ToolTipText     =   "Quantidade á ser lançado na entrada da sucata ou retalho por unidade de entrada."
         Top             =   450
         Width           =   1185
      End
      Begin DrawSuite2022.USButton btnSalvar 
         Height          =   555
         Left            =   7800
         TabIndex        =   53
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   979
         DibPicture      =   "frmEstoque_sucata.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Transformar"
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
         PicAlign        =   7
         Theme           =   4
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QT.SAIDA"
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
         Left            =   4095
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QT.ENTRADA"
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
         Left            =   5835
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Criar novo produto com código"
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
      Height          =   555
      Left            =   4350
      TabIndex        =   45
      Top             =   2040
      Width           =   5175
      Begin VB.CheckBox OPTnovo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Automático ?"
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
         Left            =   360
         TabIndex        =   47
         Top             =   270
         Width           =   1605
      End
      Begin VB.CheckBox OPTnovoman 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Manual ?"
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
         Left            =   2730
         TabIndex        =   46
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transformar em?"
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   180
      TabIndex        =   42
      Top             =   2040
      Width           =   4155
      Begin VB.CheckBox Chk_nao_baixar_RE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não baixar do RE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2610
         TabIndex        =   56
         Top             =   270
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.OptionButton optRetalho 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Retalho"
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
         Left            =   1350
         TabIndex        =   44
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optSucata 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sucata"
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
         Left            =   240
         TabIndex        =   43
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   714
      DibPicture      =   "frmEstoque_sucata.frx":2D85
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
      Icon            =   "frmEstoque_sucata.frx":146EA
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do item a transformar"
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
      Height          =   1485
      Left            =   180
      TabIndex        =   20
      Top             =   540
      Width           =   9345
      Begin VB.TextBox txtUn_com 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Unidade comercial."
         Top             =   1785
         Width           =   630
      End
      Begin VB.TextBox txtQtde_PC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7660
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Quantidade de peças."
         Top             =   1005
         Width           =   1490
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Numero da RE."
         Top             =   1005
         Width           =   1490
      End
      Begin VB.TextBox txtDesenho 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Código interno."
         Top             =   465
         Width           =   1815
      End
      Begin VB.TextBox txtDescricao 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   8
         ToolTipText     =   "Descrição."
         Top             =   465
         Width           =   7140
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1676
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Número do lote."
         Top             =   1005
         Width           =   1490
      End
      Begin VB.TextBox txtCertificado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4668
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Certificado."
         Top             =   1005
         Width           =   1490
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6164
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Quantidade."
         Top             =   1005
         Width           =   1490
      End
      Begin VB.TextBox txtCorrida 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3172
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Corrida."
         Top             =   1005
         Width           =   1490
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. com."
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
         Left            =   2655
         TabIndex        =   33
         Top             =   1590
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade PÇ"
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
         Left            =   7860
         TabIndex        =   32
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RE"
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
         Left            =   810
         TabIndex        =   31
         Top             =   810
         Width           =   225
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
         Left            =   5235
         TabIndex        =   26
         Top             =   270
         Width           =   690
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
         Left            =   630
         TabIndex        =   25
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° lote"
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
         Left            =   2160
         TabIndex        =   24
         Top             =   810
         Width           =   525
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
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
         Left            =   5010
         TabIndex        =   23
         Top             =   810
         Width           =   795
      End
      Begin VB.Label Label8 
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
         Left            =   6495
         TabIndex        =   22
         Top             =   810
         Width           =   840
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
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
         Left            =   3645
         TabIndex        =   21
         Top             =   810
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do novo item (Sucata ou retalho)"
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
      Height          =   1545
      Left            =   180
      TabIndex        =   27
      Top             =   2640
      Width           =   9345
      Begin VB.TextBox txtUN_com_sucata 
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
         Left            =   8535
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "Unidade comercial."
         Top             =   1770
         Width           =   630
      End
      Begin VB.TextBox txtQtde_PC_Sucata 
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
         Left            =   4500
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Quantidade de peças."
         Top             =   2790
         Width           =   1490
      End
      Begin VB.TextBox txtDureza 
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
         Left            =   4677
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Duzera."
         Top             =   2910
         Width           =   1490
      End
      Begin VB.TextBox txtEspessura 
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Espessura (mm)."
         Top             =   2910
         Width           =   1215
      End
      Begin VB.TextBox txtLargura 
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
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Largura (mm)."
         Top             =   2910
         Width           =   1125
      End
      Begin VB.TextBox txtComprimento 
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
         Left            =   3178
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Comprimento (mm)."
         Top             =   2910
         Width           =   1490
      End
      Begin VB.ComboBox cmbfamilia 
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
         Left            =   4905
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   480
         Width           =   4290
      End
      Begin VB.TextBox txtRev_cod 
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
         Left            =   2370
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Revisão do produto."
         Top             =   480
         Width           =   525
      End
      Begin VB.ComboBox cmbN_ref 
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
         Left            =   2910
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Código de referência."
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtDesc_sucata 
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
         MaxLength       =   150
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   1080
         Width           =   8970
      End
      Begin VB.TextBox txtDesenho_sucata 
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
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   1785
      End
      Begin DrawSuite2022.USButton cmdDesenho 
         Height          =   315
         Left            =   2010
         TabIndex        =   41
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_sucata.frx":14A04
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
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
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade PÇ"
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
         Left            =   4695
         TabIndex        =   39
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dureza"
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
         Left            =   5160
         TabIndex        =   38
         Top             =   2700
         Width           =   510
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprimento / mm"
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
         Left            =   3255
         TabIndex        =   37
         Top             =   2700
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Largura / mm"
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
         Left            =   1590
         TabIndex        =   36
         Top             =   2700
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espessura / mm"
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
         Left            =   360
         TabIndex        =   35
         Top             =   2700
         Width           =   1125
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
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
         TabIndex        =   34
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de referência"
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
         Left            =   3165
         TabIndex        =   30
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
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
         Left            =   585
         TabIndex        =   29
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label5 
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
         Left            =   4005
         TabIndex        =   28
         Top             =   870
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmEstoque_sucata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id_sucata As String 'ok

Private Sub btnSalvar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente transformar esse item?", vbYesNo, "CAPRIND v5.0") = vbYes Then
ProcSalvar
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenho_Click()
On Error GoTo tratar_erro

frmEstoque_localizar_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
EstoqueSaida = 0
EstoqueEntrada = 0
EstoqueSaldo = 0

quantestoque = 0
Total = 0
ValorTotal = 0
Qtde = 0
QtdePC = 0
Qtd = 0
quantidade = 0
QuantidadePC = 0

Acao = "salvar"
If OPTnovo.Value = 0 And txtDesenho_sucata = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtDesenho_sucata.SetFocus
    Exit Sub
End If

If txtDesc_sucata.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDesc_sucata.SetFocus
    Exit Sub
End If
If cmbfamilia.Text = "" Then
    NomeCampo = "a familia"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If

EstoqueEntrada = IIf(txtEntrada = "", 0, txtEntrada)
EstoqueSaida = IIf(txtSaida = "", 0, txtSaida)

If EstoqueEntrada <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtEntrada.SetFocus
    Exit Sub
End If

If EstoqueEmpenho > 0 Then
    USMsgBox ("Não é permitido criar " & IIf(optSucata.Value = True, "sucata", "retalho") & ", pois o RE possui empenho."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If optSucata.Value = True Then
    EstoqueSaldo = txtQtde
    
    If (EstoqueSaida > EstoqueSaldo) And saida_sucata = True Then
        USMsgBox ("A quantidade de saida não pode ser maior que a quantidade do lote."), vbExclamation, "CAPRIND v5.0"
        txtSaida.SetFocus
        Exit Sub
    End If
End If

'If txtQtde_PC <> "" And txtQtde_PC <> "0,0000" And txtQtde_PC_Sucata = "" Then
'    NomeCampo = "a quantidade pç"
'    ProcVerificaAcao
'    txtQtde_PC_Sucata.SetFocus
'    Exit Sub
'End If

'Se for novo produto
If OPTnovo.Value = 1 Then ProcNovoProduto
If OPTnovoman.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select desenho from projproduto where desenho = '" & txtDesenho_sucata.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtDesenho_sucata.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoMan
End If

'================================================================
'Executa entrada da sucata no estoque
'================================================================
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "select * from Estoque_Controle where idestoque = " & txtID, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    Qtde = 0
    QtdePC = 0
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from estoque_controle where idlote_sucata = " & txtID & " and Desenho = '" & txtDesenho_sucata & "' and Lote = '" & txtLote & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        TBGravar!ID_empresa = TBEstoque!ID_empresa
    Else
        EstoqueSaldo = IIf(IsNull(TBGravar!estoque_real), 0, TBGravar!estoque_real)
        QtdePC = IIf(IsNull(TBGravar!estoque_real_PC), 0, TBGravar!estoque_real_PC)
    End If
'==========================================================================
' Atualiza saldo da sucata no estoque
'==========================================================================
 EstoqueSaldo = Format(EstoqueSaldo + EstoqueEntrada, "###,##0.0000")
 QuantidadePC = QtdePC + QtdeSaida
'================================================
' Busca dados do item "Sucata" no cadastro
'================================================
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "select classe, PConsumo from projproduto where desenho = '" & txtDesenho_sucata & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        TBGravar!Classe = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
        TBGravar!valor_unitario = IIf(IsNull(TBItem!PConsumo), "", TBItem!PConsumo)
        ValorTotal = IIf(IsNull(TBItem!PConsumo), "", TBItem!PConsumo)
        TBGravar!Valor_total = Format(quantidade * ValorTotal, "###,##0.00")
    End If
    TBItem.Close
'===============================================
    TBGravar!idLote_sucata = txtID
    TBGravar!Desenho_sucata = txtdesenho
    If optSucata.Value = True Then
        TBGravar!LOTE = txtLote
        TBGravar!status = "ENTRADA_SUCATA"
        Evento = "Nova sucata"
    Else
        TBGravar!LOTE = txtID
        TBGravar!status = "ENTRADA_RETALHO"
        Evento = "Novo retalho"
    End If
    TBGravar!estoque_real = Format(EstoqueEntrada, "###,##0.0000")
    TBGravar!estoque_real_PC = QuantidadePC
    TBGravar!Qtde = TBGravar!estoque_real
    TBGravar!estoque_venda = TBGravar!estoque_real
    TBGravar!Desenho = txtDesenho_sucata
    TBGravar!Descricao = txtDesc_sucata
    TBGravar!Un = txtUnEntrada
    TBGravar!Data = Format(Date, "dd/mm/yy")
    TBGravar!Responsavel = pubUsuario
    TBGravar!Fornecedor = IIf(IsNull(TBEstoque!Fornecedor), "", TBEstoque!Fornecedor)
    TBGravar!Certificado = IIf(IsNull(TBEstoque!Certificado), "", TBEstoque!Certificado)
    TBGravar!Classe = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
    TBGravar!descricaotecnica = TBEstoque!descricaotecnica
    TBGravar!local_armaz = IIf(IsNull(TBEstoque!local_armaz), "", TBEstoque!local_armaz)
    TBGravar!ID_Cliente = IIf(IsNull(TBEstoque!ID_Cliente), "", TBEstoque!ID_Cliente)
    TBGravar!Cliente = TBEstoque!Cliente
    TBGravar!Ref = IIf(cmbN_ref = "", Null, cmbN_ref)
    TBGravar!Corrida = IIf(IsNull(TBEstoque!Corrida), "", TBEstoque!Corrida)
    If TBEstoque!consignacao = True Then TBGravar!consignacao = True
    TBGravar.Update
    id_sucata = TBGravar!IDEstoque
    
    ProcMovimentacao IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario), ValorTotal
    
    If Chk_nao_baixar_RE.Visible = False Or Chk_nao_baixar_RE.Value = 0 Then
        TBEstoque!estoque_real = IIf(TBEstoque!estoque_real - txtSaida.Text < 0, 0, Format(TBEstoque!estoque_real - txtSaida.Text, "###,##0.00"))
        TBEstoque!estoque_real_PC = TBEstoque!estoque_real 'IIf(IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - txtSaida < 0, 0, IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - txtSaida)
        TBEstoque!estoque_venda = TBEstoque!estoque_real
        TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!estoque_real), 0, TBEstoque!estoque_real) * IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario), "###,##0.00")
        TBEstoque.Update
    End If
    TBGravar.Close
End If
TBEstoque.Close
If optSucata.Value = True Then USMsgBox ("Nova sucata cadastrada com sucesso."), vbInformation, "CAPRIND v5.0" Else USMsgBox ("Novo retalho cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Estoque/Movimentação"
ID_documento = txtID
Documento = "Cód. interno de: " & txtdesenho
Documento1 = "Cód. interno para: " & txtDesenho_sucata
ProcGravaEvento
'==================================

'Se for retalho já faz o cadastro de similar
If optRetalho.Value = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID_similar from Projproduto where Desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        IDlista = IIf(IsNull(TBAbrir!ID_similar), 0, TBAbrir!ID_similar)
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID_similar from Projproduto where Desenho = '" & txtDesenho_sucata & "' and ID_similar IS NOT NULL and ID_similar <> 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Projproduto_similar where ID = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = True Then
            TBGravar.AddNew
            TBGravar.Update
            IDlista = TBGravar!ID
        End If
        TBGravar.Close
    Else
        IDlista = TBAbrir!ID_similar
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select ID_similar from Projproduto where Desenho = '" & txtdesenho & "' and ID_similar IS NOT NULL and ID_similar <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Conexao.Execute "UPDATE Projproduto Set ID_similar = " & IDlista & " where ID_similar = " & TBFI!ID_similar
            IDlista = FunVerifApagaIDSimilar(TBFI!ID_similar)
        End If
        TBFI.Close
    End If
    TBAbrir.Close
    Conexao.Execute "UPDATE Projproduto Set ID_similar = " & IDlista & " where Desenho = '" & txtdesenho & "'"
    Conexao.Execute "UPDATE Projproduto Set ID_similar = " & IDlista & " where Desenho = '" & txtDesenho_sucata & "'"
End If
Unload Me

Exit Sub
tratar_erro:
    'If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMovimentacao(VlrUnit_Lote As Double, VlrUnit_Sucata As Double)
On Error GoTo tratar_erro

'================================================================
'Cria movimentação de saída do item do estoque
'================================================================
If TBEstoque!estoque_real > 0 And (Chk_nao_baixar_RE.Visible = False Or Chk_nao_baixar_RE.Value = 0) Then
    Set TBMaterial = CreateObject("adodb.recordset")
    TBMaterial.Open "select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
    TBMaterial.AddNew
    TBMaterial!Destino = "Interno"
    TBMaterial!Terceiros = False
    TBMaterial!IDEstoque = txtID
    TBMaterial!Operacao = IIf(optSucata.Value = True, "SAIDA_SUCATA", "SAIDA_RETALHO")
    TBMaterial!Requisitante = pubUsuario
    TBMaterial!Desenho = txtdesenho
    TBMaterial!Descricao = txtdescricao
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "select Classe from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        TBMaterial!Familia = TBItem!Classe
    End If
    TBItem.Close
    TBMaterial!Data = Date
    'If optSucata.Value = True Then
        TBMaterial!Saida = Format(txtSaida.Text, "0.000")
        TBMaterial!Saida_PC = Format(IIf(txtQtde_PC_Sucata = "", 0, txtQtde_PC_Sucata), "0.000")
    'Else
        'TBMaterial!Saida = Format(TBEstoque!estoque_real, "0.000")
        'TBMaterial!Saida_PC = Format(TBEstoque!estoque_real_PC, "0.000")
    'End If
    TBMaterial!Entrada = 0
    TBMaterial!Entrada_PC = 0
    TBMaterial!Responsavel = pubUsuario
    'TBMaterial!Cliente = frmestoque_item.Lista.SelectedItem.SubItems(12)
    TBMaterial!LOTE = txtLote
    TBMaterial!Documento = txtDesenho_sucata
    TBMaterial!DtEmissao = Date
    TBMaterial!VlrUnit = VlrUnit_Lote
    TBMaterial!VlrTotal = Format(TBMaterial!Saida * TBMaterial!VlrUnit, "0.00")
    TBMaterial.Update
    TBMaterial.Close
End If

'================================================================
'Cria movimentação de entrada da Sucata/Retalho no estoque
'================================================================
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBProduto.AddNew
TBProduto!Destino = "Interno"
TBProduto!Terceiros = False
TBProduto!IDEstoque = id_sucata
If optSucata.Value = True Then
    TBProduto!Operacao = "ENTRADA_SUCATA"
    TBProduto!LOTE = txtLote
Else
    TBProduto!Operacao = "ENTRADA_RETALHO"
    TBProduto!LOTE = txtID
End If
TBProduto!Desenho = txtDesenho_sucata
TBProduto!Descricao = txtDesc_sucata

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "select Classe from projproduto where desenho = '" & txtDesenho_sucata & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then TBProduto!Familia = TBItem!Classe
TBItem.Close
TBProduto!Data = Date
TBProduto!Saida = 0
TBProduto!Saida_PC = 0
TBProduto!Entrada = Format(txtEntrada.Text, "0.000")
TBProduto!Entrada_PC = IIf(txtQtde_PC_Sucata = "", 0, Format(txtQtde_PC_Sucata, "0.000"))
TBProduto!Responsavel = pubUsuario
'TBProduto!Cliente = frmestoque_item.Lista.SelectedItem.SubItems(12)
TBProduto!Documento = txtDesenho_sucata
TBProduto!DtEmissao = Date
TBProduto!VlrUnit = VlrUnit_Sucata
TBProduto!VlrTotal = Format(TBProduto!Entrada * TBProduto!VlrUnit, "0.00")
TBProduto.Update
TBProduto.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

'ProcCarregaToolBar1 Me, 9345, 5, True
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
txtID = RE
'valor = 0 'Kg/Un
'Valor3 = 0 'Peso bruto

'=====================================================================
'Carrega os dados do produto a ser transformado em sucata/Retalho
'=====================================================================
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select EC.LOTE, EC.Corrida, EC.Certificado, PP.Unidade, EC.Estoque_real_PC, EC.Estoque_real, PP.descricaotecnica, EC.Desenho, PP.Unidade_com, PP.Descricao from Estoque_controle as EC inner join Projproduto as PP on EC.Desenho = PP.Desenho where IdEstoque = " & RE, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtLote = IIf(IsNull(TBProduto!LOTE), "", TBProduto!LOTE)
    txtcorrida = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
    txtCertificado = IIf(IsNull(TBProduto!Certificado), "", TBProduto!Certificado)
    txtdesenho = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
    txtUnSaida = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
    txtQtde = IIf(IsNull(TBProduto!estoque_real), "0,0000", Format(TBProduto!estoque_real, "###,##0.0000"))
    txtQtde_PC = IIf(IsNull(TBProduto!estoque_real_PC), "0,0000", Format(TBProduto!estoque_real_PC, "###,##0.0000"))
    txtUn_com = IIf(IsNull(TBProduto!Unidade_com), "", TBProduto!Unidade_com)
    txtdescricao = IIf(IsNull(TBProduto!descricaotecnica), "", TBProduto!descricaotecnica)
End If
TBProduto.Close

'====================================================================================
' Verifica se a empresa movimenta estoque por peça
'====================================================================================
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select Movimentar_estoque_pc from Empresa where Codigo = " & IDempresa & " and Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False And txtQtde_PC > 0 Then
  With txtEntrada
      .Locked = True
      .TabStop = False
  End With
  With txtQtde_PC_Sucata
      .Locked = False
      .TabStop = True
  End With
Else
  With txtEntrada
      .Locked = False
      .TabStop = True
  End With
  With txtQtde_PC_Sucata
      .Locked = True
      .TabStop = False
  End With
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPTnovo_Click()
On Error GoTo tratar_erro

If OPTnovo.Value = 1 Then
    OPTnovoman.Value = 0
    txtDesenho_sucata.Locked = True
    txtDesenho_sucata.TabStop = False
    txtDesenho_sucata.Text = ""
    Procliberacampos
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPTnovoman_Click()
On Error GoTo tratar_erro

If OPTnovoman.Value = 1 Then
    OPTnovo.Value = 0
    Procliberacampos
    USMsgBox ("Informe o código interno do produto."), vbInformation, "CAPRIND v5.0"
    txtDesenho_sucata.Locked = False
    txtDesenho_sucata.TabStop = True
    txtDesenho_sucata.Text = ""
    txtDesenho_sucata.SetFocus
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optRetalho_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSucata_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_Change()
On Error GoTo tratar_erro

If txtComprimento.Text <> "" Then
    VerifNumero = txtComprimento.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComprimento.Text = ""
        txtComprimento.SetFocus
        Exit Sub
    End If
    procCalcula
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_LostFocus()
On Error GoTo tratar_erro

txtComprimento = Format(txtComprimento, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_sucata_Change()
On Error GoTo tratar_erro

ProcCarregaComboCodRef cmbN_ref, "P.desenho = '" & txtDesenho_sucata & "'", 0, "", False, True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_Change()
On Error GoTo tratar_erro

If txtespessura.Text <> "" Then
    VerifNumero = txtespessura.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtespessura.Text = ""
        txtespessura.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_LostFocus()
On Error GoTo tratar_erro

txtespessura = Format(txtespessura, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_Change()
On Error GoTo tratar_erro

If txtLargura.Text <> "" Then
    VerifNumero = txtLargura.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLargura.Text = ""
        txtLargura.SetFocus
        Exit Sub
    End If
    procCalcula
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_LostFocus()
On Error GoTo tratar_erro

txtLargura = Format(txtLargura, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_PC_Sucata_Change()
On Error GoTo tratar_erro

If txtQtde_PC_Sucata.Text <> "" Then
    VerifNumero = txtQtde_PC_Sucata.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_PC_Sucata.Text = ""
        txtQtde_PC_Sucata.SetFocus
        Exit Sub
    End If
    If txtQtde_PC_Sucata.Locked = False Then procCalcula
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_PC_Sucata_LostFocus()
On Error GoTo tratar_erro

txtQtde_PC_Sucata = Format(txtQtde_PC_Sucata, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtentrada_Change()
On Error GoTo tratar_erro

If txtEntrada.Text <> "" Then
    VerifNumero = txtEntrada.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtEntrada.Text = ""
        txtEntrada.SetFocus
        Exit Sub
    End If
    'If TxtEntrada.Locked = False Then procCalcula
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtentrada_LostFocus()
On Error GoTo tratar_erro

txtEntrada = Format(txtEntrada, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With txtDesenho_sucata
    .Locked = True
    .TabStop = False
End With
With txtRev_cod
    .Locked = True
    .TabStop = False
End With
With txtDesc_sucata
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
With txtespessura
    .Locked = True
    .TabStop = False
End With
With txtComprimento
    .Locked = True
    .TabStop = False
End With
With txtLargura
    .Locked = True
    .TabStop = False
End With
With txtDureza
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
txtUnEntrada = ""
txtUN_com_sucata = ""
txtespessura = ""
txtDureza = ""
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procliberacampos()
On Error GoTo tratar_erro

With txtRev_cod
    .Locked = False
    .TabStop = True
End With
With txtDesc_sucata
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
With txtespessura
    .Locked = False
    .TabStop = True
End With
With txtComprimento
    .Locked = False
    .TabStop = True
End With
With txtLargura
    .Locked = False
    .TabStop = True
End With
With txtDureza
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
txtUnEntrada = txtUnSaida
txtUN_com_sucata = txtUn_com
If optSucata.Value = True Then
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Vendas = 'True'", False
Else
    ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
End If
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select Espessura, Dureza, classe from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    txtespessura = IIf(IsNull(TBMaterial!Espessura), "", Format(TBMaterial!Espessura, "###,##0.00"))
    txtDureza = IIf(IsNull(TBMaterial!Dureza), "", TBMaterial!Dureza)
    cmbfamilia = IIf(IsNull(TBMaterial!Classe), "", TBMaterial!Classe)
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then Exit Sub
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcLimpaCampos()
On Error GoTo tratar_erro

If optSucata.Value = True Then Chk_nao_baixar_RE.Visible = False Else Chk_nao_baixar_RE.Visible = True
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
OPTnovo.Value = 0
OPTnovoman.Value = 0
txtDesenho_sucata = ""
txtRev_cod = 0
cmbN_ref.Clear
txtDesc_sucata = ""
txtUnEntrada = ""
txtUN_com_sucata = ""
txtespessura = ""
txtLargura = ""
txtComprimento = ""
txtDureza = ""
txtEntrada = ""
txtQtde_PC_Sucata = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProduto()
On Error GoTo tratar_erro

If optSucata Then
    txtDesenho_sucata = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtDesenho_sucata, cmbN_ref, 0, txtDesc_sucata, txtDesc_sucata, cmbfamilia, 0, 0, 0, txtUnEntrada, txtUN_com_sucata, 0, False, True, True, False, 1, "P", "", IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, 0, "", "")
Else
    txtDesenho_sucata = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtDesenho_sucata, cmbN_ref, 0, txtDesc_sucata, txtDesc_sucata, cmbfamilia, 0, 0, 0, txtUnEntrada, txtUN_com_sucata, 0, False, False, True, False, 0, "P", "", IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, 0, "", "")
End If
procSalvaPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoMan()
On Error GoTo tratar_erro

If optSucata Then
    txtDesenho_sucata = FunCriaNovoProdServ(True, "", txtDesenho_sucata, cmbN_ref, 0, txtDesc_sucata, txtDesc_sucata, cmbfamilia, 0, 0, 0, txtUnEntrada, txtUN_com_sucata, 0, False, True, True, False, 1, "P", "", IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, 0, "", "")
Else
    txtDesenho_sucata = FunCriaNovoProdServ(True, "", txtDesenho_sucata, cmbN_ref, 0, txtDesc_sucata, txtDesc_sucata, cmbfamilia, 0, 0, 0, txtUnEntrada, txtUN_com_sucata, 0, False, False, True, False, 0, "P", "", IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, 0, "", "")
End If
procSalvaPesoBruto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procSalvaPesoBruto()
On Error GoTo tratar_erro

Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select peso_metro, un_kg, peso_metro from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "select peso_metro, un_kg, peso_metro, PBruto from projproduto where desenho = '" & txtDesenho_sucata & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        TBItem!peso_metro = valor
        TBItem!Un_Kg = TBMaterial!Un_Kg
        TBItem!PBruto = Valor3
        TBItem.Update
    End If
    TBItem.Close
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCalcula()
On Error GoTo tratar_erro

valor = 0 'Kg/Un
Valor1 = 0 'Largura
Valor2 = 0 'Comprimento
Valor3 = 0 'Peso bruto
Qtde = 0
Qtd = 0
quantnovo = 0
QuantSolicitado = 0
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select peso_metro, un_kg, peso_metro, Unidade, SubTipoItem, Comprimento, Largura, PBruto from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    If OPTnovo.Value = 1 Or OPTnovoman.Value = 1 Then
        'Calcular Peso bruto da RE de origem
        quantnovo = IIf(txtQtde = "", 0, txtQtde)
        QuantSolicitado = IIf(txtQtde_PC = "", 0, txtQtde_PC)
        If QuantSolicitado <> 0 Then Qtde = quantnovo / QuantSolicitado Else Qtde = IIf(TBMaterial!PBruto = "", 0, TBMaterial!PBruto)
        
        'Calcular Kg/ un da RE de origem
        If Qtde <> 0 Then
            If TBMaterial!Un_Kg = "Mt²" Or TBMaterial!Un_Kg = "Mt/L" Then
                Valor1 = IIf(IsNull(TBMaterial!Comprimento), 0, TBMaterial!Comprimento)
                If TBMaterial!Un_Kg = "Mt/L" Then
                    If Valor1 <> 0 Then Qtd = Valor1 / 1000
                ElseIf Valor1 <> 0 Then
                        Valor2 = IIf(IsNull(TBMaterial!Largura), 0, TBMaterial!Largura)
                        If Valor2 <> 0 Then Qtd = (Valor1 / 1000) * (Valor2 / 1000)
                End If
                If Qtd <> 0 Then valor = Format(Qtde / Qtd, "###,##0.0000000000")
            Else
                valor = Qtde
            End If
            
            'Calculo peso bruto
            Valor1 = IIf(txtLargura = "", 0, txtLargura)
            Valor2 = IIf(txtComprimento = "", 0, txtComprimento)
            If TBMaterial!Un_Kg = "Mt²" Then
                If Valor1 <> 0 And Valor2 <> 0 Then Valor3 = Format(((valor * Valor2) / 1000) * (Valor1 / 1000), "###,##0.000000")
            ElseIf TBMaterial!Un_Kg = "Mt/L" Then
                    If Valor2 <> 0 Then Valor3 = Format(((valor * Valor2) / 1000), "###,##0.000000")
                Else
                    Valor3 = Format(valor, "###,##0.000000")
            End If
        End If
    Else
        'Kg/Un e peso bruto da RE
        If TBMaterial!peso_metro <> "" Or IsNull(TBMaterial!peso_metro) = False Then valor = TBMaterial!peso_metro Else valor = 0
        If TBMaterial!PBruto <> "" Or IsNull(TBMaterial!PBruto) = False Then Valor3 = TBMaterial!PBruto Else Valor3 = 0
    End If
    
    'Calcula quantidade PÇ ou quantidade
    If TBMaterial!Unidade = "KG" Or TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then
        If TBMaterial!Unidade = "KG" And (TBMaterial!Un_Kg = "Mt²" Or TBMaterial!Un_Kg = "Mt/L") Then
            If Valor3 > 0 Then
                If txtEntrada.Locked = False Then
                    txtQtde_PC_Sucata = 0
                Else
                    If txtQtde_PC_Sucata <> "" Then txtEntrada = Format(txtQtde_PC_Sucata * Valor3, "###,##0.0000") Else txtEntrada = 0
                End If
            Else
                If txtEntrada.Locked = False Then txtQtde_PC_Sucata = 0 Else txtEntrada = 0
            End If
        Else
            If TBMaterial!SubTipoItem = 1 Or TBMaterial!SubTipoItem = 2 Or TBMaterial!SubTipoItem = 3 Then
                If txtEntrada.Locked = False Then
                    txtQtde_PC_Sucata = 0
                Else
                    txtEntrada = IIf(txtQtde_PC_Sucata = "", 0, txtQtde_PC_Sucata)
                End If
            Else
                If txtEntrada.Locked = False Then txtQtde_PC_Sucata = 0 Else txtEntrada = 0
            End If
        End If
    End If
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSaida_Change()
On Error GoTo tratar_erro
'=================================================================
' Busca Unidade do item que será transofmrado em sucata ou retalho
'=================================================================
' Se for unidade de entrada em quilo busca o peso por item
If txtUnSaida = "PÇ" And txtUnEntrada = "KG" Then
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "select PBruto, Pliquido from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then

'=================================================================
EstoqueSaida = IIf(txtSaida.Text = "", 0, txtSaida.Text)

If TBMaterial!PBruto <> 0 Then
EstoqueEntrada = EstoqueSaida * TBMaterial!PBruto
Else
EstoqueEntrada = EstoqueSaida * TBMaterial!PLiquido
End If

txtEntrada = EstoqueEntrada
End If
TBMaterial.Close
End If
'=================================================================
' Se a unidade de saida for igual a unidade de entrada
'=================================================================
If txtUnSaida = txtUnEntrada Then
txtEntrada = txtSaida
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
