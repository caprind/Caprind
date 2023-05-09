VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_Pedido_WEB 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Vendas | Detalhes do pedido WEB"
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   150
      TabIndex        =   23
      Top             =   2700
      Width           =   10575
      Begin VB.TextBox txtid 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtvlrTotal 
         Alignment       =   2  'Center
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
         Left            =   7710
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   1110
      End
      Begin VB.TextBox txtQuantidade 
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
         Height          =   285
         Left            =   6840
         MaxLength       =   50
         TabIndex        =   30
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   870
      End
      Begin VB.TextBox txtvlrUnit 
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
         Height          =   285
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   29
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   870
      End
      Begin VB.TextBox txtDescricao 
         Alignment       =   2  'Center
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
         Left            =   1890
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   4080
      End
      Begin VB.TextBox txtUnidade 
         Alignment       =   2  'Center
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
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   360
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
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
         Left            =   750
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   480
         Width           =   720
      End
      Begin DrawSuite2022.USButton btnNovo 
         Height          =   525
         Left            =   8850
         TabIndex        =   36
         Top             =   270
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   926
         DibPicture      =   "frmVendas_Pedido_WEB.frx":0000
         Caption         =   "Novo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   8
         PicSize         =   1
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton btnGravar 
         Height          =   525
         Left            =   9405
         TabIndex        =   37
         Top             =   270
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   926
         DibPicture      =   "frmVendas_Pedido_WEB.frx":62E4
         Caption         =   "Gravar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   8
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton btnExcluir 
         Height          =   525
         Left            =   9960
         TabIndex        =   38
         Top             =   270
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
         DibPicture      =   "frmVendas_Pedido_WEB.frx":ECE9
         Caption         =   "Excluir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   8
         ShowFocusRect   =   0   'False
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "vlr Total"
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
         Left            =   7973
         TabIndex        =   35
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
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
         Left            =   7050
         TabIndex        =   34
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "vlr Unitário"
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
         Left            =   6030
         TabIndex        =   33
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label12 
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
         Height          =   195
         Left            =   3390
         TabIndex        =   32
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UN"
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
         Left            =   1590
         TabIndex        =   27
         Top             =   270
         Width           =   210
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   885
         TabIndex        =   25
         Top             =   270
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   150
      TabIndex        =   16
      Top             =   8070
      Width           =   10575
      Begin VB.TextBox txtValorTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   18
         ToolTipText     =   "Código interno."
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total :"
         Height          =   285
         Left            =   8190
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   8895
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   688
      DibPicture      =   "frmVendas_Pedido_WEB.frx":18535
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
      Icon            =   "frmVendas_Pedido_WEB.frx":1F6B5
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do pedido WEB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   150
      TabIndex        =   0
      Top             =   510
      Width           =   10575
      Begin VB.TextBox txtUF 
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
         Height          =   285
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox txtCPF_CNPJ 
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
         Height          =   285
         Left            =   6600
         MaxLength       =   50
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   540
         Width           =   1670
      End
      Begin VB.TextBox txtCondicoes 
         Alignment       =   2  'Center
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
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   1140
         Width           =   4260
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
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
         Left            =   9480
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
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
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   540
         Width           =   1190
      End
      Begin VB.TextBox txtObservacoes 
         Alignment       =   2  'Center
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
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   1710
         Width           =   10170
      End
      Begin VB.TextBox txtContato 
         Alignment       =   2  'Center
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
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   1140
         Width           =   2270
      End
      Begin VB.TextBox txtVendedor 
         Alignment       =   2  'Center
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
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   1140
         Width           =   3620
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
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
         Left            =   1110
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   540
         Width           =   5205
      End
      Begin VB.TextBox txtPedidoWEB 
         Alignment       =   2  'Center
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
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   540
         Width           =   890
      End
      Begin VB.Label Label17 
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
         Left            =   6390
         TabIndex        =   43
         Top             =   330
         Width           =   195
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF | CNPJ"
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
         Left            =   7035
         TabIndex        =   41
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   8648
         TabIndex        =   19
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condições de pagamento"
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
         Left            =   7350
         TabIndex        =   14
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato"
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
         Left            =   4710
         TabIndex        =   12
         Top             =   930
         Width           =   585
      End
      Begin VB.Label Label7 
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
         Height          =   195
         Left            =   9758
         TabIndex        =   11
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
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
         Left            =   4748
         TabIndex        =   7
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   930
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   3465
         TabIndex        =   5
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "n° Pedido"
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
         Left            =   345
         TabIndex        =   4
         Top             =   330
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView ListaProdutos 
      Height          =   4440
      Left            =   150
      TabIndex        =   3
      Top             =   3630
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7832
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
         Text            =   "ID"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Un"
         Object.Width           =   884
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   7057
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   2295
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Quantidade"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valor total"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmVendas_Pedido_WEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExcluir_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente excluir esse item da lista de produtos do pedido?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    FunAbreBDSite
    
    If ConexaoMySql.State = 1 Then
        ConexaoMySql.Execute ("Delete * FROM Vendas_Pedido_Lista Where ID_Lista = '" & txtId.Text & "'")
        USMsgBox "Item excluido com sucesso!", vbInformation, "CAPRIND v5.0"
        ProcCarregaListaProdutos
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnGravar_Click()
On Error GoTo tratar_erro

ProcSalvarProduto
ProcCarregaListaProdutos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaPedidoWEB
ProcCarregaListaProdutos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvarProduto()
On Error GoTo tratar_erro

'=================================================================
' Abrir BD na WEB
'=================================================================
FunAbreBDSite

If ConexaoMySql.State = 1 Then

StrSql = "SELECT * FROM Vendas_Pedido_Lista Where ID_Lista = '" & txtId.Text & "'"

Set TBLISTA = New ADODB.Recordset
'=================================================================
' Buscar produtos do pedido na WEB
'=================================================================
TBLISTA.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
If TBLISTA.EOF = True Then
TBLISTA.AddNew
End If

Dim VlrUnit As String
Dim vlrTotal As String
Dim QtrTotal As String

VlrUnit = Replace(txtvlrUnit.Text, ".", "")
VlrUnit = Replace(VlrUnit, ",", ".")

QtrTotal = Replace(txtQuantidade.Text, ".", "")
QtrTotal = Replace(QtrTotal, ",", ".")

vlrTotal = Replace(txtvlrTotal, ".", "")
vlrTotal = Replace(vlrTotal, ",", ".")

TBLISTA!CODIGO = txtCodigo.Text
TBLISTA!Unidade = txtunidade.Text
TBLISTA!Descricao = txtdescricao.Text
TBLISTA!vlr_unit = VlrUnit
TBLISTA!qt = QtrTotal
TBLISTA!vlr_Total = vlrTotal
TBLISTA.Update
TBLISTA.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaPedidoWEB()
On Error GoTo tratar_erro

FunAbreBDSite

With frmVendas_Pedidos_WEB


txtPedidoWEB.Text = .Lista.SelectedItem
txtStatus = .Lista.SelectedItem.ListSubItems.Item(1).Text
txtCliente = .Lista.SelectedItem.ListSubItems.Item(2).Text
txtValorTotal = Format(.Lista.SelectedItem.ListSubItems.Item(3).Text, "###,##0.00")
txtData = .Lista.SelectedItem.ListSubItems.Item(4).Text
txtVendedor = .Lista.SelectedItem.ListSubItems.Item(5).Text
txtCondicoes = .Lista.SelectedItem.ListSubItems.Item(6).Text
txtContato = .Lista.SelectedItem.ListSubItems.Item(7).Text
txtObservacoes = .Lista.SelectedItem.ListSubItems.Item(8).Text

Documento = .Lista.SelectedItem.ListSubItems.Item(9).Text

If Len(Documento) = 14 Then
    txtCPF_CNPJ = Format(Documento, "@@.@@@.@@@/@@@@-@@")
    StrSql = "SELECT * FROM Vendas_Clientes Where CNPJ = '" & Documento & "'"
Else
    txtCPF_CNPJ = Format(Documento, "@@@.@@@.@@@-@@")
    StrSql = "SELECT * FROM Vendas_Clientes Where CPF = '" & Documento & "'"
End If


If ConexaoMySql.State = 1 Then
    Set TBLISTA = New ADODB.Recordset
    '=================================================================
    ' Buscar dados do cliente na WEB
    '=================================================================
    TBLISTA.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
        If TBLISTA.EOF = False Then
        txtuf.Text = TBLISTA!UF
        End If
        TBLISTA.Close
End If

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaListaProdutos()
On Error GoTo tratar_erro

'=================================================================
' Abrir BD na WEB
'=================================================================
FunAbreBDSite

vlrTotal = 0

ListaProdutos.ListItems.Clear

If ConexaoMySql.State = 1 Then

StrSql = "SELECT * FROM Vendas_Pedido_Lista Where ID_Pedido = '" & txtPedidoWEB.Text & "'"
'Debug.print StrSql

Set TBLISTA = New ADODB.Recordset
'=================================================================
' Buscar produtos do pedido na WEB
'=================================================================
TBLISTA.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
    
        With ListaProdutos.ListItems
            .Add , , TBLISTA!ID_lista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!vlr_unit), "", Format(TBLISTA!vlr_unit, "###,##0.00"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!qt), "", Format(TBLISTA!qt, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!vlr_Total), "", Format(TBLISTA!vlr_Total, "###,##0.00"))
            vlrTotal = vlrTotal + TBLISTA!vlr_Total
        End With
        TBLISTA.MoveNext
    Loop
End If

TBLISTA.Close
txtValorTotal = Format(vlrTotal, "###,##0.00")
ConexaoMySql.Execute ("Update Vendas_Pedidos set ValorTotal = " & Replace(vlrTotal, ",", ".") & " Where ID_Pedido = '" & txtPedidoWEB.Text & "'")
End If



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaProdutos_Click()
On Error GoTo tratar_erro

If ListaProdutos.ListItems.Count > 0 Then
    txtId.Text = ListaProdutos.SelectedItem
    txtCodigo.Text = ListaProdutos.SelectedItem.ListSubItems.Item(1).Text
    txtunidade.Text = ListaProdutos.SelectedItem.ListSubItems.Item(2).Text
    txtdescricao.Text = ListaProdutos.SelectedItem.ListSubItems.Item(3).Text
    txtQuantidade.Text = Format(ListaProdutos.SelectedItem.ListSubItems.Item(5).Text, "###,##0.00")
    txtvlrUnit.Text = Format(ListaProdutos.SelectedItem.ListSubItems.Item(4).Text, "###,##0.00")
    txtvlrTotal.Text = Format(ListaProdutos.SelectedItem.ListSubItems.Item(6).Text, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCalculaTotal()
On Error GoTo tratar_erro
Dim VlrUnit As Double
Dim quantidade As Double

If IsNumeric(txtvlrUnit) = True And IsNumeric(txtQuantidade) = True Then
    VlrUnit = txtvlrUnit.Text
    quantidade = txtQuantidade.Text
    
    txtvlrTotal.Text = Format(VlrUnit * quantidade, "###,##0.00")
    Else
    txtvlrTotal.Text = Format(0, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

ProcCalculaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtVlrunit_Change()
On Error GoTo tratar_erro

ProcCalculaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
