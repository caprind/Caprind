VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Orcamento_Conjunto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Vendas | Orçamento | Conjuntos (Materiais)"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
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
   ScaleHeight     =   7350
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   6945
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   873
      DibPicture      =   "frm_Orcamento_Conjunto.frx":0000
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
      Icon            =   "frm_Orcamento_Conjunto.frx":1CAD
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Valor total"
      Top             =   600
      Width           =   8115
      Begin VB.ComboBox cmbcodref 
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
         ItemData        =   "frm_Orcamento_Conjunto.frx":1FC7
         Left            =   210
         List            =   "frm_Orcamento_Conjunto.frx":1FC9
         Sorted          =   -1  'True
         TabIndex        =   29
         ToolTipText     =   "Código de referência."
         Top             =   2265
         Width           =   2340
      End
      Begin VB.TextBox cmbfamilia 
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   1590
         Width           =   6675
      End
      Begin VB.TextBox txtLargura2 
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
         Height          =   315
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Largura do item"
         ToolTipText     =   "Data e hora da validação."
         Top             =   2280
         Width           =   990
      End
      Begin VB.TextBox txtComprimento2 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4800
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Comprimento do item"
         ToolTipText     =   "Comprimento do item"
         Top             =   2280
         Width           =   1035
      End
      Begin VB.TextBox txtpeso 
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
         Left            =   2550
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Quilograma por unidade."
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtkgpc 
         Alignment       =   2  'Center
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
         Left            =   4035
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Peso por unidade."
         Top             =   2280
         Width           =   765
      End
      Begin VB.ComboBox cmbunkg 
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
         ItemData        =   "frm_Orcamento_Conjunto.frx":1FCB
         Left            =   3300
         List            =   "frm_Orcamento_Conjunto.frx":1FDB
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Unidade por kilograma."
         Top             =   2280
         Width           =   765
      End
      Begin VB.TextBox txtdimensao 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3300
         MaxLength       =   30
         TabIndex        =   14
         Text            =   "0,00000"
         ToolTipText     =   "Dimensão a ser utilizada no conjunto."
         Top             =   450
         Width           =   1125
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
         Height          =   315
         Left            =   5340
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   450
         Width           =   1515
      End
      Begin VB.TextBox txtvlrUnitario 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Valor unitario do item"
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox txtcodigoproduto 
         Alignment       =   2  'Center
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtunidade 
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
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   450
         Width           =   555
      End
      Begin VB.TextBox txtdescricao 
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
         Height          =   315
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1020
         Width           =   6690
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
         Left            =   4440
         TabIndex        =   3
         ToolTipText     =   "Quantidade"
         Top             =   450
         Width           =   885
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   31
         Top             =   240
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
         Index           =   4
         Left            =   1635
         TabIndex        =   32
         Top             =   240
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
         Left            =   3113
         TabIndex        =   33
         Top             =   810
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
         Index           =   0
         Left            =   2265
         TabIndex        =   34
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   344
         Caption         =   "Valor Unitário"
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
         NoHTMLCaption   =   "Valor Unitário"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   1
         Left            =   5775
         TabIndex        =   35
         Top             =   240
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   344
         Caption         =   "valor Total"
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
         NoHTMLCaption   =   "valor Total"
      End
      Begin DrawSuite2022.USButton btnNovo 
         Height          =   795
         Left            =   6990
         TabIndex        =   11
         Top             =   180
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1402
         DibPicture      =   "frm_Orcamento_Conjunto.frx":1FF3
         Caption         =   "Novo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Height          =   795
         Left            =   6990
         TabIndex        =   12
         Top             =   990
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1402
         DibPicture      =   "frm_Orcamento_Conjunto.frx":82D7
         Caption         =   "Gravar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Height          =   825
         Left            =   6990
         TabIndex        =   13
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1455
         DibPicture      =   "frm_Orcamento_Conjunto.frx":10CDC
         Caption         =   "Excluir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
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
         Left            =   570
         TabIndex        =   30
         Top             =   2070
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   3277
         TabIndex        =   28
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Largura"
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
         Left            =   6068
         TabIndex        =   27
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comprimento"
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
         Left            =   4830
         TabIndex        =   24
         Top             =   2070
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/un*"
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
         Left            =   2670
         TabIndex        =   22
         Top             =   2070
         Width           =   510
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   3270
         TabIndex        =   21
         Top             =   2100
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/pç"
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
         Left            =   4230
         TabIndex        =   20
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Un/Kg*"
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
         Left            =   3495
         TabIndex        =   19
         Top             =   2070
         Width           =   525
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensão"
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
         Left            =   3570
         TabIndex        =   15
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   4470
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3375
      Left            =   150
      TabIndex        =   10
      Top             =   3420
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "T"
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Codigo"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Un"
         Object.Width           =   705
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5847
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Quant"
         Object.Width           =   1146
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Valor"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frm_Orcamento_Conjunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ID_Conjunto As Integer

Private Sub btnExcluir_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
    If USMsgBox("Deseja realmente excluir o item " & Lista.SelectedItem.ListSubItems.Item(1).Text & "?", vbYesNo, "CAPRIND  v5.0") = vbYes Then
        Conexao.Execute ("Delete from Vendas_Orcamento_conjunto where id_conjunto = '" & Lista.SelectedItem & "'")
        USMsgBox "Item excluida com sucesso!", vbInformation, "CAPRIND v5.0"
        ProcCarregaLista_conjunto
    End If
End If

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnGravar_Click()
On Error GoTo tratar_erro

ProcgravarItem

Set TBFases = CreateObject("adodb.recordset")
StrSql = "Select SUM(vlrtotal) as Total from Vendas_Orcamento_Conjunto where ID_Orcamento = '" & frm_orcamento.txtId.Text & "'"
TBFases.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

    If TBFases.EOF = False Then
        frm_orcamento.txtv2.Text = Format(TBFases!Total, "###,##0.00")
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcgravarItem()
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
StrSql = "Select * from Vendas_Orcamento_Conjunto where ID_Orcamento = '" & frm_orcamento.txtId.Text & "' and ID_Conjunto = '" & ID_Conjunto & "'"
TBItem.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = True Then
TBItem.AddNew
End If
TBItem!ID_orcamento = frm_orcamento.txtId
TBItem!Codproduto = Cod_produto
TBItem!CODIGO = txtcodigoproduto.Text
TBItem!Unidade = txtunidade.Text
TBItem!Descricao = txtdescricao.Text
TBItem!quantidade = txtLote.Text
TBItem!Dimensao = txtdimensao.Text
TBItem!vlrUnitario = Format((txtvlrTotal.Text / txtLote), "###,##0.0000")
TBItem!vlrTotal = Format(txtvlrTotal.Text, "###,##0.0000")
TBItem.Update
TBItem.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND  V5.0"

StrSql = "update projproduto set comprimento = " & Replace(txtComprimento2.Text, ",", ".") & " WHERE codproduto = " & Cod_produto

Conexao.Execute (StrSql)

ProcCarregaLista_conjunto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_conjunto()
On Error GoTo tratar_erro

valor = 0
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")

StrSql = "Select * from Vendas_Orcamento_Conjunto where ID_Orcamento = '" & frm_orcamento.txtId.Text & "'"
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID_Conjunto
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!quantidade), "", TBLISTA!quantidade)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!vlrUnitario), "", "R$ " & Format(TBLISTA!vlrUnitario, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!vlrTotal), "", "R$ " & Format(TBLISTA!vlrTotal, "###,##0.00"))
            valor = valor + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop
    Lista.ListItems.Add , , 1 'TBLISTA!ID_Conjunto
    Lista.ListItems.Item(Contador + 1).SubItems(5) = "TOTAL :"
    Lista.ListItems.Item(Contador + 1).ListSubItems.Item(4).ForeColor = vbRed
    Lista.ListItems.Item(Contador + 1).SubItems(6) = "R$ " & Format(valor, "###,##0.00")
    Lista.ListItems.Item(Contador + 1).ListSubItems.Item(6).ForeColor = vbRed
    

frm_orcamento.txtv2.Text = Format(valor, "###,##0.00")

End If
TBLISTA.Close

        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnNovo_Click()
On Error GoTo tratar_erro

ProcLimparCampos
frm_Orcamento_conjunto_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcLimparCampos

If frm_orcamento.txtcodproduto <> "" Then
    ProcCarregaLista_conjunto
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaValor()
On Error GoTo tratar_erro


If txtvlrUnitario.Text <> "" And txtLote.Text <> "" And txtdimensao.Text <> "" Then
    Select Case txtunidade
        Case "KG": txtvlrTotal = Format(txtvlrUnitario * txtPesototal, "###,##0.00")
        Case "MM": txtvlrTotal = Format((txtvlrUnitario * txtdimensao) * txtLote, "###,##0.00")
        Case "MT": txtvlrTotal = Format(((txtvlrUnitario / 1000) * txtdimensao) * txtLote, "###,##0.00")
    End Select
End If

If IsNumeric(txtdimensao.Text) = True And IsNumeric(txtComprimento2.Text) = True And IsNumeric(txtLote.Text) = True And txtunidade <> "KG" And txtunidade <> "MM" And txtunidade <> "MT" Then
        txtvlrTotal = Format(((txtvlrUnitario / txtComprimento2) * txtdimensao) * txtLote, "###,##0.00")
End If

If txtunidade <> "KG" And txtunidade <> "MM" And txtunidade <> "MT" And IsNumeric(txtLote) = True Then
    If IsNumeric(txtdimensao.Text) = False Or txtdimensao.Text = "" Or txtdimensao.Text = "0,00" Then
        txtvlrTotal = Format(txtvlrUnitario * txtLote, "###,##0.00")
    End If
End If
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro
      
      ID_Conjunto = 0
      Cod_produto = 0
      txtcodigoproduto = ""
      txtdescricao.Text = ""
      txtLote = ""
      txtunidade = ""
      txtvlrUnitario.Text = ""
      txtvlrTotal.Text = ""
      
      txtdimensao.Text = Format(0, "###,##0.00")
      cmbfamilia.Text = ""
      txtvlrTotal = Format(0, "###,##0.00")
      txtComprimento2.Text = Format(0, "###,##0.00")
      txtLargura2.Text = Format(0, "###,##0.00")
      cmbunkg.Text = "N/a"
      txtpeso.Text = Format(0, "###,##0.00")
      cmbfamilia = ""


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
  If Lista.SelectedItem <> "" Then
  
  Set TBLISTA = CreateObject("adodb.recordset")
  
  StrSql = "select VOC.*, PP.Classe,PP.codproduto,PP.Un_Kg,PP.Comprimento,PP.Largura,PP.Peso_metro from Vendas_Orcamento_Conjunto VOC inner join projproduto PP on PP.codproduto = VOC.codProduto where VOC.ID_conjunto = '" & Lista.SelectedItem & "'"
  'Debug.print StrSql
  
  TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBLISTA.EOF = False Then
    ID_Conjunto = Lista.SelectedItem
      Cod_produto = TBLISTA!Codproduto
      txtcodigoproduto.Text = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
      txtdescricao.Text = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
      txtvlrUnitario = IIf(IsNull(TBLISTA!vlrUnitario), "", Format(TBLISTA!vlrUnitario, "###,##0.00"))
      txtLote.Text = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.00"))
      txtunidade.Text = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
      txtdimensao.Text = IIf(IsNull(TBLISTA!Dimensao), "", Format(TBLISTA!Dimensao, "###,##0.00"))
      cmbfamilia.Text = IIf(IsNull(TBLISTA!Classe), "", TBLISTA!Classe)
      txtvlrTotal = IIf(IsNull(TBLISTA!vlrTotal), "", Format(TBLISTA!vlrTotal, "###,##0.00"))
      txtComprimento2.Text = IIf(IsNull(TBLISTA!Comprimento), "0,0000", Format(TBLISTA!Comprimento, "###,##0.00"))
      txtLargura2.Text = IIf(IsNull(TBLISTA!Largura), "0,0000", Format(TBLISTA!Largura, "###,##0.00"))
      cmbunkg.Text = IIf(IsNull(TBLISTA!Un_Kg), "", TBLISTA!Un_Kg)
      txtpeso.Text = IIf(IsNull(TBLISTA!peso_metro), "0,0000", Format(TBLISTA!peso_metro, "###,##0.00"))
  End If
  TBLISTA.Close
  
  End If
End If

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from item_aplicacoes where codproduto = " & Cod_produto & "", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
 cmbcodref.Text = TBItem!N_referencia
End If
TBItem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_materiais()
On Error GoTo tratar_erro

valor = 0
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")

StrSql = "Select * from Vendas_Orcamento_Conjunto where ID_Orcamento = '" & frm_orcamento.txtId.Text & "'"
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID_Conjunto
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.00"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!vlrUnitario), "", "R$ " & Format(TBLISTA!vlrUnitario, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!vlrTotal), "", "R$ " & Format(TBLISTA!vlrTotal, "###,##0.00"))
            valor = valor + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop
    Lista.ListItems.Add , , 1 'TBLISTA!ID_fases
    Lista.ListItems.Item(Contador + 1).SubItems(4) = "TOTAL :"
    Lista.ListItems.Item(Contador + 1).ListSubItems.Item(4).ForeColor = vbRed
    Lista.ListItems.Item(Contador + 1).SubItems(7) = "R$ " & Format(valor, "###,##0.00")
    Lista.ListItems.Item(Contador + 1).ListSubItems.Item(7).ForeColor = vbRed
    

frm_orcamento.txtv3.Text = Format(valor, "###,##0.00")

End If
TBLISTA.Close

        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdimensao_Change()
On Error GoTo tratar_erro

If txtdimensao.Text <> "" Then
    VerifNumero = txtdimensao.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtdimensao.Text = ""
        txtdimensao.SetFocus
        Exit Sub
    End If
End If

ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdimensao_LostFocus()
On Error GoTo tratar_erro

If txtdimensao = "" Then txtdimensao = 0

If IsNumeric(txtdimensao) Then
txtdimensao = Format(txtdimensao, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLote_Change()
On Error GoTo tratar_erro

If txtLote.Text <> "" Then
    VerifNumero = txtLote.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLote.Text = ""
        txtLote.SetFocus
        Exit Sub
    End If
End If

ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
