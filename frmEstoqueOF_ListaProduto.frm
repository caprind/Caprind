VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoque_Ordem_Faturamento_ListaProduto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque  | Ordem de faturamento - Localizar produtos"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12030
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   66
      Top             =   9345
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   873
      DibPicture      =   "frmEstoqueOF_ListaProduto.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmEstoqueOF_ListaProduto.frx":3650
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8790
      Left            =   30
      TabIndex        =   25
      Top             =   510
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   15505
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Localizar produtos vendidos"
      TabPicture(0)   =   "frmEstoqueOF_ListaProduto.frx":396A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PBLista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1(25)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lista_carteira"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1(20)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1(23)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Localizar produtos cadastrados"
      TabPicture(1)   =   "frmEstoqueOF_ListaProduto.frx":3986
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(1)=   "PBLista1"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(3)=   "Frame9"
      Tab(1).Control(4)=   "Frame1(0)"
      Tab(1).ControlCount=   5
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
         Height          =   855
         Index           =   23
         Left            =   3600
         TabIndex        =   37
         Top             =   2070
         Width           =   8295
         Begin VB.ComboBox cmbfiltrarpor 
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
            Height          =   330
            ItemData        =   "frmEstoqueOF_ListaProduto.frx":39A2
            Left            =   300
            List            =   "frmEstoqueOF_ListaProduto.frx":39BB
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   3195
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
            Height          =   315
            Left            =   3480
            TabIndex        =   2
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   4545
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
            ItemData        =   "frmEstoqueOF_ListaProduto.frx":3A26
            Left            =   3480
            List            =   "frmEstoqueOF_ListaProduto.frx":3A28
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Visible         =   0   'False
            Width           =   4545
         End
         Begin VB.Label Label1 
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
            Index           =   24
            Left            =   1500
            TabIndex        =   39
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
            Index           =   2
            Left            =   4890
            TabIndex        =   38
            Top             =   210
            Width           =   1875
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar por :                                                     "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   20
         Left            =   65
         TabIndex        =   45
         Top             =   2070
         Width           =   3525
         Begin DrawSuite2022.USOptionButton chkData 
            Height          =   225
            Index           =   0
            Left            =   1110
            TabIndex        =   62
            Top             =   30
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   397
            Caption         =   "Data venda"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
         End
         Begin MSComCtl2.DTPicker msk_data 
            Height          =   315
            Index           =   1
            Left            =   2070
            TabIndex        =   59
            ToolTipText     =   "Data final para pesquisa."
            Top             =   420
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
            Format          =   181075969
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker msk_data 
            Height          =   315
            Index           =   0
            Left            =   510
            TabIndex        =   60
            ToolTipText     =   "Data início para pesquisa."
            Top             =   420
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
            Format          =   181075971
            CurrentDate     =   39057
         End
         Begin DrawSuite2022.USOptionButton chkData 
            Height          =   225
            Index           =   1
            Left            =   2220
            TabIndex        =   63
            Top             =   30
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            Caption         =   "Prazo entrega"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "De                                à"
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
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   61
            Top             =   510
            Width           =   1845
         End
      End
      Begin MSComctlLib.ListView Lista_carteira 
         Height          =   4785
         Left            =   60
         TabIndex        =   4
         Top             =   2940
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   8440
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   24
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Cotacao"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cód. ref."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Ped. cliente"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "N. item"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "D"
            Text            =   "Pr. final"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Ped. int."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   926
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Tag             =   "T"
            Text            =   "Programa"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   935
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Object.Tag             =   "T"
            Text            =   "Antec."
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   15
            Object.Tag             =   "T"
            Text            =   "Parc."
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Tag             =   "T"
            Text            =   "Moeda"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Object.Tag             =   "N"
            Text            =   "Qtde. vend."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Object.Tag             =   "N"
            Text            =   "Qtde. faturar"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Object.Tag             =   "N"
            Text            =   "Qtde. faturada"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Object.Tag             =   "N"
            Text            =   "Saldo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Object.Tag             =   "N"
            Text            =   "Emp. est."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Object.Tag             =   "N"
            Text            =   "Emp. prod."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Object.Tag             =   "T"
            Text            =   "Observações"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Carregar a lista com o campo descrição"
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
         Height          =   750
         Left            =   8070
         TabIndex        =   52
         Top             =   1320
         Width           =   3825
         Begin VB.OptionButton Optdescricao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   54
            Top             =   390
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optespecificacao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição comercial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1620
            TabIndex        =   53
            Top             =   390
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opções para filtrar produtos"
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
         Index           =   25
         Left            =   65
         TabIndex        =   46
         Top             =   1320
         Width           =   7995
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7260
            TabIndex        =   58
            Top             =   390
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   57
            Top             =   390
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3510
            TabIndex        =   56
            Top             =   390
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6090
            TabIndex        =   55
            Top             =   390
            Width           =   1155
         End
         Begin VB.CheckBox Chk_tem_estoque 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vendidos com saldo em estoque"
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
            Height          =   195
            Left            =   150
            TabIndex        =   0
            Top             =   390
            Value           =   1  'Checked
            Width           =   3165
         End
      End
      Begin VB.Frame Frame3 
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
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   65
         TabIndex        =   40
         Top             =   7710
         Width           =   11835
         Begin VB.TextBox txtNreg 
            Alignment       =   2  'Center
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
            Left            =   2970
            TabIndex        =   5
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr 
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
            Left            =   5790
            TabIndex        =   6
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   8010
            TabIndex        =   10
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":3A2A
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagAnt 
            Height          =   315
            Left            =   7470
            TabIndex        =   9
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":71CE
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagIr 
            Height          =   315
            Left            =   6360
            TabIndex        =   7
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            Caption         =   "Ir"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagPrim 
            Height          =   315
            Left            =   6930
            TabIndex        =   8
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":ACD7
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagUlt 
            Height          =   315
            Left            =   8550
            TabIndex        =   11
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":EDC6
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               registros por página"
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
            Index           =   22
            Left            =   2280
            TabIndex        =   43
            Top             =   240
            Width           =   2760
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
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
            Left            =   180
            TabIndex        =   42
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
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
            Left            =   10155
            TabIndex        =   41
            Top             =   240
            Width           =   1095
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
         Height          =   1515
         Index           =   0
         Left            =   -74935
         TabIndex        =   30
         Top             =   1320
         Width           =   11805
         Begin VB.Frame Frame4 
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
            Height          =   510
            Left            =   3660
            TabIndex        =   47
            Top             =   210
            Width           =   7965
            Begin VB.CheckBox Check1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Filtrar vendidos em estoque"
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
               Height          =   195
               Left            =   5040
               TabIndex        =   64
               Top             =   210
               Value           =   1  'Checked
               Width           =   2775
            End
            Begin VB.OptionButton optIgual1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3930
               TabIndex        =   51
               Top             =   180
               Width           =   705
            End
            Begin VB.OptionButton optMeio1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1470
               TabIndex        =   50
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optInicio1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início frase"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   49
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton optFim1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2760
               TabIndex        =   48
               Top             =   180
               Width           =   1155
            End
         End
         Begin VB.TextBox txtTexto1 
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
            TabIndex        =   13
            ToolTipText     =   "Texto para pesquisa."
            Top             =   1050
            Width           =   7305
         End
         Begin VB.ComboBox cmbfiltrarpor1 
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
            Height          =   330
            ItemData        =   "frmEstoqueOF_ListaProduto.frx":12652
            Left            =   180
            List            =   "frmEstoqueOF_ListaProduto.frx":1267A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   3375
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Carregar no campo descrição"
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
            Height          =   615
            Left            =   7650
            TabIndex        =   31
            Top             =   780
            Width           =   3975
            Begin VB.OptionButton optespecificacao1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Descrição comercial"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1740
               TabIndex        =   24
               Top             =   300
               Width           =   2055
            End
            Begin VB.OptionButton Optdescricao1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Descrição"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   390
               TabIndex        =   23
               Top             =   300
               Value           =   -1  'True
               Width           =   1125
            End
         End
         Begin VB.ComboBox cmbfamilia1 
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
            Left            =   180
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Texto para pesquisa."
            Top             =   1050
            Visible         =   0   'False
            Width           =   7305
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filtrar por"
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
            Left            =   1440
            TabIndex        =   33
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   3097
            TabIndex        =   32
            Top             =   840
            Width           =   1470
         End
      End
      Begin VB.Frame Frame9 
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
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74935
         TabIndex        =   26
         Top             =   7710
         Width           =   11805
         Begin VB.TextBox txtNreg1 
            Alignment       =   2  'Center
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
            Left            =   2970
            TabIndex        =   16
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr1 
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
            Left            =   5790
            TabIndex        =   17
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx1 
            Height          =   315
            Left            =   8010
            TabIndex        =   21
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":12718
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagAnt1 
            Height          =   315
            Left            =   7470
            TabIndex        =   20
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":15EBF
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagIr1 
            Height          =   315
            Left            =   6360
            TabIndex        =   18
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            Caption         =   "Ir"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagPrim1 
            Height          =   315
            Left            =   6930
            TabIndex        =   19
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":199CD
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagUlt1 
            Height          =   315
            Left            =   8550
            TabIndex        =   22
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoqueOF_ListaProduto.frx":1DABF
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar               registros por página"
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
            Left            =   2280
            TabIndex        =   29
            Top             =   240
            Width           =   2760
         End
         Begin VB.Label lblRegistros1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
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
            Left            =   180
            TabIndex        =   28
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
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
            Left            =   10155
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74935
         TabIndex        =   34
         Top             =   330
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   1720
         ButtonCount     =   5
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Filtrar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Filtrar (F2)"
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
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonAlignment2=   2
         ButtonType2     =   1
         ButtonStyle2    =   -1
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState2    =   -1
         ButtonLeft2     =   40
         ButtonTop2      =   4
         ButtonWidth2    =   2
         ButtonHeight2   =   54
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Ajuda"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Ajuda (F1)"
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
         ButtonLeft3     =   44
         ButtonTop3      =   2
         ButtonWidth3    =   36
         ButtonHeight3   =   21
         ButtonCaption4  =   "Sair"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Sair (Esc)"
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
         ButtonLeft4     =   82
         ButtonTop4      =   2
         ButtonWidth4    =   26
         ButtonHeight4   =   21
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   5
         ButtonLeft5     =   110
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   5970
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmEstoqueOF_ListaProduto.frx":2134C
            Count           =   1
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista1 
         Height          =   255
         Left            =   -74940
         TabIndex        =   35
         Top             =   8340
         Width           =   11805
         _ExtentX        =   20823
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
         SearchText      =   ""
         Value           =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4845
         Left            =   -74940
         TabIndex        =   15
         Top             =   2850
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   8546
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Cód."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   9375
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   5644
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   36
         Top             =   330
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   1720
         ButtonCount     =   6
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Filtrar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Filtrar (F2)"
         ButtonKey1      =   "1"
         ButtonAlignment1=   2
         BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   42
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Adicionar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Adicionar selecionados (F3)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   46
         ButtonTop2      =   2
         ButtonWidth2    =   61
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonAlignment3=   2
         ButtonType3     =   1
         ButtonStyle3    =   -1
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   -1
         ButtonLeft3     =   109
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
         ButtonCaption4  =   "Ajuda"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Ajuda (F1)"
         ButtonKey4      =   "4"
         ButtonAlignment4=   2
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   113
         ButtonTop4      =   2
         ButtonWidth4    =   41
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Sair"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Sair (Esc)"
         ButtonKey5      =   "5"
         ButtonAlignment5=   2
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   188
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   5970
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmEstoqueOF_ListaProduto.frx":23534
            Count           =   1
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   60
         TabIndex        =   44
         Top             =   8340
         Width           =   11805
         _ExtentX        =   20823
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
   End
End
Attribute VB_Name = "frmEstoque_Ordem_Faturamento_ListaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_data_Click(index As Integer)
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear
If chkData(0).Value = 1 Then
    chkData(1).Value = 0
    Frame1(20).Enabled = True
    msk_data(0).SetFocus
ElseIf chkData(1).Value = 1 Then
        chkData(0).Value = 0
        Frame1(20).Enabled = True
        msk_data(0).SetFocus
    Else
        Frame1(20).Enabled = False
        msk_data(1).Value = Date
        msk_data(0).Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_tem_estoque_Click()
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia1_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia1 <> "" Then txtTexto1 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear
If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    With cmbTexto
        .Clear
        Set TBAbrir = CreateObject("adodb.recordset")
        If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 1 Then NomeViewFiltro = "Carteira_ordem_fat_PC" Else NomeViewFiltro = "Carteira_ordem_fat"
        TBAbrir.Open "Select Familia from " & NomeViewFiltro & " where Familia is not null group by Familia", Conexao, adOpenKeyset, adLockReadOnly
        If TBAbrir.EOF = False Then
            .AddItem ""
            Do While TBAbrir.EOF = False
                .AddItem TBAbrir!Familia
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    End With
Else
    txtTexto.Visible = True
    cmbTexto.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor1_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor1 = "Família" Or cmbfiltrarpor1 = "Cliente" Or cmbfiltrarpor1 = "Fornecedor" Then
    txtTexto1.Visible = False
    With cmbfamilia1
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor1 = "Família" Then
            ProcCarregaComboFamilia cmbfamilia1, "familia <> 'Null'", True
        ElseIf cmbfiltrarpor1 = "Cliente" Then
                Permitido = False
                Set TBClientes = CreateObject("adodb.recordset")
                TBClientes.Open "Select IDCliente, NomeRazao from Clientes where NomeRazao <> 'Null' order by NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then
                    Do While TBClientes.EOF = False
                        .AddItem TBClientes!NomeRazao
                        .ItemData(.NewIndex) = TBClientes!IDCliente
                        
                        If TBClientes!NomeRazao = frmEstoque_Ordem_Faturamento.txt_Razao Then Permitido = True
                        TBClientes.MoveNext
                    Loop
                End If
                TBClientes.Close
                If Permitido = True Then .Text = frmEstoque_Ordem_Faturamento.txt_Razao
            Else
                Permitido = False
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select IDCliente, Nome_Razao from Compras_fornecedores where Nome_Razao <> 'Null' order by Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    Do While TBFornecedor.EOF = False
                        .AddItem TBFornecedor!Nome_Razao
                        .ItemData(.NewIndex) = TBFornecedor!IDCliente
                        
                        If TBFornecedor!Nome_Razao = frmEstoque_Ordem_Faturamento.txt_Razao Then Permitido = True
                        TBFornecedor.MoveNext
                    Loop
                End If
                TBFornecedor.Close
                If Permitido = True Then .Text = frmEstoque_Ordem_Faturamento.txt_Razao
        End If
    End With
Else
    txtTexto1.Visible = True
    cmbfamilia1.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarCarteira()
On Error GoTo tratar_erro

With msk_data(1)
    If FunVerificaDataFinal(msk_data(0).Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

If Formulario <> "Estoque/Ordem de faturamento" Then

With frmEstoque_Ordem_Faturamento
    If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
        If Faturamento_ListaProdudos = True Then TipoFiltro = "P" Else TipoFiltro = "S"
        FiltroPadrao = "ID_empresa = " & IDempresa & " and Tipo = '" & TipoFiltro & "' and IDcliente = " & .txtIDcliente
        If Chk_tem_estoque.Value = 1 Then EstoqueFiltro = "and Qtde_empenhada_est > 0" Else EstoqueFiltro = ""
        If chkData(0).Value = 1 Then DataTexto = "Datavendas" Else DataTexto = "prazofinal"
        If chkData(0).Value = 1 Or chkData(1).Value = 1 Then DataFiltro = "and " & DataTexto & " Between '" & Format(msk_data(0).Value, "Short Date") & "' And '" & Format(msk_data(1).Value, "Short Date") & "'" Else DataFiltro = ""
        
        If txtTexto <> "" Or cmbTexto <> "" Then
            If cmbfiltrarpor = "Família" Then
                StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and Familia = '" & cmbTexto & "' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência": TextoFiltro = "n_referencia"
                    Case "Código interno": TextoFiltro = "Desenho"
                    Case "Descrição": TextoFiltro = "Descricao_tecnica"
                    Case "Pedido do cliente": TextoFiltro = "PCcliente"
                    Case "Pedido interno": TextoFiltro = "Ncotacao"
                    Case "Programa": TextoFiltro = "Programatexto"
                End Select
                If Optinicio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " like '" & txtTexto & "%' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
                If Optmeio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "%' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
                If Optfim.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
                If optIgual.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " = '" & txtTexto & "' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
            End If
        Else
            StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
        End If
    Else
        FiltroPadrao = "ID_empresa = " & IDempresa & " and IDfornecedor = " & .txtIDcliente
        If chkData(0).Value = 1 Then DataTexto = "Data" Else DataTexto = "Prazo"
        If chkData(0).Value = 1 Or chkData(1).Value = 1 Then DataFiltro = "and " & DataTexto & " Between '" & Format(msk_data(0).Value, "Short Date") & "' And '" & Format(msk_data(1).Value, "Short Date") & "'" Else DataFiltro = ""
        
        If txtTexto <> "" Or cmbTexto <> "" Then
            If cmbfiltrarpor = "Família" Then
                StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and Familia = '" & cmbTexto & "' " & DataFiltro & " order by " & DataTexto & ", Desenho"
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência": TextoFiltro = "n_referencia"
                    Case "Código interno": TextoFiltro = "Desenho"
                    Case "Descrição": TextoFiltro = "Descricao"
                    Case "Pedido de compra": TextoFiltro = "Pedido"
                End Select
                If Optinicio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " like '" & txtTexto & "%' " & DataFiltro & " order by " & DataTexto & ", Desenho"
                If Optmeio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "%' " & DataFiltro & " order by " & DataTexto & ", Desenho"
                If Optfim.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "' " & DataFiltro & " order by " & DataTexto & ", Desenho"
                If optIgual.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " = '" & txtTexto & "' " & DataFiltro & " order by " & DataTexto & ", Desenho"
            End If
        Else
            StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " " & DataFiltro & " order by " & DataTexto & ", Desenho"
        End If
    End If
End With
ProcCarregaListaCarteira (1)
Else
With frmEstoque_Ordem_Faturamento
    If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
        If Faturamento_ListaProdudos = True Then TipoFiltro = "P" Else TipoFiltro = "S"
        FiltroPadrao = "ID_empresa = " & IDempresa & " and Tipo = '" & TipoFiltro & "' and IDcliente = " & .txtIDcliente
        If Chk_tem_estoque.Value = 1 Then EstoqueFiltro = "and Qtde_empenhada_est > 0" Else EstoqueFiltro = ""
        If chkData(0).Value = 1 Then DataTexto = "Datavendas" Else DataTexto = "prazofinal"
        If chkData(0).Value = 1 Or chkData(1).Value = 1 Then DataFiltro = "and " & DataTexto & " Between '" & Format(msk_data(0).Value, "Short Date") & "' And '" & Format(msk_data(1).Value, "Short Date") & "'" Else DataFiltro = ""
        
        If txtTexto <> "" Or cmbTexto <> "" Then
            If cmbfiltrarpor = "Família" Then
                StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and Familia = '" & cmbTexto & "' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência": TextoFiltro = "n_referencia"
                    Case "Código interno": TextoFiltro = "Desenho"
                    Case "Descrição": TextoFiltro = "Descricao_tecnica"
                    Case "Pedido do cliente": TextoFiltro = "PCcliente"
                    Case "Pedido interno": TextoFiltro = "Ncotacao"
                    Case "Programa": TextoFiltro = "Programatexto"
                End Select
                If Optinicio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " like '" & txtTexto & "%' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
                If Optmeio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "%' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
                If Optfim.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
                If optIgual.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " and " & TextoFiltro & " = '" & txtTexto & "' " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
            End If
        Else
            StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat where " & FiltroPadrao & " " & DataFiltro & " " & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
        End If
    Else
        FiltroPadrao = "ID_empresa = " & IDempresa & " and IDfornecedor = " & .txtIDcliente
        If chkData(0).Value = 1 Then DataTexto = "Data" Else DataTexto = "Prazo"
        If chkData(0).Value = 1 Or chkData(1).Value = 1 Then DataFiltro = "and " & DataTexto & " Between '" & Format(msk_data(0).Value, "Short Date") & "' And '" & Format(msk_data(1).Value, "Short Date") & "'" Else DataFiltro = ""
        
        If txtTexto <> "" Or cmbTexto <> "" Then
            If cmbfiltrarpor = "Família" Then
                StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and Familia = '" & cmbTexto & "' " & DataFiltro & " order by " & DataTexto & ", Desenho"
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência": TextoFiltro = "n_referencia"
                    Case "Código interno": TextoFiltro = "Desenho"
                    Case "Descrição": TextoFiltro = "Descricao"
                    Case "Pedido de compra": TextoFiltro = "Pedido"
                End Select
                If Optinicio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " like '" & txtTexto & "%' " & DataFiltro & " order by " & DataTexto & ", Desenho"
                If Optmeio.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "%' " & DataFiltro & " order by " & DataTexto & ", Desenho"
                If Optfim.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " like '%" & txtTexto & "' " & DataFiltro & " order by " & DataTexto & ", Desenho"
                If optIgual.Value = True Then StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " and " & TextoFiltro & " = '" & txtTexto & "' " & DataFiltro & " order by " & DataTexto & ", Desenho"
            End If
        Else
            StrSqlLocProdPadrao = "Select * from Carteira_ordem_fat_PC where " & FiltroPadrao & " " & DataFiltro & " order by " & DataTexto & ", Desenho"
        End If
    End If
End With
ProcCarregaListaCarteira (1)
End If



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaCarteira(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista_carteira.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador = 0


Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)

If Formulario <> "Estoque/Ordem de faturamento" Then
    With Lista_carteira.ListItems
        If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
            .Add , , TBLocalizar_produto_padrao!CODIGO
            .Item(.Count).SubItems(1) = TBLocalizar_produto_padrao!Cotacao
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!Rev_codinterno), "", TBLocalizar_produto_padrao!Rev_codinterno)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!N_referencia), "", TBLocalizar_produto_padrao!N_referencia)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!descricao_tecnica), "", TBLocalizar_produto_padrao!descricao_tecnica)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLocalizar_produto_padrao!PCCliente), "", TBLocalizar_produto_padrao!PCCliente)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_produto_padrao!N_item), "", TBLocalizar_produto_padrao!N_item)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao!PrazoFinal), "", (Format(TBLocalizar_produto_padrao!PrazoFinal, "dd/mm/yy")))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLocalizar_produto_padrao!Ncotacao), "", TBLocalizar_produto_padrao!Ncotacao)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLocalizar_produto_padrao!Revisao), "", TBLocalizar_produto_padrao!Revisao)
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLocalizar_produto_padrao!programatexto), "", TBLocalizar_produto_padrao!programatexto)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLocalizar_produto_padrao!Rev), "", TBLocalizar_produto_padrao!Rev)
            .Item(.Count).SubItems(14) = IIf(TBLocalizar_produto_padrao!Antecipacao_fat = True, "SIM", "NÃO")
            .Item(.Count).SubItems(15) = IIf(TBLocalizar_produto_padrao!Faturamento_parcial = True, "SIM", "NÃO")
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLocalizar_produto_padrao!Moeda), "", TBLocalizar_produto_padrao!Moeda)
            .Item(.Count).SubItems(17) = IIf(IsNull(TBLocalizar_produto_padrao!quantidade), "", Format(TBLocalizar_produto_padrao!quantidade, "###,##0.0000"))
            .Item(.Count).SubItems(18) = IIf(IsNull(TBLocalizar_produto_padrao!qtdeliberada), "", Format(TBLocalizar_produto_padrao!qtdeliberada, "###,##0.0000"))
            .Item(.Count).SubItems(19) = IIf(IsNull(TBLocalizar_produto_padrao!QtdeFaturada), "", Format(TBLocalizar_produto_padrao!QtdeFaturada, "###,##0.0000"))
            .Item(.Count).SubItems(20) = IIf(IsNull(TBLocalizar_produto_padrao!Saldo), "", Format(TBLocalizar_produto_padrao!Saldo, "###,##0.0000"))
            .Item(.Count).SubItems(21) = IIf(IsNull(TBLocalizar_produto_padrao!Qtde_empenhada_est), "", Format(TBLocalizar_produto_padrao!Qtde_empenhada_est, "###,##0.0000"))
            .Item(.Count).SubItems(22) = IIf(IsNull(TBLocalizar_produto_padrao!Qtde_empenhada_prod), "", Format(TBLocalizar_produto_padrao!Qtde_empenhada_prod, "###,##0.0000"))
            .Item(.Count).SubItems(23) = IIf(IsNull(TBLocalizar_produto_padrao!Obs_faturamento), "", TBLocalizar_produto_padrao!Obs_faturamento)
        Else
            .Add , , TBLocalizar_produto_padrao!IDlista
            .Item(.Count).SubItems(1) = TBLocalizar_produto_padrao!IDpedido
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!N_referencia), "", TBLocalizar_produto_padrao!N_referencia)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao!Prazo), "", (Format(TBLocalizar_produto_padrao!Prazo, "dd/mm/yy")))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLocalizar_produto_padrao!Pedido), "", TBLocalizar_produto_padrao!Pedido)
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLocalizar_produto_padrao!Moeda), "", TBLocalizar_produto_padrao!Moeda)
            .Item(.Count).SubItems(17) = IIf(IsNull(TBLocalizar_produto_padrao!Quant_Comp), "", Format(TBLocalizar_produto_padrao!Quant_Comp, "###,##0.0000"))
            .Item(.Count).SubItems(19) = IIf(IsNull(TBLocalizar_produto_padrao!QtdeFaturada), "", Format(TBLocalizar_produto_padrao!QtdeFaturada, "###,##0.0000"))
            .Item(.Count).SubItems(20) = IIf(IsNull(TBLocalizar_produto_padrao!Saldo), "", Format(TBLocalizar_produto_padrao!Saldo, "###,##0.0000"))
            .Item(.Count).SubItems(23) = IIf(IsNull(TBLocalizar_produto_padrao!Obs_pedido), "", TBLocalizar_produto_padrao!Obs_pedido)
        End If
    End With
    Else
    With Lista_carteira.ListItems
        If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
            .Add , , TBLocalizar_produto_padrao!CODIGO
            .Item(.Count).SubItems(1) = TBLocalizar_produto_padrao!Cotacao
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!Rev_codinterno), "", TBLocalizar_produto_padrao!Rev_codinterno)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!N_referencia), "", TBLocalizar_produto_padrao!N_referencia)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!descricao_tecnica), "", TBLocalizar_produto_padrao!descricao_tecnica)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLocalizar_produto_padrao!PCCliente), "", TBLocalizar_produto_padrao!PCCliente)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_produto_padrao!N_item), "", TBLocalizar_produto_padrao!N_item)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao!PrazoFinal), "", (Format(TBLocalizar_produto_padrao!PrazoFinal, "dd/mm/yy")))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLocalizar_produto_padrao!Ncotacao), "", TBLocalizar_produto_padrao!Ncotacao)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLocalizar_produto_padrao!Revisao), "", TBLocalizar_produto_padrao!Revisao)
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLocalizar_produto_padrao!programatexto), "", TBLocalizar_produto_padrao!programatexto)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLocalizar_produto_padrao!Rev), "", TBLocalizar_produto_padrao!Rev)
            .Item(.Count).SubItems(14) = IIf(TBLocalizar_produto_padrao!Antecipacao_fat = True, "SIM", "NÃO")
            .Item(.Count).SubItems(15) = IIf(TBLocalizar_produto_padrao!Faturamento_parcial = True, "SIM", "NÃO")
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLocalizar_produto_padrao!Moeda), "", TBLocalizar_produto_padrao!Moeda)
            .Item(.Count).SubItems(17) = IIf(IsNull(TBLocalizar_produto_padrao!quantidade), "", Format(TBLocalizar_produto_padrao!quantidade, "###,##0.0000"))
            .Item(.Count).SubItems(18) = IIf(IsNull(TBLocalizar_produto_padrao!qtdeliberada), "", Format(TBLocalizar_produto_padrao!qtdeliberada, "###,##0.0000"))
            .Item(.Count).SubItems(19) = IIf(IsNull(TBLocalizar_produto_padrao!QtdeFaturada), "", Format(TBLocalizar_produto_padrao!QtdeFaturada, "###,##0.0000"))
            .Item(.Count).SubItems(20) = IIf(IsNull(TBLocalizar_produto_padrao!Saldo), "", Format(TBLocalizar_produto_padrao!Saldo, "###,##0.0000"))
            .Item(.Count).SubItems(21) = IIf(IsNull(TBLocalizar_produto_padrao!Qtde_empenhada_est), "", Format(TBLocalizar_produto_padrao!Qtde_empenhada_est, "###,##0.0000"))
            .Item(.Count).SubItems(22) = IIf(IsNull(TBLocalizar_produto_padrao!Qtde_empenhada_prod), "", Format(TBLocalizar_produto_padrao!Qtde_empenhada_prod, "###,##0.0000"))
            .Item(.Count).SubItems(23) = IIf(IsNull(TBLocalizar_produto_padrao!Obs_faturamento), "", TBLocalizar_produto_padrao!Obs_faturamento)
        Else
            .Add , , TBLocalizar_produto_padrao!IDlista
            .Item(.Count).SubItems(1) = TBLocalizar_produto_padrao!IDpedido
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!N_referencia), "", TBLocalizar_produto_padrao!N_referencia)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLocalizar_produto_padrao!Prazo), "", (Format(TBLocalizar_produto_padrao!Prazo, "dd/mm/yy")))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLocalizar_produto_padrao!Pedido), "", TBLocalizar_produto_padrao!Pedido)
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLocalizar_produto_padrao!Moeda), "", TBLocalizar_produto_padrao!Moeda)
            .Item(.Count).SubItems(17) = IIf(IsNull(TBLocalizar_produto_padrao!Quant_Comp), "", Format(TBLocalizar_produto_padrao!Quant_Comp, "###,##0.0000"))
            .Item(.Count).SubItems(19) = IIf(IsNull(TBLocalizar_produto_padrao!QtdeFaturada), "", Format(TBLocalizar_produto_padrao!QtdeFaturada, "###,##0.0000"))
            .Item(.Count).SubItems(20) = IIf(IsNull(TBLocalizar_produto_padrao!Saldo), "", Format(TBLocalizar_produto_padrao!Saldo, "###,##0.0000"))
            .Item(.Count).SubItems(23) = IIf(IsNull(TBLocalizar_produto_padrao!Obs_pedido), "", TBLocalizar_produto_padrao!Obs_pedido)
        End If
    End With
    
    End If
    
    
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

If Faturamento_ListaProdudos = True Then TipoProduto = "P.Tipo <> 'S'" Else TipoProduto = "P.Tipo = 'S'"
CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.Unidade_com, P.classe"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto) LEFT JOIN Projproduto_clientes PC ON PC.codproduto = P.Codproduto) LEFT JOIN Projproduto_fornecedor PF ON PF.Codproduto = P.Codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
TextoFiltroPadrao = TipoProduto & " and P.DtValidacao IS NOT NULL and P.Bloqueado = 'False' group by " & CamposFiltro & " order by P.desenho"

If txtTexto1.Visible = True And txtTexto1 <> "" Or cmbfamilia1.Visible = True And cmbfamilia1 <> "" Then
    If cmbfiltrarpor1 = "Cliente" Or cmbfiltrarpor1 = "Fornecedor" Then
        If cmbfiltrarpor1 = "Cliente" Then TextoFiltro = "PC.IDCliente" Else TextoFiltro = "PF.IDfornecedor"
        StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = " & cmbfamilia1.ItemData(cmbfamilia1.ListIndex) & " and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor1 = "Família" Then
            StrSqlLocProdPadrao = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia1 & "' and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor1 = "Comprimento" Or cmbfiltrarpor1 = "Largura" Or cmbfiltrarpor1 = "Espessura" Then
                Select Case cmbfiltrarpor1
                    Case "Comprimento": TextoFiltro = "P.Comprimento"
                    Case "Largura": TextoFiltro = "P.Largura"
                    Case "Espessura": TextoFiltro = "P.Espessura"
                End Select
                valor = txtTexto1
                NovoValor = Replace(valor, ",", ".")
                StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
            Else
                Select Case cmbfiltrarpor1
                    Case "Código interno": TextoFiltro = "P.desenho"
                    Case "Código de referência": TextoFiltro = "IA.N_referencia"
                    Case "Descrição": TextoFiltro = "P.descricao"
                    Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                    Case "Dureza": TextoFiltro = "P.Dureza"
                    Case "Part number": TextoFiltro = "PFAB.Part_number"
                End Select
                StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio1, optMeio1, optFim1, optIgual1, txtTexto1) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregaListaCad

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_produto_padrao.PageCount - 1)
    Else
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_produto_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao1.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao1.AbsolutePage = -3 Then
        ProcExibePagina1 (TBLocalizar_produto_padrao1.PageCount - 1)
    Else
        TBLocalizar_produto_padrao1.AbsolutePage = TBLocalizar_produto_padrao1.AbsolutePage - 2
        ProcExibePagina1 (TBLocalizar_produto_padrao1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr1_Click()
On Error GoTo tratar_erro

If txtPagIr1 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas1.Caption, 4))
If Quant <= 1 Or txtPagIr1 > Quant Then Exit Sub
If txtPagIr1.Text >= 1 And txtPagIr1.Text <= Quant Then
    TBLocalizar_produto_padrao1.AbsolutePage = txtPagIr1.Text
    ProcExibePagina1 (TBLocalizar_produto_padrao1.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao1.AbsolutePage = 1
ProcExibePagina1 (TBLocalizar_produto_padrao1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao1.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao1.AbsolutePage = 1 Then
        ProcExibePagina1 (2)
    Else
        ProcExibePagina1 (TBLocalizar_produto_padrao1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (TBLocalizar_produto_padrao1.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao1.AbsolutePage = TBLocalizar_produto_padrao1.PageCount
ProcExibePagina1 (TBLocalizar_produto_padrao1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyReturn: ListView1_DblClick
            Case vbKeyF2: ProcFiltrarCarteira
            Case vbKeyF3: ProcAdicionar
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyReturn: ListView1_DblClick
            Case vbKeyF2: ProcFiltrar1
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11805, 6, True
ProcCarregaToolBar2 Me, 11805, 5, True

If Formulario <> "Estoque/Ordem de faturamento" Then
With SSTab1
    If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 1 Then .TabCaption(0) = "Pedidos de compra (remessa)"
    If Faturamento_ListaProdudos = False Then Texto = "Localizar serviços" Else Texto = "Localizar produtos"
    If Formulario = "Faturamento/Nota fiscal/Própria" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Própria - " & Texto
        .Tab = 0
    ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - " & Texto
            .TabVisible(0) = False
            .TabsPerRow = 1
            .Tab = 1
        ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                Caption = "Estoque - Ordem de faturamento - " & Texto
                .Tab = 0
            Else
                Caption = "Estoque - Nota fiscal - " & Texto
                .TabVisible(0) = False
                .TabsPerRow = 1
                .Tab = 1
    End If
End With

msk_data(0).Value = Date
msk_data(1).Value = Date
With frmEstoque_Ordem_Faturamento
    ProcFiltroPadrao cmbfiltrarpor1, optMeio1, optFim1, optIgual1, frmEstoque_Ordem_Faturamento.txtIDEmpresa, "Produtos/Serviços", "F", True
    If Permitido = False Then cmbfiltrarpor1 = "Código interno"
    If FunVerifNFProdServSemCad(IDempresa) = True Then optespecificacao.Enabled = False
    ProcCorrigeFormPedIntCompra
End With
Else
With SSTab1
    If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 1 Then .TabCaption(0) = "Pedidos de compra (remessa)"
    If Faturamento_ListaProdudos = False Then Texto = "Localizar serviços" Else Texto = "Localizar produtos"
    If Formulario = "Faturamento/Nota fiscal/Própria" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Própria - " & Texto
        .Tab = 0
    ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - " & Texto
            .TabVisible(0) = False
            .TabsPerRow = 1
            .Tab = 1
        ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                Caption = "Estoque - Ordem de faturamento - " & Texto
                .Tab = 0
            Else
                Caption = "Estoque - Nota fiscal - " & Texto
                .TabVisible(0) = False
                .TabsPerRow = 1
                .Tab = 1
    End If
End With

msk_data(0).Value = Date
msk_data(1).Value = Date
With frmEstoque_Ordem_Faturamento
    ProcFiltroPadrao cmbfiltrarpor1, optMeio1, optFim1, optIgual1, frmEstoque_Ordem_Faturamento.txtIDEmpresa, "Produtos/Serviços", "F", True
    If Permitido = False Then cmbfiltrarpor1 = "Código interno"
    If FunVerifNFProdServSemCad(IDempresa) = True Then optespecificacao.Enabled = False
    ProcCorrigeFormPedIntCompra
End With
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_carteira_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

MoedaNF = ""
If ColumnHeader = "" Then
    With Lista_carteira
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
            
            If Formulario <> "Estoque/Ordem de faturamento" Then
                If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then Familiatext = "produto" Else Familiatext = "serviço"
                If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
                    If frmEstoque_Ordem_Faturamento.FunVerifDadosPedido(.ListItems.Item(InitFor).SubItems(1), Familiatext, False) = False Then GoTo Proximo
                End If
            
            Else
                If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then Familiatext = "produto" Else Familiatext = "serviço"
                If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
                    If frmEstoque_Ordem_Faturamento.FunVerifDadosPedido(.ListItems.Item(InitFor).SubItems(1), Familiatext, False) = False Then GoTo Proximo
                End If
            
            End If
            
                If MoedaNF = "" Then MoedaNF = .ListItems.Item(InitFor).SubItems(16)
                If MoedaNF <> .ListItems.Item(InitFor).SubItems(16) Then GoTo Proximo
                If MoedaNF <> "REAL" Then
                    OutraMoeda = True
                    Moeda = MoedaNF
                End If
                
                .ListItems.Item(InitFor).Checked = True
            End If
Proximo:
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_carteira, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_carteira_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Formulario <> "Estoque/Ordem de faturamento" Then
With Lista_carteira
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
                If Opt_produto_filtrar = True Then Familiatext = "produto" Else Familiatext = "serviço"
            Else
                Familiatext = "produto"
            End If
            If MoedaNF = "" Then MoedaNF = .ListItems.Item(InitFor).SubItems(16)
            If MoedaNF <> .ListItems.Item(InitFor).SubItems(16) Then
                USMsgBox ("Só é permitido selecionar " & Familiatext & " de pedido com a moeda " & MoedaNF & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If MoedaNF <> "REAL" Then
                OutraMoeda = True
                Moeda = MoedaNF
            End If
            If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
                If frmEstoque_Ordem_Faturamento.FunVerifDadosPedido(.ListItems.Item(InitFor).SubItems(1), Familiatext, True) = False Then .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With
Else
With Lista_carteira
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
                If Opt_produto_filtrar = True Then Familiatext = "produto" Else Familiatext = "serviço"
            Else
                Familiatext = "produto"
            End If
            If MoedaNF = "" Then MoedaNF = .ListItems.Item(InitFor).SubItems(16)
            If MoedaNF <> .ListItems.Item(InitFor).SubItems(16) Then
                USMsgBox ("Só é permitido selecionar " & Familiatext & " de pedido com a moeda " & MoedaNF & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If MoedaNF <> "REAL" Then
                OutraMoeda = True
                Moeda = MoedaNF
            End If
            If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
                If frmEstoque_Ordem_Faturamento.FunVerifDadosPedido(.ListItems.Item(InitFor).SubItems(1), Familiatext, True) = False Then .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With
End If

If Permitido1 = False Then MoedaNF = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
With frmEstoque_Ordem_Faturamento
    If Optdescricao1.Value = True Then Permitido1 = False Else Permitido1 = True
    If Faturamento_ListaProdudos = True Then
        .ProcLimpaCamposProd1
        .txtCod_Produto.Text = ListView1.SelectedItem.ListSubItems(1)
        .ProcCarregaDadosProduto
    Else
        .ProcLimpaCamposServicos1
        .txtcodServ.Text = ListView1.SelectedItem.ListSubItems(1)
        .ProcCarregaDadosServico
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaCad()
On Error GoTo tratar_erro

lblRegistros1.Caption = "Nº de reg.: 0"
lblPaginas1.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub
Set TBLocalizar_produto_padrao1 = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao1.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao1.EOF = False Then ProcExibePagina1 (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina1(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_produto_padrao1.PageSize = IIf(txtNreg1 = "", 30, txtNreg1)
TBLocalizar_produto_padrao1.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao1.PageSize
ContadorReg = 1

If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
    PBLista1.Min = 0
    PBLista1.Max = TBLocalizar_produto_padrao1.PageSize
    PBLista1.Value = 1
    Contador = 0
End If
Do While TBLocalizar_produto_padrao1.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_produto_padrao1!Codproduto
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao1!Desenho), "", TBLocalizar_produto_padrao1!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao1!Descricao), "", TBLocalizar_produto_padrao1!Descricao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao1!Unidade), "", TBLocalizar_produto_padrao1!Unidade)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao1!Unidade_com), "", TBLocalizar_produto_padrao1!Unidade_com)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao1!Classe), "", TBLocalizar_produto_padrao1!Classe)
    End With
    TBLocalizar_produto_padrao1.MoveNext
    ContadorReg = ContadorReg + 1
    If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
        Contador = Contador + 1
        PBLista1.Value = Contador
    End If
Loop
lblRegistros1.Caption = "Nº de reg.: " & ContadorReg
If TBLocalizar_produto_padrao1.AbsolutePage = adPosBOF Then
   lblPaginas1.Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao1.PageCount
ElseIf TBLocalizar_produto_padrao1.AbsolutePage = adPosEOF Then
        lblPaginas1.Caption = "Pág.: " & TBLocalizar_produto_padrao1.PageCount & " de: " & TBLocalizar_produto_padrao1.PageCount
    Else
        lblPaginas1.Caption = "Pág.: " & TBLocalizar_produto_padrao1.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao1.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_data_CallbackKeyDown(index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeFormPedIntCompra()
On Error GoTo tratar_erro

If Formulario <> "Estoque/Ordem de faturamento" Then
If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 1 Then
    Frame1(25).Visible = False
    Frame5.Visible = False
    Frame1(23).Top = 1320
    Frame1(20).Top = 1320
'    ChkData(0).Top = 3150
'    ChkData(1).Top = 3150
    chkData(0).Caption = "Dt. compra"
    
    With Lista_carteira
        .Top = 2160
        .Height = 5535
    End With
End If

With cmbfiltrarpor
    .Clear
    If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
        .AddItem "Pedido do cliente"
        .AddItem "Pedido interno"
        .AddItem "Programa"
    Else
        .AddItem "Pedido de compra"
    End If
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    
    ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, frmEstoque_Ordem_Faturamento.txtIDEmpresa.Text, "Produtos/Serviços", "F", True
    If Permitido = False Then .Text = "Código interno"
End With

With Lista_carteira
    If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
        .ColumnHeaders(4).Width = 500
        .ColumnHeaders(8).Width = 900
        .ColumnHeaders(9).Width = 800
        .ColumnHeaders(11).Text = "Ped. int."
        .ColumnHeaders(11).Width = 850
        .ColumnHeaders(12).Width = 530
        .ColumnHeaders(13).Width = 1100
        .ColumnHeaders(14).Width = 530
        .ColumnHeaders(15).Width = 650
        .ColumnHeaders(16).Width = 650
        .ColumnHeaders(18).Text = "Qtde. vend."
        .ColumnHeaders(19).Width = 1000
        .ColumnHeaders(22).Width = 900
        .ColumnHeaders(23).Width = 900
    Else
        .ColumnHeaders(4).Width = 0
        .ColumnHeaders(8).Width = 0
        .ColumnHeaders(9).Width = 0
        .ColumnHeaders(11).Text = "Ped. compra"
        .ColumnHeaders(11).Width = 1000
        .ColumnHeaders(12).Width = 0
        .ColumnHeaders(13).Width = 0
        .ColumnHeaders(14).Width = 0
        .ColumnHeaders(15).Width = 0
        .ColumnHeaders(16).Width = 0
        .ColumnHeaders(18).Text = "Qtde. comp."
        .ColumnHeaders(19).Width = 0
        .ColumnHeaders(22).Width = 0
        .ColumnHeaders(23).Width = 0
    End If
End With
Else
If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 1 Then
    Frame1(25).Visible = False
    Frame5.Visible = False
    Frame1(23).Top = 1320
    Frame1(20).Top = 1320
'    ChkData(0).Top = 3150
'    ChkData(1).Top = 3150
    chkData(0).Caption = "Dt. compra"
    
    With Lista_carteira
        .Top = 2160
        .Height = 5535
    End With
End If

With cmbfiltrarpor
    .Clear
    If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
        .AddItem "Pedido do cliente"
        .AddItem "Pedido interno"
        .AddItem "Programa"
    Else
        .AddItem "Pedido de compra"
    End If
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    
    ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, frmEstoque_Ordem_Faturamento.txtIDEmpresa.Text, "Produtos/Serviços", "F", True
    If Permitido = False Then .Text = "Código interno"
End With

With Lista_carteira
    If Len(frmEstoque_Ordem_Faturamento.txttipocliente) = 2 And frmEstoque_Ordem_Faturamento.opt_Saida.Value = True Then
        .ColumnHeaders(4).Width = 500
        .ColumnHeaders(8).Width = 900
        .ColumnHeaders(9).Width = 800
        .ColumnHeaders(11).Text = "Ped. int."
        .ColumnHeaders(11).Width = 850
        .ColumnHeaders(12).Width = 530
        .ColumnHeaders(13).Width = 1100
        .ColumnHeaders(14).Width = 530
        .ColumnHeaders(15).Width = 650
        .ColumnHeaders(16).Width = 650
        .ColumnHeaders(18).Text = "Qtde. vend."
        .ColumnHeaders(19).Width = 1000
        .ColumnHeaders(22).Width = 900
        .ColumnHeaders(23).Width = 900
    Else
        .ColumnHeaders(4).Width = 0
        .ColumnHeaders(8).Width = 0
        .ColumnHeaders(9).Width = 0
        .ColumnHeaders(11).Text = "Ped. compra"
        .ColumnHeaders(11).Width = 1000
        .ColumnHeaders(12).Width = 0
        .ColumnHeaders(13).Width = 0
        .ColumnHeaders(14).Width = 0
        .ColumnHeaders(15).Width = 0
        .ColumnHeaders(16).Width = 0
        .ColumnHeaders(18).Text = "Qtde. comp."
        .ColumnHeaders(19).Width = 0
        .ColumnHeaders(22).Width = 0
        .ColumnHeaders(23).Width = 0
    End If
End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptFim1_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual1_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio1_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio1_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change()
On Error GoTo tratar_erro

If txtNreg <> "" Then
    VerifNumero = txtNreg
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg = ""
        txtNreg.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg1_Change()
On Error GoTo tratar_erro

If txtNreg1 <> "" Then
    VerifNumero = txtNreg1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg1 = ""
        txtNreg1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change()
On Error GoTo tratar_erro

If txtPagIr <> "" Then
    VerifNumero = txtPagIr
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr = ""
        txtPagIr.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr1_Change()
On Error GoTo tratar_erro

If txtPagIr1 <> "" Then
    VerifNumero = txtPagIr1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr1 = ""
        txtPagIr1.SetFocus
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

Lista_carteira.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto1_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtTexto1 <> "" Then cmbfamilia1.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAdicionar()
On Error GoTo tratar_erro

Permitido = False
Permitido1 = False

If Formulario <> "Estoque/Ordem de faturamento" Then
    With frmEstoque_Ordem_Faturamento
        If .txtNFiscal = "" Then NomeFormulario = "ordem de faturamento" Else NomeFormulario = "nota fiscal"
        If Len(.txttipocliente) = 2 Then
            If Faturamento_ListaProdudos = True Then NomeCampo = "produto" Else NomeCampo = "serviço"
        Else
            NomeCampo = "produto"
        End If
        
        For InitFor = 1 To Lista_carteira.ListItems.Count
            If Lista_carteira.ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If USMsgBox("Deseja realmente adicionar esse(s) " & NomeCampo & "(s) nesta " & NomeFormulario & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                        If OutraMoeda = True Then
                            If Moeda <> "REAL" Then
Mensagem:
                                Dolar = InputBox("Favor informar o valor do " & Moeda & ".")
                                If Dolar = "" Then Exit Sub
                                If IsNumeric(Dolar) = False Then
                                    USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                                    GoTo Mensagem
                                End If
                                ValorMoeda = Dolar
                            End If
                        End If
                        If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
                            .Faturamento_Vendas_PI = True
                            Conexao.Execute "UPDATE tbl_Dados_Nota_Fiscal Set Pedido_interno = 'True' where ID = " & .txtId
                        End If
                    Else
                        Exit Sub
                    End If
                End If
                
                Permitido = True
                If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
                    .ProcEnviaDadosPedido Lista_carteira.ListItems.Item(InitFor).ListSubItems(1), True
                Else
                    .ProcEnviaDadosPedido Lista_carteira.ListItems.Item(InitFor).ListSubItems(1), False
                End If
                
    '===============================================================================================
    ' Busca dados do pedido interno
    '===============================================================================================
    'If Pedido_int = True Then
    ID_pedido = Lista_carteira.ListItems.Item(InitFor).ListSubItems(1)
        
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select * from vendas_comercial As VC inner join Vendas_proposta as VP on VC.Cotacao = VP.Cotacao where VC.Cotacao = " & ID_pedido, Conexao, adOpenKeyset, adLockOptimistic
    Var = ID_pedido
    frmEstoque_Ordem_Faturamento.ProcSalvarTransporte
    '================================================================================================
                    
                If Faturamento_ListaProdudos = True Then
                    If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
                        If Permitido1 = False Then
                            If USMsgBox("Deseja agrupar o(s) produto(s) com o mesmo código interno, pedido do cliente e número do item?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido2 = True Else Permitido2 = False
                        End If
                        Permitido1 = True
                        .Procenviadadoslista Lista_carteira.ListItems.Item(InitFor), optespecificacao.Value, Permitido2, True, Lista_carteira.ListItems.Item(InitFor).ListSubItems(20)
                    Else
                        .Procenviadadoslista Lista_carteira.ListItems.Item(InitFor), optespecificacao.Value, False, False, Lista_carteira.ListItems.Item(InitFor).ListSubItems(20)
                    End If
                Else
                    .ProcEnviadadosListaServicos Lista_carteira.ListItems.Item(InitFor), optespecificacao.Value, Lista_carteira.ListItems.Item(InitFor).ListSubItems(20)
                End If
            End If
        Next InitFor
        
        If Permitido = False Then
            USMsgBox ("Informe o(s) " & NomeCampo & "(s) antes de gerar a ordem de faturamento."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        Else
        
     If Faturamento_ListaProdudos = True Then
     .ProcCarregaLista
     Else
     .ProcCarregaListaServicos
     End If
    
     .ProcGravarTotaisNota
     .ProcAtualizarDadosAdicionais
    
    .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
     ProcCarregaListaCarteira (1)
    .ProcCarregaDadosNota IIf(.txtId = "", 0, .txtId)
    
    End If
    End With
Else

    With frmEstoque_Ordem_Faturamento
        If .txtNFiscal = "" Then NomeFormulario = "ordem de faturamento" Else NomeFormulario = "nota fiscal"
        If Len(.txttipocliente) = 2 Then
            If Faturamento_ListaProdudos = True Then NomeCampo = "produto" Else NomeCampo = "serviço"
        Else
            NomeCampo = "produto"
        End If
        
        For InitFor = 1 To Lista_carteira.ListItems.Count
            If Lista_carteira.ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If USMsgBox("Deseja realmente adicionar esse(s) " & NomeCampo & "(s) nesta " & NomeFormulario & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                        If OutraMoeda = True Then
                            If Moeda <> "REAL" Then
Mensagem2:
                                Dolar = InputBox("Favor informar o valor do " & Moeda & ".")
                                If Dolar = "" Then Exit Sub
                                If IsNumeric(Dolar) = False Then
                                    USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                                    GoTo Mensagem2
                                End If
                                ValorMoeda = Dolar
                            End If
                        End If
                        If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
                            .Faturamento_Vendas_PI = True
                            Conexao.Execute "UPDATE tbl_Dados_Nota_Fiscal Set Pedido_interno = 'True' where ID = " & .txtId
                        End If
                    Else
                        Exit Sub
                    End If
                End If
                
                Permitido = True
                If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
                    .ProcEnviaDadosPedido Lista_carteira.ListItems.Item(InitFor).ListSubItems(1), True
                Else
                    .ProcEnviaDadosPedido Lista_carteira.ListItems.Item(InitFor).ListSubItems(1), False
                End If
                
    '===============================================================================================
    ' Busca dados do pedido interno
    '===============================================================================================
    'If Pedido_int = True Then
    ID_pedido = Lista_carteira.ListItems.Item(InitFor).ListSubItems(1)
        
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select * from vendas_comercial As VC inner join Vendas_proposta as VP on VC.Cotacao = VP.Cotacao where VC.Cotacao = " & ID_pedido, Conexao, adOpenKeyset, adLockOptimistic
    Var = ID_pedido
    frmEstoque_Ordem_Faturamento.ProcSalvarTransporte
    '================================================================================================
                    
                If Faturamento_ListaProdudos = True Then
                    If Len(.txttipocliente) = 2 And .opt_Saida.Value = True Then
                        If Permitido1 = False Then
                            If USMsgBox("Deseja agrupar o(s) produto(s) com o mesmo código interno, pedido do cliente e número do item?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido2 = True Else Permitido2 = False
                        End If
                        Permitido1 = True
                        .Procenviadadoslista Lista_carteira.ListItems.Item(InitFor), optespecificacao.Value, Permitido2, True, Lista_carteira.ListItems.Item(InitFor).ListSubItems(20)
                    Else
                        .Procenviadadoslista Lista_carteira.ListItems.Item(InitFor), optespecificacao.Value, False, False, Lista_carteira.ListItems.Item(InitFor).ListSubItems(20)
                    End If
                Else
                    .ProcEnviadadosListaServicos Lista_carteira.ListItems.Item(InitFor), optespecificacao.Value, Lista_carteira.ListItems.Item(InitFor).ListSubItems(20)
                End If
            End If
        Next InitFor
        
        If Permitido = False Then
            USMsgBox ("Informe o(s) " & NomeCampo & "(s) antes de gerar a ordem de faturamento."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        Else
        
     If Faturamento_ListaProdudos = True Then
     .ProcCarregaLista
     Else
     .ProcCarregaListaServicos
     End If
    
     .ProcGravarTotaisNota
     '.ProcAtualizarDadosAdicionais
    
    .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
     ProcCarregaListaCarteira (1)
    .ProcCarregaDadosNota IIf(.txtId = "", 0, .txtId)
    
    End If
    End With

End If



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrarCarteira
    Case 2: ProcAdicionar
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar1
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
