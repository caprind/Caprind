VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProd_Lista_Produto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Gerenciamento de ordem - Localizar pedidos interno"
   ClientHeight    =   8205
   ClientLeft      =   1680
   ClientTop       =   1365
   ClientWidth     =   14505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
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
      Height          =   1515
      Left            =   8760
      TabIndex        =   16
      Top             =   990
      Width           =   5655
      Begin VB.TextBox Txt_qtde_total_disp_produzindo 
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
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade disponível."
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox Txt_qtde_total_emp_produzindo 
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
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Quatidade total empenhada."
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox Txt_qtde_total_produzindo 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total produzindo."
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. empenhada"
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
         Index           =   1
         Left            =   2055
         TabIndex        =   20
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. produzindo"
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
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. disponível"
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
         Index           =   2
         Left            =   4005
         TabIndex        =   18
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-                                       ="
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
         Index           =   26
         Left            =   1860
         TabIndex        =   17
         Top             =   810
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   60
      TabIndex        =   21
      Top             =   990
      Width           =   8685
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   32
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
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
            TabIndex        =   28
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
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
            TabIndex        =   26
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
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
            TabIndex        =   25
            Top             =   180
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   27
            Top             =   180
            Width           =   1155
         End
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
         ItemData        =   "frmProd_Lista_Produto.frx":0000
         Left            =   180
         List            =   "frmProd_Lista_Produto.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.TextBox txtTexto 
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
         TabIndex        =   23
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
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
         ItemData        =   "frmProd_Lista_Produto.frx":0004
         Left            =   180
         List            =   "frmProd_Lista_Produto.frx":001D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3652
         TabIndex        =   30
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label45 
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
         Left            =   1522
         TabIndex        =   29
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   7290
      Width           =   14355
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
         Left            =   8880
         TabIndex        =   5
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
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
         Left            =   3780
         TabIndex        =   4
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11100
         TabIndex        =   9
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmProd_Lista_Produto.frx":0087
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
         Left            =   10560
         TabIndex        =   8
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmProd_Lista_Produto.frx":382B
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
         Left            =   9450
         TabIndex        =   6
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
         Left            =   10020
         TabIndex        =   7
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmProd_Lista_Produto.frx":7334
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
         Left            =   11640
         TabIndex        =   10
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmProd_Lista_Produto.frx":B423
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
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12390
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar               registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3090
         TabIndex        =   13
         Top             =   240
         Width           =   2760
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4755
      Left            =   60
      TabIndex        =   3
      Top             =   2520
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   8387
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Cód. cart."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Pedido int."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. int."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. ref."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5089
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Qt. vend."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Emp. est."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Emp. prod."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Necessidade"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "D"
         Text            =   "Pr. final"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Tag             =   "T"
         Text            =   "Ped. cliente"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "N. item"
         Object.Width           =   1587
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   7920
      Width           =   14355
      _ExtentX        =   25321
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4620
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmProd_Lista_Produto.frx":ECAF
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   31
      Top             =   0
      Width           =   14355
      _ExtentX        =   25321
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
   End
End
Attribute VB_Name = "frmProd_Lista_Produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLISTA_Ordem_PI     As ADODB.Recordset 'OK
Dim StrSql_PI_Localizar_empenho As String 'OK

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
Else
    txtTexto.Visible = True
    cmbTexto.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Ordem_PI.AbsolutePage <> 2 Then
    If TBLISTA_Ordem_PI.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Ordem_PI.PageCount - 1)
    Else
        TBLISTA_Ordem_PI.AbsolutePage = TBLISTA_Ordem_PI.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Ordem_PI.AbsolutePage)
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
    TBLISTA_Ordem_PI.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Ordem_PI.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Ordem_PI.AbsolutePage = 1
ProcExibePagina (TBLISTA_Ordem_PI.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Ordem_PI.AbsolutePage <> -3 Then
    If TBLISTA_Ordem_PI.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Ordem_PI.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Ordem_PI.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Ordem_PI.AbsolutePage = TBLISTA_Ordem_PI.PageCount
ProcExibePagina (TBLISTA_Ordem_PI.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
    Case vbKeyReturn: Lista_DblClick
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 14355, 5, True
cmbfiltrarpor = "Pedido interno"
If Compras_Requisicao = True Then
    With frmCompras_Requisicao
        Caption = "Outros - Solicitação - Localizar pedidos interno (Cód. interno: " & .txtN_Estoque & ")"
        Label1(0).Caption = "Qtde. solicitada"
        Txt_qtde_total_disp_produzindo.ToolTipText = "Quantidade total solicitada."
        Txt_qtde_total_produzindo = .Txt_qtde_total_solicitada
        Txt_qtde_total_emp_produzindo = .Txt_qtde_total_emp
        Txt_qtde_total_disp_produzindo = .Txt_qtde_total_disp
    End With
ElseIf Compras_Cotacao = True Then
        With frmcompras_reqcot
            Caption = "Compras - Cotação - Localizar pedidos interno (Cód. interno: " & .txtdesenho & ")"
            Label1(0).Caption = "Qtde. cotada"
            Txt_qtde_total_disp_produzindo.ToolTipText = "Quantidade total cotada."
            Txt_qtde_total_produzindo = .Txt_qtde_total_cotada
            Txt_qtde_total_emp_produzindo = .Txt_qtde_total_emp
            Txt_qtde_total_disp_produzindo = .Txt_qtde_total_disp
        End With
    ElseIf Compras_Pedido = True Then
            With frmCompras_Pedido
                Caption = "Compras - Pedido - Localizar pedidos interno (Cód. interno: " & IIf(Sit_REG = 0, .txtNomenclatura, .txtCodigo) & ")"
                Label1(0).Caption = "Qtde. comprada"
                Txt_qtde_total_disp_produzindo.ToolTipText = "Quantidade total comprada."
                Txt_qtde_total_produzindo = .Txt_qtde_total_comprada(Sit_REG)
                Txt_qtde_total_emp_produzindo = .Txt_qtde_total_emp(Sit_REG)
                Txt_qtde_total_disp_produzindo = .Txt_qtde_total_disp(Sit_REG)
            End With
        Else
            With frmprod
                Caption = "PCP - Gerenciamento de ordem - Localizar pedidos interno (Ordem: " & .txtof & ")"
                Label1(0).Caption = "Qtde. produzindo"
                Txt_qtde_total_disp_produzindo.ToolTipText = "Quantidade total produzindo."
                Txt_qtde_total_produzindo = .Txt_qtde_total_produzindo
                Txt_qtde_total_emp_produzindo = .Txt_qtde_total_emp_produzindo
                Txt_qtde_total_disp_produzindo = .Txt_qtde_total_disp_produzindo
                
                'Arruma tipo de filtro que vai aparecer
                cmbfiltrarpor.Clear
                If .Txt_ID_cliente = "" Or .txtCliente = "" Then cmbfiltrarpor.AddItem "Cliente"
                If .Txt_cod_prod = "" Then
                    cmbfiltrarpor.AddItem "Código de referência"
                    cmbfiltrarpor.AddItem "Código interno"
                    cmbfiltrarpor.AddItem "Descrição"
                    cmbfiltrarpor.AddItem "Família"
                End If
                cmbfiltrarpor.AddItem "Pedido do cliente"
                cmbfiltrarpor.AddItem "Pedido interno"
                cmbfiltrarpor = "Pedido interno"
            End With
End If
ProcCarregaComboFamilia cmbTexto, "familia <> 'Null' and vendas = 'True'", True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaPedidos(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Set TBLISTA_Ordem_PI = CreateObject("adodb.recordset")
TBLISTA_Ordem_PI.Open StrSql_PI_Localizar_empenho, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Ordem_PI.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Ordem_PI.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Ordem_PI.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Ordem_PI.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Ordem_PI.RecordCount - IIf(Pagina > 1, (TBLISTA_Ordem_PI.PageSize * (Pagina - 1)), 0), TBLISTA_Ordem_PI.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Ordem_PI.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems.Add(, , TBLISTA_Ordem_PI!CODIGO)
        .SubItems(1) = IIf(IsNull(TBLISTA_Ordem_PI!Ncotacao), "", TBLISTA_Ordem_PI!Ncotacao)
        .SubItems(2) = IIf(IsNull(TBLISTA_Ordem_PI!Revisao), "", TBLISTA_Ordem_PI!Revisao)
        .SubItems(3) = IIf(IsNull(TBLISTA_Ordem_PI!Desenho), "", TBLISTA_Ordem_PI!Desenho)
        .SubItems(4) = IIf(IsNull(TBLISTA_Ordem_PI!Rev_codinterno), "", TBLISTA_Ordem_PI!Rev_codinterno)
        .SubItems(5) = IIf(IsNull(TBLISTA_Ordem_PI!N_referencia), "", TBLISTA_Ordem_PI!N_referencia)
        .SubItems(6) = IIf(IsNull(TBLISTA_Ordem_PI!descricao_tecnica), "", TBLISTA_Ordem_PI!descricao_tecnica)
        .SubItems(7) = TBLISTA_Ordem_PI!Qtde_Faturar
        .SubItems(8) = TBLISTA_Ordem_PI!Qtde_emp_est
        .SubItems(9) = TBLISTA_Ordem_PI!Qtde_emp_prod
        .SubItems(10) = TBLISTA_Ordem_PI!Necessidade
        .SubItems(11) = IIf(IsNull(TBLISTA_Ordem_PI!PrazoFinal), "", Format(TBLISTA_Ordem_PI!PrazoFinal, "dd/mm/yy"))
        .SubItems(12) = IIf(IsNull(TBLISTA_Ordem_PI!Tipo), "", TBLISTA_Ordem_PI!Tipo)
        .SubItems(13) = IIf(IsNull(TBLISTA_Ordem_PI!PCCliente), "", TBLISTA_Ordem_PI!PCCliente)
        .SubItems(14) = IIf(IsNull(TBLISTA_Ordem_PI!N_item), "", TBLISTA_Ordem_PI!N_item)
    End With
    TBLISTA_Ordem_PI.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Ordem_PI.RecordCount
If TBLISTA_Ordem_PI.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Ordem_PI.PageCount
ElseIf TBLISTA_Ordem_PI.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Ordem_PI.PageCount & " de: " & TBLISTA_Ordem_PI.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Ordem_PI.AbsolutePage - 1 & " de: " & TBLISTA_Ordem_PI.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If Lista.ListItems.Count = 0 Then Exit Sub
valor = 0
Valor1 = 0
qtde_solicitada = ""
Mensagem1:

If Compras_Requisicao = True Then
    With frmCompras_Requisicao
        qtde_solicitada = IIf(.OptQS_est.Value = True, .txtQS_est, .txtQS_com)
    End With
ElseIf Compras_Cotacao = True Then
        qtde_solicitada = frmcompras_reqcot.Txt_quantidade
    ElseIf Compras_Pedido = True Then
            With frmCompras_Pedido
                qtde_solicitada = IIf(Sit_REG = 0, .txtQuantidade, .txtQtde_serv)
            End With
        Else
            qtde_solicitada = Lista.SelectedItem.SubItems(10)
End If
    
qtde_solicitada = InputBox("Favor informar a quantidade empenhada para este pedido.", , qtde_solicitada)
If qtde_solicitada = "" Then
    Unload Me
    Exit Sub
End If
If IsNumeric(qtde_solicitada) = False Then
    USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
    GoTo Mensagem1
End If
valor = qtde_solicitada
If valor <= 0 Then
    USMsgBox ("So é permitido quantidade maior que 0."), vbExclamation, "CAPRIND v5.0"
    GoTo Mensagem1
End If
Valor1 = Lista.SelectedItem.SubItems(10)
Valor2 = Txt_qtde_total_disp_produzindo
If Compras_Requisicao = False And Compras_Cotacao = False And Compras_Pedido = False Then
    With frmprod
        If .Txt_cod_prod = .txtdesenho And valor > Valor1 Then
            USMsgBox ("A quantidade empenhada não pode ser maior que a necessidade."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
    End With
End If
If Valor2 < valor Then
    USMsgBox ("Não é permitido empenhar, pois a quantidade disponível é menor que a quantidade empenhada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Compras_Requisicao = True Then
    With frmCompras_Requisicao
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Compras_pedido_lista_empenhos where IDlista = " & .TXTIDLista & " and IDCarteira = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            USMsgBox ("O produto/serviço já foi empenhado para este pedido interno."), vbExclamation, "CAPRIND v5.0"
            TBGravar.Close
            Exit Sub
        Else
            TBGravar.AddNew
        End If
        ProcEnviaDadosEmpenho 0, .TXTIDLista
        TBGravar.Update
        
        USMsgBox ("Produto/serviço empenhado para o pedido interno com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Outros/Solicitação"
        Evento = "Empenhar produto/serviço"
        ID_documento = TBGravar!ID
        Documento = "Nº solicitação: " & .txtNumero & " - Cód. interno: " & .txtN_Estoque
        Documento1 = "Pedido int.: " & Lista.SelectedItem.ListSubItems(1) & " - Rev.: " & Lista.SelectedItem.ListSubItems(2) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems(3) & " - Rev.: " & Lista.SelectedItem.ListSubItems(4)
        ProcGravaEvento
        '==================================
        .ProcCarregaListaEmpenhos
    End With
ElseIf Compras_Cotacao = True Then
        With frmcompras_reqcot
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Compras_pedido_lista_empenhos where IDlista = " & .TXTIDLista & " and IDCarteira = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                USMsgBox ("O produto/serviço já foi empenhado para este pedido interno."), vbExclamation, "CAPRIND v5.0"
                TBGravar.Close
                Exit Sub
            Else
                TBGravar.AddNew
            End If
            ProcEnviaDadosEmpenho 0, .TXTIDLista
            TBGravar.Update
            
            USMsgBox ("Produto/serviço empenhado para o pedido interno com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Empenhar produto/serviço"
            ID_documento = TBGravar!ID
            Documento = "Nº cotação: " & .txtidcotacao & " - Cód. interno: " & .txtdesenho
            Documento1 = "Pedido int.: " & Lista.SelectedItem.ListSubItems(1) & " - Rev.: " & Lista.SelectedItem.ListSubItems(2) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems(3) & " - Rev.: " & Lista.SelectedItem.ListSubItems(4)
            ProcGravaEvento
            '==================================
            .ProcCarregaListaEmpenhos
        End With
    ElseIf Compras_Pedido = True Then
            With frmCompras_Pedido
                If Sit_REG = 0 Then
                    IDlista = .TXTIDLista
                    TextoMsg = "produto"
                    TextoMsg1 = "Produto"
                    Desenho = .txtNomenclatura
                Else
                    IDlista = .txtIDLista_serv
                    TextoMsg = "serviço"
                    TextoMsg1 = "Serviço"
                    Desenho = .txtCodigo
                End If
                                
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from Compras_pedido_lista_empenhos where IDlista = " & IDlista & " and IDCarteira = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = False Then
                    USMsgBox ("O " & TextoMsg & " já foi empenhado para este pedido interno."), vbExclamation, "CAPRIND v5.0"
                    TBGravar.Close
                    Exit Sub
                Else
                    TBGravar.AddNew
                End If
                ProcEnviaDadosEmpenho 0, IDlista
                TBGravar.Update
                
                USMsgBox (TextoMsg1 & " empenhado para o pedido interno com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "Compras/Cotação"
                Evento = "Empenhar " & TextoMsg
                ID_documento = TBGravar!ID
                Documento = "Nº pedido: " & .txtPedido & " - Cód. interno: " & Desenho
                Documento1 = "Pedido int.: " & Lista.SelectedItem.ListSubItems(1) & " - Rev.: " & Lista.SelectedItem.ListSubItems(2) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems(3) & " - Rev.: " & Lista.SelectedItem.ListSubItems(4)
                ProcGravaEvento
                '==================================
                If Sit_REG = 0 Then .ProcCarregaListaEmpenhosProd Else .ProcCarregaListaEmpenhosServ
            End With
        Else
            With frmprod
Mensagem2:
                Familiatext = InputBox("Favor informar para qual ordem deseja empenhar.")
                If Familiatext <> "" Then
                    If IsNumeric(Familiatext) = False Then
                        USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                        GoTo Mensagem2
                    End If
                    Valor1 = Familiatext
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select NOF from Producao where Ordem = " & Valor1, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        USMsgBox ("Não foi encontrado nenhuma ordem com este número, favor alterar."), vbExclamation, "CAPRIND v5.0"
                        TBAbrir.Close
                        GoTo Mensagem2
                    End If
                    TBAbrir.Close
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select IDmateriaprima from Producaomaterial where Ordem = " & Valor1 & " and Codigo = '" & .txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        USMsgBox ("Não é permitido empenhar para esta ordem, pois a mesma não requisita o produto " & .txtdesenho & "."), vbExclamation, "CAPRIND v5.0"
                        TBAbrir.Close
                        GoTo Mensagem2
                    End If
                    TBAbrir.Close
                Else
                    Familiatext = ""
                End If
                
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from Producao_pedidos where Ordem = " & .txtof & " and IDCarteira = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = False Then
                    USMsgBox ("A ordem já foi empenhada para este pedido interno."), vbExclamation, "CAPRIND v5.0"
                    TBGravar.Close
                    Exit Sub
                Else
                    TBGravar.AddNew
                End If
                ProcEnviaDadosEmpenho .txtof, 0
                If Familiatext <> "" Then TBGravar!OrdemEmpenho = Familiatext
                TBGravar.Update
                
                Conexao.Execute "Update vendas_carteira Set Tem_ordem = 'True' where Codigo = " & Lista.SelectedItem
                
                USMsgBox ("Ordem empenhada para o pedido interno com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "PCP/Gerenciamento de ordem"
                Evento = "Empenhar ordem"
                ID_documento = TBGravar!ID
                Documento = "Ordem: " & .txtof.Text & " - Cód. interno: " & .txtdesenho
                Documento1 = "Pedido int.: " & Lista.SelectedItem.ListSubItems(1) & " - Rev.: " & Lista.SelectedItem.ListSubItems(2) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems(3) & " - Rev.: " & Lista.SelectedItem.ListSubItems(4)
                ProcGravaEvento
                '==================================
                .ProcCarregaListaPedidos
            End With
End If
ProcCarregaListaPedidos (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosEmpenho(Ordem As Long, IDlista As Long)
On Error GoTo tratar_erro

TBGravar!Data = Date
TBGravar!Responsavel = pubUsuario
If Ordem = 0 Then TBGravar!IDlista = IDlista Else TBGravar!Ordem = Ordem
TBGravar!IDcarteira = Lista.SelectedItem
TBGravar!Qtde_empenho = valor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

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

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If Sit_REG = 0 Then TextoFiltroPadrao_modulo = " and Necessidade > 0" Else TextoFiltroPadrao_modulo = " and Qtde_faturar > 0"
If Compras_Requisicao = True Then
    IDempresa = frmCompras_Requisicao.Cmb_empresa.ItemData(frmCompras_Requisicao.Cmb_empresa.ListIndex)
ElseIf Compras_Cotacao = True Then
        IDempresa = frmcompras_reqcot.Cmb_empresa.ItemData(frmcompras_reqcot.Cmb_empresa.ListIndex)
    ElseIf Compras_Pedido = True Then
            IDempresa = frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex)
        Else
            With frmprod
                TextoFiltroPadrao_modulo = ""
                If .Txt_ID_cliente = "" Or .txtCliente = "" Then TextoFiltroPadrao_modulo = "" Else TextoFiltroPadrao_modulo = "and IDcliente = " & .Txt_ID_cliente
                If .Txt_cod_prod <> "" Then
                    If TextoFiltroPadrao_modulo <> "" Then TextoFiltroPadrao_modulo = TextoFiltroPadrao_modulo & " and Desenho = '" & .Txt_cod_prod & "'" Else TextoFiltroPadrao_modulo = " and Desenho = '" & .Txt_cod_prod & "'"
                    If .Txt_cod_prod = .txtdesenho Then TextoFiltroPadrao_modulo = TextoFiltroPadrao_modulo & " and Necessidade > 0" Else TextoFiltroPadrao_modulo = TextoFiltroPadrao_modulo & " and Qtde_faturar > 0"
                End If
                IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            End With
End If

CamposFiltro = "CODIGO, Ncotacao, Revisao, Desenho, Rev_codinterno, N_Referencia, descricao_tecnica, Qtde_Faturar, Qtde_emp_est, Qtde_emp_prod, Necessidade, PrazoFinal, Tipo, PCCliente, N_item"
INNERJOINTEXTO = "Select " & CamposFiltro & " from Carteira_producao"
TextoFiltroPadrao = "ID_empresa = " & IDempresa & TextoFiltroPadrao_modulo

If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
    If cmbfiltrarpor = "Família" Then
        FiltroTexto = INNERJOINTEXTO & " where Familia = '" & cmbTexto & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor
            Case "Cliente": TextoFiltro = "cliente"
            Case "Código de referência": TextoFiltro = "n_referencia"
            Case "Código interno": TextoFiltro = "desenho"
            Case "Descrição": TextoFiltro = "descricao_tecnica"
            Case "Pedido do cliente": TextoFiltro = "PCcliente"
            Case "Pedido interno": TextoFiltro = "Ncotacao"
        End Select
        FiltroTexto = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    FiltroTexto = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If

StrSql_PI_Localizar_empenho = FiltroTexto
'Debug.print StrSql_PI_Localizar_empenho
ProcCarregaListaPedidos (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
