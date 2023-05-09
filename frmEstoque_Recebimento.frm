VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoque_Recebimento 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Recebimento"
   ClientHeight    =   10200
   ClientLeft      =   1050
   ClientTop       =   1665
   ClientWidth     =   15450
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
   ForeColor       =   &H00000000&
   Icon            =   "frmEstoque_Recebimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   15450
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista de produtos/serviços "
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
      Height          =   4065
      Left            =   55
      TabIndex        =   47
      Top             =   1950
      Width           =   11895
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5850
         Top             =   870
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView listprod 
         Height          =   3675
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   6482
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   7991
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Vlr. unitário"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Qtde. PÇ"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "D"
            Text            =   "Prazo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "IDLista"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da nota fiscal"
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
      Height          =   945
      Left            =   11970
      TabIndex        =   117
      Top             =   4170
      Width           =   3375
      Begin VB.TextBox txtSerie 
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
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   119
         ToolTipText     =   "Série."
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtnotafiscal 
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
         Left            =   90
         MaxLength       =   9
         TabIndex        =   118
         ToolTipText     =   "Número da nota fiscal."
         Top             =   480
         Width           =   945
      End
      Begin DrawSuite2022.USButton imgCalendario 
         Height          =   315
         Left            =   2910
         TabIndex        =   120
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":014A
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
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin MSMask.MaskEdBox txtdataemissao 
         Height          =   315
         Left            =   1920
         TabIndex        =   121
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   480
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin DrawSuite2022.USButton Cmd_salvar 
         Height          =   315
         Left            =   1560
         TabIndex        =   122
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":6F17
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
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1140
         TabIndex        =   125
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Data emissão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1935
         TabIndex        =   124
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Nota fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   123
         Top             =   285
         Width           =   750
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   11970
      TabIndex        =   115
      Top             =   5070
      Width           =   2025
      Begin VB.Label lblStatusNF 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Á emitir"
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
         Height          =   315
         Left            =   180
         TabIndex        =   116
         Top             =   450
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do pedido de compra"
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
      Height          =   975
      Left            =   55
      TabIndex        =   42
      Top             =   990
      Width           =   11895
      Begin VB.TextBox txtfornecedor 
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
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Fornecedor."
         Top             =   510
         Width           =   5910
      End
      Begin VB.TextBox txtProg_pedido 
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Programação de compra."
         Top             =   510
         Width           =   825
      End
      Begin VB.TextBox Txt_ID_pedido 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Left            =   120
         TabIndex        =   84
         ToolTipText     =   "ID do pedido"
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox txtuf 
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
         Left            =   9825
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "UF."
         Top             =   1260
         Width           =   360
      End
      Begin VB.TextBox Txt_ID_forn 
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
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Código do fornecedor."
         Top             =   510
         Width           =   555
      End
      Begin VB.TextBox txtEmpresa 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8010
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Empresa."
         Top             =   510
         Width           =   3645
      End
      Begin VB.TextBox txtdata 
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
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data."
         Top             =   510
         Width           =   705
      End
      Begin DrawSuite2022.USButton cmdPedido 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   510
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":F91C
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
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicSize         =   1
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin VB.TextBox txtID_empresa 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9810
         Locked          =   -1  'True
         MaxLength       =   255
         MouseIcon       =   "frmEstoque_Recebimento.frx":12F6C
         MousePointer    =   99  'Custom
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1290
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9525
         TabIndex        =   62
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblPedido 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "N° pedido"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4613
         TabIndex        =   44
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1522
         TabIndex        =   43
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados comerciais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   11970
      TabIndex        =   103
      Top             =   990
      Width           =   3375
      Begin VB.TextBox txtvalortotal 
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
         Height          =   345
         Left            =   2130
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   108
         TabStop         =   0   'False
         ToolTipText     =   "Valor total."
         Top             =   1260
         Width           =   1080
      End
      Begin VB.TextBox txtcondpagamento 
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
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   107
         TabStop         =   0   'False
         ToolTipText     =   "Condições de pagamento."
         Top             =   1980
         Width           =   2715
      End
      Begin VB.TextBox txtReferencia 
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   106
         TabStop         =   0   'False
         ToolTipText     =   "Condições de pagamento."
         Top             =   2610
         Width           =   3105
      End
      Begin VB.TextBox txtMoeda 
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
         Height          =   345
         Left            =   2130
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   105
         TabStop         =   0   'False
         ToolTipText     =   "Moeda do pedido de compra"
         Top             =   330
         Width           =   1080
      End
      Begin VB.TextBox txtvlrMoeda 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   104
         TabStop         =   0   'False
         ToolTipText     =   "Valor da moeda do dia."
         Top             =   810
         Width           =   1080
      End
      Begin DrawSuite2022.USButton cmdPagamento 
         Height          =   315
         Left            =   2910
         TabIndex        =   109
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":13276
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
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Condições comerciais"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   810
         TabIndex        =   114
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total pedido : "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   113
         Top             =   1290
         Width           =   1410
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1275
         TabIndex        =   112
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido de compras em :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   345
         TabIndex        =   111
         Top             =   390
         Width           =   1710
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor moeda :"
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
         Left            =   840
         TabIndex        =   110
         Top             =   870
         Width           =   1155
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar itens do pedido na lista"
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   60
      TabIndex        =   85
      Top             =   6000
      Width           =   3900
      Begin DrawSuite2022.USButton btnTodos 
         Height          =   345
         Left            =   2550
         TabIndex        =   88
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         DibPicture      =   "frmEstoque_Recebimento.frx":3137B
         Caption         =   "  Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin DrawSuite2022.USButton btnRecebidos 
         Height          =   345
         Left            =   1290
         TabIndex        =   87
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         DibPicture      =   "frmEstoque_Recebimento.frx":349CB
         Caption         =   "  Recebidos"
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
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton btnReceber 
         Height          =   345
         Left            =   120
         TabIndex        =   86
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         DibPicture      =   "frmEstoque_Recebimento.frx":3801B
         Caption         =   "  Á receber"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicSize         =   1
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10665
      FormWidthDT     =   15570
      FormScaleHeightDT=   10200
      FormScaleWidthDT=   15450
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   60
      TabIndex        =   48
      Top             =   7230
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   5265
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Recebimento"
      TabPicture(0)   =   "frmEstoque_Recebimento.frx":3B66B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "btnReceberTodos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdReceber"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Movimentação"
      TabPicture(1)   =   "frmEstoque_Recebimento.frx":3B687
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "cmdCancelar"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo do lote"
         Height          =   915
         Left            =   30
         TabIndex        =   130
         Top             =   330
         Width           =   1275
         Begin VB.OptionButton optPedido 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pedido"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   150
            TabIndex        =   132
            Top             =   570
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optNF 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Notafiscal"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   150
            TabIndex        =   131
            Top             =   330
            Width           =   1035
         End
      End
      Begin DrawSuite2022.USButton cmdCancelar 
         Height          =   2430
         Left            =   -60990
         TabIndex        =   41
         ToolTipText     =   "Cancelar recebimento (F4)"
         Top             =   330
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   4286
         DibPicture      =   "frmEstoque_Recebimento.frx":3B6A3
         Caption         =   "Cancelar recebimento (F4)"
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
         PicAlign        =   8
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   4
      End
      Begin DrawSuite2022.USButton cmdReceber 
         Height          =   750
         Left            =   14010
         TabIndex        =   19
         ToolTipText     =   "Receber (F3)"
         Top             =   1260
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1323
         DibPicture      =   "frmEstoque_Recebimento.frx":4156A
         Caption         =   "Receber item (F3)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         HandPointer     =   0   'False
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Height          =   2450
         Left            =   -74970
         TabIndex        =   57
         Top             =   330
         Width           =   13965
         Begin MSComctlLib.ListView Lista_movimentacao 
            Height          =   1860
            Left            =   180
            TabIndex        =   40
            Top             =   195
            Width           =   13665
            _ExtentX        =   24104
            _ExtentY        =   3281
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
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "RE"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Operação"
               Object.Width           =   4912
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Entrada"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Entrada PÇ"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Documento"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "N. de série"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Inspecionado?"
               Object.Width           =   2205
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "IDestCR"
               Object.Width           =   0
            EndProperty
         End
         Begin DrawSuite2022.USProgressBar PBlista1 
            Height          =   255
            Left            =   180
            TabIndex        =   64
            Top             =   2070
            Width           =   13695
            _ExtentX        =   24156
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do item a receber no estoque"
         ForeColor       =   &H00000000&
         Height          =   915
         Left            =   1320
         TabIndex        =   49
         Top             =   330
         Width           =   11235
         Begin VB.TextBox txtEspecificacoes 
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
            Left            =   2580
            Locked          =   -1  'True
            MaxLength       =   255
            MouseIcon       =   "frmEstoque_Recebimento.frx":442EF
            MousePointer    =   99  'Custom
            MultiLine       =   -1  'True
            TabIndex        =   127
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   480
            Width           =   3765
         End
         Begin VB.TextBox txtLote 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   97
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   480
            Width           =   1185
         End
         Begin VB.TextBox txtUnCom 
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
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "Unidade comercial."
            Top             =   480
            Width           =   630
         End
         Begin VB.ComboBox Cmb_codigo_ref 
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
            Left            =   6360
            Sorted          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Codigo de referência."
            Top             =   480
            Width           =   1530
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   480
            Width           =   1245
         End
         Begin VB.TextBox txtUN 
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
            Left            =   7890
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Unidade."
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtstatus 
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
            Left            =   9165
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   480
            Width           =   1980
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "* Lote *"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   450
            TabIndex        =   133
            Top             =   270
            Width           =   585
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4117
            TabIndex        =   126
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Un est."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7950
            TabIndex        =   92
            Top             =   270
            Width           =   525
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Código de ref."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6570
            TabIndex        =   79
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1425
            TabIndex        =   52
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Un com."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8550
            TabIndex        =   51
            Top             =   270
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9930
            TabIndex        =   50
            Top             =   270
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informações para recebimento do item"
         Enabled         =   0   'False
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
         Height          =   1665
         Left            =   30
         TabIndex        =   53
         Top             =   1260
         Width           =   13965
         Begin VB.CheckBox chkdtVencimento 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7740
            TabIndex        =   98
            Top             =   900
            Width           =   195
         End
         Begin DrawSuite2022.USCheckBox chkInspecao 
            Height          =   255
            Left            =   150
            TabIndex        =   94
            Top             =   810
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   450
            Caption         =   "Inspeção recebimento?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   0   'False
         End
         Begin VB.TextBox Txt_caminho2 
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
            Height          =   335
            Left            =   4680
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Caminho do certificado."
            Top             =   450
            Width           =   4515
         End
         Begin VB.CheckBox Chk_LA 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10770
            TabIndex        =   83
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox Chk_Dt_rcbto 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9270
            TabIndex        =   82
            Top             =   900
            Width           =   195
         End
         Begin VB.TextBox Txt_numero_serie 
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
            Height          =   335
            Left            =   10650
            MaxLength       =   50
            TabIndex        =   17
            ToolTipText     =   "Número de série."
            Top             =   1110
            Width           =   1455
         End
         Begin VB.TextBox txtQuantidade_PC 
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
            Left            =   10830
            MaxLength       =   50
            TabIndex        =   39
            ToolTipText     =   "Quantidade de peças à receber."
            Top             =   -960
            Width           =   1185
         End
         Begin VB.TextBox txtOBS 
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
            Left            =   2640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            ToolTipText     =   "Observações."
            Top             =   1110
            Width           =   5085
         End
         Begin VB.TextBox txtQuantidade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   12120
            MaxLength       =   50
            TabIndex        =   18
            ToolTipText     =   "Quantidade à receber."
            Top             =   1110
            Width           =   1365
         End
         Begin VB.TextBox txtcertificado 
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
            Height          =   335
            Left            =   3675
            MaxLength       =   50
            TabIndex        =   8
            ToolTipText     =   "Número do certificado."
            Top             =   450
            Width           =   990
         End
         Begin VB.TextBox txtcorrida 
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
            Height          =   335
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   7
            ToolTipText     =   "Número da corrida."
            Top             =   450
            Width           =   1020
         End
         Begin VB.ComboBox cmbLocal_armaz 
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
            Left            =   10215
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Local de armazenamento."
            Top             =   450
            Width           =   3630
         End
         Begin MSMask.MaskEdBox Txt_data_recebimento 
            Height          =   315
            Left            =   9210
            TabIndex        =   15
            ToolTipText     =   "Data do recebimento."
            Top             =   1110
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin DrawSuite2022.USButton imgCalendario_receb 
            Height          =   315
            Left            =   10290
            TabIndex        =   16
            Top             =   1110
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Recebimento.frx":445F9
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
            BorderColor     =   4960354
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   49152
            GradientColor1  =   4960354
            GradientColor2  =   4960354
            GradientColor3  =   4960354
            GradientColor4  =   4960354
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   49152
            GradientColorOver2=   49152
            GradientColorOver3=   49152
            GradientColorOver4=   49152
            GradientColorDown1=   32768
            GradientColorDown2=   32768
            GradientColorDown3=   32768
            GradientColorDown4=   32768
            ShowFocusRect   =   0   'False
            Theme           =   3
         End
         Begin DrawSuite2022.USButton cmdImportar2 
            Height          =   315
            Left            =   9210
            TabIndex        =   10
            ToolTipText     =   "Localizar arquivo do certificado da materia prima"
            Top             =   450
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Recebimento.frx":4B3C6
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
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin DrawSuite2022.USButton Cmd_visualizar_arquivo1 
            Height          =   315
            Left            =   9870
            TabIndex        =   12
            ToolTipText     =   "Visualizar arquivo do certificado da materia prima"
            Top             =   450
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Recebimento.frx":694CB
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
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   5
         End
         Begin DrawSuite2022.USButton Cmd_limpar_caminho1 
            Height          =   315
            Left            =   9540
            TabIndex        =   11
            ToolTipText     =   "Limpar caminho do certificado"
            Top             =   450
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Recebimento.frx":6CB1B
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
            BorderColor     =   4960354
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   49152
            GradientColor1  =   4960354
            GradientColor2  =   4960354
            GradientColor3  =   4960354
            GradientColor4  =   4960354
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   49152
            GradientColorOver2=   49152
            GradientColorOver3=   49152
            GradientColorOver4=   49152
            GradientColorDown1=   32768
            GradientColorDown2=   32768
            GradientColorDown3=   32768
            GradientColorDown4=   32768
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   3
         End
         Begin DrawSuite2022.USButton cmdcalc_peso 
            Height          =   315
            Left            =   13530
            TabIndex        =   90
            Top             =   1110
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Recebimento.frx":76367
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
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   5
         End
         Begin DrawSuite2022.USCheckBox chkEstoque 
            Height          =   255
            Left            =   150
            TabIndex        =   95
            Top             =   570
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   450
            Caption         =   "Movimenta estoque?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USCheckBox chkretorno 
            Height          =   255
            Left            =   150
            TabIndex        =   96
            Top             =   1050
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   450
            Caption         =   "Retorno de industrialização?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   0   'False
         End
         Begin MSMask.MaskEdBox txtVencimento 
            Height          =   315
            Left            =   7740
            TabIndex        =   99
            ToolTipText     =   "Informe a data de vencimento do lote"
            Top             =   1110
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin DrawSuite2022.USButton btnVencimento 
            Height          =   315
            Left            =   8820
            TabIndex        =   100
            ToolTipText     =   "Informe a data de vencimento do lote"
            Top             =   1110
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Recebimento.frx":90D13
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
            BorderColor     =   4960354
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   49152
            GradientColor1  =   4960354
            GradientColor2  =   4960354
            GradientColor3  =   4960354
            GradientColor4  =   4960354
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   49152
            GradientColorOver2=   49152
            GradientColorOver3=   49152
            GradientColorOver4=   49152
            GradientColorDown1=   32768
            GradientColorDown2=   32768
            GradientColorDown3=   32768
            GradientColorDown4=   32768
            ShowFocusRect   =   0   'False
            Theme           =   3
         End
         Begin DrawSuite2022.USCheckBox chkPerecivel 
            Height          =   255
            Left            =   150
            TabIndex        =   102
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   450
            Caption         =   "Perecível?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USCheckBox chkICMS 
            Height          =   255
            Left            =   150
            TabIndex        =   135
            Top             =   1290
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   450
            Caption         =   "Valor unitário com ICMS?"
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
            ShowFocusRect   =   0   'False
            Value           =   1
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. vencimento"
            DragMode        =   1  'Automatic
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7995
            TabIndex        =   101
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho do certificado"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6120
            TabIndex        =   89
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Número de série"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10740
            TabIndex        =   80
            Top             =   900
            Width           =   1170
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. receber PÇ"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10830
            TabIndex        =   75
            Top             =   -1170
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4710
            TabIndex        =   74
            Top             =   900
            Width           =   945
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. receb."
            DragMode        =   1  'Automatic
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9525
            TabIndex        =   73
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. receber"
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
            Left            =   12255
            TabIndex        =   61
            Top             =   900
            Width           =   1050
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Certificado"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3810
            TabIndex        =   56
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Corrida"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2925
            TabIndex        =   55
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Local armazenamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11040
            TabIndex        =   54
            Top             =   240
            Width           =   1560
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do recebimento"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   915
         Left            =   12570
         TabIndex        =   58
         Top             =   330
         Width           =   2685
         Begin VB.TextBox txtrecebida_PC 
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
            Left            =   6090
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade recebida em peça."
            Top             =   1590
            Width           =   1455
         End
         Begin VB.TextBox txtrequisitado_PC 
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
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade comprada em peça."
            Top             =   1590
            Width           =   1455
         End
         Begin VB.TextBox txtSaldo_PC 
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
            Left            =   7560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Saldo em peça."
            Top             =   1590
            Width           =   1485
         End
         Begin VB.TextBox txtSaldo 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1770
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Saldo."
            Top             =   480
            Width           =   825
         End
         Begin VB.TextBox txtrequisitado 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   90
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade comprada."
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtrecebida 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade recebida."
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Comprada PÇ"
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
            Left            =   4785
            TabIndex        =   78
            Top             =   1380
            Width           =   1125
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Recebida PÇ"
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
            Left            =   6300
            TabIndex        =   77
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo PÇ"
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
            Left            =   7935
            TabIndex        =   76
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2010
            TabIndex        =   71
            Top             =   270
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Recebida"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1035
            TabIndex        =   60
            Top             =   270
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Comprada"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   135
            TabIndex        =   59
            Top             =   270
            Width           =   735
         End
      End
      Begin DrawSuite2022.USButton btnReceberTodos 
         Height          =   870
         Left            =   14010
         TabIndex        =   134
         ToolTipText     =   "Receber selecionados na lista (F3)"
         Top             =   2040
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1535
         DibPicture      =   "frmEstoque_Recebimento.frx":97AE0
         Caption         =   "Receber itens selecionados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         HandPointer     =   0   'False
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
   Begin VB.TextBox txtidlista 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
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
      Left            =   360
      TabIndex        =   46
      ToolTipText     =   "ID da lista"
      Top             =   5130
      Visible         =   0   'False
      Width           =   1335
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   63
      Top             =   0
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   1720
      ButtonCount     =   8
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
      ButtonCaption2  =   "Relatório"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Relatório (F5)"
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
      ButtonWidth2    =   60
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Status"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Alterar status do(s) produto(s)/serviço(s) (F7)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   108
      ButtonTop3      =   2
      ButtonWidth3    =   45
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Atualizar"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Utilizado pelo administrador do sistema."
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
      ButtonLeft4     =   155
      ButtonTop4      =   2
      ButtonWidth4    =   59
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonAlignment5=   2
      ButtonType5     =   1
      ButtonStyle5    =   -1
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   -1
      ButtonLeft5     =   216
      ButtonTop5      =   4
      ButtonWidth5    =   2
      ButtonHeight5   =   54
      ButtonCaption6  =   "Ajuda"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Ajuda (F1)"
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   220
      ButtonTop6      =   2
      ButtonWidth6    =   41
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Sair"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Sair (Esc)"
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   263
      ButtonTop7      =   2
      ButtonWidth7    =   30
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState8    =   5
      ButtonLeft8     =   295
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5820
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmEstoque_Recebimento.frx":A1C8D
         Count           =   1
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   3960
      TabIndex        =   68
      Top             =   6000
      Width           =   11385
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
         TabIndex        =   21
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   300
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
         Left            =   6420
         TabIndex        =   22
         ToolTipText     =   "Número da página."
         Top             =   300
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   8640
         TabIndex        =   26
         ToolTipText     =   "Próxima página."
         Top             =   300
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":A5DCB
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
         Left            =   8100
         TabIndex        =   25
         ToolTipText     =   "Página anterior."
         Top             =   300
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":A956F
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
         Left            =   6990
         TabIndex        =   23
         Top             =   300
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
         Left            =   7560
         TabIndex        =   24
         ToolTipText     =   "Primeira página."
         Top             =   300
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":AD078
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
         Left            =   9180
         TabIndex        =   27
         ToolTipText     =   "Última página."
         Top             =   300
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_Recebimento.frx":B1167
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
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4410
         TabIndex        =   81
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3090
         TabIndex        =   72
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9960
         TabIndex        =   70
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   69
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   55
      TabIndex        =   65
      Top             =   6630
      Width           =   15285
      Begin VB.TextBox txtQtde_total 
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
         Height          =   315
         Left            =   13440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total."
         Top             =   180
         Width           =   1560
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   66
         Top             =   210
         Width           =   12105
         _ExtentX        =   21352
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
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total :"
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
         Left            =   12390
         TabIndex        =   67
         Top             =   180
         Width           =   2415
         WordWrap        =   -1  'True
      End
   End
   Begin DrawSuite2022.USButton cmdNota 
      Height          =   885
      Left            =   14040
      TabIndex        =   128
      ToolTipText     =   "Emitir nota fiscal."
      Top             =   5130
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1561
      DibPicture      =   "frmEstoque_Recebimento.frx":B49F3
      Caption         =   "Emitir nota Fiscal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USButton cmdDup 
      Height          =   855
      Left            =   14040
      TabIndex        =   129
      ToolTipText     =   "Gerar conta(s) a pagar do pedido de compras"
      Top             =   5130
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1508
      DibPicture      =   "frmEstoque_Recebimento.frx":BE4A0
      Caption         =   "Duplicatas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   1154291
      BorderColorDisabled=   13160660
      BorderColorDown =   16576
      BorderColorOver =   8438015
      GradientColor1  =   1154291
      GradientColor2  =   1154291
      GradientColor3  =   1154291
      GradientColor4  =   1154291
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   8438015
      GradientColorOver2=   8438015
      GradientColorOver3=   8438015
      GradientColorOver4=   8438015
      GradientColorDown1=   16576
      GradientColorDown2=   16576
      GradientColorDown3=   16576
      GradientColorDown4=   16576
      PicAlign        =   8
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   5
      ToolTipIcon     =   1
      ToolTipTitle    =   "CAPRIND v5.0"
   End
End
Attribute VB_Name = "frmEstoque_Recebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLISTA_Estoque_RecebimentoPedido As ADODB.Recordset 'OK
Public StrSql_Estoque_Recebimento_Localizar As String 'OK
Public StrSql_Estoque_Recebimento_LocalizarTotal As String 'OK
Public FormulaRel_Estoque_Recebimento As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=BZk-gwHpncU&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=47&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaMovimentacao()
On Error GoTo tratar_erro

Lista_Movimentacao.ListItems.Clear
If txtCodigo = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
StrSql = "select EM.*, ECR.recebido, ECR.ID, EC.Numero_serie,EC.Vencimento from (estoque_controle_recebimento ECR INNER JOIN estoque_movimentacao EM on ECR.ID = EM.idestoque_recebimento) INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where ECR.IDLista = " & IIf(TXTIDLista = "", 0, TXTIDLista) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " order by EM.Idoperacao"
'Debug.print StrSql

TBProduto.Open "select EM.*, ECR.recebido, ECR.ID, EC.Numero_serie,EC.vencimento from (estoque_controle_recebimento ECR INNER JOIN estoque_movimentacao EM on ECR.ID = EM.idestoque_recebimento) INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where ECR.IDLista = " & IIf(TXTIDLista = "", 0, TXTIDLista) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " order by EM.Idoperacao", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    PBLista1.Min = 0
    PBLista1.Max = TBProduto.RecordCount
    PBLista1.Value = 1
    Contador = 0
    Do While TBProduto.EOF = False
        With Lista_Movimentacao.ListItems
            .Add , , TBProduto!IDoperacao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProduto!IDEstoque), 0, TBProduto!IDEstoque)
            .Item(.Count).SubItems(2) = TBProduto!Operacao
            .Item(.Count).SubItems(3) = Format(TBProduto!Data, "dd/mm/yy")
            .Item(.Count).SubItems(4) = Format(TBProduto!Entrada, "###,##0.0000")
            .Item(.Count).SubItems(5) = IIf(IsNull(TBProduto!Entrada_PC), "", Format(TBProduto!Entrada_PC, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBProduto!Documento), "", TBProduto!Documento)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBProduto!Numero_serie), "", TBProduto!Numero_serie)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select ID, Laudo from Compras_recebimento where IDestoque = " & TBProduto!IDEstoque & " and DtValidacao is not null and Laudo = 'APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then .Item(.Count).SubItems(9) = "Sim" Else .Item(.Count).SubItems(9) = "Não"
            .Item(.Count).SubItems(10) = IIf(IsNull(TBProduto!ID), 0, TBProduto!ID)
        End With
        txtvencimento.Text = IIf(IsNull(TBProduto!Vencimento), "__/__/____", TBProduto!Vencimento)
        TBProduto.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnReceber_Click()
On Error GoTo tratar_erro

StatusItem = "and (Status_Item = 'N_RECEBIDO' or Status_Item = 'APROVADO' or Status_Item = 'PARCIAL')"
StatusItemRel = "and ({Estoque_recebimento_pedido.Status_Item} = 'N_RECEBIDO' or {Estoque_recebimento_pedido.Status_Item} = 'APROVADO' or {Estoque_recebimento_pedido.Status_Item} = 'PARCIAL')"

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnReceberTodos_Click()
On Error GoTo tratar_erro

If txtnotafiscal.Text = "" Then
    txtnotafiscal.SetFocus
    USMsgBox "Digite o numero da nota fiscal para receber os itens", vbInformation, "CAPRIND v5.0"
    Exit Sub
Else
NotaFiscal = txtnotafiscal.Text
End If


If txtSerie.Text = "" Then
    txtSerie.SetFocus
    USMsgBox "Digite a série da nota fiscal para receber os itens", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

If IsDate(txtDataemissao.Text) = False Then
    txtDataemissao.SetFocus
    USMsgBox "Digite a data da nota fiscal para receber os itens", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

IDpedido = Txt_ID_pedido.Text


If Chk_LA.Value <> 1 Then
If USMsgBox("Atenção, após o recebimento de vários itens no estoque ao mesmo tempo todos serão armazenados no local de armazenamento padrão do sistema." & vbCrLf & "Deseja continuar?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
End If

If Listprod.ListItems.Count = 0 Then
    USMsgBox "Selecione os itens a receber na lista", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If






With Listprod
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
        Desenho = .ListItems.Item(InitFor).ListSubItems(3).Text
        IDlista = .ListItems.Item(InitFor).ListSubItems(12).Text
            FunRecebePedidoCompra
        End If
    Next InitFor
End With

'=================================================================================================
' Corrige o status do pedido de compras
'=================================================================================================
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & IDpedido & " and (status_item = 'N_RECEBIDO' or status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then
    Status_pedido = "ENCERRADO"
Else
    Status_pedido = "PARCIAL"
End If

'Grava status do produto na ordem de compra se pedido em Aberto = False
Conexao.Execute "Update compras_pedido Set Status_pedido = '" & Status_pedido & "' where IDpedido = " & IDpedido

TBCompras.Close

btnRecebidos_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRecebidos_Click()
On Error GoTo tratar_erro

StatusItem = "and (Status_Item = 'PARCIAL' or Status_Item = 'RECEBIDO')"
StatusItemRel = "and ({Estoque_recebimento_pedido.Status_Item} = 'PARCIAL' or {Estoque_recebimento_pedido.Status_Item} = 'RECEBIDO')"

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub btnTodos_Click()
On Error GoTo tratar_erro

StatusItem = "and (Status_Item = 'N_RECEBIDO' or Status_Item = 'APROVADO' or Status_Item = 'PARCIAL' or Status_Item = 'RECEBIDO')"
StatusItemRel = "and ({Estoque_recebimento_pedido.Status_Item} = 'N_RECEBIDO' or {Estoque_recebimento_pedido.Status_Item} = 'APROVADO' or {Estoque_recebimento_pedido.Status_Item} = 'PARCIAL' or {Estoque_recebimento_pedido.Status_Item} = 'RECEBIDO')"

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnVencimento_Click()
On Error GoTo tratar_erro

Frmvencimento.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho1_Click()
On Error GoTo tratar_erro

Txt_caminho2 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_Click()
On Error GoTo tratar_erro

Acao = "salvar"
If TXTIDLista = "" Or TXTIDLista = "0" Then
    NomeCampo = "o produto/serviço"
    ProcVerificaAcao
    Exit Sub
End If
If Lista_Movimentacao.ListItems.Count = 0 Then
    NomeCampo = "a movimentação na lista"
    ProcVerificaAcao
    Exit Sub
End If
If Lista_Movimentacao.SelectedItem = False Then
    NomeCampo = "a movimentação na lista"
    ProcVerificaAcao
    Exit Sub
End If
If txtStatus = "NÃO_RECEBIDO" Then Exit Sub
If IsDate(txtDataemissao) = False Then
    NomeCampo = "a data de emissão da nota fiscal"
    ProcVerificaAcao
    txtDataemissao.SetFocus
    Exit Sub
End If
If txtnotafiscal = "" Then
    NomeCampo = "o número da nota fiscal"
    ProcVerificaAcao
    txtnotafiscal.SetFocus
    Exit Sub
End If
If txtSerie = "" Then
    NomeCampo = "a série"
    ProcVerificaAcao
    txtSerie.SetFocus
    Exit Sub
End If
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_movimentacao where Idoperacao = " & Lista_Movimentacao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    If IsNull(TBEstoque!IDEstoque_recebimento) = False And TBEstoque!IDEstoque_recebimento <> "" Then
        'Atualiza dados da nota na tabela de estoque_controle_recebimento
        Conexao.Execute "Update estoque_controle_recebimento Set Nota_fiscal = '" & txtnotafiscal & "', Serie = '" & txtSerie & "', Data_emissao =  '" & Format(txtDataemissao, "Short Date") & "' where Id = " & TBEstoque!IDEstoque_recebimento
        
        TBEstoque!Documento = txtnotafiscal
        TBEstoque.Update
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Recebimento/Pedido de compra"
        Evento = "Salvar nota fiscal"
        ID_documento = TXTIDLista
        Documento = "Cód. interno: " & txtCodigo & " - Nº lote: " & txtProg_pedido & " - Nº corrida: " & txtcorrida & " - Nº certificado: " & txtCertificado & " - Local armaz.: " & cmbLocal_armaz
        Documento1 = "Operação: " & Lista_Movimentacao.SelectedItem.SubItems(2) & " - Documento: " & Lista_Movimentacao.SelectedItem.SubItems(6)
        ProcGravaEvento
        '==================================
        ProcAtualizaVlrEntradaEstoque True
        ProcCarregaListaMovimentacao
    Else
        USMsgBox ("Não é possivel salvar os dados da nota fiscal, pois não foi encontrato o id do recebimento na movimentação"), vbExclamation, "CAPRIND v5.0"
    End If
End If
TBEstoque.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo1_Click()
On Error GoTo tratar_erro

If Txt_caminho2 <> "" Then ProcAbrirArquivo Txt_caminho2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDup_Click()
On Error GoTo tratar_erro


With frmEstoque_Recebimento
    If USMsgBox("Deseja realmente gerar contas a pagar do pedido de compra n° " & .txtProg_pedido.Text & "?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
    
    QtdeSaida = 0
    Contador = 1
    Contador2 = 0
    Controle = 1
    nPagto = 0
    Valor_Duplicatas = 0
    If Not IsNumeric(Left(.txtcondpagamento, 1)) Then
        USMsgBox ("O campo condições de pagamento tem que ser em dias, favor alterar."), vbExclamation, "CAPRIND v5.0"
        Unload Me
        Exit Sub
    End If
    
    QtdeSaida = Len(.txtcondpagamento)
            
    TextoCond = ""
    Do While Contador <= QtdeSaida
        If Mid(.txtcondpagamento, Contador, 1) = "/" Or Mid(.txtcondpagamento, Contador, 1) = "," Or IsNumeric(Mid(.txtcondpagamento, Contador, 1)) = True Then
            If TextoCond = "" Then TextoCond = Mid(.txtcondpagamento, Contador, 1) Else TextoCond = TextoCond & Mid(.txtcondpagamento, Contador, 1)
        End If
        Contador = Contador + 1
    Loop
    
        Qtd = 1
    
    'Verifica qtde. de parcelas
    Contador = 1
    QtdeSaida = Len(TextoCond)
    Do While Contador <= QtdeSaida
       Do While Mid(TextoCond, Contador, 1) <> "/" And Contador <= QtdeSaida
            Contador2 = Contador2 + 1
            Contador = Contador + 1
        Loop
        nPagto = nPagto + 1
        Contador = Contador + 1
    Loop
        
    TotalProduto = .txtValorTotal
    mxValorPag = Format(TotalProduto / nPagto, "###,##0.00")
    
    Contador = 1
    Contador3 = 1
    Dataini = Date
    Controle = 0
    Do While Contador <= QtdeSaida
        
        Contador2 = 0
        Do While Mid(.txtcondpagamento, Contador, 1) <> "/" And Contador <= QtdeSaida
            Contador2 = Contador2 + 1
            Contador = Contador + 1
        Loop
        
        mxCondpag = ReturnNumbersOnly(Mid(.txtcondpagamento, Contador3, Contador2))
        Contador3 = Contador3 + Contador2 + 1
        
        Controle = Controle + 1
        Par1 = Controle
        Par2 = nPagto
        If Len(Par1) = 1 Then
            Par1 = "00" & Par1
        ElseIf Len(Par1) = 1 Then
                Par1 = "0" & Par1
        End If
        If Len(Par2) = 1 Then
            Par2 = "00" & Par2
        ElseIf Len(Par2) = 1 Then
            Par2 = "0" & Par2
        End If
        
        Set TBReceber = CreateObject("adodb.recordset")
        TBReceber.Open "Select * from tbl_ContasPagar where txt_pedido = '" & .txtProg_pedido & "' and txt_parcela = '" & Par1 & "/" & Par2 & "' order by dt_pagamento", Conexao, adOpenKeyset, adLockOptimistic
        If TBReceber.EOF = True Then
            TBReceber.AddNew
        End If
        TBReceber!Data_transacao = Date
        TBReceber!Parcial = False
        TBReceber!impresso = False
        TBReceber!Bloqueado = False
        TBReceber!Logsit = "N"
        TBReceber!Despesas_NF = False
        TBReceber!Antecipacao = False
        TBReceber!Devolucao = False
        TBReceber!status = "TÍTULO EM ABERTO"
        TBReceber!Responsavel = pubUsuario
        TBReceber!ID_nota = 0
        TBReceber!txt_ndocumento = IIf(txtnotafiscal.Text <> "", txtnotafiscal.Text, "")
        TBReceber!dt_Pagamento = Dataini + mxCondpag
        TBReceber!Txt_pedido = .txtProg_pedido.Text
        TBReceber!Dt_emissao = Date
        TBReceber!int_codforn = .Txt_ID_forn.Text
        TBReceber!txt_condpag = ""
        TBReceber!dbl_valorpagto = mxValorPag * Qtd
        TBReceber!txt_Parcela = Par1 & "/" & Par2
        TBReceber!Txt_fornecedor = .txtFornecedor.Text
        TBReceber!ID_empresa = .txtID_empresa.Text
        TBReceber!Tipo = "FO"
        TBReceber!Class_conta = "REQ"
        TBReceber.Update
        
        'Fluxo de Caixa
        Set TBFluxo = CreateObject("adodb.recordset")
        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBReceber!IDFluxo), 0, TBReceber!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
        If TBFluxo.EOF = True Then TBFluxo.AddNew
        TBFluxo!Operacao = "À Debitar"
        TBFluxo!Data = TBReceber!dt_Pagamento
        TBFluxo!valor = TBReceber!dbl_valorpagto
        TBFluxo!Descricao = TBReceber!Txt_fornecedor
        TBFluxo!status = "N"
        TBFluxo!Documento = TBReceber!Txt_pedido
        TBFluxo!Bloqueado = False
        TBFluxo!ID_empresa = .txtID_empresa
        TBFluxo!IDintconta = TBReceber!IDintconta
        
        TBFluxo.Update
        Conexao.Execute "Update tbl_ContasPagar Set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBFluxo!IDintconta
        TBFluxo.Close
        
        'ProcCriaFamiliaFinanceiro .txtValorTotal.Text, .txtIDPedido
        TBReceber.Close
        Contador = Contador + 1
    Loop
        
    USMsgBox ("Nova(s) conta(s) enviada(s) para o financeiro com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Estoque/Recebimento/Pedido"
    Evento = "Enviar p/ financeiro"
    ID_documento = .txtProg_pedido
    Documento = "Nº pedido: " & .txtProg_pedido.Text
    Documento1 = ""
    ProcGravaEvento
    '==================================
End With
'Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar2_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_caminho2 = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNota_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

ID_nota = 0
Acao = "emitir a nota"
Servicos = False
Prod = False


With txtvlrMoeda

If txtMoeda.Text <> "REAL" And txtMoeda.Text <> "" And txtvlrMoeda.Text = "" Then
    .Locked = False
    .Enabled = True
    .BorderStyle = flex3D
    .BackColor = vbWhite
    .ForeColor = &H80&
    USMsgBox "Obrigatório confirmar o valor do dólar em reais", vbCritical, "CAPRIND v5.0"
    .SetFocus
    Exit Sub
End If

End With





If txtProg_pedido = "" Then
    NomeCampo = "o pedido"
    ProcVerificaAcao
    txtProg_pedido.SetFocus
    Exit Sub
End If
If txtuf <> "EX" Then
    If IsDate(txtDataemissao) = False Then
        NomeCampo = "a data de emissão da nota fiscal"
        ProcVerificaAcao
        txtDataemissao.SetFocus
        Exit Sub
    End If
    If txtnotafiscal = "" Then
        NomeCampo = "a nota fiscal"
        ProcVerificaAcao
        txtnotafiscal.SetFocus
        Exit Sub
    End If
    If txtSerie = "" Then
        NomeCampo = "a série"
        ProcVerificaAcao
        txtSerie.SetFocus
        Exit Sub
    End If
End If

TextoFiltro = ""
If txtuf <> "EX" Then TextoFiltro = " and ECR.Nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Data_emissao = '" & Format(txtDataemissao, "Short Date") & "'"

'Verifica se tem algum produto/serviço recebido para o pedido
Prodpedido = False
ServPedido = False
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select CPL.Tipo from (compras_pedido_lista CPL INNER JOIN Estoque_controle_recebimento ECR ON CPL.IDPedido = ECR.IDPedido and CPL.IDLista = ECR.IDLista and CPL.Desenho = ECR.Desenho) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDpedido where CP.Pedido = '" & txtProg_pedido & "'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = True Then
    USMsgBox ("É necessário receber o(s) produto(s)/serviço(s) deste pedido antes de emitir a nota."), vbExclamation, "CAPRIND v5.0"
    TBPedido.Close
    Exit Sub
Else
    'Verifica tipo da nota
    Do While TBPedido.EOF = False
        If TBPedido!Tipo = "P" Then Prodpedido = True Else ServPedido = True
        TBPedido.MoveNext
    Loop
End If
TBPedido.Close
If Prodpedido = True And ServPedido = True Then
    TipoNF = "M1SA"
ElseIf Prodpedido = True Then
        TipoNF = "M1"
    Else
        TipoNF = "SA"
End If

strPedido = txtProg_pedido
DataEmissao = txtDataemissao
frmEstoque_Recebimento_Menu.Show 1


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'Private Sub ProcSalvarCSTLista()
'On Error GoTo tratar_erro
'
'CST_ICMS = False
'CST_IPI = False
'CST_PIS = False
'CST_Cofins = False
'
''ICMS
'If IsNull(TBAbrir!txt_CST) = False And TBAbrir!txt_CST <> "" Then
'    InicioCST = Left(TBAbrir!txt_CST, 1)
'    If Len(TBAbrir!txt_CST) = 4 Then FimCST = Right(TBAbrir!txt_CST, 3) Else FimCST = Right(TBAbrir!txt_CST, 2)
'    CST_ICMS = True
'    CST_Cofins = False
'    CST_IPI = False
'    CST_PIS = False
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
''IPI
'If IsNull(TBAbrir!CST_IPI) = False And TBAbrir!CST_IPI <> "" Then
'    FimCST = TBAbrir!CST_IPI
'    CST_ICMS = False
'    CST_Cofins = False
'    CST_IPI = True
'    CST_PIS = False
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_IPI where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
''PIS
'If IsNull(TBAbrir!CST_PIS) = False And TBAbrir!CST_PIS <> "" Then
'    FimCST = TBAbrir!CST_PIS
'    CST_ICMS = False
'    CST_Cofins = False
'    CST_IPI = False
'    CST_PIS = True
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_PIS where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
''Cofins
'If IsNull(TBAbrir!CST_Cofins) = False And TBAbrir!CST_Cofins <> "" Then
'    FimCST = TBAbrir!CST_Cofins
'    CST_ICMS = False
'    CST_Cofins = True
'    CST_IPI = False
'    CST_PIS = False
'
'    Set TBCST = CreateObject("adodb.recordset")
'    TBCST.Open "select * from tbl_Detalhes_Nota_CST_Cofins where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
'    ProcEnviadadosListaCST
'    TBCST.Close
'End If
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
'
'Private Sub ProcEnviadadosListaCST()
'On Error GoTo tratar_erro
'
'If TBCST.EOF = True Then TBCST.AddNew
''ICMS
'If CST_ICMS = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Origem_mercadoria = InicioCST
'    TBCST!Tributacao_ICMS = FimCST
'    If FimCST <> "40" And FimCST <> "41" And FimCST <> "50" And FimCST <> "60" And FimCST <> "101" And FimCST <> "102" And FimCST <> "103" And FimCST <> "300" And FimCST <> "400" And FimCST <> "500" Then
'        If FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
'            If FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
'                Set TBMaquinas = CreateObject("adodb.recordset")
'                TBMaquinas.Open "Select * from regioes where uf = '" & UF & "'", Conexao, adOpenKeyset, adLockOptimistic
'                If TBMaquinas.EOF = False Then
'                    Set TBAfericao = CreateObject("adodb.recordset")
'                    TBAfericao.Open "Select * from tbl_ClassificacaoFiscal where txt_Ref = '" & TBAbrir!Txt_CF & "'", Conexao, adOpenKeyset, adLockOptimistic
'                    If TBAfericao.EOF = False Then
'                        Select Case TBMaquinas!regiao
'                            Case "DE":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTDE
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTDE
'                            Case "SS":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTSS
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTSS
'                            Case "NN":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTNN
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTNN
'                            Case "CO":
'                                TBCST!Percentual_reducao_BC = TBAfericao!CTCO
'                                If cbo_UF <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTCO
'                        End Select
'                    End If
'                    TBAfericao.Close
'                End If
'            End If
'        End If
'
'        If FimCST <> "201" And FimCST <> "202" And FimCST <> "203" Then
'            IntICMS = IIf(IsNull(TBAbrir!int_ICMS), 0, TBAbrir!int_ICMS)
'            If IntICMS <> 0 Then
'                TBCST!Valor_BC = Format(TBAbrir!dbl_ValorTotal, "###,##0.00")
'                Valortotal = TBCST!Valor_BC
'                TBCST!Valor_ICMS = Format((Valortotal * IntICMS) / 100, "###,##0.00")
'            Else
'                TBCST!Valor_BC = 0
'                TBCST!Valor_ICMS = 0
'            End If
'        End If
'    End If
'
'    If FimCST = "101" Or FimCST = "201" Or FimCST = "900" Then
'        TBCST!ICMS_SN = 0
'        TBCST!Valor_ICMS_SN = 0
'
'        IntICMS = IIf(IsNull(TBAbrir!ICMS_SN), 0, TBAbrir!ICMS_SN)
'        If IntICMS <> 0 Then
'            Valortotal = Format(TBAbrir!dbl_ValorTotal, "###,##0.00")
'            TBCST!ICMS_SN = IntICMS
'            TBCST!Valor_ICMS_SN = Format((Valortotal * IntICMS) / 100, "###,##0.00")
'        End If
'    End If
'End If
''IPI
'If CST_IPI = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Codigo_situacaoTributaria = FimCST
'    If FimCST = "00" Or FimCST = "49" Or FimCST = "50" Or FimCST = "99" Then TBCST!Valor_BC = TBAbrir!dbl_ValorTotal
'End If
''PIS
'If CST_PIS = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Codigo_situacaoTributaria = FimCST
'    If FimCST = "01" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBCST!Valor_BC = TBAbrir!dbl_ValorTotal
'End If
''Cofins
'If CST_Cofins = True Then
'    TBCST!Id_Item = TBAbrir!Int_codigo
'    TBCST!Codigo_situacaoTributaria = FimCST
'    If FimCST = "01" Or FimCST = "02" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBCST!Valor_BC = TBAbrir!dbl_ValorTotal
'End If
'TBCST.Update
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage <> 2 Then
    If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.PageCount - 1)
    Else
        TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = TBLISTA_Estoque_RecebimentoPedido.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)
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
    TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = 1
ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage <> -3 Then
    If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = TBLISTA_Estoque_RecebimentoPedido.PageCount
ProcExibePagina (TBLISTA_Estoque_RecebimentoPedido.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

ProcLimpar
Listprod.ListItems.Clear
ProcLimparCamposNF
ProcLimparCamposReq False
Lista_Movimentacao.ListItems.Clear
ProcCarregaPedido
ProcBloqueiaFrame

StatusItem = "and (Status_Item = 'N_RECEBIDO' or Status_Item = 'APROVADO' or Status_Item = 'PARCIAL')"
StatusItemRel = "and ({Estoque_recebimento_pedido.Status_Item} = 'N_RECEBIDO' or {Estoque_recebimento_pedido.Status_Item} = 'APROVADO' or {Estoque_recebimento_pedido.Status_Item} = 'PARCIAL')"

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCamposNF()
On Error GoTo tratar_erro

txtDataemissao = "__/__/____"
txtnotafiscal = ""
txtSerie = ""

lblStatusNF.Caption = "Á EMITIR"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

Permitido = False
Permitido1 = False
With Listprod
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido1 = False Then
                If USMsgBox("Deseja realmente alterar o status do(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            If Permitido1 = False Then
                frmEstoque_Recebimento_aut.Show 1
            End If
            If Permitido = False Then Exit Sub

            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from compras_pedido where pedido = '" & .ListItems.Item(InitFor).SubItems(2) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                Txt_ID_pedido = TBCompras_Pedido!IDpedido
                IDempresa = TBCompras_Pedido!ID_empresa
            End If
            TBCompras_Pedido.Close
            
            If .ListItems.Item(InitFor).SubItems(10) <> "RECEBIDO" Then
                status = "Status_item = 'RECEBIDO'"
            Else
                'Verifica se o status do produto esta aberto, parcial ou recebido
                quantnovo = 0
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select Sum(Recebido) as quantnovo from estoque_controle_recebimento where idpedido = " & Txt_ID_pedido & " and idlista = " & .ListItems.Item(InitFor) & " and Programacao = 'False' and id_empresa = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    quantnovo = IIf(IsNull(TBEstoque!quantnovo), 0, TBEstoque!quantnovo)
                End If
                TBEstoque.Close
                
                quantestoque = .ListItems.Item(InitFor).ListSubItems(7)
                If quantnovo = 0 Then
                    If FunVerifStatusAprovadoPC(txtID_empresa) = True Then status = "Status_item = 'APROVADO'" Else status = "Status_item = 'N_RECEBIDO'"
                ElseIf quantnovo < quantestoque Then
                        status = "Status_item = 'PARCIAL'"
                    Else
                        If quantnovo >= quantestoque Then status = "Status_item = 'RECEBIDO'"
                End If
            End If
            
            Conexao.Execute "Update compras_pedido_lista Set " & status & " where idpedido = " & Txt_ID_pedido & " and idlista = " & .ListItems.Item(InitFor)
            'Verifica status do item
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and (status_item = 'RECEBIDO' or status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = True Then
                If FunVerifStatusAprovadoPC(txtID_empresa) = True Then Status_pedido = "APROVADO" Else Status_pedido = "ABERTO"
            Else
                Set TBCompras = CreateObject("adodb.recordset")
                TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and (status_item = 'N_RECEBIDO' or status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                If TBCompras.EOF = True Then
                    Status_pedido = "ENCERRADO"
                Else
                    Status_pedido = "PARCIAL"
                End If
            End If
            TBCompras.Close
                
            'Grava status do produto na ordem de compra se pedido em Aberto = False
            Conexao.Execute "Update compras_pedido Set Status_pedido = '" & Status_pedido & "' where IDpedido = " & Txt_ID_pedido
            '==================================
            Modulo = "Estoque/Recebimento/Pedido de compra"
            Evento = "Alterar o status"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "N° pedido: " & .ListItems.Item(InitFor).SubItems(2) & " - Cód. interno: " & .ListItems.Item(InitFor).SubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Permitido1 = True
        End If
    Next InitFor
End With
If Permitido1 = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) na lista antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparCamposReq False
    ProcCarregaListaFiltro
    cmdReceber.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Estoque/Recebimento/Pedido de compra"
Direitos
ProcLimpaVariaveisPrincipais
Imprimir = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362ER" Then
    If USMsgBox("Deseja realmente atualizar os dados dos recebimentos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Estoque_controle_recebimento order by IDPedido, IDLista", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBAbrir.EOF = False
                If IsNull(TBEstoque!local_armaz) = True Or TBEstoque!local_armaz = "" Then TBEstoque!local_armaz = "N/A"
                
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from Compras_programa_item where ID = " & TBAbrir!IDpedido & " and ID_item = " & TBAbrir!IDlista, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then TBAbrir!Programacao = True Else TBAbrir!Programacao = False
                TBAbrir.Update
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBAbrir.Close
        
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "select * from estoque_movimentacao where operacao = 'ENTRADA_NOTA_FISCAL'", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBGravar.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBGravar.EOF = False
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "select * from estoque_controle where idestoque = " & TBGravar!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    'verifica se é pedido de compra
                    Set TBPedido = CreateObject("adodb.recordset")
                    TBPedido.Open "select * from compras_pedido where Pedido = '" & TBGravar!LOTE & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBPedido.EOF = False Then
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "select estoque_controle_recebimento.Id,estoque_controle_recebimento.Idlista from estoque_controle_recebimento inner join compras_pedido on estoque_controle_recebimento.IDPedido = compras_pedido.IDPedido where estoque_controle_recebimento.Desenho = '" & TBEstoque!Desenho & "' and estoque_controle_recebimento.certificado = '" & TBEstoque!Certificado & "' and estoque_controle_recebimento.corrida = '" & TBEstoque!Corrida & "' and estoque_controle_recebimento.Local_armaz = '" & TBEstoque!local_armaz & "' and compras_pedido.Pedido = '" & TBEstoque!LOTE & "' and estoque_controle_recebimento.Nota_fiscal = '" & TBGravar!Documento & "' and estoque_controle_recebimento.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            TBGravar!IDEstoque_recebimento = TBItem!ID
                            TBGravar!idlista_recebimento = TBItem!IDlista
                            TBGravar.Update
                        End If
                        TBItem.Close
                    End If
                    TBPedido.Close
                    'verifica se é programação de compra
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "select * from Compras_programa where Programatexto = '" & TBGravar!LOTE & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "select estoque_controle_recebimento.Id from estoque_controle_recebimento inner join Compras_programa on estoque_controle_recebimento.IDPedido = Compras_programa.ID where estoque_controle_recebimento.Desenho = '" & TBEstoque!Desenho & "' and estoque_controle_recebimento.certificado = '" & TBEstoque!Certificado & "' and estoque_controle_recebimento.corrida = '" & TBEstoque!Corrida & "' and estoque_controle_recebimento.Local_armaz = '" & TBEstoque!local_armaz & "' and Compras_programa.Programatexto = '" & TBEstoque!LOTE & "' and estoque_controle_recebimento.Nota_fiscal = '" & TBGravar!Documento & "' and estoque_controle_recebimento.ID_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            TBGravar!IDEstoque_recebimento = TBItem!ID
                            TBGravar.Update
                        End If
                        TBItem.Close
                    End If
                    TBAbrir.Close
                End If
                TBEstoque.Close
                TBGravar.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBGravar.Close
        
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal where int_TipoNota = 2 and Serie <> 'Null' order by int_NotaFiscal", Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBGravar.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBGravar.EOF = False
                If TBGravar!Serie <> "" Then Conexao.Execute "Update Estoque_controle_recebimento Set Serie = '" & TBGravar!Serie & "' where Nota_fiscal = '" & TBGravar!int_NotaFiscal & "' and id_empresa = " & txtID_empresa
                TBGravar.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBGravar.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Recebimento/Pedido de compra"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = True
Sit_Data = 1
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_receb_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = True
Sit_Data = 2
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_movimentacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

Permitido1 = False
If ColumnHeader = "" Then
    With Lista_Movimentacao
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifExcluirMovimentacao(.ListItems.Item(InitFor), .ListItems.Item(InitFor).SubItems(1), .ListItems.Item(InitFor).SubItems(9), .ListItems.Item(InitFor).SubItems(10), False) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_Movimentacao, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifExcluirMovimentacao(IDoperacao As Long, IDEstoque As Long, Inspecionado As String, IDestCR As Long, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifExcluirMovimentacao = True
'Verifica se houve saída
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select idoperacao from estoque_movimentacao where Idestoque = " & IDEstoque & " and Saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois já houve movimentação de saída neste RE."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If

'Verifica se tem nota fiscal emitida
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select NF.ID from (Estoque_controle_recebimento ECR INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = ECR.Nota_fiscal and NF.dt_DataEmissao = '" & txtDataemissao.Text & "' AND NF.Serie = ECR.Serie and NF.txt_Razao_Nome = '" & txtFornecedor & "' and NF.int_TipoNota = 2) INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_nota where ECR.ID = " & IDestCR & " and NFP.int_cod_produto = '" & txtCodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo está sendo utilizado no módulo Estoque/Nota fiscal."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If
TBAbrir.Close

If Inspecionado = "Sim" Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo já foi inspecionado."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    Exit Function
End If

'Verifica se esta amarrado alguma ordem
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select PNC.Ordem from estoque_movimentacao EM INNER JOIN Producao_NF_Consignada PNC ON EM.IDestoque = PNC.Idestoque where EM.idoperacao = " & IDoperacao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo está sendo vinculado a ordem " & TBAbrir!Ordem & "."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If

'Verifica se foi criado instrumento pelo RE
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select I.Codigo from estoque_movimentacao EM INNER JOIN Instrumentos I ON EM.IDestoque = I.Idestoque where EM.idoperacao = " & IDoperacao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido excluir este recebimento, pois o mesmo já foi cadastrado como instrumento."), vbExclamation, "CAPRIND v5.0"
    FunVerifExcluirMovimentacao = False
    TBAbrir.Close
    Exit Function
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Lista_movimentacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Permitido1 = False
With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifExcluirMovimentacao(.ListItems.Item(InitFor), .ListItems.Item(InitFor).SubItems(1), .ListItems.Item(InitFor).SubItems(9), .ListItems.Item(InitFor).SubItems(10), True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Listprod
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
    Frame6.Enabled = True
     ProcCarregaComboLA cmbLocal_armaz, False, False
Else
    ProcOrdenaListView Listprod, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Listprod.ListItems.Count = 0 Then Exit Sub
ProcLimpar
ProcLimparCamposReq True
ProcHabilitaFrame


With Listprod
    txtProg_pedido.Text = .SelectedItem.ListSubItems(2)
    
    If optNF.Value = True Then
        txtLote.Text = txtnotafiscal
    Else
        txtLote.Text = txtProg_pedido
    End If
    
    TXTIDLista = .SelectedItem
    
End With

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from compras_pedido_lista where idlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtCodigo = TBProduto!Desenho
    Cmb_codigo_ref = IIf(IsNull(TBProduto!N_referencia), "", TBProduto!N_referencia)
    txtEspecificacoes = TBProduto!Descricao
    txtUN = TBProduto!Un
    txtUnCom = TBProduto!Unidade_com
    
    If TBProduto!Un <> TBProduto!Unidade_com And IsNull(TBProduto!Qtde_estoque) = False And TBProduto!Qtde_estoque <> 0 Then
        txtrequisitado = Format(TBProduto!Qtde_estoque, "###,##0.0000")
    Else
        txtrequisitado = Format(TBProduto!Quant_Comp, "###,##0.0000")
    End If
    txtrequisitado_PC = IIf(IsNull(TBProduto!Quant_Comp_PC), "", TBProduto!Quant_Comp_PC)
    
    If TBProduto!Status_Item = "N_RECEBIDO" Or TBProduto!Status_Item = "APROVADO" Then
        StatusItem = "NÃO_RECEBIDO"
    Else
        StatusItem = TBProduto!Status_Item
    End If
    txtStatus = StatusItem
'==================================================================
' Carrega local de armazenamento por tipo
'==================================================================
    'If Chk_Dt_rcbto.Value = 0 Then
    Proccarregalocarm
    'End If
End If

'==================================================================
' Carrega local de armazenamento por CFOP
'==================================================================

Set TBCFOP = CreateObject("adodb.recordset")

cmbLocal_armaz.Enabled = True

TBCFOP.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = '" & TBProduto!ID_CFOP & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then

If TBCFOP!ID_CFOP = "5.124" Or TBCFOP!ID_CFOP = "6.124" Or TBCFOP!ID_CFOP = "1.124" Or TBCFOP!ID_CFOP = "2.124" Then
cmbLocal_armaz.Text = "INDUSTRIALIZAÇÃO"
cmbLocal_armaz.Enabled = False
End If

If TBCFOP!ID_CFOP = "5.901" Or TBCFOP!ID_CFOP = "6.901" Or TBCFOP!ID_CFOP = "1.901" Or TBCFOP!ID_CFOP = "2.901" Then
cmbLocal_armaz.Text = "RETORNO DE MERCADORIA"
cmbLocal_armaz.Enabled = False
End If

If TBCFOP!ID_CFOP = "5.902" Or TBCFOP!ID_CFOP = "6.902" Or TBCFOP!ID_CFOP = "1.902" Or TBCFOP!ID_CFOP = "2.902" Then
cmbLocal_armaz.Text = "RETORNO DE MERCADORIA"
cmbLocal_armaz.Enabled = False
End If

Else
cmbLocal_armaz.Enabled = True
End If
TBCFOP.Close
'================================================================================================================================

ProcCarregaPedido

'Verifica status do produto
Set TBProduto = CreateObject("adodb.recordset")

TBProduto.Open "Select * from estoque_controle_recebimento where idlista = " & TXTIDLista & " and Programacao = 'False' and id_empresa = " & txtID_empresa & " order by Id desc", Conexao, adOpenKeyset, adLockOptimistic

If TBProduto.EOF = False Then
    If TBProduto.RecordCount > 1 Then
        ProcLimparCamposNF
    Else
        txtDataemissao = IIf(IsNull(TBProduto!Data_emissao), "__/__/____", Format(TBProduto!Data_emissao, "dd/mm/yyyy"))
        txtnotafiscal = IIf(IsNull(TBProduto!Nota_fiscal), "", TBProduto!Nota_fiscal)
        txtSerie = IIf(IsNull(TBProduto!Serie), "", TBProduto!Serie)
        
        If txtMoeda.Text = "DOLAR" And IsDate(txtDataemissao) And IsNumeric(txtID_empresa.Text) And IsNumeric(Txt_ID_forn.Text) And txtSerie.Text <> "" Then
                TipoNF = "M1"
            
            '==============================================================================
            ' Verifica se existe cadastro da nota fiscal
            '==============================================================================
                        Set TBGravar = CreateObject("adodb.recordset")
                        StrSql = "Select * from tbl_Dados_Nota_Fiscal where dt_DataEmissao = '" & txtDataemissao.Text & "'  and ID_empresa = " & txtID_empresa.Text & " and Id_Int_Cliente = " & Txt_ID_forn.Text & " and int_NotaFiscal = '" & txtnotafiscal.Text & "' and Serie = '" & txtSerie & "' and int_TipoNota = 2 and TipoNF = '" & TipoNF & "'"
                        'Debug.print StrSql
                        
                        
                        TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                            If TBGravar.EOF = False Then
                            
                            txtvlrMoeda.Text = TBGravar!ValorMoeda
                                lblStatusNF.Caption = "EMITIDA"
                            
                            TBGravar.Close
                            Else
                                lblStatusNF.Caption = "Á EMITIR"
                            End If
        End If

    End If
    txtcorrida = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
    txtCertificado = IIf(IsNull(TBProduto!Certificado), "", TBProduto!Certificado)
    txtcorrida = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
    txtObs = IIf(IsNull(TBProduto!Obs), "", TBProduto!Obs)
    If IsNull(TBProduto!local_armaz) = False And TBProduto!local_armaz <> "" Then
        NomeCampo = "o local de armazenamento (" & TBProduto!local_armaz & ")"
        cmbLocal_armaz = TBProduto!local_armaz
    End If
Else
'ProcLimparCamposNF
End If
'==================================================================================================
' Dados do caminho do certificado
'==================================================================================================
 '   Txt_caminho2.Text = IIf(IsNull(TBProduto!CaminhoCertificado), "", TBProduto!CaminhoCertificado)
'==================================================================================================

TBProduto.Close

1:
'============================================================================
'Carrega qtde. recebida e atualiza o saldo
'============================================================================
    qtdeliberada = 0
    qtdeliberar = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(ISNULL(Recebido, 0)) as qtdeliberada, Sum(ISNULL(Recebido_PC, 0)) as qtdeliberar from estoque_controle_recebimento where idlista = " & TXTIDLista & " and Programacao = 'False' and id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
        qtdeliberar = IIf(IsNull(TBAbrir!qtdeliberar), 0, TBAbrir!qtdeliberar)
    End If
    TBAbrir.Close
    
    txtrecebida.Text = Format(qtdeliberada, "###,##0.0000")
    Qtde = txtrequisitado.Text
    qt = Qtde - qtdeliberada
    txtSaldo = IIf(qt < 0, 0, Format(qt, "###,##0.0000"))
    
    If txtrequisitado_PC <> "" Then
        txtrecebida_PC = qtdeliberar
        Qtde = txtrequisitado_PC
        qt = Qtde - qtdeliberar
        txtSaldo_PC = IIf(qt < 0, 0, qt)
    Else
        txtrecebida_PC = ""
        txtSaldo_PC = ""
    End If
    txtQuantidade = txtSaldo
    
    ProcCarregaListaMovimentacao
    Estoquereal = 0
    qtdeliberada = 0

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " desta movimentação, favor revisar."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAreceber_Click()
On Error GoTo tratar_erro

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optRecebidos_Click()
On Error GoTo tratar_erro

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opttodos_Click()
On Error GoTo tratar_erro

ProcCarregaListaFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaFiltro()
On Error GoTo tratar_erro

If txtProg_pedido.Text = "" Then
    USMsgBox ("Informe o número do pedido de compras antes de filtrar."), vbExclamation, "CAPRIND v5.0"
    txtProg_pedido.SetFocus
    Exit Sub
End If

CamposFiltro = "IDlista, ID_empresa, Pedido, Desenho, Descricao, UN, Unidade_com, preco_unitario, Quant_Comp, Quant_Comp_PC, Prazo, Status_item, Ordem, Qtde_estoque"
StrSql_Estoque_Recebimento_Localizar = "SELECT " & CamposFiltro & " FROM Estoque_recebimento_pedido where Pedido = '" & txtProg_pedido.Text & "'" & StatusItem & " group by " & CamposFiltro & " order by Pedido, Desenho"
StrSql_Estoque_Recebimento_LocalizarTotal = "SELECT Sum(Quant_Comp) as TotContas, IDlista FROM Estoque_recebimento_pedido where Pedido = '" & txtProg_pedido.Text & "'" & StatusItem & " group by IDlista, Data_recebimento, Nota_fiscal order by IDlista"
FormulaRel_Estoque_Recebimento = "{Estoque_recebimento_pedido.Pedido} = '" & txtProg_pedido.Text & "' " & StatusItemRel
ProcCarregaLista
ProcGravarDataFiltroRel Date, Date, False, 0, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub optNF_Click()
On Error GoTo tratar_erro

    If optNF.Value = True Then
        txtLote.Text = txtnotafiscal
    Else
        txtLote.Text = txtProg_pedido
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPedido_Click()
On Error GoTo tratar_erro

    If optPedido.Value = True Then
        txtLote.Text = txtProg_pedido
    Else
        txtLote.Text = txtnotafiscal
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

If txtCodigo.Text <> "" Then
Set TBInspecao = CreateObject("adodb.recordset")
TBInspecao.Open "Select estoque, insp_recebimento,Perecivel from projproduto where Desenho = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBInspecao.EOF = False Then
chkPerecivel.Value = IIf(IsNull(TBInspecao!perecivel), False, TBInspecao!perecivel)
chkEstoque.Value = TBInspecao!Estoque
chkInspecao.Value = TBInspecao!Insp_recebimento
End If
TBInspecao.Close
End If

If txtCodigo.Text <> "" Then
Set TBInspecao = CreateObject("adodb.recordset")
TBInspecao.Open "Select TNO.MaoObra from Compras_pedido_lista CPL inner join tbl_NaturezaOperacao TNO on TNO.IDCountCfop = CPL.ID_CFOP where IdLista = '" & Listprod.SelectedItem & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBInspecao.EOF = False Then
chkRetorno.Value = TBInspecao!MaoObra
End If
TBInspecao.Close
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEmpresa_Change()
On Error GoTo tratar_erro

Set TBEmpresa = CreateObject("adodb.recordset")
TBEmpresa.Open "Select Codigo from empresa where empresa = '" & txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBEmpresa.EOF = False Then
IDempresa = TBEmpresa!CODIGO
End If
TBEmpresa.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtID_empresa_Change()
On Error GoTo tratar_erro

IDempresa = txtID_empresa.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtMoeda_Change()
On Error GoTo tratar_erro

With txtvlrMoeda

If txtMoeda.Text <> "REAL" And txtMoeda.Text <> "" Then
    .Text = ""
    .Locked = False
    .Enabled = True
    .BorderStyle = flex3D
    .BackColor = vbWhite
    .ForeColor = &H80&
   ' .SetFocus
    Exit Sub
End If

If txtMoeda.Text = "REAL" And txtMoeda.Text <> "" Then
    .Text = "1,00"
    .Locked = True
    .Enabled = False
    .BackColor = &HE0E0E0
End If

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNotaFiscal_LostFocus()
On Error GoTo tratar_erro

If txtnotafiscal <> "" Then txtnotafiscal = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(txtnotafiscal), 9)

If txtMoeda.Text = "DOLAR" And IsDate(txtDataemissao) And IsNumeric(txtID_empresa.Text) And IsNumeric(Txt_ID_forn.Text) And txtSerie.Text <> "" Then
    TipoNF = "M1"

'==============================================================================
' Verifica se existe cadastro da nota fiscal
'==============================================================================
            Set TBGravar = CreateObject("adodb.recordset")
            StrSql = "Select * from tbl_Dados_Nota_Fiscal where dt_DataEmissao = '" & txtDataemissao.Text & "'  and ID_empresa = " & txtID_empresa.Text & " and Id_Int_Cliente = " & Txt_ID_forn.Text & " and int_NotaFiscal = '" & txtnotafiscal.Text & "' and Serie = '" & txtSerie & "' and int_TipoNota = 2 and TipoNF = '" & TipoNF & "'"
            'Debug.print StrSql
            
            
            TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = False Then
                
                txtvlrMoeda.Text = TBGravar!ValorMoeda
                    lblStatusNF.Caption = "EMITIDA"
                
                TBGravar.Close
                Else
                    lblStatusNF.Caption = "Á EMITIR"
                End If
End If

    If optNF.Value = True Then
        txtLote.Text = txtnotafiscal
    Else
        txtLote.Text = txtProg_pedido
    End If

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

Private Sub txtProg_pedido_Change()
On Error GoTo tratar_erro

If txtData <> "" Then Listprod.ListItems.Clear
ProcLimpar
ProcLimparCamposReq False
Lista_Movimentacao.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

StrSql_Estoque_Recebimento_Localizar = ""
frmEstoque_Recebimento_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcalc_peso_Click()
On Error GoTo tratar_erro

If txtCodigo = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & txtCodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = False
    Compras_Requisicao = False
    Compras_Cotacao = False
    Compras_Pedido = False
    Estoque_recebimento = True
    Vendas_Proposta = False
    Vendas_PI = False
    FrmCalculo_Peso.Show 1
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo tratar_erro
Dim Numero2     As Long 'OK
Dim ID_estoque  As Long 'OK

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido1 = False
With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido1 = False Then
                If USMsgBox("Deseja realmente excluir esta(s) movimentação(ões) do produto " & txtCodigo.Text & "?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido1 = True
            quantidade = 0
            'Verifica registro na tabela estoque_movimentacao/estoque_controle
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select idestoque, entrada, Entrada_PC, VlrUnit, IDEstoque_recebimento from estoque_movimentacao where idoperacao = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                ID_estoque = TBAbrir!IDEstoque
                qt = TBAbrir!Entrada
                Quant = IIf(IsNull(TBAbrir!Entrada_PC), 0, TBAbrir!Entrada_PC)
                valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
                
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select * from estoque_controle where idestoque = " & TBAbrir!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    'Altera qtde recebida na tabela estoque_controle_recebimento
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from compras_pedido where pedido = '" & txtProg_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then
                        NovoValor = Replace(qt, ",", ".")
                        If IsNull(TBAbrir!IDEstoque_recebimento) = False And TBAbrir!IDEstoque_recebimento <> 0 Then
                            TextoFiltro = "Id = " & TBAbrir!IDEstoque_recebimento
                        Else
                            TextoFiltro = "idpedido = " & TBCompras!IDpedido & " and idlista = " & TXTIDLista & " and Certificado = '" & txtCertificado & "' and Corrida = '" & txtcorrida & "' and local_armaz = '" & cmbLocal_armaz & "' and Programacao = 'False'"
                            If txtnotafiscal <> "" Then TextoFiltro = TextoFiltro & " and Nota_fiscal = '" & txtnotafiscal & "'"
                        End If
                        Set TBCompras_Lista = CreateObject("adodb.recordset")
                        TBCompras_Lista.Open "Select ID, Recebido, Recebido_PC, IDPedido, IDlista from estoque_controle_recebimento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCompras_Lista.EOF = False Then
                            TBCompras_Lista!Recebido = IIf(qt <> 0, TBCompras_Lista!Recebido - qt, 0)
                            TBCompras_Lista!Recebido_PC = IIf(Quant <> 0, TBCompras_Lista!Recebido_PC - Quant, 0)
                            TBCompras_Lista.Update
                            
                            ProcAtualizaQtdeRecebEmp TBCompras_Lista!IDpedido, TBCompras_Lista!IDlista
                            
                            If TBCompras_Lista!Recebido <= 0 Then Conexao.Execute "DELETE from estoque_controle_recebimento where ID = " & TBCompras_Lista!ID
                        End If
                        TBCompras_Lista.Close
                        
                        'Verifica se existe algum produto já recebido para definir o status do item
                        Set TBCompras_Lista = CreateObject("adodb.recordset")
                        TBCompras_Lista.Open "Select ID from estoque_controle_recebimento where idlista = " & TXTIDLista & " and idpedido = " & TBCompras!IDpedido & " and Programacao = 'False' and id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCompras_Lista.EOF = False Then
                            Status_pedido = "PARCIAL"
                        Else
                            If FunVerifStatusAprovadoPC(txtID_empresa) = True Then Status_pedido = "APROVADO" Else Status_pedido = "N_RECEBIDO"
                        End If
                        TBCompras_Lista.Close
                        
                        'Altera status do item
                        Conexao.Execute "Update compras_pedido_lista Set Status_item = '" & Status_pedido & "' where idlista = " & TXTIDLista
                        Conexao.Execute "Update Compras_Programacao set Compras_Programacao.Status_prog = '" & Status_pedido & "' from Compras_Programacao INNER JOIN compras_pedido_lista ON Compras_Programacao.ID_prog = compras_pedido_lista.ID_programacao where compras_pedido_lista.idlista = " & TXTIDLista
                        
                        'Altera status do pedido
                        Set TBPedido = CreateObject("adodb.recordset")
                        TBPedido.Open "Select * from compras_pedido_lista where idpedido = " & TBCompras!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
                        If TBPedido.EOF = False Then
                            Do While TBPedido.EOF = False
                                If TBPedido!Status_Item = "PARCIAL" Or TBPedido!Status_Item = "RECEBIDO" Then
                                    TBCompras!Status_pedido = "PARCIAL"
                                    GoTo 1
                                Else
                                    If FunVerifStatusAprovadoPC(txtID_empresa) = True Then TBCompras!Status_pedido = "APROVADO" Else TBCompras!Status_pedido = "ABERTO"
                                End If
                                TBPedido.MoveNext
                            Loop
                        End If
                        TBPedido.Close
                        
                        'Atualiza qtde recebida e status na programação
                        ProcAtualizaQtdeRecebidaProg "Select Compras_Programacao.* from Compras_Programacao INNER JOIN Compras_pedido_lista ON Compras_Programacao.ID_prog = Compras_pedido_lista.ID_programacao where Compras_pedido_lista.IDlista = " & TXTIDLista & " and Compras_Programacao.qtderecebida <> 0 order by Compras_Programacao.id_prog desc", True
                        ProcAlteraStatus_prog
                    End If
1:
                    TBCompras.Update
                    TBCompras.Close
                    
                    'Altera a qtde recebida na tabela estoque_controle
                    Permitido = False
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where desenho = '" & txtCodigo & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Permitido = True
                    End If
                    
                    'Verifica se o produto do pedido é remessa
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select CPL.* from Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho where CPL.Idlista = " & TXTIDLista & " and CPL.Remessa = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then Permitido = False
                    
                    'Verifica se o produto do pedido é mão de obra e se é a ultima fase da ordem e a ordem não controla estoque automaticamente, se for entra no estoque
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select CPL.OS, CPL.Ordem from (Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho) LEFT JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = CPL.ID_CFOP where CPL.Idlista = " & TXTIDLista & " and CFOP.MaoObra = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockReadOnly
                    If TBProduto.EOF = False Then
                        Permitido = False
                        If IsNull(TBProduto!OS) = False And TBProduto!OS <> "" And IsNull(TBProduto!Ordem) = False And TBProduto!Ordem <> "" Then
                            Set TBOrdemServico = CreateObject("adodb.recordset")
                            TBOrdemServico.Open "Select OS.idproducao from OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.Ordem = " & TBProduto!Ordem & " and P.Entrar_estoque <> 'True' ORDER BY OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockReadOnly
                            If TBOrdemServico.EOF = False Then
                                TBOrdemServico.MoveLast
                                If TBOrdemServico!IDProducao = TBProduto!OS Then Permitido = True
                            End If
                            TBOrdemServico.Close
                        End If
                    End If
                    TBProduto.Close
                    
                    'Verifica se tem centro de custo amarrado ao produto, se tiver não controla estoque
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Compras_pedido_lista_custo.* from Compras_pedido_lista_custo INNER JOIN Compras_pedido ON Compras_pedido_lista_custo.IDPedido = Compras_pedido.IDPedido where Compras_pedido_lista_custo.IDLista = " & TXTIDLista.Text & " and Compras_pedido.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        Permitido = False
                    End If
                    TBProduto.Close
                    
                    If Permitido = True Then
                        TBEstoque!estoque_real = TBEstoque!estoque_real - qt
                        TBEstoque!estoque_real_PC = TBEstoque!estoque_real_PC - Quant
                        TBEstoque!estoque_venda = TBEstoque!estoque_real
                        TBEstoque.Update
                    End If
                End If
                TBEstoque.Close
            End If
            
            Conexao.Execute "DELETE from estoque_movimentacao where idoperacao = " & .ListItems(InitFor)
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select Idoperacao from estoque_movimentacao where idestoque = " & ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = True Then
                Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & ID_estoque
                Conexao.Execute "DELETE from estoque_controle where idestoque = " & ID_estoque
            End If
            ' Apagar o registro da tabela estoque_Controle_Recebimento
            StrSql = "DELETE from Estoque_controle_recebimento where IdLista = " & TXTIDLista.Text & " and Nota_Fiscal = '" & txtnotafiscal & "' and Recebido = '" & Replace(txtrecebida.Text, ",", ".") & "'"
            'Debug.print StrSql
            
            Conexao.Execute StrSql
            '====================================================

            
            'Centro de custo
            Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & .ListItems(InitFor)
            
            Set TBNivel2 = CreateObject("adodb.recordset")
            TBNivel2.Open "Select sum(Saida) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtCodigo & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
            If TBNivel2.EOF = False Then
                Valor1 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
            End If
            TBNivel2.Close
            Set TBNivel2 = CreateObject("adodb.recordset")
            TBNivel2.Open "Select sum(entrada) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtCodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBNivel2.EOF = False Then
                Valor2 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
            End If
            TBNivel2.Close
            If Valor1 > Valor2 Then Conexao.Execute "UPDATE estoque_movimentacao set Terceiros = 'True' where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtCodigo & "' and destino = 'Terceiros'"
            
            '==================================
            Modulo = "Estoque/Recebimento/Pedido de compra"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & txtCodigo & " - Nº lote: " & txtProg_pedido & " - Nº corrida: " & txtcorrida & " - Nº certificado: " & txtCertificado & " - Local armaz.: " & cmbLocal_armaz
            Documento1 = "Operação: " & .ListItems(InitFor).SubItems(2) & " - Documento: " & .ListItems(InitFor).SubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido1 = False Then
    USMsgBox ("Informe a(s) movimentação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Movimentação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
    ProcLimparCamposNF
    ProcLimparCamposReq False
    ProcCarregaListaMovimentacao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagamento_Click()
On Error GoTo tratar_erro

If txtProg_pedido.Text = "" And txtProg_pedido = "" Then Exit Sub
Aplic = 1
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = True
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdReceber_Click()
On Error GoTo tratar_erro

If optNF.Value = False And optPedido.Value = False Then
    USMsgBox "Escolha um tipo de lote para entrada no estoque", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If chkICMS.Value = False Then
    USMsgBox ("Atenção " & pubUsuario & " o valor unitário a ser gravado no estoque será o valor unitário do pedido de compras sem o valor do ICMS."), vbInformation, "CAPRIND v5.0"
Else
    USMsgBox ("Atenção " & pubUsuario & " o valor unitário a ser gravado no estoque será o valor unitário do pedido de compras."), vbInformation, "CAPRIND v5.0"
End If

quantnovo = 0
quantestoque = 0
qt = 0
valor = 0
ValorTotal = 0
IDEstoque = 0
Acao = "receber no estoque"
If txtProg_pedido = "" Then
    NomeCampo = "o número do pedido/programa"
    ProcVerificaAcao
    txtProg_pedido.SetFocus
    Exit Sub
End If

If IsDate(txtvencimento) = False And chkPerecivel.Value = True Then
    NomeCampo = "a data de vencimento do lote"
    ProcVerificaAcao
    txtvencimento.SetFocus
    Exit Sub
End If

If TXTIDLista.Text = "" Or txtCodigo.Text = "" Then
    NomeCampo = "o produto/serviço"
    ProcVerificaAcao
    Exit Sub
End If

If txtStatus = "RECEBIDO" Then
    USMsgBox ("Este produto já foi recebido no estoque."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtuf <> "EX" Then
    If txtnotafiscal = "" Then
        If USMsgBox("O número da nota fiscal não foi informado, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            txtnotafiscal.SetFocus
            Exit Sub
        End If
    Else
        If IsDate(txtDataemissao) = False Then
            NomeCampo = "a data de emissão da nota fiscal"
            ProcVerificaAcao
            txtDataemissao.SetFocus
            Exit Sub
        End If
        If txtSerie = "" Then
            NomeCampo = "o número de série da nota fiscal"
            ProcVerificaAcao
            txtSerie.SetFocus
            Exit Sub
        End If
    End If
End If

StrSql = "UPDATE Estoque_Controle SET  NF = EM.Documento from  Estoque_Controle EC INNER JOIN Estoque_movimentacao EM ON EC.IdEstoque = EM.IdEstoque Where EM.Operacao = 'ENTRADA_NOTA_FISCAL'"

'Debug.print StrSql

Conexao.Execute (StrSql)

TextoLocal = cmbLocal_armaz
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Remessa, Tipo, ID_CFOP from Compras_pedido_lista where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!Tipo = "S" Then
        'Permitido = False
        TextoLocal = "SERVIÇOS"
    ElseIf TBProduto!Remessa = True Then
        Permitido = False
        TextoLocal = "RETORNO DE MERCADORIA"
    ElseIf IsNull(TBProduto!ID_CFOP) = False And TBProduto!ID_CFOP <> "" And TBProduto!ID_CFOP <> "0" Then
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select IDCountCfop from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP & " and MaoObra = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            Permitido = False
            TextoLocal = "INDUSTRIALIZAÇÃO"
        End If
        TBCFOP.Close
    End If
End If
TBProduto.Close

If Permitido = True Then
    If txtcorrida = "" Then txtcorrida = 0
    If txtCertificado = "" Then txtCertificado = 0
    If cmbLocal_armaz = "" Then
        NomeCampo = "o local de armazenamento"
        ProcVerificaAcao
        cmbLocal_armaz.SetFocus
        Exit Sub
    End If
End If

If IsDate(Txt_data_recebimento) = False Then
    NomeCampo = "a data do recebimento"
    ProcVerificaAcao
    Txt_data_recebimento.SetFocus
    Exit Sub
End If
'pega quantidade recebida na caixa de texto
valor = IIf(txtQuantidade = "", 0, txtQuantidade)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQuantidade.SetFocus
    Exit Sub
End If

Valor_Cofins_Prod = IIf(txtQuantidade_PC = "", 0, txtQuantidade_PC)
If txtQuantidade_PC <> "" Then
    If Valor_Cofins_Prod - Int(Valor_Cofins_Prod) > 0 Then
        USMsgBox ("Só é permitido número inteiro na quantidade de peças a receber."), vbExclamation, "CAPRIND v5.0"
        txtQuantidade_PC.SetFocus
        Exit Sub
    End If
Else
    Valor_Cofins_Prod = FunCalculaQtdePC(txtCodigo, txtQuantidade, True, txtUN)
End If

Valor1 = IIf(txtrequisitado_PC = "", 0, txtrequisitado_PC)
If Valor1 > 0 And Valor_Cofins_Prod <= 0 Then
    NomeCampo = "a quantidade de peças"
    ProcVerificaAcao
    txtQuantidade_PC.SetFocus
    Exit Sub
End If

If txtnotafiscal <> "" Then txtnotafiscal = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(txtnotafiscal), 9)
If txtProg_pedido.Text <> "" Then
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from compras_pedido where pedido = '" & txtProg_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        Txt_ID_pedido = TBCompras_Pedido!IDpedido
    End If
    
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from Compras_pedido_lista where IDpedido = " & Txt_ID_pedido & " and IDlista <> " & IDlista & " and Desenho = '" & txtCodigo & "' and Prazo < '" & Listprod.SelectedItem.ListSubItems(9) & "' and Status_item <> 'RECEBIDO' and Status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        USMsgBox ("Não é permitido receber este produto/serviço, pois existe(m) outro(s) em aberto com o código interno " & txtCodigo & " e prazo de entrega menor que " & Format(Listprod.SelectedItem.ListSubItems(9), "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
        TBCompras_Pedido.Close
        Exit Sub
    End If
    TBCompras_Pedido.Close
End If

'Verifica se o produto já foi recebido
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select IDlista from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and idlista = " & TXTIDLista.Text & " and status_item = '" & "RECEBIDO" & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    USMsgBox ("Este produto/serviço já foi recebido no estoque " & vbCrLf & "Código interno: " & txtCodigo.Text & vbCrLf & " descrição : " & txtEspecificacoes & "."), vbExclamation, "CAPRIND v5.0"
    txtnotafiscal.SetFocus
    TBCompras.Close
    Exit Sub
End If
TBCompras.Close

'Grava movimentação na tabela estoque_controle
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!Desenho = txtCodigo.Text
TBEstoque!Ref = IIf(Cmb_codigo_ref = "", Null, Cmb_codigo_ref)
TBEstoque!Descricao = IIf(txtEspecificacoes.Text = "", Null, txtEspecificacoes.Text)
TBEstoque!LOTE = txtLote.Text
TBEstoque!Data = Txt_data_recebimento
TBEstoque!Responsavel = pubUsuario
TBEstoque!Certificado = IIf(txtCertificado = "", 0, txtCertificado)
TBEstoque!Corrida = IIf(txtcorrida = "", 0, txtcorrida)
TBEstoque!local_armaz = cmbLocal_armaz
TBEstoque!Fornecedor = txtFornecedor
TBEstoque!Un = txtUN.Text
TBEstoque!Numero_serie = Txt_numero_serie
TBEstoque!NF = txtnotafiscal.Text

If chkPerecivel.Value = True And txtvencimento.Text <> "__/__/____" Then
    TBEstoque!Vencimento = txtvencimento.Text
End If

TBEstoque!ID_empresa = txtID_empresa
TBEstoque!Consignacao = "False"
TBEstoque!Bloqueado = "False"
'=========================================================
' Novo - Caminho para o certifcado
'=========================================================
TBEstoque!CaminhoCertificado = IIf(Txt_caminho2.Text <> "", Txt_caminho2.Text, "0")

'===================================================
'Verifica se o produto controla estoque
'===================================================
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtCodigo & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
Permitido = True
Else
Permitido = False
End If
TBProduto.Close

'===================================================================================================================================================
'Verifica se o produto do pedido é remessa para industrialização
'===================================================================================================================================================
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select CPL.IDlista from Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho where CPL.Idlista = " & TXTIDLista & " and CPL.Remessa = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
Permitido = False
Else
Permitido = True
End If
TBProduto.Close

'===================================================================================================================================================
'Verifica se o produto do pedido é mão de obra e se é a ultima fase da ordem e a ordem não controla estoque automaticamente, se for entra no estoque
'====================================================================================================================================================
Set TBProduto = CreateObject("adodb.recordset")
StrSql = "Select CPL.OS, CPL.Ordem from (Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho) LEFT JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = CPL.ID_CFOP where CPL.Idlista = " & TXTIDLista & " and CFOP.MaoObra = 'True' and P.Subtipoitem <> 4"
'Debug.print StrSql

TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
If TBProduto.EOF = False Then
    Permitido = False
    If IsNull(TBProduto!OS) = False And TBProduto!OS <> "" And IsNull(TBProduto!Ordem) = False And TBProduto!Ordem <> "" Then
        Set TBOrdemServico = CreateObject("adodb.recordset")
        TBOrdemServico.Open "Select OS.idproducao from OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.Ordem = " & TBProduto!Ordem & " and P.Entrar_estoque <> 'True' ORDER BY OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdemServico.EOF = False Then
            TBOrdemServico.MoveLast
            If TBOrdemServico!IDProducao = TBProduto!OS Then Permitido = True
        End If
        TBOrdemServico.Close
    End If
End If
TBProduto.Close

'Se for serviço, mão de obra ou remessa cria um local de armz padrão

If cmbLocal_armaz = "SERVIÇOS" Or cmbLocal_armaz = "RETORNO DE MERCADORIA" Or cmbLocal_armaz = "INDUSTRIALIZAÇÃO" Then
Permitido = False
Else
Permitido = True
End If

Qtd = IIf(txtQuantidade.Text = "", 0, txtQuantidade.Text)


'============================================================
' Se for pra entrar no estoque grava as quantidades na RE
'============================================================
If Permitido = True Then
TBEstoque!estoque_venda = txtQuantidade.Text
TBEstoque!estoque_real = txtQuantidade.Text
TBEstoque!estoque_real_PC = Valor_Cofins_Prod
TBEstoque!Qtde = Qtd
Else
TBEstoque!estoque_venda = 0
TBEstoque!estoque_real = 0
TBEstoque!estoque_real_PC = 0
TBEstoque!Qtde = 0
End If
'============================================================

TBEstoque!status = "ENTRADA_NOTA_FISCAL"

IDFase = 0
IDPlano = 0

'============================================================================================
'Grava familia do produto na tabela estoque_controle e Atualiza valor do material no estoque
'============================================================================================
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select classe, Codproduto, ID_PC from projproduto where desenho = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBEstoque!Classe = TBAbrir!Classe
    IDFase = TBAbrir!Codproduto
    IDPlano = IIf(IsNull(TBAbrir!ID_PC), 0, TBAbrir!ID_PC)
    Familiatext = TBAbrir!Classe
'========================================================
'Grava código de referência no produto
'========================================================
    If Cmb_codigo_ref <> "" Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from item_aplicacoes where Codproduto = " & IDFase & " and n_referencia = '" & Cmb_codigo_ref & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then TBProduto.AddNew
        TBProduto!Codproduto = IDFase
        TBProduto!N_referencia = Cmb_codigo_ref
        TBProduto!Descricao = IIf(txtEspecificacoes.Text = "", Null, txtEspecificacoes.Text)
        TBProduto!ID_cliente_forn = Txt_ID_forn
        TBProduto!Tipo = "F"
        TBProduto.Update
        TBProduto.Close
    End If
End If
TBAbrir.Close

TBEstoque.Update
'===========================================================
'Cria o empenho no RE para o pedido interno se o produto controlar estoque
'===========================================================
If Permitido = True Then
    qt = txtQuantidade
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IDcarteira, Qtde_empenho - ISNULL(Qtde_recebida, 0) AS Qtde_empenhada from Compras_pedido_lista_empenhos where IDlista = " & TXTIDLista & " and Qtde_empenho - ISNULL(Qtde_recebida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False And qt > 0
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Estoque_Controle_Empenho_Vendas", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Data = Txt_data_recebimento
            TBGravar!Responsavel = pubUsuario
            If qt >= TBAbrir!Qtde_empenhada Then
                TBGravar!Qtde_empenhada = TBAbrir!Qtde_empenhada
                qt = qt - TBAbrir!Qtde_empenhada
            Else
                TBGravar!Qtde_empenhada = qt
                qt = 0
            End If
            TBGravar.Update
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End If
frmEstoque_Recebimento.Tag = TBEstoque!IDEstoque
IDEstoque = TBEstoque!IDEstoque
TBEstoque.Close

ValorTotal = 0

'=============================================================================
' Grava valor unitário do item no estoque controle com ou sem ICMS
'=============================================================================
Set TBPedido = CreateObject("adodb.recordset")
StrSql = "Select CPL.Remessa, CPL.preco_unitario_desconto, CPL.preco_total, CPL.vlrICMS, CPL.Quant_Comp, CPL.UN, CPL.Unidade_com, ISNULL(CPL.Qtde_estoque, 0) AS Qtde_estoque from compras_pedido_lista CPL inner join compras_pedido CP on CPL.idpedido = CP.idpedido where CP.pedido = '" & txtProg_pedido & "' and CPL.idlista = " & TXTIDLista
'Debug.print StrSql

TBPedido.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then

Dim QTEstoque As Double
Dim qtComprado As Double
Dim ValorICMSUNitario As Double
Dim valorComICMS As Double
Dim valorSemICMS As Double
Dim valortotalComICMS As Double
Dim valorTotalSemICMS As Double


Remessa = TBPedido!Remessa
    QTEstoque = 1
    ValorICMSUNitario = Format(IIf(IsNull(TBPedido!vlrICMS), 0, TBPedido!vlrICMS) / IIf(IsNull(TBPedido!Quant_Comp), 0, TBPedido!Quant_Comp), "0.000000000")
    valorComICMS = Format(IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto))
    valorSemICMS = valorComICMS - ValorICMSUNitario
End If

qtComprado = txtQuantidade

valorTotalSemICMS = qtComprado * valorSemICMS
NovoValorSemICMS = Replace(valorSemICMS, ",", ".")
NovoValorTotalSemICMS = Replace(valorTotalSemICMS, ",", ".")


If chkICMS.Value = 0 Then ' Valor sem ICMS
    ValorTotal = qtComprado * valorSemICMS
    NovoValorUnitario = Replace(valorSemICMS, ",", ".")
    NovoValorTotal = Replace(valorTotalSemICMS, ",", ".")
Else 'Valor com ICMS
    ValorTotal = qtComprado * valorComICMS
    NovoValorUnitario = Replace(valorComICMS, ",", ".")
    NovoValorTotal = Replace(ValorTotal, ",", ".")
End If

Conexao.Execute "Update estoque_controle Set valor_unitario = " & NovoValorUnitario & " where IDestoque = " & IDEstoque
Conexao.Execute "Update estoque_controle set Valor_total = " & NovoValorTotal & " where IDestoque = " & IDEstoque



quantestoque = 0

'===========================================================
'Grava movimentação na tabela estoque_controle_recebimento
'===========================================================
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle_recebimento", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!Data_recebimento = Txt_data_recebimento
TBEstoque!IDpedido = Txt_ID_pedido
TBEstoque!IDlista = TXTIDLista.Text
TBEstoque!Desenho = txtCodigo.Text
TBEstoque!Certificado = IIf(txtCertificado = "", 0, txtCertificado)
TBEstoque!Corrida = IIf(txtcorrida = "", 0, txtcorrida)
TBEstoque!local_armaz = cmbLocal_armaz
TBEstoque!Nota_fiscal = txtnotafiscal.Text
TBEstoque!Serie = txtSerie
If txtnotafiscal <> "" And txtDataemissao <> "__/__/____" Then TBEstoque!Data_emissao = txtDataemissao Else TBEstoque!Data_emissao = Null
TBEstoque!Responsavel = pubUsuario
TBEstoque!Recebido = Format(txtQuantidade.Text, "###.##0.000")
TBEstoque!Recebido_PC = Format(Valor_Cofins_Prod, "###.##0.000")

 
If Valor_Cofins_Prod = 0 Then
    quantestoque = Format(txtrequisitado.Text, "###.##0.000")
    quantnovo = Format(txtQuantidade.Text, "###.##0.000")
Else
    quantestoque = Format(IIf(txtrequisitado_PC = "", txtrequisitado, txtrequisitado_PC), "###.##0.000")
    quantnovo = Format(Valor_Cofins_Prod, "###.##0.000")
End If
If quantnovo < quantestoque Then TBEstoque!Parcial = True Else TBEstoque!Parcial = False
TBEstoque!Programacao = False
TBEstoque!ID_empresa = txtID_empresa
TBEstoque!Obs = txtObs
TBEstoque.Update
IDEstoque_recebimento = TBEstoque!ID
TBEstoque.Close

ProcAtualizaQtdeRecebEmp Txt_ID_pedido, TXTIDLista

quantestoque = 0
quantnovo = 0

'===================================================
' Abaixo tem que verificar
'===================================================
'Grava movimentação na tabela estoque_movimentacao
'===================================================
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBEstoque.AddNew
TBEstoque!Destino = "Interno"
TBEstoque!Terceiros = False
TBEstoque!Operacao = "ENTRADA_NOTA_FISCAL"
TBEstoque!IDEstoque = IDEstoque
TBEstoque!Documento = txtnotafiscal.Text
TBEstoque!DtEmissao = Txt_data_recebimento
TBEstoque!LOTE = txtLote.Text
TBEstoque!Responsavel = pubUsuario
TBEstoque!Data = Txt_data_recebimento
TBEstoque!Descricao = txtEspecificacoes.Text
TBEstoque!Desenho = txtCodigo.Text
TBEstoque!Bloqueado = "False"
TBEstoque!ID_empresa = txtID_empresa
'===================================================
' Se for item de remessa grava quantidade de entrada igual a 0
'===================================================
If Permitido = False Then
    TBEstoque!estoque_venda = Format("0", "###.##0.000")
    TBEstoque!Entrada = Format("0", "###.##0.000")
    TBEstoque!Entrada_PC = Format("0", "###.##0.000")
Else
    TBEstoque!estoque_venda = Format(txtQuantidade.Text, "###.##0.000")
    TBEstoque!Entrada = Format(txtQuantidade.Text, "###.##0.000")
    TBEstoque!Entrada_PC = Format(Valor_Cofins_Prod, "###.##0.000")

End If
TBEstoque!QT_inspecionar = Format(txtQuantidade.Text, "###.##0.000")
'===================================================
TBEstoque!Familia = Familiatext
TBEstoque!Obs = txtObs
quantestoque = txtQuantidade

'Atualiza valor do material no estoque

TBEstoque!VlrUnit = Format(valorSemICMS, "###.##0.00000")
TBEstoque!vlrTotal = Format(valorTotalSemICMS, "###.##0.00")

TBEstoque!IDEstoque_recebimento = IDEstoque_recebimento
TBEstoque!idlista_recebimento = TXTIDLista.Text
TBEstoque!Destino = "Interno"
TBEstoque!Terceiros = False

TBEstoque!Saida = Format(0, "###.##0.000")

Set TBNivel1 = CreateObject("adodb.recordset")
TBNivel1.Open "Select * from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtCodigo & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
If TBNivel1.EOF = False Then
    Set TBNivel2 = CreateObject("adodb.recordset")
    TBNivel2.Open "Select sum(Saida) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtCodigo & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel2.EOF = False Then
        Valor1 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
    End If
    TBNivel2.Close
    Set TBNivel2 = CreateObject("adodb.recordset")
    TBNivel2.Open "Select sum(entrada) as quantidade from estoque_movimentacao where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtCodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBNivel2.EOF = False Then
        Valor2 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
    End If
    TBNivel2.Close
    Valor2 = Valor2 + Qtd
    If Valor1 <= Valor2 Then
        Conexao.Execute "UPDATE estoque_movimentacao set Terceiros = 'False' where pedidocompra = '" & txtProg_pedido & "' and desenho = '" & txtCodigo & "' and destino = 'Terceiros'"
    End If
    TBEstoque!Pedidocompra = txtProg_pedido
    TBEstoque!IDpedido = Txt_ID_pedido
Else
    TBEstoque!Pedidocompra = Null
    TBEstoque!IDpedido = Null
End If
TBNivel1.Close
TBEstoque.Update

'Atualiza o RE no instrumento
Conexao.Execute "Update I set I.IDestoque = " & TBEstoque!IDEstoque & " from Instrumentos I INNER JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where EC.Desenho = '" & txtCodigo & "' and EC.Numero_serie = '" & Txt_numero_serie & "'"

'Centro de custo
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CPLC.Valor, CPLC.ID_CC, CP.Data from Compras_pedido_lista_custo CPLC INNER JOIN Compras_pedido CP ON CPLC.IDPedido = CP.IDPedido where CPLC.IDLista = " & TXTIDLista.Text & " and CP.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Valor3 = TBAbrir!valor
        
        qt = txtrequisitado
        Qtde = TBEstoque!Entrada
        
        'Calcula quantidade se for com unidade diferente (inverte a un de estoque com a un comercial porque preciso voltar a conversão)
        Set TBPedido = CreateObject("adodb.recordset")
        TBPedido.Open "Select CPL.UN, CPL.Unidade_com, CPL.Qtde_estoque, CPL.Quant_comp from compras_pedido_lista CPL INNER JOIN compras_pedido CP on CPL.idpedido = CP.idpedido where CP.pedido = '" & txtProg_pedido & "' and CPL.idlista = " & TXTIDLista & " and CPL.Qtde_Estoque IS NOT NULL and CPL.Qtde_estoque <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBPedido.EOF = False Then
            If TBPedido!Un <> TBPedido!Unidade_com Then
                qt = TBPedido!Quant_Comp
                If FunVerifUNConversao(TBPedido!Un, TBPedido!Unidade_com) = True Then
                    Qtde = FunConverteUN(TBPedido!Unidade_com, TBPedido!Un, TBEstoque!Entrada, txtCodigo)
                Else
                    Qtde = TBEstoque!Entrada * FunVerificaTabelaConversaoUnidade(TBPedido!Un, TBPedido!Unidade_com)
                End If
            End If
        End If
        valor = Format((Valor3 / qt) * Qtde, "###,##0.00")
        
        'Verifica se tem CC amarrado no produto, se for diferente do informado no pedido cria débito e crédito no CC do produto
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select ID_CC from projproduto where codproduto = " & IDFase & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            If TBProduto!ID_CC <> "" Then
                If TBAbrir!ID_CC <> TBProduto!ID_CC Then
                    ProcSalvarCCRealizado TBAbrir!Data, txtID_empresa, "Débito", TBProduto!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TXTIDLista, valor, True, False
                    
                    'Grava movimentação no centro consolidado
                    ProcSalvarRealCCConsolidado TBProduto!ID_CC, "Débito", False, False, False
                    
                    ProcSalvarCCRealizado TBAbrir!Data, txtID_empresa, "Crédito", TBProduto!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TXTIDLista, valor, True, False
                    
                    'Grava movimentação no centro consolidado
                    ProcSalvarRealCCConsolidado TBProduto!ID_CC, "Crédito", True, True, False
                End If
            End If
        End If
        TBProduto.Close
        
        ProcSalvarCCRealizado TBAbrir!Data, txtID_empresa, "Débito", TBAbrir!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TXTIDLista, valor, False, False
        
        'Grava movimentação no centro consolidado
        ProcSalvarRealCCConsolidado TBAbrir!ID_CC, "Débito", False, False, False
        
        TBAbrir.MoveNext
    Loop
Else
    'Verifica se tem CC amarrado no produto e cria um débito no CC do produto
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select ID_CC from projproduto where codproduto = " & IDFase & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        qt = txtrequisitado
        Qtde = TBEstoque!Entrada
        valor = Format((Valor3 / qt) * Qtde, "###,##0.00")
        
        ProcSalvarCCRealizado Txt_data_recebimento, txtID_empresa, "Débito", TBProduto!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TXTIDLista, valor, False, False
        
        'Grava movimentação no centro consolidado
        ProcSalvarRealCCConsolidado TBProduto!ID_CC, "Débito", False, False, False
    End If
    TBProduto.Close
End If
TBAbrir.Close
TBEstoque.Close

'Atualiza o status do pedido, status e qtde recebida da programação
ProcAtualizaQtdeRecebidaProg "Select CP.qtderecebida, CP.Quantidade from Compras_Programacao CP INNER JOIN Compras_pedido_lista CPL ON CP.ID_prog = CPL.ID_programacao where CPL.IDlista = " & TXTIDLista & " order by CP.data_inicio", False
ProcAlteraStatus_pedido
ProcAlteraStatus_prog

USMsgBox ("Produto recebido ao estoque com sucesso."), vbInformation, "CAPRIND v5.0"
cmdReceber.Enabled = False
ProcBloqueiaFrame
ProcCarregaListaMovimentacao
'==================================
Modulo = "Estoque/Recebimento/Pedido de compra"
Evento = "Receber"
ID_documento = TXTIDLista
Documento = "Cód. interno: " & txtCodigo & " - Nº lote: " & txtProg_pedido & " - Nº corrida: " & IIf(txtcorrida = "", 0, txtcorrida) & " - Nº certificado: " & IIf(txtCertificado = "", 0, txtCertificado) & " - Local armaz.: " & cmbLocal_armaz
Documento1 = "Operação: " & Lista_Movimentacao.SelectedItem.SubItems(2) & " - Documento: " & Lista_Movimentacao.SelectedItem.SubItems(6)
ProcGravaEvento
'==================================
If txtnotafiscal <> "" Then ProcAtualizaVlrEntradaEstoque True
btnReceber_Click

'===========================================================================================
'Gravar Pedido de compras na nota
'===========================================================================================

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_proposta_nota where NF = '" & txtnotafiscal & "' and proposta = '" & txtProg_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
End If
    TBGravar!ID_nota = 0
    TBGravar!Proposta = txtProg_pedido
    TBGravar!NF = txtnotafiscal
    TBGravar!Revisao = 0 'frmEstoque_Recebimento.txtrev.Text
TBGravar.Update
TBGravar.Close

'ProcCarregaListaFiltro
ProcLimparCamposReq True
Lista_Movimentacao.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaQtdeRecebEmp(IDpedido As Long, IDlista As Long)
On Error GoTo tratar_erro

valor = 0
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select ROUND(Sum(ISNULL(Recebido, 0)), 4) as Valor from estoque_controle_recebimento where idpedido = " & IDpedido & " and idlista = " & IDlista & " and Programacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    valor = TBCFOP!valor
End If
TBCFOP.Close
NovoValor = Replace(valor, ",", ".")
Conexao.Execute "Update Compras_pedido_lista_empenhos Set Qtde_recebida = " & NovoValor & " where IDlista = " & IDlista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaQtdeRecebidaProg(TextoFiltro As String, Excluir As Boolean)
On Error GoTo tratar_erro

If Excluir = False Then
    QuantEmpenho = qt
    quantidade = 0
    ValorTotal = 0
    Set TBProgramas = CreateObject("adodb.recordset")
    TBProgramas.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBProgramas.EOF = False Then
        Do While TBProgramas.EOF = False
            If TBProgramas!Qtderecebida <> TBProgramas!quantidade Then
                If QuantEmpenho > TBProgramas!quantidade - TBProgramas!Qtderecebida Then  'Se qtde recebida for maior q a qtde programada
                    quantidade = TBProgramas!quantidade - TBProgramas!Qtderecebida
                    TBProgramas!Qtderecebida = TBProgramas!quantidade
                    QuantEmpenho = QuantEmpenho - quantidade
                Else
                    If TBProgramas!Qtderecebida = 0 Then 'Se qtde recebida for = 0
                        TBProgramas!Qtderecebida = QuantEmpenho
                        QuantEmpenho = QuantEmpenho - QuantEmpenho
                    Else
                        'Se qtde recebida for menor q a qtde programada
                        TBProgramas!Qtderecebida = QuantEmpenho + TBProgramas!Qtderecebida
                        QuantEmpenho = QuantEmpenho - QuantEmpenho
                    End If
                End If
            End If
            TBProgramas.Update
            If QuantEmpenho <= 0 Then GoTo Sair
            TBProgramas.MoveNext
        Loop
    End If
Sair:
        TBProgramas.Close
    
Else
    QuantEmpenho = qt
    quantidade = 0
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        Do While TBCompras_Lista.EOF = False
            If QuantEmpenho <= 0 Then GoTo Sair1
            quantidade = TBCompras_Lista!Qtderecebida
            If QuantEmpenho >= TBCompras_Lista!Qtderecebida Then
                TBCompras_Lista!Qtderecebida = TBCompras_Lista!Qtderecebida - TBCompras_Lista!Qtderecebida
            Else
                TBCompras_Lista!Qtderecebida = TBCompras_Lista!Qtderecebida - QuantEmpenho
            End If
            If TBCompras_Lista!Qtderecebida = 0 Then
                 If TBCompras_Lista!Firme = True Then
                    TBCompras_Lista!Status_prog = "ABERTO"
                    TBCompras_Lista!Ordenar = 2
                Else
                    TBCompras_Lista!Status_prog = "PREVISÃO FUTURA"
                    TBCompras_Lista!Ordenar = 3
                End If
            End If
            If TBCompras_Lista!Qtderecebida <> 0 And TBCompras_Lista!Qtderecebida < TBCompras_Lista!quantidade Then
                TBCompras_Lista!Status_prog = "PARCIAL"
                TBCompras_Lista!Ordenar = 1
            End If
            If TBCompras_Lista!Qtderecebida >= TBCompras_Lista!quantidade Then
                TBCompras_Lista!Status_prog = "RECEBIDO"
                TBCompras_Lista!Ordenar = 4
            End If
            TBCompras_Lista.Update
            QuantEmpenho = QuantEmpenho - quantidade
Sair1:
            TBCompras_Lista.MoveNext
        Loop
    End If
    TBCompras_Lista.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: cmdReceber_Click
    Case vbKeyF4: cmdCancelar_Click
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyF7: ProcStatus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 8, True
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Formulario = "Estoque/Recebimento/Pedido de compra"
Direitos

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaPedido()
On Error GoTo tratar_erro

If txtProg_pedido.Text = "" Then Exit Sub
ValorTotal = 0
Qtd = 0
Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select ID_empresa, Data, IDpedido, idfornecedor, Fornecedor, Estado,N_referencia from compras_pedido where pedido = '" & txtProg_pedido & "' and (Status_pedido = 'ABERTO' or Status_pedido = 'APROVADO' or Status_pedido = 'PARCIAL' or Status_pedido = 'ENCERRADO')", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    Txt_ID_pedido = TBCompras_Pedido!IDpedido
    Txt_ID_forn = TBCompras_Pedido!IDFornecedor
    txtFornecedor.Text = IIf(IsNull(TBCompras_Pedido!Fornecedor) = False, TBCompras_Pedido!Fornecedor, "")
    txtuf = IIf(IsNull(TBCompras_Pedido!Estado), "", TBCompras_Pedido!Estado)
    
    Set TBExecucao = CreateObject("adodb.recordset")
    TBExecucao.Open "Select CODIGO, Empresa from Empresa where codigo = " & TBCompras_Pedido!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBExecucao.EOF = False Then
        txtID_empresa.Text = IIf(IsNull(TBExecucao!CODIGO), "", TBExecucao!CODIGO)
        txtEmpresa.Text = IIf(IsNull(TBExecucao!Empresa), "", TBExecucao!Empresa)
    End If
    TBExecucao.Close
    
    txtData.Text = Format(TBCompras_Pedido!Data, "dd/mm/yy")
    txtreferencia = IIf(IsNull(TBCompras_Pedido!N_referencia), "", TBCompras_Pedido!N_referencia)
End If
TBCompras_Pedido.Close
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from compras_comercial where idpedido = " & Txt_ID_pedido, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtcondpagamento.Text = IIf(IsNull(TBFornecedor!condicoes) = False, TBFornecedor!condicoes, "")
    txtMoeda.Text = IIf(IsNull(TBFornecedor!Moeda) = False, TBFornecedor!Moeda, "REAL")
    txtvlrMoeda = IIf(IsNull(TBFornecedor!Valor_moeda) = False, Format(TBFornecedor!Valor_moeda, "###,##0.00"), "1,00")
End If
TBFornecedor.Close
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Sum(preco_total) as Valortotal from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ValorTotal = IIf(IsNull(TBLISTA!ValorTotal), 0, TBLISTA!ValorTotal)
End If
TBLISTA.Close
txtValorTotal.Text = Format(ValorTotal, "###,##0.00")

If txtMoeda.Text <> "" Then
    Set TBMoeda = CreateObject("adodb.recordset")
    TBMoeda.Open "Select * from Moeda where Moeda = '" & txtMoeda.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMoeda.EOF = False Then
            txtValorTotal.Text = TBMoeda!Simbolo & txtValorTotal.Text
        End If
    TBMoeda.Close
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmEstoque_Recebimento_Imprimir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro
  
TXTIDLista = ""
txtcondpagamento = ""
txtData = ""
txtValorTotal = ""
Txt_ID_forn = ""
txtFornecedor = ""
txtuf = ""
Txt_ID_pedido = 0
txtMoeda.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparCamposReq(Receber As Boolean)
On Error GoTo tratar_erro

If chkdtVencimento = 0 Then txtvencimento = "__/__/____"
    
txtCodigo.Text = ""
Cmb_codigo_ref.Clear
txtUN.Text = ""
txtStatus.Text = ""
txtEspecificacoes = ""
txtCertificado = ""
txtcorrida = ""
If Receber = True Then
    If Chk_LA.Value = 0 Then cmbLocal_armaz.ListIndex = -1
Else
    cmbLocal_armaz.ListIndex = -1
End If
txtQuantidade = ""
txtQuantidade_PC = ""
Txt_numero_serie = ""
If Receber = True Then
    If Chk_Dt_rcbto.Value = 0 Then Txt_data_recebimento = "__/__/____"
Else
    Txt_data_recebimento = "__/__/____"
End If
txtrequisitado = "0,0000"
txtrequisitado_PC = "0,0000"
txtrecebida = "0,0000"
txtrecebida_PC = "0,0000"
txtSaldo = "0,0000"
txtObs = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_movimentacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ECR.*, EM.Data, EC.Numero_serie, EC.Caminhocertificado from (Estoque_controle_recebimento ECR INNER JOIN Estoque_movimentacao EM ON ECR.Id = EM.IDEstoque_recebimento) INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.Idoperacao = " & Lista_Movimentacao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Nota_fiscal) = False And TBAbrir!Nota_fiscal <> "" Then
        txtnotafiscal.Text = TBAbrir!Nota_fiscal
        txtSerie = IIf(IsNull(TBAbrir!Serie), "", TBAbrir!Serie)
        txtDataemissao = Format(TBAbrir!Data_emissao, "dd/mm/yyyy")
    End If
    txtCertificado.Text = IIf(IsNull(TBAbrir!Certificado), "", TBAbrir!Certificado)
    Txt_data_recebimento = Format(TBAbrir!Data, "dd/mm/yyyy")
    txtQuantidade.Text = Format(Lista_Movimentacao.SelectedItem.ListSubItems(4), "###,##0.0000")
    txtQuantidade_PC.Text = Lista_Movimentacao.SelectedItem.ListSubItems(5)
    txtcorrida.Text = IIf(IsNull(TBAbrir!Corrida), "", TBAbrir!Corrida)
    Txt_numero_serie = IIf(IsNull(TBAbrir!Numero_serie), "", TBAbrir!Numero_serie)
    
        If txtMoeda.Text = "DOLAR" And IsDate(txtDataemissao) And IsNumeric(txtID_empresa.Text) And IsNumeric(Txt_ID_forn.Text) And txtSerie.Text <> "" Then
                TipoNF = "M1"
            
            '==============================================================================
            ' Verifica se existe cadastro da nota fiscal
            '==============================================================================
                        Set TBGravar = CreateObject("adodb.recordset")
                        StrSql = "Select * from tbl_Dados_Nota_Fiscal where dt_DataEmissao = '" & txtDataemissao.Text & "'  and ID_empresa = " & txtID_empresa.Text & " and Id_Int_Cliente = " & Txt_ID_forn.Text & " and int_NotaFiscal = '" & txtnotafiscal.Text & "' and Serie = '" & txtSerie & "' and int_TipoNota = 2 and TipoNF = '" & TipoNF & "'"
                        'Debug.print StrSql
                        
                        
                        TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                            If TBGravar.EOF = False Then
                            
                            txtvlrMoeda.Text = TBGravar!ValorMoeda
                                lblStatusNF.Caption = "EMITIDA"
                            
                            TBGravar.Close
                            Else
                                lblStatusNF.Caption = "Á EMITIR"
                            End If
        End If
    
'==================================================================================================
' Dados do caminho do certificado
'==================================================================================================
    Txt_caminho2.Text = IIf(IsNull(TBAbrir!CaminhoCertificado), "", TBAbrir!CaminhoCertificado)
'==================================================================================================
    If TBAbrir!local_armaz <> "" Then cmbLocal_armaz.Text = TBAbrir!local_armaz
1:
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o local de armazenamento (" & TBAbrir!local_armaz & ") desta movimentação, favor revisar."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNotaFiscal_Change()
On Error GoTo tratar_erro
    
If txtnotafiscal.Text <> "" Then
    VerifNumero = txtnotafiscal.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtnotafiscal.Text = ""
        txtnotafiscal.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

If txtQuantidade.Text <> "" Then
    VerifNumero = txtQuantidade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade.Text = ""
        txtQuantidade.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantidade_LostFocus()
On Error GoTo tratar_erro

txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAlteraStatus_pedido()
On Error GoTo tratar_erro
quantnovo = 0
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from compras_pedido where pedido = '" & txtProg_pedido.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    Txt_ID_pedido = TBCompras!IDpedido
    
    quantnovo = 0
    Valor_Cofins_Prod = 0
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select Sum(Recebido) as quantnovo, Sum(ISNULL(Recebido_PC, 0)) as Valor_Cofins_Prod from estoque_controle_recebimento where idpedido = " & Txt_ID_pedido & " and idlista = " & TXTIDLista.Text & " and Programacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        quantnovo = IIf(IsNull(TBEstoque!quantnovo), 0, TBEstoque!quantnovo)
        Valor_Cofins_Prod = IIf(IsNull(TBEstoque!Valor_Cofins_Prod), 0, TBEstoque!Valor_Cofins_Prod)
    End If
    TBEstoque.Close
    
    If Valor_Cofins_Prod > 0 Then
        quantnovo = Valor_Cofins_Prod
        quantestoque = IIf(txtrequisitado_PC = "", txtrequisitado, txtrequisitado_PC)
    Else
        quantestoque = txtrequisitado.Text
    End If
    
'=======================================================
'Acerta Status dos itns recebidos de forma parcial
'=======================================================
    If quantnovo < quantestoque Then
        If USMsgBox("Este produto/serviço será recebido parcialmente, deseja encerrar o mesmo no pedido de compra?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Status_Item = "RECEBIDO"
        Else
            Status_Item = "PARCIAL"
        End If
    End If
'======================================================
'Acerta Status dos itens do pedido
'======================================================
    If quantnovo >= quantestoque Then Status_Item = "RECEBIDO"
    Conexao.Execute "Update compras_pedido_lista Set Status_item = '" & Status_Item & "' where idpedido = " & Txt_ID_pedido & " and idlista = " & TXTIDLista.Text
    Conexao.Execute "Update Compras_Programacao set Compras_Programacao.Status_prog = '" & Status_Item & "' from Compras_Programacao INNER JOIN compras_pedido_lista ON Compras_Programacao.ID_prog = compras_pedido_lista.ID_programacao where compras_pedido_lista.idpedido = " & Txt_ID_pedido & " and compras_pedido_lista.idlista = " & TXTIDLista.Text
'=====================================================
'Verifica Status dos itens no pedido
'=====================================================
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & Txt_ID_pedido & " and status_item <> 'RECEBIDO' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = True Then
        Status_pedido = "ENCERRADO"
    Else
        Status_pedido = "PARCIAL"
    End If
    TBCompras.Close
'======================================================
'Acerta Status do Pedido de compra
'======================================================
    Conexao.Execute "Update compras_pedido Set Status_pedido = '" & Status_pedido & "' where idpedido = '" & Txt_ID_pedido.Text & "'"
End If
'======================================================
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAlteraStatus_prog()
On Error GoTo tratar_erro

'Produto
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select compras_programa_item.* from (Compras_Programacao INNER JOIN Compras_pedido_lista ON Compras_Programacao.ID_prog = Compras_pedido_lista.ID_programacao) INNER JOIN compras_programa_item ON compras_programa_item.ID_item = Compras_Programacao.Id_item where Compras_pedido_lista.IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!Status_Item = "PREVISÃO FUTURA"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from compras_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!Status_Item = "N_RECEBIDO"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from compras_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!Status_Item = "RECEBIDO"
            Else
                TBItem!Status_Item = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    Txt_ID_pedido = TBItem!ID
    TBItem.Update
End If

'Programa
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from compras_programa where id = " & Txt_ID_pedido, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_programa_item where id = " & Txt_ID_pedido & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!status = "PREVISÃO FUTURA"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from compras_programa_item where id = " & Txt_ID_pedido & " and Status_Item <> 'N_RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!status = "ABERTO"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from compras_programa_item where id = " & Txt_ID_pedido & " and Status_Item <> 'RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!status = "RECEBIDO"
            Else
                TBItem!status = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    TBItem.Update
End If
TBItem.Close
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

'Debug.print StrSql_Estoque_Recebimento_Localizar

If StrSql_Estoque_Recebimento_Localizar = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Listprod.ListItems.Clear
Set TBLISTA_Estoque_RecebimentoPedido = CreateObject("adodb.recordset")
TBLISTA_Estoque_RecebimentoPedido.Open StrSql_Estoque_Recebimento_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Estoque_RecebimentoPedido.EOF = False Then ProcExibePagina (1)
ProcCarregaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Listprod.ListItems.Clear
TBLISTA_Estoque_RecebimentoPedido.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_RecebimentoPedido.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_RecebimentoPedido.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_RecebimentoPedido.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_RecebimentoPedido.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Estoque_RecebimentoPedido.EOF = False And (ContadorReg <= TamanhoPagina)
    With Listprod.ListItems
        .Add , , TBLISTA_Estoque_RecebimentoPedido!IDlista
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Empresa from Empresa where codigo = " & TBLISTA_Estoque_RecebimentoPedido!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!Empresa), "", TBFI!Empresa)
        TBFI.Close
        .Item(.Count).SubItems(2) = TBLISTA_Estoque_RecebimentoPedido!Pedido
        .Item(.Count).SubItems(3) = TBLISTA_Estoque_RecebimentoPedido!Desenho
        .Item(.Count).SubItems(4) = TBLISTA_Estoque_RecebimentoPedido!Descricao
        .Item(.Count).SubItems(5) = TBLISTA_Estoque_RecebimentoPedido!Un
        
        If TBLISTA_Estoque_RecebimentoPedido!Un <> TBLISTA_Estoque_RecebimentoPedido!Unidade_com And IsNull(TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque) = False And TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque <> 0 Then
            valor = TBLISTA_Estoque_RecebimentoPedido!Quant_Comp / TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque
             
            .Item(.Count).SubItems(6) = Format(valor * TBLISTA_Estoque_RecebimentoPedido!preco_unitario, "###,##0.0000")
            .Item(.Count).SubItems(7) = Format(TBLISTA_Estoque_RecebimentoPedido!Qtde_estoque, "###,##0.0000")
        Else
            .Item(.Count).SubItems(6) = Format(TBLISTA_Estoque_RecebimentoPedido!preco_unitario, "###,##0.0000")
            .Item(.Count).SubItems(7) = Format(TBLISTA_Estoque_RecebimentoPedido!Quant_Comp, "###,##0.0000")
        End If
        
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Estoque_RecebimentoPedido!Quant_Comp_PC), "", TBLISTA_Estoque_RecebimentoPedido!Quant_Comp_PC)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Estoque_RecebimentoPedido!Prazo), "", Format(TBLISTA_Estoque_RecebimentoPedido!Prazo, "dd/mm/yy"))
        
        If TBLISTA_Estoque_RecebimentoPedido!Status_Item = "N_RECEBIDO" Or TBLISTA_Estoque_RecebimentoPedido!Status_Item = "APROVADO" Then
            StatusItem = "NÃO_RECEBIDO"
        Else
            StatusItem = TBLISTA_Estoque_RecebimentoPedido!Status_Item
        End If
        .Item(.Count).SubItems(10) = StatusItem
        
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Estoque_RecebimentoPedido!Ordem), "", TBLISTA_Estoque_RecebimentoPedido!Ordem)
        .Item(.Count).SubItems(12) = TBLISTA_Estoque_RecebimentoPedido!IDlista
    End With
    TBLISTA_Estoque_RecebimentoPedido.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_RecebimentoPedido.RecordCount
If TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_RecebimentoPedido.PageCount
ElseIf TBLISTA_Estoque_RecebimentoPedido.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_RecebimentoPedido.PageCount & " de: " & TBLISTA_Estoque_RecebimentoPedido.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_RecebimentoPedido.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_RecebimentoPedido.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

TotContas = 0
IDlista = 0
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open StrSql_Estoque_Recebimento_LocalizarTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Do While TBContas.EOF = False
        Permitido = False
        If IDlista <> TBContas!IDlista Then Permitido = True
        If Permitido = True Then TotContas = TotContas + IIf(IsNull(TBContas!TotContas), 0, TBContas!TotContas)
        IDlista = TBContas!IDlista
        TBContas.MoveNext
    Loop
End If
TBContas.Close
txtQtde_total.Text = Format(TotContas, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proccarregalocarm()
On Error GoTo tratar_erro

cmbLocal_armaz.Clear
TextoLocal = ""
Set TBAliquota = CreateObject("adodb.recordset")
StrSql = "Select ELC.Descricao from Estoque_Localarmazenamento_criar ELC INNER JOIN Estoque_Localarmazenamento EL ON ELC.ID = EL.idemb_locarm where EL.codinterno = '" & txtCodigo & "' and ELC.Descricao is not null and EL.padrao = '1' "

'Debug.print StrSql


TBAliquota.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    Do While TBAliquota.EOF = False
        If IsNull(TBAliquota!Descricao) = False Then cmbLocal_armaz.AddItem TBAliquota!Descricao
        TBAliquota.MoveNext
    Loop
Else
    ProcCarregaComboLA cmbLocal_armaz, False, False
    
    'Se for serviço, mão de obra ou remessa cria um local de armz padrão
    If TBProduto!Tipo = "S" Then
        TextoLocal = "SERVIÇOS"
    ElseIf TBProduto!Remessa = True Then
        TextoLocal = "RETORNO DE MERCADORIA"
    ElseIf IsNull(TBProduto!ID_CFOP) = False And TBProduto!ID_CFOP <> "" And TBProduto!ID_CFOP <> "0" Then
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select IDCountCfop from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP & " and MaoObra = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            TextoLocal = "INDUSTRIALIZAÇÃO"
        End If
        TBCFOP.Close
    End If
End If
TBAliquota.Close
    
With cmbLocal_armaz
    If TextoLocal <> "" Then
        .Text = TextoLocal
        .Locked = True
        .TabStop = False
    Else
        .Locked = False
        .TabStop = True
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_PC_Change()
On Error GoTo tratar_erro

If txtQuantidade_PC <> "" Then
    VerifNumero = txtQuantidade_PC
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade_PC = ""
        txtQuantidade_PC.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtReferencia_Change()
On Error GoTo tratar_erro

If txtreferencia = "REQ" Then
    cmdNota.Enabled = False
    cmdDup.Visible = True
Else
    cmdNota.Enabled = True
    cmdDup.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtstatus_Change()
On Error GoTo tratar_erro

If txtStatus = "RECEBIDO" Then cmdReceber.Enabled = False Else cmdReceber.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaFrame()
On Error GoTo tratar_erro

Frame2.Enabled = False
Frame6.Enabled = False
Frame11.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaFrame()
On Error GoTo tratar_erro

Frame2.Enabled = True
Frame6.Enabled = True
Frame11.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarRealCCConsolidado(ID_CC As Long, Operacao As String, Credito As Boolean, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & ID_CC, Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
    Do While TBAfericao.EOF = False
        ProcSalvarCCRealizado Txt_data_recebimento, txtID_empresa, Operacao, TBAfericao!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TXTIDLista, valor, CC_produto, Bloqueado
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                ProcSalvarCCRealizado Txt_data_recebimento, txtID_empresa, Operacao, TBCiclo!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TXTIDLista, valor, CC_produto, Bloqueado
                TBCiclo.MoveNext
            Loop
        End If
        TBCiclo.Close
        
        TBAfericao.MoveNext
    Loop
End If
TBAfericao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarCCRealizado(Data1 As Date, ID_empresa As Integer, Operacao As String, ID_CC As Long, Cod_produto As Long, ID_plano_contas As Long, ID_estoque As Long, ID_lista As Long, valor As Double, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

NovoValor = Replace(valor, ",", ".")
ProcINSERTINTO "CC_realizado", "Data, Responsavel, ID_empresa, Operacao, ID_CC, Cod_produto, ID_PC, ID_estoque, ID_lista, Valor, Bloqueado", "'" & Data & "', '" & pubUsuario & "', " & ID_empresa & ", '" & Operacao & "', " & ID_CC & ", " & Cod_produto & ", " & ID_plano_contas & ", " & IIf(ID_estoque = 0, "NULL", ID_estoque) & ", " & ID_lista & ", " & NovoValor & ", " & IIf(Bloqueado = True, 1, 0) & ""

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select ID from CC_realizado where ID_estoque = " & ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If CC_produto = True And Operacao = "Crédito" Then Conexao.Execute "Update CC_realizado Set ID_ref_debito = " & TBGravar!ID - 1 & " where ID = " & TBGravar!ID
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrMoeda_LostFocus()
On Error GoTo tratar_erro



If IsNumeric(txtvlrMoeda.Text) Then txtvlrMoeda.Text = Format(txtvlrMoeda, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 3: ProcStatus
    Case 4: procAtualiza
    Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

