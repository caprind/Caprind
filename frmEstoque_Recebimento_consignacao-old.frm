VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoque_Recebimento_consignacao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Recebimento - Material de terceiro"
   ClientHeight    =   10035
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15360
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
   Icon            =   "frmEstoque_Recebimento_consignacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximizado
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1360
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.ComboBox Cmb_empresa1 
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
      ItemData        =   "frmEstoque_Recebimento_consignacao.frx":014A
      Left            =   270
      List            =   "frmEstoque_Recebimento_consignacao.frx":014C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Empresa."
      Top             =   1710
      Visible         =   0   'False
      Width           =   4245
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   390
      Top             =   -1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":014E
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":02C0
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":0432
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":05A4
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":0716
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":09C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":0C7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Cmb_empresa 
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
      ItemData        =   "frmEstoque_Recebimento_consignacao.frx":0D8C
      Left            =   270
      List            =   "frmEstoque_Recebimento_consignacao.frx":0D8E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1710
      Width           =   4215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Notas fiscais recebidas"
      TabPicture(0)   =   "frmEstoque_Recebimento_consignacao.frx":0D90
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "USToolBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "PBLista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Receber nota de terceiro"
      TabPicture(1)   =   "frmEstoque_Recebimento_consignacao.frx":0DAC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "PBLista1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Lista_movimentacao"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtid_estoque"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtId_produto_lista"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdNota"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Txt_tipodest"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   3585
         Left            =   90
         TabIndex        =   39
         Top             =   2190
         Width           =   15195
         Begin VB.ComboBox Cmb_cod_ref 
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
            ItemData        =   "frmEstoque_Recebimento_consignacao.frx":0DC8
            Left            =   4290
            List            =   "frmEstoque_Recebimento_consignacao.frx":0DCA
            Sorted          =   -1  'True
            TabIndex        =   57
            Text            =   "Cmb_cod_ref"
            ToolTipText     =   "Código de referência."
            Top             =   390
            Width           =   2385
         End
         Begin VB.TextBox Txt_qtde_PC 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   7320
            MaxLength       =   50
            TabIndex        =   56
            ToolTipText     =   "Quantidade de peças."
            Top             =   2295
            Width           =   1305
         End
         Begin VB.TextBox txtVlr_total 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   13710
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Valor total."
            Top             =   2295
            Width           =   1305
         End
         Begin VB.TextBox txtVlr_unit 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   4680
            MaxLength       =   50
            TabIndex        =   54
            ToolTipText     =   "Valor unitário."
            Top             =   2295
            Width           =   1305
         End
         Begin VB.ComboBox cmbLocal_armaz 
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
            Left            =   180
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   53
            ToolTipText     =   "Local de armazenamento."
            Top             =   2295
            Width           =   4505
         End
         Begin VB.TextBox txtun 
            Alignment       =   2  'Centralizar
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
            Left            =   6690
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Unidade."
            Top             =   390
            Width           =   735
         End
         Begin VB.TextBox txtpeso 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   8640
            MaxLength       =   50
            TabIndex        =   51
            ToolTipText     =   "Peso unitário."
            Top             =   2295
            Width           =   855
         End
         Begin VB.TextBox txtdesctecnica 
            Alignment       =   2  'Centralizar
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
            Height          =   420
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "Descrição técnica."
            Top             =   1575
            Width           =   14835
         End
         Begin VB.TextBox txtCertificado 
            Alignment       =   2  'Centralizar
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
            Left            =   11760
            MaxLength       =   50
            TabIndex        =   49
            ToolTipText     =   "Certificado."
            Top             =   2295
            Width           =   1935
         End
         Begin VB.TextBox txtfamilia 
            Alignment       =   2  'Centralizar
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
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Fámilia."
            Top             =   390
            Width           =   7575
         End
         Begin VB.TextBox txtcodproduto 
            Alignment       =   2  'Centralizar
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            MaxLength       =   50
            MouseIcon       =   "frmEstoque_Recebimento_consignacao.frx":0DCC
            MousePointer    =   99  'Custom
            TabIndex        =   47
            ToolTipText     =   "Número da nota fiscal."
            Top             =   5100
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtdesc 
            Alignment       =   2  'Centralizar
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
            MaxLength       =   255
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   990
            Width           =   14835
         End
         Begin VB.TextBox txtcorrida 
            Alignment       =   2  'Centralizar
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
            MaxLength       =   50
            TabIndex        =   45
            ToolTipText     =   "Corrida."
            Top             =   2295
            Width           =   1935
         End
         Begin VB.TextBox txtqtde 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   6000
            MaxLength       =   50
            TabIndex        =   44
            ToolTipText     =   "Quantidade."
            Top             =   2295
            Width           =   1305
         End
         Begin VB.TextBox txtobs 
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
            Height          =   555
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            ToolTipText     =   "Observações."
            Top             =   2880
            Width           =   14835
         End
         Begin VB.CommandButton CmdEscolher_item 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3885
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":10D6
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Localizar produtos."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton cmdfiltrar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3570
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":11D8
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Filtrar por código interno."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtdesenho 
            Alignment       =   2  'Centralizar
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
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   40
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1965
         End
         Begin MSComCtl2.DTPicker txtdata 
            Height          =   315
            Left            =   180
            TabIndex        =   58
            ToolTipText     =   "Data do recebimento."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   255
            Format          =   104464385
            CurrentDate     =   39057
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Qtde. PÇ"
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
            Left            =   7620
            TabIndex        =   75
            Top             =   2100
            Width           =   705
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Valor total"
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
            Left            =   13920
            TabIndex        =   74
            Top             =   2100
            Width           =   885
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Valor unitário*"
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
            Left            =   4710
            TabIndex        =   73
            Top             =   2100
            Width           =   1245
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Local de armazenamento*"
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
            Left            =   1307
            TabIndex        =   72
            Top             =   2100
            Width           =   2250
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Data receb."
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
            Left            =   390
            TabIndex        =   71
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Unidade"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6765
            TabIndex        =   70
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Kg"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9540
            TabIndex        =   69
            Top             =   2370
            Width           =   180
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Peso unit."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8700
            TabIndex        =   68
            Top             =   2100
            Width           =   720
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Descrição técnica"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6975
            TabIndex        =   67
            Top             =   1380
            Width           =   1245
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Família"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10987
            TabIndex        =   66
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Descrição"
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
            Left            =   7185
            TabIndex        =   65
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Corrida"
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
            Left            =   10470
            TabIndex        =   64
            Top             =   2100
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Quantidade*"
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
            Left            =   6112
            TabIndex        =   63
            Top             =   2100
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7125
            TabIndex        =   62
            Top             =   2670
            Width           =   945
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Cód. interno*"
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
            Left            =   2010
            TabIndex        =   61
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Cód. de referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4807
            TabIndex        =   60
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Certificado"
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
            Left            =   12270
            TabIndex        =   59
            Top             =   2100
            Width           =   915
         End
      End
      Begin VB.TextBox Txt_tipodest 
         Alignment       =   2  'Centralizar
         BackColor       =   &H80000014&
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
         Left            =   3510
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Tipo do destinatário."
         Top             =   6300
         Visible         =   0   'False
         Width           =   825
      End
      Begin DrawSuite2014.USButton cmdNota 
         Height          =   855
         Left            =   14415
         TabIndex        =   17
         ToolTipText     =   "Emitir nota fiscal."
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         DibPicture      =   "frmEstoque_Recebimento_consignacao.frx":15F3
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Emitir NF"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicAlign        =   8
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Height          =   1485
         Left            =   -74925
         TabIndex        =   21
         Top             =   1320
         Width           =   15195
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   10200
            TabIndex        =   38
            Top             =   210
            WhatsThisHelpID =   210
            Width           =   4785
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
               TabIndex        =   7
               Top             =   180
               Width           =   1155
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
               TabIndex        =   5
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
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
               TabIndex        =   6
               Top             =   180
               Width           =   1275
            End
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
               TabIndex        =   8
               Top             =   180
               Width           =   705
            End
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
            ItemData        =   "frmEstoque_Recebimento_consignacao.frx":7C46
            Left            =   6030
            List            =   "frmEstoque_Recebimento_consignacao.frx":7C5C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   4065
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
            Height          =   315
            Left            =   180
            TabIndex        =   2
            ToolTipText     =   "Texto para pesquisa."
            Top             =   1020
            Width           =   14805
         End
         Begin VB.ComboBox cmbfamilia 
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
            TabIndex        =   3
            ToolTipText     =   "Texto para pesquisa."
            Top             =   1020
            Width           =   14805
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Empresa"
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
            Left            =   2700
            TabIndex        =   31
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
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
            Left            =   7642
            TabIndex        =   23
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Texto para pesquisa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6877
            TabIndex        =   22
            Top             =   810
            Width           =   1470
         End
      End
      Begin VB.TextBox txtId_produto_lista 
         Alignment       =   2  'Centralizar
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2910
         MaxLength       =   50
         TabIndex        =   20
         Text            =   "0"
         Top             =   6150
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtid_estoque 
         Alignment       =   2  'Centralizar
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2225
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "0"
         Top             =   6150
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   75
         TabIndex        =   24
         Top             =   1320
         Width           =   14325
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   1050
            Top             =   180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox Txt_serie 
            Alignment       =   2  'Centralizar
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
            Height          =   315
            Left            =   5820
            MaxLength       =   3
            TabIndex        =   11
            ToolTipText     =   "Série."
            Top             =   390
            Width           =   645
         End
         Begin VB.Frame Frame_dt_emissao 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'Nenhum
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
            Height          =   345
            Left            =   6480
            TabIndex        =   33
            Top             =   390
            Width           =   1395
            Begin MSComCtl2.DTPicker txtemissao 
               Height          =   315
               Left            =   0
               TabIndex        =   12
               ToolTipText     =   "Data de emissão da nota fiscal."
               Top             =   0
               Width           =   1395
               _ExtentX        =   2461
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
               Format          =   140378113
               CurrentDate     =   39057
            End
         End
         Begin VB.CommandButton cmdcliente 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   13830
            Picture         =   "frmEstoque_Recebimento_consignacao.frx":7CB5
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Localizar destinatário."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtcliente 
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
            Left            =   8580
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Emitente."
            Top             =   390
            Width           =   5235
         End
         Begin VB.TextBox txtnotafiscal 
            Alignment       =   2  'Centralizar
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
            Left            =   4470
            MaxLength       =   9
            TabIndex        =   10
            ToolTipText     =   "Número da nota fiscal."
            Top             =   390
            Width           =   1335
         End
         Begin VB.TextBox txtid_cliente 
            Alignment       =   2  'Centralizar
            BackColor       =   &H80000014&
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
            Left            =   7890
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Código do cliente."
            Top             =   390
            Width           =   675
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Série*"
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
            Left            =   5985
            TabIndex        =   36
            Top             =   180
            Width           =   540
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Empresa*"
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
            Left            =   1800
            TabIndex        =   32
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Emissão"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6892
            TabIndex        =   28
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Alinhar à Direita
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparente
            Caption         =   "Nota fiscal*"
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
            Left            =   4650
            TabIndex        =   27
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparente
            Caption         =   "Emitente*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10837
            TabIndex        =   26
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparente
            Caption         =   "ID"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8145
            TabIndex        =   25
            Top             =   180
            Width           =   165
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   6885
         Left            =   -74925
         TabIndex        =   4
         Top             =   2820
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12144
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Nota fiscal"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Série"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Destinatário"
            Object.Width           =   18706
         EndProperty
      End
      Begin DrawSuite2014.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   29
         Top             =   330
         Width           =   15192
         _ExtentX        =   26802
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
         Begin DrawSuite2014.USImageList USImageList1 
            Left            =   7890
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmEstoque_Recebimento_consignacao.frx":7DB7
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_movimentacao 
         Height          =   3930
         Left            =   75
         TabIndex        =   16
         Top             =   5780
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6932
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "RE"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cód. de ref."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   4771
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde. PÇ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Corrida"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Certificado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Observações"
            Object.Width           =   4410
         EndProperty
      End
      Begin DrawSuite2014.USToolBar USToolBar2 
         Height          =   975
         Left            =   75
         TabIndex        =   30
         Top             =   330
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1720
         ButtonCount     =   9
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
         ButtonCaption2  =   "XML"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Importar XML"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   30
         ButtonHeight2   =   21
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
         ButtonLeft3     =   69
         ButtonTop3      =   2
         ButtonWidth3    =   38
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Excluir"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Excluir (F4)"
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
         ButtonLeft4     =   109
         ButtonTop4      =   2
         ButtonWidth4    =   39
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Excluir lote"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Excluir lote (F7)"
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
         ButtonLeft5     =   150
         ButtonTop5      =   2
         ButtonWidth5    =   60
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonAlignment6=   2
         ButtonType6     =   1
         ButtonStyle6    =   -1
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   -1
         ButtonLeft6     =   212
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   54
         ButtonCaption7  =   "Ajuda"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Ajuda (F1)"
         ButtonKey7      =   "6"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   216
         ButtonTop7      =   2
         ButtonWidth7    =   36
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Sair"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Sair (Esc)"
         ButtonKey8      =   "7"
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
         ButtonLeft8     =   254
         ButtonTop8      =   2
         ButtonWidth8    =   26
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   282
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         Begin DrawSuite2014.USImageList USImageList2 
            Left            =   7620
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmEstoque_Recebimento_consignacao.frx":9F9F
            Count           =   1
         End
      End
      Begin DrawSuite2014.USProgressBar PBLista 
         Height          =   255
         Left            =   -74925
         TabIndex        =   34
         Top             =   9720
         Width           =   15195
         _ExtentX        =   26802
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
      Begin DrawSuite2014.USProgressBar PBLista1 
         Height          =   255
         Left            =   75
         TabIndex        =   35
         Top             =   9720
         Width           =   15195
         _ExtentX        =   26802
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
   End
End
Attribute VB_Name = "frmEstoque_Recebimento_consignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Consignacao As Boolean 'OK
Dim StrSql_Localizar_Consignacao As String 'OK

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public PosicaoBase As Integer
Public lngPosicaoInicial As Long
Public lngPosicaoFinal As Long
Public lngPosicaoAuxiliar As Long

Public i As Integer
Public n As Long
Public lLinha As Integer

Public Function ProcCarregacampo(V1 As String, V2 As String, V3 As Integer)
On Error GoTo tratar_erro
    
        lngPosicaoInicial = InStr(IIf(PosicaoBase > 0, PosicaoBase, 1), strarquivo, V1, 1)
        lngPosicaoFinal = InStr(IIf(PosicaoBase > 0, PosicaoBase, 1), strarquivo, V2, 1)
        
    If lngPosicaoFinal > 0 And lngPosicaoInicial > 0 Then
        If lngPosicaoFinal > lngPosicaoInicial Then
            ProcCarregacampo = Mid(strarquivo, lngPosicaoInicial + V3, (lngPosicaoFinal - (lngPosicaoInicial + V3)))
            PosicaoBase = lngPosicaoFinal
            Debug.Print PosicaoBase
        End If
    End If
    
Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function ImportarXML(strCaminho) As Long
On Error GoTo tratar_erro
'PosicaoBase = 0
'
'    Lista.ListItems.Clear
'    lLinha = 1
'    XML.Text = strCaminho
'    ' Ler arquivo XML
'    n = FreeFile()
'    Open strCaminho For Input As #n
'    strarquivo = Input(LOF(n), n)
''    txtxml.Text = Replace(strarquivo, "ï»¿", "")
'    strarquivo = Replace(strarquivo, "ï»¿", "")
'    Close #n
'
'    infNFe = ProcCarregacampo("<infNFe", "/infNFe>", Len("<infNFe"))
'    infNFe = Left$(infNFe, 66)
'    infNFe = Right$(infNFe, 44)
'
'    lblChave.Caption = infNFe
'    lblnrChave.Caption = infNFe
'
''Dados da nota fiscal
'    V1 = "ide"
'    PosicaoBase = InStr(1, strarquivo, V1, 1)
'
'    'Dados da nota fiscal
'    natOp.Text = UCase(ProcCarregacampo("<natOp>", "</natOp>", Len("<natOp>")))
'
''    indPag.Text = UCase(ProcCarregacampo("<indPag>", "</indPag>", Len("<indPag>")))
''
''    Select Case indPag.Text
''        Case "0"
''        indPag.Text = "à vista"
''        Case "1"
''        indPag.Text = "a prazo"
''        Case "2"
''        indPag.Text = "Outros"
''    End Select
'
'    indmod.Text = UCase(ProcCarregacampo("<mod>", "</mod>", Len("<mod>")))
'
'    Serie.Text = ProcCarregacampo("<serie>", "</serie>", Len("<serie>"))
'    cNF.Text = ProcCarregacampo("<nNF>", "</nNF>", Len("<nNF>"))
'
'    dhEmi.Text = ProcCarregacampo("<dhEmi>", "</dhEmi>", Len("<dhEmi>"))
'    dhEmi.Text = Replace(dhEmi.Text, "T", " ")
'    dhEmi.Text = Left$(dhEmi.Text, 19)
'    dhEmi.Text = Format(dhEmi.Text, "General date")
'
'    dhSaiEnt.Text = ProcCarregacampo("<dhSaiEnt>", "</dhSaiEnt>", Len("<dhSaiEnt>"))
'    dhSaiEnt.Text = Replace(dhSaiEnt.Text, "T", " ")
'    dhSaiEnt.Text = Left$(dhSaiEnt.Text, 19)
'    dhSaiEnt.Text = Format(dhSaiEnt.Text, "General date")
'
'    finNFe.Text = UCase(ProcCarregacampo("<finNFe>", "</finNFe>", Len("<finNFe>")))
'    Select Case finNFe.Text
'        Case "1"
'        finNFe.Text = "NF-e normal"
'        Case "2"
'        finNFe.Text = "NF-e complementar"
'        Case "3"
'        finNFe.Text = "NF-e de ajuste"
'        Case "4"
'        finNFe.Text = "Devolução/Retorno"
'    End Select
'
'
'    indFinal.Text = UCase(ProcCarregacampo("<indFinal>", "</indFinal>", Len("<indFinal>")))
'    Select Case indFinal.Text
'        Case "0"
'        indFinal.Text = "Não"
'        Case "1"
'        indFinal.Text = "Consumidor final"
'    End Select
'
'    indPres.Text = UCase(ProcCarregacampo("<indPres>", "</indPres>", Len("<indPres>")))
'
'    Select Case indPres.Text
'        Case "0"
'        indPres.Text = "Não se aplica" ' (por exemplo, para a Nota Fiscal complementar ou de ajuste);
'        Case "1"
'        indPres.Text = "Operação presencial"
'        Case "2"
'        indPres.Text = "Operação não presencial, pela Internet"
'        Case "3"
'        indPres.Text = "Operação não presencial, Teleatendimento;"
'        Case "4"
'        indPres.Text = "NFC-e em operação com entrega em domicílio;"
'        Case "5"
'        indPres.Text = "Operação presencial, fora do estabelecimento"
'        Case "9"
'        indPres.Text = "Operação não presencial, outros."
'    End Select
'
'    'Dados do emitente
'    'CNPJ.Text = LerDadosXML(strarquivo, "SignatureValue", "")
'    CNPJ.Text = ProcCarregacampo("<CNPJ>", "</CNPJ>", Len("<CNPJ>"))
'    xNome.Text = UCase(ProcCarregacampo("<xNome>", "</xNome>", Len("<xNome>")))
'    xFant.Text = UCase(ProcCarregacampo("<xFant>", "</xFant>", Len("<xFant>")))
'
'    'Endereço emitente
'    xLgr.Text = UCase(ProcCarregacampo("<xLgr>", "</xLgr>", Len("<xLgr>")))
'    nro.Text = ProcCarregacampo("<nro>", "</nro>", Len("<nro>"))
'    xBairro.Text = UCase(ProcCarregacampo("<xBairro>", "</xBairro>", Len("<xBairro>")))
'    xMun.Text = UCase(ProcCarregacampo("<xMun>", "</xMun>", Len("<xMun>")))
'    UF.Text = UCase(ProcCarregacampo("<UF>", "</UF>", Len("<UF>")))
'    CEP.Text = ProcCarregacampo("<CEP>", "</CEP>", Len("<CEP>"))
'    xPais.Text = UCase(ProcCarregacampo("<xPais>", "</xPais>", Len("<xPais>")))
'    Var1 = "fone"
'    fone.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    '=====================================
'        'PosicaoBase = 1
'
'         '   V1 = "imposto"
'         '  lngPosicaoInicial = InStr(PosicaoBase, strArquivo, V1, 1)
'         '   PosicaoBase = lngPosicaoInicial
'
'        'ini:
'        '    Var1 = "CST"
'        '    CST = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))'
'
'        'Debug.Print PosicaoBase
'        'Debug.Print CST
'        'GoTo ini
'
'    '=====================================
'    'Carrega Dados lista de produtos
'    '=====================================
'
'    V1 = "prod"
'    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
'
'Inicio:
'
'    If PosicaoBase > 0 Then
'
'    Dim cProd As String, xProd As String, NCM As String, CFOP As String, uCom As String, qCom As String, vUnCom As String, vProd As String, orig As String ', CST As String
'
'    Var1 = "cProd"
'    cProd = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
'    Var1 = "xProd"
'    xProd = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
'    Var1 = "NCM"
'    NCM = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
'    Var1 = "CFOP"
'    CFOP = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
'    Var1 = "uCom"
'    uCom = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
'    Var1 = "qCom"
'    qCom = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
'    Var1 = "vUnCom"
'    vUnCom = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'    Var1 = "vProd"
'    vProd = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
''Carrega impostos do produto
'    V1 = "imposto"
'    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
'Debug.Print PosicaoBase
'If PosicaoBase > 0 Then
'    Var1 = "orig"
'    orig = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "CSOSN"
'
'    CST = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'    If CST = "" Then
'        Var1 = "CST"
'        CST = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'    End If
'
'    orig = orig & CST
'
'    Var1 = "CST"
'    V1 = "IPI"
'    PosicaoBase = InStr(lngPosicaoFinal, strarquivo, V1, 1)
'    CSTIPI = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    V1 = "PIS"
'    PosicaoBase = InStr(lngPosicaoFinal, strarquivo, V1, 1)
'    CSTPIS = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    V1 = "COFINS"
'    PosicaoBase = InStr(lngPosicaoFinal, strarquivo, V1, 1)
'    CSTCOFINS = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'If cProd <> "" Then
'
'If IsNumeric(vUnCom) Then
'    vUnCom = Replace(vUnCom, ".", ",")
'    vUnCom = "R$ " & vUnCom
'End If
'
'If IsNumeric(vProd) Then
'    vProd = Replace(vProd, ".", ",")
'    vProd = "R$ " & vProd
'End If
'
'
'    ValorTotal = 0
'
'        With Lista.ListItems
'            .Add , , lLinha
'            .Item(.Count).SubItems(1) = cProd
'            .Item(.Count).SubItems(2) = xProd
'            .Item(.Count).SubItems(3) = NCM
'            .Item(.Count).SubItems(4) = CFOP
'            .Item(.Count).SubItems(5) = orig
'            .Item(.Count).SubItems(6) = CSTIPI
'            .Item(.Count).SubItems(7) = CSTPIS
'            .Item(.Count).SubItems(8) = CSTCOFINS
'            .Item(.Count).SubItems(9) = uCom
'            .Item(.Count).SubItems(10) = vUnCom
'            .Item(.Count).SubItems(11) = qCom
'            .Item(.Count).SubItems(12) = vProd
'        End With
'
'    lLinha = lLinha + 1
'GoTo Inicio
'End If
'
'End If
'End If
'
''Carregatotais da nota
'    V1 = "total"
'    PosicaoBase = InStr(1, strarquivo, V1, 1)
'
'    Var1 = "vBC"
'    vBC.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "vICMS"
'    vICMS.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "vProd"
'    vProdTotal.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'    vProdTotal.Text = Replace(vProdTotal.Text, ".", ",")
'    vProdTotal.Text = Format(vProdTotal.Text, "###,##0.00")
'
'
'    Var1 = "vFrete"
'    vFrete.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "vIPI"
'    vIPI.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "vPIS"
'    vPIS.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'    Var1 = "vCOFINS"
'    vCOFINS.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
'
'
'Exit Function
'tratar_erro:
'    MsgBox ("Descrição do erro : " + Error()), vbCritical
'    Exit Function
'End Function
'
'Sub ProcAjuda()
'On Error GoTo tratar_erro
'
'FunAbrirVideoWeb ("http://www.youtube.com/watch?v=Akc9wt_4w8c&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=50&feature=plcp")
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
'
'Sub ProcLimpaCampos()
'On Error GoTo tratar_erro
'
'txtnotafiscal.Text = ""
'txtemissao.Value = Date
'txtcliente.Text = ""
'txtid_cliente.Text = ""
'Txt_tipodest = ""
'txtid_estoque.Text = 0

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcLimpaCamposItem()
On Error GoTo tratar_erro

txtid_estoque.Text = 0
txtId_produto_lista = 0
txtcodproduto.Text = ""
Cmb_cod_ref.Clear
txtdesc.Text = ""
txtdesctecnica.Text = ""
cmbLocal_armaz.ListIndex = -1
txtqtde.Text = ""
txtcorrida.Text = ""
txtCertificado.Text = ""
txtpeso.Text = ""
txtfamilia.Text = ""
txtobs.Text = ""
txtun.Text = ""
txtVlr_unit = ""
txtVlr_total = "0,00"
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa1_Click()
On Error GoTo tratar_erro

'ProcLimpaCampos
txtdesenho = ""
ProcLimpaCamposItem
Lista_movimentacao.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbfiltrarpor = "Família" Then
    cmbfamilia.Visible = True
    txtTexto.Visible = False
Else
    cmbfamilia.Visible = False
    txtTexto.Visible = True
    If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcliente_Click()
On Error GoTo tratar_erro

Faturamento = False
frmFaturamento_Prod_serv_cliente_forn.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEscolher_item_Click()
On Error GoTo tratar_erro

frmEstoque_Recebimento_consignacao_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then Exit Sub
If txtId_produto_lista = 0 Then
    USMsgBox ("Informe a movimentação na lista antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir esta movimentação do estoque?", vbYesNo + vbQuestion) = vbYes Then
    Mensagem = "Não é permitido excluir esta movimentação, pois a mesma está sendo utilizada no módulo"
    ProcVerificaRegistroUtilizado "Producao_NF_Consignada", "IDestoque = " & txtid_estoque, "PCP/Gerenciamento de ordem"
    If Permitido = False Then Exit Sub
    
    If Txt_serie <> "" Then TextoFiltro = " and Serie = '" & Txt_serie & "'" Else TextoFiltro = ""
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select ID from tbl_Dados_Nota_Fiscal where int_NotaFiscal = '" & txtnotafiscal & "' and txt_Razao_Nome = '" & txtcliente & "' and int_TipoNota = 2" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBPedido.EOF = False Then
        ProcVerificaRegistroUtilizado "tbl_detalhes_nota", "ID_Nota = " & TBPedido!ID & " and int_cod_produto = '" & Lista_movimentacao.SelectedItem.SubItems(2) & "'", "Estoque/Nota fiscal"
        If Permitido = False Then Exit Sub
    End If
    TBPedido.Close
    
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select IDoperacao from Estoque_movimentacao where IdEstoque = " & txtid_estoque & " and Saida > 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBPedido.EOF = False Then
        USMsgBox ("Não é permitido excluir esta movimentação, pois este RE já possui movimentação de saída."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBPedido.Close
    
    Conexao.Execute "DELETE from Estoque_controle where IdEstoque = " & txtid_estoque
    Conexao.Execute "DELETE from Estoque_movimentacao where Idoperacao = " & txtId_produto_lista

    USMsgBox ("Movimentação excluída com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Estoque/Recebimento/Consignação"
    Evento = "Excluir movimentação"
    ID_documento = Lista_movimentacao.SelectedItem
    Documento = "Nota fiscal: " & txtnotafiscal & " - Emitente: " & txtcliente
    Documento1 = "Código interno: " & txtdesenho
    ProcGravaEvento
    '==================================
    ProcLimpaCamposItem
    txtdesenho = ""
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

If txtdesenho = "" Then Exit Sub
ProcLimpaCamposItem
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Projproduto where desenho = '" & txtdesenho & "' and Tipo = 'P' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcPuxaDadosProduto
Else
    USMsgBox ("Não foi encontrado nenhum produto com esse código interno."), vbExclamation, "CAPRIND v5.0"
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosProduto()
On Error GoTo tratar_erro

txtcodproduto.Text = TBProduto!Codproduto
txtdesenho.Text = TBProduto!Desenho
txtdesc.Text = TBProduto!Descricao
txtfamilia.Text = TBProduto!Classe
txtun.Text = TBProduto!Unidade
txtdesctecnica.Text = TBProduto!descricaotecnica
txtpeso.Text = IIf(IsNull(TBProduto!peso_metro), "", Format(TBProduto!peso_metro, "###,##0.0000"))
txtVlr_unit = IIf(IsNull(TBProduto!PCusto), "", Format(TBProduto!PCusto, "###,##0.0000000000"))
ProcCarregaComboCodRef Cmb_cod_ref, "P.codproduto = " & TBProduto!Codproduto, 0, "", False, True
Proccarregalocarm

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNota_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

ID_nota = 0
Acao = "emitir a nota"
If Cmb_empresa1 = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa1.SetFocus
    Exit Sub
End If
If txtnotafiscal = "" Then
    NomeCampo = "a nota fiscal"
    ProcVerificaAcao
    txtnotafiscal.SetFocus
    Exit Sub
End If
If txtcliente = "" Then
    NomeCampo = "o emitente"
    ProcVerificaAcao
    cmdcliente_Click
End If

'Verifica se tem algum produto/serviço recebido para o pedido
If Lista_movimentacao.ListItems.Count = 0 Then
    USMsgBox ("É necessário receber o(s) produto(s)/serviço(s) antes de emitir a nota."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Txt_tipodest = "C" Then NomeTabela = "Clientes" Else NomeTabela = "Compras_fornecedores"

'Cria a nota fiscal
If Txt_serie <> "" Then TextoFiltro = " and Serie = '" & Txt_serie & "'" Else TextoFiltro = ""
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal where int_NotaFiscal = '" & txtnotafiscal & "' and Id_Int_Cliente = " & txtid_cliente & " and int_TipoNota = 2" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!TabelaSN = 0
    TBGravar!Regime = FunVerifRegimeEmpresa(Cmb_empresa1.ItemData(Cmb_empresa1.ListIndex))
    TBGravar!pedido_interno = False
    TBGravar!DtValidacaoOF = Now
    TBGravar!RespValidacaoOF = pubUsuario
    TBGravar!ID_empresa = Cmb_empresa1.ItemData(Cmb_empresa1.ListIndex)
    TBGravar!int_NotaFiscal = txtnotafiscal
    TBGravar!Serie = Txt_serie
    TBGravar!int_TipoNota = "2"
    TBGravar!Id_Int_Cliente = txtid_cliente
    TBGravar!txt_Razao_Nome = txtcliente
    TBGravar!TipoNF = "M1"
    TBGravar!dt_DataEmissao = txtemissao.Value
    
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from " & NomeTabela & " where IDCliente = " & txtid_cliente, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        
        If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
            Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
        Else
            Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
        End If
        TBGravar!txt_Endereco = Endereco
        TBGravar!Numero = IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero)
        If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
            Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
        Else
            Bairro = IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
        End If
        TBGravar!txt_Bairro = Bairro
        
        If Txt_tipodest = "C" Then
            TBGravar!txt_tipocliente = IIf(IsNull(TBClientes!Tipo), "", TBClientes!Tipo)
            If TBClientes!Tipo = "JP" Or TBClientes!Tipo = "JR" Then TBGravar!txt_IE_Cliente = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
            TBGravar!txt_UF = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
            TBGravar!txt_Fone_Fax = IIf(IsNull(TBClientes!Tel01), "", TBClientes!Tel01)
            If TBClientes!chkSuframa = True Then Suframa = True Else Suframa = False
        Else
            If TBClientes!Pessoa = "JURÍDICA" Then
                TBGravar!txt_tipocliente = "J"
                TBGravar!txt_IE_Cliente = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
            Else
                TBGravar!txt_tipocliente = "F"
            End If
            TBGravar!txt_UF = IIf(IsNull(TBClientes!Estado), "", TBClientes!Estado)
            TBGravar!txt_Fone_Fax = IIf(IsNull(TBClientes!Telefones), "", TBClientes!Telefones)
            Suframa = False
        End If
        If TBClientes!idTipoEmpresa = 1 Then TBGravar!txt_CNPJ_CPF = IIf(IsNull(TBClientes!CPF_CNPJ), "", TBClientes!CPF_CNPJ)
        TBGravar!Txt_CEP = IIf(IsNull(TBClientes!CEP), "", TBClientes!CEP)
        TBGravar!txt_Municipio = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
    End If
    
    TBGravar!txt_Hora_Saida = Format(Date, "hh:mm")
    TBGravar!Int_status = "1"
    TBGravar!Aplicacao = "T"
    TBGravar.Update
Else
    'Verifica se a NF já foi validada e não permite alteração
    If IsNull(TBGravar!DtValidacao) = False Then
        USMsgBox ("Esta nota fiscal não será alterada, pois a mesma já foi validada."), vbInformation, "CAPRIND v5.0"
        TBGravar.Close
        GoTo Validada
    End If
End If
ID_nota = TBGravar!ID
TBGravar.Close

'Cria ou altera os produtos
Desenho = ""
DesenhoProduto = ""
Valor1 = 0
Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select * from " & NomeTabela & " where IDcliente = " & txtid_cliente, Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = False Then
    Set TBReceber = CreateObject("adodb.recordset")
    TBReceber.Open "Select EM.*, EC.ref, EC.ID_cliente, EC.Tipodest_NFcons from Estoque_movimentacao EM INNER JOIN estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.documento = '" & txtnotafiscal & "' and EC.ID_cliente = " & txtid_cliente & " and EM.Operacao = 'ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO' order by EM.Desenho, EC.ref", Conexao, adOpenKeyset, adLockOptimistic
    If TBReceber.EOF = False Then
        Do While TBReceber.EOF = False
            If Desenho <> TBReceber!Desenho Or DesenhoProduto <> TBReceber!Ref Or Valor1 <> TBReceber!VlrUnit Then
            
                'Verifica a qtde. total recebida do produto na NF por código int., cód. de ref e valor
                valor = IIf(IsNull(TBReceber!VlrUnit), 0, TBReceber!VlrUnit)
                NovoValor = Replace(valor, ",", ".")
                
                Qtde = 0
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(EM.Entrada) as Qtde from Estoque_movimentacao EM INNER JOIN estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.documento = '" & txtnotafiscal & "' and EC.ID_cliente = " & txtid_cliente & " and EM.Desenho = '" & TBReceber!Desenho & "' and EC.ref = '" & TBReceber!Ref & "' and EM.VlrUnit = " & NovoValor & " and EM.Operacao = 'ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
                End If
                
                'Vincula o cod ref ao cliente/fornecedor
                If TBReceber!Ref <> "" Then
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select I.iD_cliente_forn, I.Tipo, I.N_referencia, I.Codproduto, I.Descricao, I.Aplicacao from item_aplicacoes I INNER JOIN projproduto P on I.codproduto = P.codproduto where I.n_referencia = '" & TBReceber!Ref & "' and P.desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = True Then
                        TBItem.AddNew
                        
                        TBItem!N_referencia = TBReceber!Ref
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select Codproduto, Descricao from projproduto  where desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = False Then
                            TBItem!Codproduto = TBProduto!Codproduto
                            TBItem!Descricao = TBProduto!Descricao
                        End If
                        TBProduto.Close
                    End If
                    TBItem!ID_cliente_forn = TBReceber!ID_cliente
                    TBItem!Tipo = TBReceber!Tipodest_NFcons
                    TBItem!Aplicacao = txtcliente
                    TBItem.Update
                    TBItem.Close
                End If
                
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from projproduto where desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    valor = IIf(IsNull(TBReceber!VlrUnit), 0, TBReceber!VlrUnit) / FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)
                    NovoValor = Replace(valor, ",", ".")
                
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_Detalhes_Nota where int_Cod_Produto = '" & TBReceber!Desenho & "' and N_referencia = '" & TBReceber!Ref & "' and dbl_ValorUnitario = " & NovoValor & " and id_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then TBAbrir.AddNew
                    TBAbrir!Tipo = "P"
                    TBAbrir!int_Cod_Produto = TBReceber!Desenho
                    TBAbrir!N_referencia = TBReceber!Ref
                    TBAbrir!int_NotaFiscal = txtnotafiscal
                    TBAbrir!ID_nota = ID_nota
                    TBAbrir!int_Qtd = Qtde * FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)
                    TBAbrir!Saldo = TBAbrir!int_Qtd
                    TBAbrir!Codproduto = TBItem!Codproduto
                    TBAbrir!txt_Descricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
                    TBAbrir!ID_CFOP = IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP)
                    If IsNull(TBItem!ID_CF) = False Then TBAbrir!ID_CF = TBItem!ID_CF
                    TBAbrir!txt_Unid = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
                    TBAbrir!Unidade_com = IIf(IsNull(TBItem!Unidade_com), "", TBItem!Unidade_com)
                    TBAbrir!Familia = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
                    
                    ProcControleImposto IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP), IIf(txtid_cliente = "", 0, txtid_cliente)
                    ProcBuscaTributos IIf(IsNull(TBAbrir!ID_CF), 0, TBAbrir!ID_CF)
                    If Txt_tipodest = "C" Then ProcVerificaRegiao TBCarteira!UF, txtid_cliente, txtcliente Else ProcVerificaRegiao TBCarteira!Estado, txtid_cliente, txtcliente
                    If TemIPI = "SIM" Then TBAbrir!int_IPI = vRegiao(0, 0) Else TBAbrir!int_IPI = 0
                    If TemICMS = "SIM" Then TBAbrir!int_ICMS = vRegiao(0, 1) Else TBAbrir!int_ICMS = 0
                    
                    TBAbrir!dbl_ValorUnitario = valor
                    TBAbrir!dbl_ValorTotal = Format(TBAbrir!dbl_ValorUnitario * TBAbrir!int_Qtd, "###,##0.00")
                    TBAbrir!dbl_valoripi = Format((TBAbrir!dbl_ValorTotal * TBAbrir!int_IPI) / 100, "###,##0.00")
                        
                    TBAbrir.Update
                    TBAbrir.Close
                End If
                TBItem.Close
            End If
            Desenho = TBReceber!Desenho
            DesenhoProduto = TBReceber!Ref
            Valor1 = TBReceber!VlrUnit
            TBReceber.MoveNext
        Loop
    Else
        USMsgBox ("Não há produto recebido para a nota " & txtnotafiscal & "."), vbExclamation, "CAPRIND v5.0"
    End If
End If

Validada:
    If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then Unload frmFaturamento_Prod_Serv
    Faturamento_NF_Saida = False
    Formulario = "Estoque/Nota fiscal"
    With frmFaturamento_Prod_Serv
        .Novo_Nota = False
        .Faturamento_Vendas_PI = False
        .TxtID.Text = ID_nota
        .txtNFiscal.Text = txtnotafiscal
        .ProcCarregaDadosNota .TxtID.Text
        .ProcCarregaLista
        .ProcCarregaListaServicos
        .ProcGravarTotaisNota
        .ProcCarregaDadosTransporte
        .ProcCarregaDuplicatas
        .Show
        .txt_DtEmissao.Value = txtemissao.Value
        .txtserie.Locked = False
        .txtserie.TabStop = True
        
        CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.dt_Saida_Entrada, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, TN.dbl_Valor_Total_Nota, NF.txt_Razao_Nome, NF.Int_status, NF.Imprimir, NF.ID_empresa, NF.Aplicacao, NF.DtValidacaoOF, NF.DtValidacao"
        .Strsql_Faturamento = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .TxtID
        .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .TxtID & " and NF.Int_status = 1"
        .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .TxtID & " and NF.Int_status = 2"
        .Strsql_FaturamentoNFe = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.TipoNF <> 'SA' AND  NF.ID = " & .TxtID
        .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.TipoNF = 'SA' AND NF.ID = " & .TxtID
        .ProcCarregaListaNota (1)
        
        If USMsgBox("Deseja prosseguir com o preenchimento dos dados da nota fiscal?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Unload frmFaturamento_Prod_Serv
    End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

CamposFiltro = "NF, Serie, emissaonf, ID_cliente, Cliente"
INNERJOINTEXTO = "Select " & CamposFiltro & " from Estoque_controle where "
TextoFiltroPadrao = "Consignacao = 'True' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Cliente is not null and NF is not null and Status = 'CONSIGNAÇÃO RECEBIDA' group by " & CamposFiltro & " order by Cliente, NF"
If txtTexto <> "" Or cmbfamilia <> "" Then
    If cmbfiltrarpor = "Família" Then
        StrSql_Localizar_Consignacao = INNERJOINTEXTO & "classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor
            Case "Nota fiscal":
                TextoFiltro = "nf"
                If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
            Case "Destinatário": TextoFiltro = "Cliente"
            Case "Código interno": TextoFiltro = "Desenho"
            Case "Código de referência": TextoFiltro = "Ref"
            Case "Descrição": TextoFiltro = "Descricao"
        End Select
        StrSql_Localizar_Consignacao = INNERJOINTEXTO & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Localizar_Consignacao = INNERJOINTEXTO & TextoFiltroPadrao
End If
ProcCarregaListaNotas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaNotas()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Localizar_Consignacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , IIf(IsNull(TBLISTA!NF), "", TBLISTA!NF)
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Serie), "", TBLISTA!Serie)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!emissaonf), "", Format(TBLISTA!emissaonf, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!ID_cliente), "", TBLISTA!ID_cliente)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    Select Case KeyCode
        Case vbKeyF2: ProcFiltrar
        Case vbKeyF1: ' ProcAjuda
        Case vbKeyEscape: ProcSair
    End Select
Else
    Select Case KeyCode
        Case vbKeyInsert: ProcNovo
        Case vbKeyF3: ProcGravar
        Case vbKeyF4: ProcExcluir
        Case vbKeyF7: ProcExcluirLote
        Case vbKeyF1: 'ProcAjuda
        Case vbKeyEscape: ProcSair
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 5, True
ProcCarregaToolBar2 Me, 15192, 8, True

Formulario = "Estoque/Recebimento/Consignação"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
cmbfiltrarpor = "Nota fiscal"
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboEmpresa Cmb_empresa1, False
txtdata = Date

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirLote()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If Frame1.Enabled = False Or txtnotafiscal.Text = "" Then
    USMsgBox ("Informe a nota fiscal antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir todas as movimentações dessa nota fiscal nº " & txtnotafiscal.Text & " série " & Txt_serie & "?", vbYesNo + vbQuestion) = vbYes Then
    Mensagem = "Não é permitido excluir todas as movimentações desta nota, pois a mesma está sendo utilizada no módulo"
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select IDestoque, Desenho from estoque_controle where nf = '" & txtnotafiscal.Text & "' and id_cliente = " & txtid_cliente.Text & " and consignacao = 'True'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Do While TBEstoque.EOF = False
            ProcVerificaRegistroUtilizado "Producao_NF_Consignada", "IDestoque = " & TBEstoque!IDestoque, "PCP/Gerenciamento de ordem"
            If Permitido = False Then Exit Sub
            
            If Txt_serie <> "" Then TextoFiltro = " and Serie = '" & Txt_serie & "'" Else TextoFiltro = ""
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select ID from tbl_Dados_Nota_Fiscal where int_NotaFiscal = '" & txtnotafiscal & "' and txt_Razao_Nome = '" & txtcliente & "' and int_TipoNota = 2" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                ProcVerificaRegistroUtilizado "tbl_detalhes_nota", "ID_Nota = " & TBPedido!ID & " and int_cod_produto = '" & Lista_movimentacao.SelectedItem.SubItems(2) & "'", "Estoque/Nota fiscal"
                If Permitido = False Then Exit Sub
            End If
            TBPedido.Close
            
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select IDoperacao from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDestoque & " and Saida > 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                USMsgBox ("Não é permitido excluir todas as movimentações desta nota, pois este RE já possui movimentação de saída."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
            TBPedido.Close
            
            Conexao.Execute "DELETE from estoque_movimentacao where idestoque = " & TBEstoque!IDestoque
            Conexao.Execute "DELETE from estoque_controle where idestoque = " & TBEstoque!IDestoque
            
            '==================================
            Modulo = "Estoque/Recebimento/Consignação"
            Evento = "Excluir lote"
            ID_documento = TBEstoque!IDestoque
            Documento = "Nota fiscal: " & txtnotafiscal & " - Emitente: " & txtcliente
            Documento1 = "Código interno: " & TBEstoque!Desenho
            ProcGravaEvento
            '==================================
            TBEstoque.MoveNext
        Loop
        USMsgBox ("Nota fiscal excluída com sucesso."), vbInformation, "CAPRIND v5.0"
        'ProcLimpaCampos
        txtdesenho = ""
        ProcLimpaCamposItem
        Lista_movimentacao.ListItems.Clear
        Frame1.Enabled = False
        Frame2.Enabled = False
        Novo_Consignacao = False
    End If
    TBEstoque.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
txtdesenho = ""
ProcLimpaCamposItem
Frame1.Enabled = True
ProcLiberaCamposPrinc
Frame2.Enabled = True
Novo_Consignacao = True
If txtnotafiscal = "" Then Cmb_empresa1.SetFocus Else txtdesenho.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_empresa1 = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa1.SetFocus
    Exit Sub
End If
If txtnotafiscal = "" Then
    NomeCampo = "o número da nota fiscal"
    ProcVerificaAcao
    txtnotafiscal.SetFocus
    Exit Sub
Else
    txtnotafiscal = FunTamanhoTextoZeroEsq(DS_RetornarNumeros(txtnotafiscal), 9)
End If
If Txt_serie = "" Then
    NomeCampo = "a série"
    ProcVerificaAcao
    Txt_serie.SetFocus
    Exit Sub
End If
If txtcliente = "" Then
    NomeCampo = "o destinatário"
    ProcVerificaAcao
    cmdcliente_Click
    Exit Sub
End If
If txtdesenho = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtdesenho.SetFocus
    Exit Sub
End If
If txtVlr_unit = "" Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    txtVlr_unit.SetFocus
    Exit Sub
End If
If cmbLocal_armaz = "" Then
    NomeCampo = "o local de armazenamento"
    ProcVerificaAcao
    cmbLocal_armaz.SetFocus
    Exit Sub
End If
qt = IIf(txtqtde = "", 0, txtqtde)
If qt = 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtqtde.SetFocus
    Exit Sub
End If
If txtcorrida = "" Then txtcorrida = 0
If txtCertificado = "" Then txtCertificado = 0

Desenho = ""
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Estoque_controle where idestoque = " & txtid_estoque, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Mensagem = "Não é permitido alterar esta movimentação, pois a mesma está sendo utilizada no módulo"
    ProcVerificaRegistroUtilizado "Producao_NF_Consignada", "IDestoque = " & txtid_estoque, "PCP/Gerenciamento de ordem"
    If Permitido = False Then Exit Sub
    
    If Txt_serie <> "" Then TextoFiltro = " and Serie = '" & Txt_serie & "'" Else TextoFiltro = ""
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select ID from tbl_Dados_Nota_Fiscal where int_NotaFiscal = '" & txtnotafiscal & "' and txt_Razao_Nome = '" & txtcliente & "' and int_TipoNota = 2" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBPedido.EOF = False Then
        ProcVerificaRegistroUtilizado "tbl_detalhes_nota", "ID_Nota = " & TBPedido!ID & " and int_cod_produto = '" & Lista_movimentacao.SelectedItem.SubItems(2) & "'", "Estoque/Nota fiscal"
        If Permitido = False Then Exit Sub
    End If
    TBPedido.Close
    
    Set TBPedido = CreateObject("adodb.recordset")
    TBPedido.Open "Select IDoperacao from Estoque_movimentacao where IdEstoque = " & txtid_estoque & " and Saida > 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBPedido.EOF = False Then
        USMsgBox ("Não é permitido alterar esta movimentação, pois este RE já possui movimentação de saída."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBPedido.Close
    
    Evento = "Alterar"
    Desenho = TBGravar!Desenho
Else
    TBGravar.AddNew
    Evento = "Novo"
End If

ProcEnviaDados
TBGravar.Update

Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_movimentacao where idoperacao = " & txtId_produto_lista, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
Else
    TBEstoque.AddNew
    USMsgBox ("Produto recebido ao estoque com sucesso."), vbInformation, "CAPRIND v5.0"
End If
'Grava na tabela Estoque_movimentacao
TBEstoque!Destino = "Interno"
TBEstoque!Terceiros = False
TBEstoque!IDestoque = TBGravar!IDestoque
TBEstoque!Operacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO"
TBEstoque!Desenho = txtdesenho.Text
TBEstoque!Documento = txtnotafiscal.Text
TBEstoque!LOTE = txtnotafiscal.Text
TBEstoque!Descricao = txtdesc.Text
TBEstoque!DtEmissao = txtdata
TBEstoque!Entrada = Format(txtqtde.Text, "###.##0.000")
TBEstoque!Entrada_PC = IIf(Txt_qtde_PC = "", Null, Format(Txt_qtde_PC, "###.##0.000"))
TBEstoque!Responsavel = pubUsuario
TBEstoque!Cliente = txtcliente.Text
TBEstoque!data = txtdata
TBEstoque!VlrUnit = Format(txtVlr_unit, "###.##0.00000")
TBEstoque!VlrTotal = Format(txtVlr_total, "###.##0.00")
TBEstoque!Obs = IIf(txtobs.Text = "", Null, txtobs)
TBEstoque.Update
txtId_produto_lista = TBEstoque!IDoperacao
TBEstoque.Close

'==================================
Modulo = "Estoque/Recebimento/Consignação"
ID_documento = TBGravar!IDestoque
Documento = "Nota fiscal: " & txtnotafiscal & " - Emitente: " & txtcliente
Documento1 = "Código interno: " & txtdesenho
ProcGravaEvento
'==================================
TBGravar.Close

ProcCarregaLista
If Novo_Consignacao = False Then
    If CodigoLista <> 0 And Lista_movimentacao.ListItems.Count <> 0 Then
        Lista_movimentacao.SelectedItem = Lista_movimentacao.ListItems(CodigoLista)
        Lista_movimentacao.SetFocus
    End If
End If
Novo_Consignacao = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista_movimentacao.ListItems.Clear
If Txt_serie <> "" Then TextoFiltro = " and EC.Serie = '" & Txt_serie & "'" Else TextoFiltro = ""
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select EM.idoperacao, EM.Obs, EM.Entrada, EM.Entrada_PC, EC.* from Estoque_controle EC INNER JOIN Estoque_movimentacao EM ON EC.IDestoque = EM.IDEstoque where EC.nf = '" & txtnotafiscal & "' and EC.cliente = '" & txtcliente & "' and EC.consignacao = 'True' and EM.Operacao = 'ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcBloqueiaCamposPrinc
    
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_movimentacao.ListItems
            .Add , , TBLISTA!IDoperacao
            .Item(.Count).SubItems(1) = TBLISTA!IDestoque
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Ref), "", TBLISTA!Ref)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Classe), "", TBLISTA!Classe)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Entrada), "0,0000", Format(TBLISTA!Entrada, "###,##0.0000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Entrada_PC), "", TBLISTA!Entrada_PC)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Corrida), "", TBLISTA!Corrida)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
Else
    ProcLiberaCamposPrinc
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from estoque_controle where idestoque = " & Lista_movimentacao.SelectedItem.ListSubItems(1), Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtnotafiscal = IIf(IsNull(TBLISTA!LOTE), "", TBLISTA!LOTE)
    txtemissao = TBLISTA!emissaonf
    txtid_cliente = IIf(IsNull(TBLISTA!ID_cliente), "", TBLISTA!ID_cliente)
    txtcliente = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
    ProcBloqueiaCamposPrinc
    
    txtdesenho.Text = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
    If IsNull(TBLISTA!Ref) = False And TBLISTA!Ref <> "" Then
        ProcCarregaComboCodRef Cmb_cod_ref, "P.Desenho = '" & txtdesenho & "'", 0, "", False, True
        Cmb_cod_ref = TBLISTA!Ref
    Else
        Cmb_cod_ref.Clear
    End If
    
    txtId_produto_lista = Lista_movimentacao.SelectedItem
    txtid_estoque = Lista_movimentacao.SelectedItem.ListSubItems(1)
    txtdesc.Text = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
    txtdesctecnica.Text = IIf(IsNull(TBLISTA!descricaotecnica), "", TBLISTA!descricaotecnica)
    txtpeso.Text = IIf(IsNull(TBLISTA!peso_unit), "", TBLISTA!peso_unit)
    txtqtde.Text = IIf(IsNull(TBLISTA!Qtde), "0.000", Format(TBLISTA!Qtde, "0.000"))
    txtcorrida.Text = IIf(IsNull(TBLISTA!Corrida), "", TBLISTA!Corrida)
    txtCertificado.Text = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
    txtfamilia.Text = IIf(IsNull(TBLISTA!Classe), "", TBLISTA!Classe)
    txtobs.Text = Lista_movimentacao.SelectedItem.ListSubItems(10)
    txtun.Text = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
    txtdata.Value = IIf(IsNull(TBLISTA!data), Date, Format(TBLISTA!data, "dd/mm/yyyy"))
    txtVlr_unit.Text = IIf(IsNull(TBLISTA!valor_unitario), "", Format(TBLISTA!valor_unitario, "###,##0.0000000000"))
    txtVlr_total.Text = IIf(IsNull(TBLISTA!Valor_total), "0,00", Format(TBLISTA!Valor_total, "###,##0.00"))
    Proccarregalocarm
    If (IsNull(TBLISTA!local_armaz)) = False And TBLISTA!local_armaz <> "" Then cmbLocal_armaz = TBLISTA!local_armaz
End If
TBLISTA.Close
Frame1.Enabled = True
Frame2.Enabled = True
Novo_Consignacao = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Consignacao = True Then
    If USMsgBox("A movimentação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo + vbQuestion) = vbYes Then
        ProcGravar
        If Novo_Consignacao = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Consignacao = False
Unload Me

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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    If Lista.ListItems.Count = 0 Then Exit Sub
    If .SelectedItem.ListSubItems(1) <> "" Then TextoFiltro = " and Serie = '" & .SelectedItem.ListSubItems(1) & "'" Else TextoFiltro = ""
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Estoque_controle where cliente = '" & .SelectedItem.ListSubItems(4) & "' and nf = '" & .SelectedItem & "' and Consignacao = 'True'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtnotafiscal = TBAbrir!NF
        Txt_serie = IIf(IsNull(TBAbrir!Serie), "", TBAbrir!Serie)
        txtcliente = TBAbrir!Cliente
        txtemissao = Format(TBAbrir!emissaonf, "dd/mm/yy")
        txtid_cliente = TBAbrir!ID_cliente
        Txt_tipodest = IIf(IsNull(TBAbrir!Tipodest_NFcons), "", TBAbrir!Tipodest_NFcons)
        Novo_Consignacao = False
        ProcBloqueiaCamposPrinc
        Frame2.Enabled = True
        ProcCarregaLista
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposPrinc()
On Error GoTo tratar_erro

With txtnotafiscal
    .Locked = True
    .TabStop = False
End With
With Txt_serie
    .Locked = True
    .TabStop = False
End With
Frame_dt_emissao.Enabled = False
With txtid_cliente
    .Locked = True
    .TabStop = False
End With
cmdcliente.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaCamposPrinc()
On Error GoTo tratar_erro

With txtnotafiscal
    .Locked = False
    .TabStop = True
End With
With Txt_serie
    .Locked = False
    .TabStop = True
End With
Frame_dt_emissao.Enabled = True
With txtid_cliente
    .Locked = False
    .TabStop = True
End With
cmdcliente.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_movimentacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_movimentacao, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_movimentacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_movimentacao.ListItems.Count = 0 Then Exit Sub
ProcPuxaDados
CodigoLista = Lista_movimentacao.SelectedItem.index

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

'Grava na tabela Estoque_Controle
TBGravar!ID_empresa = Cmb_empresa1.ItemData(Cmb_empresa1.ListIndex)
TBGravar!status = "CONSIGNAÇÃO RECEBIDA"
TBGravar!emissaonf = txtemissao.Value
TBGravar!consignacao = True
TBGravar!Ref = Cmb_cod_ref
TBGravar!LOTE = txtnotafiscal.Text
TBGravar!Desenho = txtdesenho.Text
TBGravar!Descricao = txtdesc.Text
TBGravar!peso_unit = txtpeso.Text
TBGravar!descricaotecnica = txtdesctecnica.Text
TBGravar!data = txtdata
TBGravar!estoque_real = Format(txtqtde.Text, "###.##0.000")
TBGravar!estoque_real_PC = IIf(Txt_qtde_PC = "", Null, Format(Txt_qtde_PC, "###.##0.000"))
TBGravar!estoque_venda = Format(txtqtde.Text, "###.##0.000")
TBGravar!Qtde = Format(txtqtde.Text, "###.##0.000")
TBGravar!Corrida = txtcorrida.Text
TBGravar!Certificado = txtCertificado.Text
TBGravar!Classe = txtfamilia.Text
TBGravar!Un = txtun.Text
TBGravar!NF = txtnotafiscal.Text
TBGravar!Serie = Txt_serie
TBGravar!ID_cliente = txtid_cliente.Text
TBGravar!Cliente = txtcliente.Text
TBGravar!Tipodest_NFcons = Txt_tipodest
TBGravar!valor_unitario = Format(txtVlr_unit, "###.##0.00000")
TBGravar!Valor_total = Format(txtVlr_total, "###.##0.00")
TBGravar!local_armaz = cmbLocal_armaz

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

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Cmb_empresa.Visible = True
        Cmb_empresa1.Visible = False
        If Lista.Visible = True Then Lista.SetFocus
    Case 1:
        Cmb_empresa.Visible = False
        Cmb_empresa1.Visible = True
        Lista_movimentacao.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_PC_Change()
On Error GoTo tratar_erro

If Txt_qtde_PC <> "" Then
    VerifNumero = Txt_qtde_PC
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_PC = ""
        Txt_qtde_PC.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

ProcLimpaCamposItem

Exit Sub
tratar_erro:
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

Private Sub txtNotaFiscal_LostFocus()
On Error GoTo tratar_erro

If txtnotafiscal <> "" Then txtnotafiscal = FunTamanhoTextoZeroEsq(DS_RetornarNumeros(txtnotafiscal), 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpeso_Change()
On Error GoTo tratar_erro

If txtpeso.Text <> "" Then
    VerifNumero = txtpeso.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtpeso.Text = ""
        txtpeso.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Change()
On Error GoTo tratar_erro

If txtqtde.Text <> "" Then
    VerifNumero = txtqtde.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtqtde.Text = ""
        txtqtde.SetFocus
        Exit Sub
    End If
End If
ProcCalculo
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtde_LostFocus()
On Error GoTo tratar_erro

txtqtde.Text = Format(txtqtde.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlr_unit_Change()
On Error GoTo tratar_erro

If txtVlr_unit.Text <> "" Then
    VerifNumero = txtVlr_unit.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtVlr_unit.Text = ""
        txtVlr_unit.SetFocus
        Exit Sub
    End If
End If
ProcCalculo
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlr_unit_LostFocus()
On Error GoTo tratar_erro

txtVlr_unit.Text = Format(txtVlr_unit.Text, "###,##0.0000000000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculo()
On Error GoTo tratar_erro

ValorTotal = 0
valor = 0
Quant = 0
valor = IIf(txtVlr_unit = "", "0", txtVlr_unit)
Qtd = IIf(txtqtde = "", "0", txtqtde)

ValorTotal = (Qtd * valor)
txtVlr_total = Format(ValorTotal, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proccarregalocarm()
On Error GoTo tratar_erro

With cmbLocal_armaz
    .Clear
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select idemb_locarm from Estoque_Localarmazenamento where codinterno = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = False Then
        Do While TBAliquota.EOF = False
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Descricao from Estoque_Localarmazenamento_criar where id = " & TBAliquota!idemb_locarm, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If IsNull(TBAbrir!Descricao) = False Then .AddItem TBAbrir!Descricao
            End If
            TBAbrir.Close
            TBAliquota.MoveNext
        Loop
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Descricao from Estoque_Localarmazenamento_criar order by descricao", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!Descricao) = False Then .AddItem TBAbrir!Descricao
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    End If
    TBAliquota.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 3: ' ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirXML()
On Error GoTo tratar_erro

    CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    strCaminho = CommonDialog1.filename
    ImportarXML (strCaminho)
    PosicaoBase = 1
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: 'ProcAbrirXML
    Case 3: ProcGravar
    Case 4: ProcExcluir
    Case 5: ProcExcluirLote
    Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
