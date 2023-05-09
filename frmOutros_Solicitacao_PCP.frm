VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutros_Solicitacao_PCP 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Outros - Solicitação de produção"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   58
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
      SearchText      =   "Atualizando..."
      Value           =   0
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
      ItemData        =   "frmOutros_Solicitacao_PCP.frx":0000
      Left            =   270
      List            =   "frmOutros_Solicitacao_PCP.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1705
      Width           =   8670
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   17754
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
      TabCaption(0)   =   "Solicitação de produção"
      TabPicture(0)   =   "frmOutros_Solicitacao_PCP.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lista_solicitacao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdStatus"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Txt_ID_req"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Lista de produtos/serviços"
      TabPicture(1)   =   "frmOutros_Solicitacao_PCP.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Framelista"
      Tab(1).Control(1)=   "txtidcarteira"
      Tab(1).Control(2)=   "txtIDLista"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "USToolBar2"
      Tab(1).Control(5)=   "Lista"
      Tab(1).ControlCount=   6
      Begin VB.Frame Framelista 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2865
         Left            =   -74925
         TabIndex        =   72
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtRev 
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
            Left            =   2100
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   88
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Revisão do produto/item."
            Top             =   390
            Width           =   525
         End
         Begin VB.TextBox txtReferencia 
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
            Left            =   3420
            MaxLength       =   50
            TabIndex        =   24
            ToolTipText     =   "Código de referência."
            Top             =   390
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.ComboBox cmbRef 
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
            ItemData        =   "frmOutros_Solicitacao_PCP.frx":003C
            Left            =   3420
            List            =   "frmOutros_Solicitacao_PCP.frx":003E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   74
            ToolTipText     =   "Código de referência."
            Top             =   390
            Width           =   2865
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Criar novo produto/serviço"
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
            Height          =   525
            Left            =   11580
            TabIndex        =   73
            Top             =   180
            Width           =   3345
            Begin VB.CheckBox chkManual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cód. manual ?"
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
               Height          =   225
               Left            =   120
               TabIndex        =   26
               Top             =   270
               Width           =   1335
            End
            Begin VB.CheckBox chkAuto 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cód. automático ?"
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
               Height          =   225
               Left            =   1620
               TabIndex        =   27
               Top             =   270
               Width           =   1605
            End
         End
         Begin VB.TextBox cmbStatus 
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
            Left            =   6300
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   5175
         End
         Begin VB.CommandButton cmdEscolher_item 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2970
            Picture         =   "frmOutros_Solicitacao_PCP.frx":0040
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Localizar produtos/itens."
            Top             =   390
            Width           =   315
         End
         Begin VB.ComboBox cmbun 
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
            ItemData        =   "frmOutros_Solicitacao_PCP.frx":0142
            Left            =   10740
            List            =   "frmOutros_Solicitacao_PCP.frx":0144
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Unidade de estoque."
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox txtQE 
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
            Left            =   12480
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade em estoque."
            Top             =   2400
            Width           =   1245
         End
         Begin VB.TextBox txtQS 
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
            Left            =   13740
            TabIndex        =   36
            ToolTipText     =   "Quantidade solicitada."
            Top             =   2400
            Width           =   1185
         End
         Begin VB.TextBox txtCodInterno 
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
            Left            =   180
            TabIndex        =   21
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1935
         End
         Begin VB.TextBox txtDescricao 
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
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   990
            Width           =   13275
         End
         Begin VB.TextBox txtObs 
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
            Height          =   465
            Left            =   9600
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            ToolTipText     =   "Observações."
            Top             =   1620
            Width           =   5325
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
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Família."
            Top             =   2400
            Width           =   10545
         End
         Begin VB.TextBox Txt_descricao_comercial 
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
            Height          =   465
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Descrição comercial."
            Top             =   1620
            Width           =   9405
         End
         Begin VB.CommandButton cmdfiltrar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2640
            Picture         =   "frmOutros_Solicitacao_PCP.frx":0146
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Filtrar por código interno."
            Top             =   390
            Width           =   315
         End
         Begin VB.ComboBox Cmb_un_com 
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
            ItemData        =   "frmOutros_Solicitacao_PCP.frx":0561
            Left            =   11610
            List            =   "frmOutros_Solicitacao_PCP.frx":0563
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Unidade comercial."
            Top             =   2400
            Width           =   855
         End
         Begin MSMask.MaskEdBox txtprazo 
            Height          =   315
            Left            =   13470
            TabIndex        =   29
            ToolTipText     =   "Prazo de entrega."
            Top             =   990
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   4102
            TabIndex        =   86
            Top             =   180
            Width           =   1500
         End
         Begin VB.Image imgCalendario 
            Height          =   360
            Left            =   14595
            Picture         =   "frmOutros_Solicitacao_PCP.frx":0565
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   960
            Width           =   330
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Codigo interno*"
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
            Left            =   465
            TabIndex        =   85
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde. solicitada*"
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
            Left            =   13770
            TabIndex        =   84
            Top             =   2190
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   5317
            TabIndex        =   83
            Top             =   780
            Width           =   690
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   5182
            TabIndex        =   82
            Top             =   2190
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Un. est."
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
            Left            =   10875
            TabIndex        =   81
            Top             =   2190
            Width           =   585
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   12577
            TabIndex        =   80
            Top             =   2190
            Width           =   1050
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observação"
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
            Left            =   11677
            TabIndex        =   79
            Top             =   1410
            Width           =   870
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo"
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
            Left            =   13830
            TabIndex        =   78
            Top             =   780
            Width           =   405
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8610
            TabIndex        =   77
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descrição comercial"
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
            Left            =   4185
            TabIndex        =   76
            Top             =   1410
            Width           =   1395
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   11715
            TabIndex        =   75
            Top             =   2190
            Width           =   645
         End
      End
      Begin VB.TextBox txtidcarteira 
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
         Left            =   -73260
         TabIndex        =   71
         Top             =   7530
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtIDLista 
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
         Height          =   330
         Left            =   -74250
         TabIndex        =   70
         Text            =   "0"
         Top             =   7530
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operação da lista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -62055
         TabIndex        =   68
         Top             =   9450
         Width           =   2310
         Begin VB.ComboBox Cmb_opcao_lista_Item 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            ItemData        =   "frmOutros_Solicitacao_PCP.frx":09E8
            Left            =   180
            List            =   "frmOutros_Solicitacao_PCP.frx":09F2
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   170
            Width           =   1965
         End
      End
      Begin VB.TextBox Txt_ID_req 
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
         Height          =   330
         Left            =   2220
         TabIndex        =   64
         Text            =   "0"
         ToolTipText     =   "IDLista."
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   75
         TabIndex        =   60
         Top             =   9090
         Width           =   15195
         Begin VB.ComboBox Cmb_opcao_lista 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            ItemData        =   "frmOutros_Solicitacao_PCP.frx":0A07
            Left            =   7020
            List            =   "frmOutros_Solicitacao_PCP.frx":0A17
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   180
            Width           =   1965
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
            Left            =   9540
            TabIndex        =   14
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
            Left            =   2880
            TabIndex        =   12
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   18
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmOutros_Solicitacao_PCP.frx":0A42
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
            Left            =   11220
            TabIndex        =   17
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmOutros_Solicitacao_PCP.frx":41E6
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
            Left            =   10110
            TabIndex        =   15
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
            Left            =   10680
            TabIndex        =   16
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmOutros_Solicitacao_PCP.frx":7CEF
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
            Left            =   12300
            TabIndex        =   19
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmOutros_Solicitacao_PCP.frx":BDDE
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
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
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
            Left            =   5670
            TabIndex        =   65
            Top             =   240
            Width           =   1260
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
            Left            =   13050
            TabIndex        =   63
            Top             =   240
            Width           =   1095
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
            TabIndex        =   62
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label17 
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
            Left            =   2190
            TabIndex        =   61
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.CommandButton cmdStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3780
         Picture         =   "frmOutros_Solicitacao_PCP.frx":F66A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Verificar dados do cancelamento."
         Top             =   2310
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   37
         Top             =   1200
         Width           =   11820
         Begin VB.TextBox txtIDContato 
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
            Height          =   315
            Left            =   1800
            MaxLength       =   60
            MouseIcon       =   "frmOutros_Solicitacao_PCP.frx":F76C
            MousePointer    =   99  'Custom
            TabIndex        =   42
            ToolTipText     =   "Digite o nome para contato."
            Top             =   240
            Visible         =   0   'False
            Width           =   950
         End
         Begin VB.TextBox txtNomeContato 
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
            Height          =   315
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmOutros_Solicitacao_PCP.frx":FA76
            MousePointer    =   99  'Custom
            TabIndex        =   41
            ToolTipText     =   "Nome do contato."
            Top             =   240
            Width           =   9855
         End
         Begin VB.TextBox txtdepartamento 
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
            Height          =   315
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmOutros_Solicitacao_PCP.frx":FD80
            MousePointer    =   99  'Custom
            TabIndex        =   40
            ToolTipText     =   "Departamento do contato."
            Top             =   630
            Width           =   9855
         End
         Begin VB.TextBox txttelcontato 
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
            Height          =   315
            Left            =   1770
            MaxLength       =   40
            MouseIcon       =   "frmOutros_Solicitacao_PCP.frx":1008A
            MousePointer    =   99  'Custom
            TabIndex        =   39
            ToolTipText     =   "Ramal do contato."
            Top             =   1020
            Width           =   9855
         End
         Begin VB.TextBox TxtEmail_Contato 
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
            Height          =   330
            Left            =   1770
            MouseIcon       =   "frmOutros_Solicitacao_PCP.frx":10394
            MousePointer    =   99  'Custom
            TabIndex        =   38
            ToolTipText     =   "E-mail do cliente."
            Top             =   1440
            Width           =   9855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
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
            Left            =   600
            TabIndex        =   46
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do contato:"
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
            Left            =   405
            TabIndex        =   45
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ramal:"
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
            Index           =   13
            Left            =   1200
            TabIndex        =   44
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
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
            Index           =   16
            Left            =   1215
            TabIndex        =   43
            Top             =   1478
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2595
         Left            =   75
         TabIndex        =   47
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtRespValidacao 
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
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   990
            Width           =   3675
         End
         Begin VB.TextBox txtDtValidacao 
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
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   990
            Width           =   1755
         End
         Begin VB.TextBox txtData 
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
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox txtData_Autorizacao 
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
            Left            =   9570
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da aprovação."
            Top             =   990
            Width           =   1755
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   11340
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3675
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Número da solicitação de compra."
            Top             =   390
            Width           =   1425
         End
         Begin VB.TextBox txtAutorizado 
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
            Left            =   11340
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela aprovação."
            Top             =   990
            Width           =   3675
         End
         Begin VB.TextBox txtObservacao 
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
            Height          =   825
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            ToolTipText     =   "Observações da solicitação."
            Top             =   1620
            Width           =   14835
         End
         Begin VB.TextBox txtStatus 
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Status da solicitação."
            Top             =   990
            Width           =   3495
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável pela aprovação"
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
            Left            =   12142
            TabIndex        =   67
            Top             =   780
            Width           =   2070
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data/hora aprovação"
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
            Left            =   9675
            TabIndex        =   66
            Top             =   780
            Width           =   1545
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   4290
            TabIndex        =   56
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   12720
            TabIndex        =   55
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   10650
            TabIndex        =   54
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável pela validação"
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
            Left            =   6727
            TabIndex        =   53
            Top             =   780
            Width           =   1980
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data/hora validação"
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
            Left            =   4260
            TabIndex        =   52
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000001&
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -11580
            TabIndex        =   51
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº solicitação"
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
            Left            =   9022
            TabIndex        =   50
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observação"
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
            Left            =   7162
            TabIndex        =   49
            Top             =   1410
            Width           =   870
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1635
            TabIndex        =   48
            Top             =   780
            Width           =   585
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   57
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   15
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Filtrar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Filtrar (F2)"
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
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   42
         ButtonHeight2   =   21
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   84
         ButtonTop3      =   2
         ButtonWidth3    =   44
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   130
         ButtonTop4      =   2
         ButtonWidth4    =   45
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Relatório"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Relatório (F5)"
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
         ButtonLeft5     =   177
         ButtonTop5      =   2
         ButtonWidth5    =   60
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Anterior"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Registro anterior."
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
         ButtonLeft6     =   239
         ButtonTop6      =   2
         ButtonWidth6    =   55
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Próximo"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Próximo registro."
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
         ButtonLeft7     =   296
         ButtonTop7      =   2
         ButtonWidth7    =   55
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Status"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Status (F7)"
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   353
         ButtonTop8      =   2
         ButtonWidth8    =   45
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Copiar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Copiar (F8)"
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   400
         ButtonTop9      =   2
         ButtonWidth9    =   44
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Validação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Validar/Cancelar validação (F9)"
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   446
         ButtonTop10     =   2
         ButtonWidth10   =   62
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Aprovação"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Aprovar/cancelar aprovação (F10)"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   510
         ButtonTop11     =   2
         ButtonWidth11   =   69
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonAlignment12=   2
         ButtonType12    =   1
         ButtonStyle12   =   -1
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   -1
         ButtonLeft12    =   581
         ButtonTop12     =   4
         ButtonWidth12   =   2
         ButtonHeight12  =   54
         ButtonCaption13 =   "Ajuda"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Ajuda (F1)"
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   585
         ButtonTop13     =   2
         ButtonWidth13   =   41
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonCaption14 =   "Sair"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Sair (Esc)"
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft14    =   628
         ButtonTop14     =   2
         ButtonWidth14   =   30
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonKey15     =   "15"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState15   =   5
         ButtonLeft15    =   660
         ButtonTop15     =   2
         ButtonWidth15   =   24
         ButtonHeight15  =   24
         ButtonUseMaskColor15=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12090
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOutros_Solicitacao_PCP.frx":1069E
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   59
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   11
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Salvar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Salvar (F3)"
         ButtonKey2      =   "3"
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
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   44
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Excluir"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir (F4)"
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
         ButtonLeft3     =   86
         ButtonTop3      =   2
         ButtonWidth3    =   45
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   133
         ButtonTop4      =   2
         ButtonWidth4    =   60
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   195
         ButtonTop5      =   2
         ButtonWidth5    =   55
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   252
         ButtonTop6      =   2
         ButtonWidth6    =   55
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Status"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Status (F7)"
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
         ButtonLeft7     =   309
         ButtonTop7      =   2
         ButtonWidth7    =   45
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonAlignment8=   2
         ButtonType8     =   1
         ButtonStyle8    =   -1
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState8    =   -1
         ButtonLeft8     =   356
         ButtonTop8      =   4
         ButtonWidth8    =   2
         ButtonHeight8   =   54
         ButtonCaption9  =   "Ajuda"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Ajuda (F1)"
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   360
         ButtonTop9      =   2
         ButtonWidth9    =   41
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Sair"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Sair (Esc)"
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   403
         ButtonTop10     =   2
         ButtonWidth10   =   30
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   5
         ButtonLeft11    =   435
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   8970
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOutros_Solicitacao_PCP.frx":19057
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_solicitacao 
         Height          =   5145
         Left            =   75
         TabIndex        =   11
         Top             =   3930
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   9075
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
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   8475
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nº solicitação"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Data aprov."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Responsável aprov."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Validada"
            Object.Width           =   1499
         EndProperty
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   5235
         Left            =   -74940
         TabIndex        =   87
         Top             =   4200
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   9234
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   10590
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un. estoque"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un. comercial"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Quantidade"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "D"
            Text            =   "Prazo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2205
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOutros_Solicitacao_PCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_solicitacaoPCP      As Boolean 'OK
Dim Novo_solicitacaoPCP1        As Boolean 'OK
Public StrSql_solicitacaoPCP    As String 'OK
Dim TBLISTA_solicitacaoPCP      As ADODB.Recordset 'OK

Private Sub chkAuto_Click()
On Error GoTo tratar_erro

If chkAuto.Value = 1 Then
    chkManual.Value = 0
    Procliberacampos
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkManual_Click()
On Error GoTo tratar_erro

If chkManual.Value = 1 Then
    chkAuto.Value = 0
    Procliberacampos
    USMsgBox ("Informe o código interno do produto."), vbInformation, "CAPRIND v5.0"
    txtCodinterno.Text = ""
    txtCodinterno.SetFocus
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Procliberacampos()
On Error GoTo tratar_erro

With txtdescricao
    .Locked = False
    .TabStop = True
End With
With Txt_descricao_comercial
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
With cmbun
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com
    .Locked = False
    .TabStop = True
End With
If chkAuto.Value = 1 Or chkManual.Value = 1 Then
    cmbRef.Visible = False
    txtreferencia.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With txtdescricao
    .Locked = True
    .TabStop = False
End With
With Txt_descricao_comercial
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
With cmbun
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com
    .Locked = True
    .TabStop = False
End With
cmbRef.Visible = True
txtreferencia.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtNumero = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Requisicaotexto from Outros_SolicitacaoPCP order by Requisicaotexto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Requisicaotexto = '" & txtNumero & "'")
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        txtNumero = TBLISTA!Requisicaotexto
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from Outros_SolicitacaoPCP where Requisicaotexto = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaDados
        procCarregalista_Itens
    Else
        USMsgBox ("Fim dos cadastros de solicitação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_solicitacaoPCP1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Item_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar2
    If Cmb_opcao_lista_Item = "Excluir" Then
        .ButtonState(3) = 0
        .ButtonState(7) = 5
    Else
        .ButtonState(3) = 5
        .ButtonState(7) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Lista_solicitacao
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    Select Case Cmb_opcao_lista
        Case "Excluir"
            .ButtonState(4) = 0
            .ButtonState(8) = 5
            .ButtonState(10) = 5
            .ButtonState(11) = 5
        Case "Status"
            .ButtonState(4) = 5
            .ButtonState(8) = 0
            .ButtonState(10) = 5
            .ButtonState(11) = 5
        Case "Validação"
            .ButtonState(4) = 5
            .ButtonState(8) = 5
            .ButtonState(10) = 0
            .ButtonState(11) = 5
        Case "Aprovação"
            .ButtonState(4) = 5
            .ButtonState(8) = 5
            .ButtonState(10) = 5
            .ButtonState(11) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro
Dim ContAntigo As Integer 'OK

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtNumero = "" Then
    Acao = "copiar"
    NomeCampo = "a solicitação"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_solicitacaoPCP = True Then
    USMsgBox ("Salve a solicitação antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar esta solicitação?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    ContAntigo = Txt_ID_req
    Set TBSolicitacao = CreateObject("adodb.recordset")
    TBSolicitacao.Open "Select * from Outros_SolicitacaoPCP", Conexao, adOpenKeyset, adLockOptimistic
    TBSolicitacao.AddNew
    TBSolicitacao!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBSolicitacao!Responsavel = pubUsuario
    TBSolicitacao!Data = Format(Date, "dd/mm/yy")
    TBSolicitacao!Observacao = IIf(txtObservacao = "", Null, txtObservacao)
    TBSolicitacao!status = "ABERTA"
    ProcCriarNovoNumero
    TBSolicitacao!Requisicaotexto = a
    TBSolicitacao.Update
    Txt_ID_req = TBSolicitacao!ID
    TBSolicitacao.Close
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Vendas_carteira where ID_solicitacao = " & ContAntigo & " and Liberacao <> 'CANCELADO' order by codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Vendas_carteira", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!ID_solicitacao = Txt_ID_req
            TBGravar!Tipo = TBAbrir!Tipo
            TBGravar!Liberacao = "REQUISITADO"
            TBGravar!Unidade = TBAbrir!Unidade
            TBGravar!Unidade_com = TBAbrir!Unidade_com
            TBGravar!Familia = TBAbrir!Familia
            TBGravar!Descricao = TBAbrir!Descricao
            TBGravar!descricao_tecnica = TBAbrir!descricao_tecnica
            TBGravar!Qtde_produzir = TBAbrir!Qtde_produzir
            TBGravar!Desenho = TBAbrir!Desenho
            TBGravar!N_referencia = TBAbrir!N_referencia
            TBGravar!PrazoFinal = IIf(IsNull(TBAbrir!PrazoFinal), Null, Format(TBAbrir!PrazoFinal, "dd/mm/yy"))
            TBGravar!Observacoes = TBAbrir!Observacoes
            TBGravar.Update
            
            TBGravar.Close
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from Outros_SolicitacaoPCP where ID = " & Txt_ID_req, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = False Then
        ProcLimpaCampos
        ProcAbrir
    End If
    ProcCarregalista_Solicitacao (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista_solicitacao.ListItems.Count <> 0 Then
        Lista_solicitacao.SelectedItem = Lista_solicitacao.ListItems(CodigoLista)
        Lista_solicitacao.SetFocus
    End If
    USMsgBox ("Solicitação copiada com sucesso."), vbInformation, "CAPRIND v5.0"
    '============================================
    Modulo = "Outros/Solicitação de produção"
    Evento = "Novo"
    ID_documento = Txt_ID_req
    Documento = "Nº solicitação: " & txtNumero
    Documento1 = ""
    ProcGravaEvento
    '============================================
    Frame1.Enabled = True
    With txtObservacao
        .Locked = False
        .TabStop = True
    End With
    Novo_solicitacaoPCP = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEscolher_item_Click()
On Error GoTo tratar_erro
  
If txtResponsavel <> pubUsuario Then
    USMsgBox ("Só é possível modificação na solicitação pelo usuário " & txtResponsavel.Text & "."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    Exit Sub
End If
If txtStatus <> "ABERTA" Then
    USMsgBox ("Não é permitido alterar este produto/serviço, pois a solicitação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If cmbStatus <> "REQUISITADO" Then
    USMsgBox ("Não é permitido alterar este produto/serviço, pois o mesmo está " & cmbStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Outros_solicitacaoPCP = True
frmcompras_Req_EscolherProduto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoItem()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel <> pubUsuario Then
    USMsgBox ("Só é possível modificação na solicitação pelo usuário " & txtResponsavel.Text & "."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    Exit Sub
End If
If txtStatus <> "ABERTA" Then
    USMsgBox ("Não é permitido criar um novo produto/serviço, pois a solicitação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "solicitação", "produto/serviço", False) = False Then Exit Sub
ProcLimpaCampos_Itens
Novo_solicitacaoPCP1 = True
Framelista.Enabled = True
txtCodinterno.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtNumero = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Requisicaotexto from Outros_SolicitacaoPCP order by Requisicaotexto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Requisicaotexto = '" & txtNumero & "'")
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        txtNumero = TBLISTA!Requisicaotexto
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from Outros_SolicitacaoPCP where Requisicaotexto = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaDados
        procCarregalista_Itens
    Else
        USMsgBox ("Fim dos cadastros de solicitação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_solicitacaoPCP1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

ProcCarregaProduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaProduto()
On Error GoTo tratar_erro

If txtCodinterno <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto WHERE desenho = '" & txtCodinterno.Text & "' and Bloqueado = 'False' and (Compras = 'True' or Producao = 'True')", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtCodinterno = TBProduto!Desenho
        txtRev = TBProduto!RevDesenho
        txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        Txt_descricao_comercial = IIf(IsNull(TBProduto!descricaotecnica), "", TBProduto!descricaotecnica)
        If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun.Text = TBProduto!Unidade
        If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com.Text = TBProduto!Unidade_com
        If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbfamilia.Text = TBProduto!Classe
2:
        txtQE = Format(FunVerificaQtdeEstoque(txtCodinterno, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
        ProcCarregaComboCodRef cmbRef, "P.Desenho = '" & txtCodinterno & "'", 0, "", False, True
        
        ProcBloqueiaCampos
        With cmbun
            If TBProduto!Estoque = True Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
    Else
        txtdescricao.Text = ""
        Txt_descricao_comercial = ""
        cmbun.ListIndex = -1
        Cmb_un_com.ListIndex = -1
        cmbfamilia.ListIndex = -1
        txtQE.Text = "0,0000"
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos
    End If
Else
    If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a unidade ou familia desse registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_solicitacaoPCP.AbsolutePage <> 2 Then
    If TBLISTA_solicitacaoPCP.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_solicitacaoPCP.PageCount - 1)
    Else
        TBLISTA_solicitacaoPCP.AbsolutePage = TBLISTA_solicitacaoPCP.AbsolutePage - 2
        ProcExibePagina (TBLISTA_solicitacaoPCP.AbsolutePage)
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
    TBLISTA_solicitacaoPCP.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_solicitacaoPCP.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_solicitacaoPCP.AbsolutePage = 1
ProcExibePagina (TBLISTA_solicitacaoPCP.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_solicitacaoPCP.AbsolutePage <> -3 Then
    If TBLISTA_solicitacaoPCP.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_solicitacaoPCP.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_solicitacaoPCP.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_solicitacaoPCP.AbsolutePage = TBLISTA_solicitacaoPCP.PageCount
ProcExibePagina (TBLISTA_solicitacaoPCP.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Data_cancelamento, Resp_cancelamento, Motivo from Outros_SolicitacaoPCP where Requisicaotexto = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Dados do cancelamento: " & vbCrLf & "Data: " & Format(TBAbrir!Data_cancelamento, "dd/mm/yy") & " " & vbCrLf & "Responsável: " & TBAbrir!Resp_cancelamento & " " & vbCrLf & "Motivo: " & TBAbrir!motivo), vbInformation, "CAPRIND v5.0"
End If
TBAbrir.Close

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
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF8: ProcCopiar
            Case vbKeyF9: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista_solicitacao, "Outros/Solicitação de produção"
            Case vbKeyF10: If Cmb_opcao_lista = "Aprovação" Then ProcValidarRegistros Lista_solicitacao, "Outros/Solicitação de produção/Autorizar solicitação"
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoItem
            Case vbKeyF3: ProcSalvarItem
            Case vbKeyF4: If Cmb_opcao_lista_Item = "Excluir" Then ProcExcluirItem
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista_Item = "Status" Then ProcStatusItem
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Sub ProcAbrir()
On Error GoTo tratar_erro

ProcCarregaDados
ProcHabilitaCamposSolic
If TBCompras!status = "CANCELADA" Then
    Frame1.Enabled = False
Else
    If TBCompras!status = "LIBERADA" Then
        With txtObservacao
            .Locked = True
            .TabStop = False
        End With
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBCompras!ID_empresa) = False And TBCompras!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBCompras!ID_empresa
Txt_ID_req = TBCompras!ID
Caption = "Outros - Solicitação de produção - (Solicitação : " & IIf(IsNull(TBCompras!Requisicaotexto), "", TBCompras!Requisicaotexto) & ")"
txtNumero = IIf(IsNull(TBCompras!Requisicaotexto), "", TBCompras!Requisicaotexto)
If TBCompras!status = "LIBERADA" Then
    txtAutorizado.Text = IIf(IsNull(TBCompras!Autorizado), "", (TBCompras!Autorizado))
    txtData_Autorizacao = IIf(IsNull(TBCompras!Data_autorizacao), "", TBCompras!Data_autorizacao)
End If
txtDtValidacao = IIf(IsNull(TBCompras!DtValidacao), "", (TBCompras!DtValidacao))
txtRespValidacao = IIf(IsNull(TBCompras!RespValidacao), "", TBCompras!RespValidacao)
txtResponsavel.Text = IIf(IsNull(TBCompras!Responsavel), "", (TBCompras!Responsavel))
txtData.Text = IIf(IsNull(TBCompras!Data), "", (Format(TBCompras!Data, "dd/mm/yy")))
txtObservacao.Text = IIf(IsNull(TBCompras!Observacao), "", TBCompras!Observacao)
txtStatus.Text = IIf(IsNull(TBCompras!status), "", TBCompras!status)
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCarregalista_Itens()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Vendas_carteira Where ID_solicitacao = " & Txt_ID_req & " order by codigo desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Qtde_produzir), "", Format(TBLISTA!Qtde_produzir, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Familia), "", TBLISTA!Familia)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Liberacao), "", TBLISTA!Liberacao)
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

Sub ProcLimpaCampos_Itens()
On Error GoTo tratar_erro

TXTIDLista = 0
txtCodinterno.Text = ""
txtRev = 0
chkAuto.Value = 0
chkManual.Value = 0
cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
txtQE.Text = ""
txtQS.Text = ""
txtdescricao.Text = ""
Txt_descricao_comercial = ""
txtprazo.Text = "__/__/____"
txtObs = ""
cmbStatus = "REQUISITADO"
cmbRef.Clear
txtreferencia = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID_req = 0
txtNumero.Text = ""
txtStatus.Text = "ABERTA"
txtResponsavel.Text = pubUsuario
txtData.Text = Format(Date, "dd/mm/yy")
txtData_Autorizacao.Text = ""
txtAutorizado.Text = ""
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtObservacao.Text = ""
CodigoLista = 0
Caption = "Outros - Solicitação de produção"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadados_Itens()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Tipo from Projproduto where Desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then If TBProduto!Tipo = "P" Or TBProduto!Tipo = "I" Then TBCompras!Tipo = "P" Else TBCompras!Tipo = "S"
TBProduto.Close

TBCompras!ID_solicitacao = Txt_ID_req
If cmbStatus = "REQUISITADO" Then TBCompras!Liberacao = "REQUISITADO"
TBCompras!Unidade = cmbun.Text
TBCompras!Unidade_com = Cmb_un_com.Text
TBCompras!Familia = cmbfamilia.Text
TBCompras!descricao_tecnica = txtdescricao.Text
TBCompras!Descricao = Txt_descricao_comercial
TBCompras!Qtde_produzir = txtQS
TBCompras!Desenho = txtCodinterno.Text
TBCompras!Rev_codinterno = txtRev
TBCompras!N_referencia = IIf(cmbRef.Text = "", Null, cmbRef.Text)
If txtprazo.Text <> "__/__/____" Then TBCompras!PrazoFinal = txtprazo.Text Else TBCompras!PrazoFinal = Null
TBCompras!Observacoes = txtObs
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaCamposSolic()
On Error GoTo tratar_erro

Frame1.Enabled = True
With txtObservacao
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarItem()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel <> pubUsuario Then
    USMsgBox ("Só é possível modificação na solicitação pelo usuário " & txtResponsavel.Text & "."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "solicitação", "o produto/serviço", False) = False Then Exit Sub
If txtStatus <> "ABERTA" Then
    USMsgBox ("Não é permitido alterar este produto/serviço, pois a solicitação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If cmbStatus <> "REQUISITADO" Then
    USMsgBox ("Não é permitido alterar este produto/serviço, pois o mesmo está cancelado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Framelista.Enabled = False Then
    ProcVerificaSalvar
    cmdnovo_lista.SetFocus
    Exit Sub
End If
Acao = "salvar"
If chkAuto.Value = 0 And txtCodinterno.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodinterno.SetFocus
    Exit Sub
End If
If txtdescricao.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
If txtprazo <> "__/__/____" Then
    If IsDate(txtprazo) = False Then
        USMsgBox ("A data foi digitada incorretamente."), vbExclamation, "CAPRIND v5.0"
        txtprazo.SetFocus
        Exit Sub
    End If
End If
If Txt_descricao_comercial = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    Txt_descricao_comercial.SetFocus
    Exit Sub
End If
If cmbfamilia.Text = "" Then
    NomeCampo = "a familia"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
If cmbun.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun.SetFocus
    Exit Sub
End If
If Cmb_un_com.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
Qtd = IIf(txtQS = "", 0, txtQS)
If Qtd = 0 Then
    NomeCampo = "a quantidade solicitada"
    ProcVerificaAcao
    txtQS.SetFocus
    Exit Sub
End If

If Novo_solicitacaoPCP1 = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Vendas_carteira where desenho = '" & txtCodinterno.Text & "' And Liberacao = 'REQUISITADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        If USMsgBox("Já existe uma solicitação em aberto para este produto/serviço, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            TBProduto.Close
            Exit Sub
        End If
    End If
    TBProduto.Close
End If
If chkAuto.Value = 1 Then
    ProcNovoProdutoAuto
    If txtreferencia <> "" Then
        cmbRef.AddItem txtreferencia
        cmbRef = txtreferencia
    End If
    chkAuto.Value = 0
End If
If chkManual.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select desenho from projproduto where desenho = '" & txtCodinterno.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto/serviço cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtCodinterno.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoManual
    If txtreferencia <> "" Then
        cmbRef.AddItem txtreferencia
        cmbRef = txtreferencia
    End If
    chkManual.Value = 0
End If

'Verifica se o produto está cadastrado
If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select desenho from projproduto where desenho = '" & txtCodinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then
        USMsgBox ("Não é permitido salvar este produto/serviço, pois o mesmo não está cadastrado."), vbExclamation, "CAPRIND v5.0"
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Vendas_carteira WHERE Codigo = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then TBCompras.AddNew
ProcEnviadados_Itens
TBCompras.Update
TXTIDLista = TBCompras!CODIGO
TBCompras.Close
procCarregalista_Itens
If Novo_solicitacaoPCP1 = True Then
    USMsgBox ("Novo produto/serviço cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto/serviço"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto/serviço"
    If CodigoLista1 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista1)
        Lista.SetFocus
    End If
End If
Novo_solicitacaoPCP1 = False
'==================================
Modulo = "Outros/Solicitação de produção"
ID_documento = TXTIDLista
Documento = "Nº solicitação: " & txtNumero
Documento1 = "Cód. interno: " & txtCodinterno
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoAuto()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtCodinterno = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtCodinterno, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, False, True, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtCodinterno = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtCodinterno, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, False, True, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirItem()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Vendas_carteira where Codigo = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "Outros/Solicitação de produção"
            Evento = "Excluir produto/serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº solicitação: " & txtNumero
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Itens
    procCarregalista_Itens
    Framelista.Enabled = False
    Novo_solicitacaoPCP1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 15, True
ProcCarregaToolBar2 Me, 15195, 11, True

Formulario = "Outros/Solicitação de produção"
Direitos
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista = "Validação"
Cmb_opcao_lista_Item = "Excluir"
SSTab1.Tab = 0
ProcCarregaFamiliaUN
ProcCarregaComboEmpresa Cmb_empresa, False

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Outros/Solicitação de produção"
Direitos
ProcCarregaFamiliaUN
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Familia, Unidade, Unidade_com from Vendas_carteira where codigo = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
    If IsNull(TBAbrir!Unidade) = False And TBAbrir!Unidade <> "" Then cmbun = TBAbrir!Unidade
    If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFamiliaUN()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamilia, "Familia <> 'Null' and (Vendas = 'True' or Fabricacao = 'True')", False
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False
ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmOutros_Solicitacao_PCP_abrir.Show 1

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
Outros_solicitacaoPCP = True
Estoque_recebimento = False
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_solicitacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) solicitação(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Outros_SolicitacaoPCP where ID = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Vendas_carteira where ID_solicitacao = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Outros/Solicitação de produção"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº solicitação: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) solicitação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Solicitação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcLimparTudo
    ProcCarregalista_Solicitacao (1)
    Novo_solicitacaoPCP = False
    Frame1.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_solicitacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o status desta(s) solicitação(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBCotacao = CreateObject("adodb.recordset")
            TBCotacao.Open "Select * from Outros_SolicitacaoPCP where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCotacao.EOF = False Then
                If TBCotacao!status = "CANCELADA" Then
                    Conexao.Execute "Update Vendas_carteira Set Liberacao = 'REQUISITADO' where ID_Solicitacao = " & .ListItems(InitFor)
                    TBCotacao!status = "ABERTA"
                    TBCotacao!Resp_cancelamento = ""
                    TBCotacao!Data_cancelamento = Null
                    TBCotacao!motivo = ""
                    TBCotacao.Update
                    '==================================
                    Modulo = "Outros/Solicitação de produção"
                    Evento = "Alterar status"
                    ID_documento = .ListItems(InitFor)
                    Documento = "Nº solicitação: " & .ListItems(InitFor).SubItems(2)
                    Documento1 = ""
                    ProcGravaEvento
                    '==================================
                Else
                    IDlista = .ListItems(InitFor)
                    Familiatext = .ListItems(InitFor).SubItems(2)
                    Outros_solicitacaoPCP = True
                    frmCompras_Requisicao_cancelar.Show 1
                End If
                TBCotacao.Close
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) solicitação(ões) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Solicitação(ões) alterada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregalista_Solicitacao (1)
    Novo_solicitacaoPCP = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatusItem()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o status deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select Liberacao from Vendas_carteira WHERE codigo = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = False Then
                If IsNull(TBCompras!Liberacao) = False Then
                    If TBCompras!Liberacao = "REQUISITADO" Then TBCompras!Liberacao = "CANCELADO" Else TBCompras!Liberacao = "REQUISITADO"
                    TBCompras.Update
                    
                    '==================================
                    Modulo = "Outros/Solicitação de produção"
                    Evento = "Alterar status"
                    ID_documento = .ListItems(InitFor)
                    Documento = "Nº solicitação: " & txtNumero
                    Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(2)
                    ProcGravaEvento
                    '==================================
                End If
            End If
            TBCompras.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("produto(s)/serviço(s) alterada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Itens
    procCarregalista_Itens
    Framelista.Enabled = False
    Novo_solicitacaoPCP1 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro

If txtNumero.Text <> "" Then
    NomeRel = "Outros_SolicitacaoPCP.rpt"
    ProcImprimirRel "{Outros_SolicitacaoPCP.ID} = " & Txt_ID_req & " and {Vendas_carteira.liberacao} <> 'CANCELADO'", ""
Else
    USMsgBox ("Informe a solicitação antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
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
ProcLimpaCampos
Frame1.Enabled = True
txtObservacao.SetFocus
Novo_solicitacaoPCP = True
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Framelista.Enabled = False
ProcLimpaCampos_Itens
Lista.ListItems.Clear
Novo_solicitacaoPCP1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_solicitacaoPCP = True Then
    If USMsgBox("A solicitação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_solicitacaoPCP = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_solicitacaoPCP1 = True Then
    If USMsgBox("O produto/serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarItem
        If Novo_solicitacaoPCP1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_solicitacaoPCP = False
Novo_solicitacaoPCP1 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtStatus = "CANCELADA" Then
    USMsgBox ("Não é permitido alterar solicitação cancelada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtAutorizado.Text <> "" Then
    USMsgBox ("Não é permitido alterar solicitação validada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Outros_SolicitacaoPCP where ID = " & Txt_ID_req, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then
    TBCompras.AddNew
    TBCompras!status = "ABERTA"
    ProcCriarNovoNumero
    TBCompras!Requisicaotexto = a
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "a mesma", "solicitação", False) = False Then Exit Sub
End If
TBCompras!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBCompras!Data = IIf(txtData = "", Date, txtData)
TBCompras!Responsavel = IIf(txtResponsavel = "", pubUsuario, txtResponsavel)
TBCompras!Observacao = IIf(txtObservacao = "", Null, txtObservacao)
TBCompras.Update
ProcAbrir
TBCompras.Close

If Novo_solicitacaoPCP = True Then
    USMsgBox ("Nova solicitação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_solicitacaoPCP = "Select * from Outros_SolicitacaoPCP where Requisicaotexto = '" & txtNumero & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    ProcCarregalista_Solicitacao (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregalista_Solicitacao (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista_solicitacao.ListItems.Count <> 0 Then
        Lista_solicitacao.SelectedItem = Lista_solicitacao.ListItems(CodigoLista)
        Lista_solicitacao.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Solicitação de produção"
ID_documento = Txt_ID_req
Documento = "Nº solicitação: " & txtNumero
Documento1 = ""
ProcGravaEvento
'==================================
Novo_solicitacaoPCP = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista_Item = "Excluir" Then
                    If txtDtValidacao <> "" Then GoTo Proximo
                    If txtStatus <> "ABERTA" Then GoTo Proximo
                    If .ListItems.Item(InitFor).SubItems(8) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(8) <> "CANCELADO" Then GoTo Proximo
                Else
                    If txtStatus = "CANCELADA" Then GoTo Proximo
                    If .ListItems.Item(InitFor).SubItems(8) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(8) <> "CANCELADO" Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista_Item = "Excluir" Then
                If txtDtValidacao <> "" Then
                    USMsgBox ("Não é permitido excluir este produto/serviço, pois a solicitação já foi validada."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If txtStatus <> "ABERTA" Then
                    USMsgBox ("Não é permitido excluir este produto/serviço, pois a solicitação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If .ListItems.Item(InitFor).SubItems(8) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(8) <> "CANCELADO" Then
                    USMsgBox ("Não é permitido excluir este produto/serviço, pois o mesmo está " & .ListItems.Item(InitFor).SubItems(11) & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
            Else
                If txtStatus = "CANCELADA" Then
                    USMsgBox ("Não é permitido alterar o status deste produto/serviço, pois a solicitação está cancelada."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                If .ListItems.Item(InitFor).SubItems(8) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(8) <> "CANCELADO" Then
                    USMsgBox ("Não é permitido alterar o status deste produto/serviço, pois o mesmo está " & .ListItems.Item(InitFor).SubItems(11) & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
txtCodinterno.Locked = False
txtCodinterno.TabStop = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_carteira where codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_Itens
    TXTIDLista = Lista.SelectedItem
    procCarregaDados_Itens
    
    CodigoLista1 = Lista.SelectedItem.index
    Novo_solicitacaoPCP1 = False
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto WHERE desenho = '" & txtCodinterno.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        ProcBloqueiaCampos
        With cmbun
            If TBProduto!Estoque = True Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
    Else
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_solicitacao
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Aprovação" Then
                    If .ListItems.Item(InitFor).SubItems(7) = "CANCELADA" Then GoTo Proximo
                    If .ListItems.Item(InitFor).SubItems(8) = "Não" Then GoTo Proximo
                ElseIf Cmb_opcao_lista = "Validação" Then
                        If .ListItems.Item(InitFor).SubItems(6) <> "" Then GoTo Proximo
                    Else
                        If .ListItems.Item(InitFor).SubItems(8) = "Sim" And Cmb_opcao_lista = "Excluir" Then GoTo Proximo
                        If .ListItems.Item(InitFor).SubItems(7) <> "ABERTA" And .ListItems.Item(InitFor).SubItems(7) <> "CANCELADA" Then GoTo Proximo
                End If

                If .ListItems.Item(InitFor).SubItems(6) <> "" Then
                    Set TBAcessos = CreateObject("adodb.recordset")
                    TBAcessos.Open "Select ID_Solicitacao from Vendas_carteira where ID_Solicitacao = " & .ListItems.Item(InitFor) & " and Liberacao <> 'REQUISITADO' and Liberacao <> 'CANCELADO' and Liberacao <> 'NÃO APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAcessos.EOF = False Then
                        GoTo Proximo
                    End If
                    TBAcessos.Close
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_solicitacao, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Select Case Cmb_opcao_lista
    Case "Excluir": TextoLista = "excluir"
    Case "Status": TextoLista = "alterar status"
    Case "Aprovação": TextoLista = "cancelar aprovação"
    Case "Validação": TextoLista = "cancelar validação"
End Select

With Lista_solicitacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Aprovação" Then
                If .ListItems.Item(InitFor).SubItems(8) = "Não" Then
                    USMsgBox ("Não é permitido aprovar solicitação, pois a mesma ainda não foi validada."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If .ListItems.Item(InitFor).SubItems(7) = "CANCELADA" Then
                    USMsgBox ("Não é permitido autorizar/cancelar aprovação, pois o status da solicitação está cancelada."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            ElseIf Cmb_opcao_lista = "Validação" Then
                    If .ListItems.Item(InitFor).SubItems(6) <> "" Then
                        USMsgBox ("Não é permitido cancelar validação, pois a solicitação já foi aprovada."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                Else
                    If .ListItems.Item(InitFor).SubItems(8) = "Sim" And Cmb_opcao_lista = "Excluir" Then
                        USMsgBox ("Não é permitido excluir solicitação, pois a mesma já foi validada."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If .ListItems.Item(InitFor).SubItems(7) <> "ABERTA" And .ListItems.Item(InitFor).SubItems(7) <> "CANCELADA" Then
                        USMsgBox ("Não é permitido " & TextoLista & ", pois o status da solicitação está " & .ListItems.Item(InitFor).SubItems(7) & "."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
            End If
            
            If .ListItems.Item(InitFor).SubItems(6) <> "" Then
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select ID_Solicitacao from Vendas_carteira where ID_Solicitacao = " & .ListItems.Item(InitFor) & " and Liberacao <> 'REQUISITADO' and Liberacao <> 'CANCELADO' and Liberacao <> 'NÃO APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    USMsgBox ("Não é permitido " & TextoLista & ", pois os produtos/serviços já sofreram alguma alteração."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBAcessos.Close
                    Exit Sub
                End If
                TBAcessos.Close
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_solicitacao.ListItems.Count = 0 Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Outros_SolicitacaoPCP where ID = " & Lista_solicitacao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    ProcLimpaCampos
    ProcAbrir
    CodigoLista = Lista_solicitacao.SelectedItem.index
    Novo_solicitacaoPCP = False
End If
TBCompras.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtNumero = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
PBLista.Width = Frame2.Width
Select Case SSTab1.Tab
    Case 0:
        Cmb_empresa.Visible = True
        If Lista_solicitacao.Visible = True Then Lista_solicitacao.SetFocus
    Case 1:
        PBLista.Width = Frame2.Width - (Frame4.Width + 100)
        Cmb_empresa.Visible = False
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos Else ProcBloqueiaCampos
        If Novo_solicitacaoPCP = True Then
            USMsgBox ("Salve a solicitação antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            Exit Sub
        End If
        Lista.SetFocus
        procCarregalista_Itens
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodInterno_Change()
On Error GoTo tratar_erro

If chkAuto.Value = 0 And chkManual.Value = 0 Then
    txtdescricao = ""
    Txt_descricao_comercial = ""
    cmbun.ListIndex = -1
    Cmb_un_com.ListIndex = -1
    cmbfamilia.ListIndex = -1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriarNovoNumero()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Requisicaotexto from Outros_SolicitacaoPCP where Year (Data) = '" & Year(Date) & "' order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    Numero = Left(TBAbrir!Requisicaotexto, Len(TBAbrir!Requisicaotexto) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBAbrir.Close

a = Numero
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: a = "SPR-0000" & Numero & "/" & Ano
    Case 2: a = "SPR-000" & Numero & "/" & Ano
    Case 3: a = "SPR-00" & Numero & "/" & Ano
    Case 4: a = "SPR-0" & Numero & "/" & Ano
    Case 5: a = "SPR-" & Numero & "/" & Ano
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregalista_Solicitacao(Pagina As Integer)
On Error GoTo tratar_erro

Lista_solicitacao.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_solicitacaoPCP = "" Then Exit Sub
Set TBLISTA_solicitacaoPCP = CreateObject("adodb.recordset")
TBLISTA_solicitacaoPCP.Open StrSql_solicitacaoPCP, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_solicitacaoPCP.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista_solicitacao.ListItems.Clear
TBLISTA_solicitacaoPCP.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_solicitacaoPCP.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_solicitacaoPCP.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_solicitacaoPCP.RecordCount - IIf(Pagina > 1, (TBLISTA_solicitacaoPCP.PageSize * (Pagina - 1)), 0), TBLISTA_solicitacaoPCP.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_solicitacaoPCP.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_solicitacao.ListItems
        .Add , , TBLISTA_solicitacaoPCP!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Empresa from Empresa where Codigo = " & IIf(IsNull(TBLISTA_solicitacaoPCP!ID_empresa), 0, TBLISTA_solicitacaoPCP!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
        End If
        TBAbrir.Close
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_solicitacaoPCP!Requisicaotexto), "", TBLISTA_solicitacaoPCP!Requisicaotexto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_solicitacaoPCP!Data), "", Format(TBLISTA_solicitacaoPCP!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_solicitacaoPCP!Responsavel), "", TBLISTA_solicitacaoPCP!Responsavel)
        If TBLISTA_solicitacaoPCP!Data_autorizacao <> "" Then
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_solicitacaoPCP!Data_autorizacao), "", Format(TBLISTA_solicitacaoPCP!Data_autorizacao, "dd/mm/yy"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_solicitacaoPCP!Autorizado), "", TBLISTA_solicitacaoPCP!Autorizado)
        End If
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_solicitacaoPCP!status), "", TBLISTA_solicitacaoPCP!status)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_solicitacaoPCP!DtValidacao), "Não", "Sim")
    End With
    TBLISTA_solicitacaoPCP.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_solicitacaoPCP.RecordCount
If TBLISTA_solicitacaoPCP.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_solicitacaoPCP.PageCount
ElseIf TBLISTA_solicitacaoPCP.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_solicitacaoPCP.PageCount & " de: " & TBLISTA_solicitacaoPCP.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_solicitacaoPCP.AbsolutePage - 1 & " de: " & TBLISTA_solicitacaoPCP.PageCount
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

Private Sub txtprazo_LostFocus()
On Error GoTo tratar_erro

If txtprazo.Text <> "__/__/____" Then
    VerifData = txtprazo.Text
    ProcVerificaData
    If VerifData = False Then
        txtprazo.Text = "__/__/____"
        txtprazo.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_Change()
On Error GoTo tratar_erro

If txtQS <> "" Then
    VerifNumero = txtQS
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQS = ""
        txtQS.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_LostFocus()
On Error GoTo tratar_erro

txtQS.Text = Format(txtQS.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoManual()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtCodinterno = FunCriaNovoProdServ(True, "", txtCodinterno, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, False, True, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtCodinterno = FunCriaNovoProdServ(True, "", txtCodinterno, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, False, True, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtstatus_Change()
On Error GoTo tratar_erro

With txtStatus
    If .Text = "CANCELADA" Then
        cmdstatus.Visible = True
        .Width = cmdstatus.Left - txtStatus.Left
    Else
        cmdstatus.Visible = False
        .Width = txtDtValidacao.Left - txtStatus.Left
    End If
    Label23.Width = (.Width / 2) + (Label23.Width / 2)
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcStatus
    Case 9: ProcCopiar
    Case 10: ProcValidarRegistros Lista_solicitacao, "Outros/Solicitação de produção"
    Case 11: ProcValidarRegistros Lista_solicitacao, "Outros/Solicitação de produção/Autorizar solicitação"
    'Case 13: ProcAjuda
    Case 14: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoItem
    Case 2: ProcSalvarItem
    Case 3: ProcExcluirItem
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcStatusItem
    'Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCarregaDados_Itens()
On Error GoTo tratar_erro

txtObs.Text = IIf(IsNull(TBAbrir!Observacoes), "", TBAbrir!Observacoes)
txtCodinterno.Text = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
ProcCarregaComboCodRef cmbRef, "P.Desenho = '" & txtCodinterno & "'", 0, "", False, True
If IsNull(TBAbrir!N_referencia) = False And TBAbrir!N_referencia <> "" Then cmbRef = TBAbrir!N_referencia Else cmbRef.ListIndex = -1
1:
    txtdescricao.Text = IIf(IsNull(TBAbrir!descricao_tecnica), "", TBAbrir!descricao_tecnica)
    Txt_descricao_comercial = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    txtprazo.Text = IIf(IsNull(TBAbrir!PrazoFinal), "__/__/____", Format(TBAbrir!PrazoFinal, "dd/mm/yyyy"))
    If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
    If IsNull(TBAbrir!Unidade) = False And TBAbrir!Unidade <> "" Then cmbun = TBAbrir!Unidade
    If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
    txtQE = Format(FunVerificaQtdeEstoque(TBAbrir!Desenho, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
    txtQS = IIf(IsNull(TBAbrir!Qtde_produzir), "", Format(TBAbrir!Qtde_produzir, "###,##0.0000"))
    
    If IsNull(TBAbrir!Liberacao) = False Then
        cmbStatus.Text = Lista.SelectedItem.ListSubItems(8)
        If cmbStatus <> "REQUISITADO" Then Framelista.Enabled = False Else Framelista.Enabled = True
    Else
        Framelista.Enabled = True
    End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o código de referência deste produto/serviço."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
