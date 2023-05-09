VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_Tele_Clientes 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Telemarketing"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmvendas_tele_clientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
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
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin MSComctlLib.ListView Lista_historico 
      Height          =   6105
      Left            =   75
      TabIndex        =   39
      Top             =   2760
      Visible         =   0   'False
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   10769
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   5569
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Contato"
         Object.Width           =   5569
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6165
      Left            =   90
      TabIndex        =   20
      Top             =   2910
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10874
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   512
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Razão social"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cidade"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "UF"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "D"
         Text            =   "Últ. contato"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "D"
         Text            =   "Prox. contato"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Histórico"
         Object.Width           =   8493
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "IDcontato"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3175
      EndProperty
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "frmvendas_tele_clientes.frx":212A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "PBLista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "USToolBar1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Chk_ultimo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Chk_proximo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "frmvendas_tele_clientes.frx":2146
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PBlista_hist"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "txtCodigo"
      Tab(1).Control(3)=   "MSComm1"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Frame4"
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(7)=   "USToolBar2"
      Tab(1).ControlCount=   8
      Begin VB.CheckBox Chk_proximo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Próximo contato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9600
         TabIndex        =   17
         Top             =   2580
         Width           =   1725
      End
      Begin VB.CheckBox Chk_ultimo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Último contato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9600
         TabIndex        =   16
         Top             =   2340
         Width           =   1575
      End
      Begin DrawSuite2022.USProgressBar PBlista_hist 
         Height          =   255
         Left            =   -74925
         TabIndex        =   96
         Top             =   8880
         Width           =   7905
         _ExtentX        =   13944
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
      Begin VB.Frame Frame3 
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
         Height          =   825
         Left            =   -72240
         TabIndex        =   62
         Top             =   9150
         Width           =   12510
         Begin VB.CommandButton cmdcontato_visita 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   5085
            Picture         =   "frmvendas_tele_clientes.frx":2162
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Localizar contatos."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_departamento_contato_visita 
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
            Left            =   5490
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Departamento."
            Top             =   390
            Width           =   3450
         End
         Begin VB.TextBox Txt_contato_visita 
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
            MaxLength       =   40
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "Nome do contato."
            Top             =   390
            Width           =   4890
         End
         Begin MSComCtl2.DTPicker mskhora_inicio 
            Height          =   315
            Left            =   10320
            TabIndex        =   58
            ToolTipText     =   "Hora início da visita."
            Top             =   390
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   198049794
            CurrentDate     =   39055
         End
         Begin MSComCtl2.DTPicker mskhora_fim 
            Height          =   315
            Left            =   11340
            TabIndex        =   59
            ToolTipText     =   "Hora final da visita."
            Top             =   390
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   198049794
            CurrentDate     =   39055
         End
         Begin MSComCtl2.DTPicker txtDataVisita 
            Height          =   315
            Left            =   8955
            TabIndex        =   57
            ToolTipText     =   "Data da visita."
            Top             =   390
            Width           =   1350
            _ExtentX        =   2381
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
            Format          =   198049793
            CurrentDate     =   39057
         End
         Begin VB.Label lblcontato 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
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
            Left            =   6698
            TabIndex        =   67
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lbldata 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   9458
            TabIndex        =   66
            Top             =   180
            Width           =   345
         End
         Begin VB.Label lblhora_fim 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fim"
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
            Left            =   11722
            TabIndex        =   65
            Top             =   180
            Width           =   240
         End
         Begin VB.Label lblhora_inicio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Início"
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
            Left            =   10635
            TabIndex        =   64
            Top             =   180
            Width           =   375
         End
         Begin VB.Label lblcontato 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   2333
            TabIndex        =   63
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   495
         Left            =   -70350
         TabIndex        =   61
         Text            =   "0"
         Top             =   4980
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   -61590
         Top             =   510
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.Frame Frame8 
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
         Left            =   75
         TabIndex        =   86
         Top             =   1320
         Width           =   1305
         Begin VB.OptionButton optFisica 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Física"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   1
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optJuridica 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Jurídica"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   0
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
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
         Left            =   1395
         TabIndex        =   83
         Top             =   1320
         Width           =   13875
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   2790
            TabIndex        =   105
            Top             =   210
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
               TabIndex        =   10
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
               TabIndex        =   8
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
               TabIndex        =   9
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
               TabIndex        =   11
               Top             =   180
               Width           =   705
            End
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
            Left            =   7650
            TabIndex        =   3
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6045
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
            ItemData        =   "frmvendas_tele_clientes.frx":2264
            Left            =   180
            List            =   "frmvendas_tele_clientes.frx":2280
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2535
         End
         Begin VB.ComboBox cmbstatus 
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
            ItemData        =   "frmvendas_tele_clientes.frx":22DE
            Left            =   7650
            List            =   "frmvendas_tele_clientes.frx":22EB
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6045
         End
         Begin VB.ComboBox cmbFamilia 
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
            Left            =   7650
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6045
         End
         Begin MSMask.MaskEdBox txtCpf 
            Height          =   315
            Left            =   7650
            TabIndex        =   6
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Visible         =   0   'False
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtcnpj 
            Height          =   315
            Left            =   7650
            TabIndex        =   7
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
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
            Left            =   9937
            TabIndex        =   85
            Top             =   180
            Width           =   1470
         End
         Begin VB.Label Label8 
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
            Left            =   1027
            TabIndex        =   84
            Top             =   180
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   1440
         Left            =   -74925
         TabIndex        =   79
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton imgemail 
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
            Left            =   14685
            Picture         =   "frmvendas_tele_clientes.frx":230D
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Enviar e-mail para o cliente (F8)"
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   7380
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "E-mail."
            Top             =   990
            Width           =   7320
         End
         Begin VB.TextBox txtfax 
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
            Left            =   5940
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Fax."
            Top             =   990
            Width           =   1420
         End
         Begin VB.TextBox txttel01 
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
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Telefone."
            Top             =   990
            Width           =   1420
         End
         Begin VB.TextBox txttel02 
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
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Telefone."
            Top             =   990
            Width           =   1420
         End
         Begin VB.TextBox txttel03 
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
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Telefone."
            Top             =   990
            Width           =   1420
         End
         Begin VB.TextBox txttel04 
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
            Left            =   4505
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Telefone."
            Top             =   990
            Width           =   1420
         End
         Begin VB.TextBox txtidcliente 
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
            TabIndex        =   28
            ToolTipText     =   "Código do cliente."
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtUF 
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
            Left            =   14550
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "UF."
            Top             =   390
            Width           =   450
         End
         Begin VB.TextBox txtnomerazao 
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
            Left            =   1110
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Razão social."
            Top             =   390
            Width           =   9585
         End
         Begin VB.TextBox txtcidade 
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
            Left            =   10710
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Cidade."
            Top             =   390
            Width           =   3820
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone 2"
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
            Index           =   6
            Left            =   1948
            TabIndex        =   92
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone 3"
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
            Index           =   5
            Left            =   3388
            TabIndex        =   91
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone 4"
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
            Left            =   4833
            TabIndex        =   90
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            Left            =   6510
            TabIndex        =   89
            Top             =   780
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
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
            Left            =   10830
            TabIndex        =   88
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone 1"
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
            Left            =   508
            TabIndex        =   87
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
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
            Left            =   12373
            TabIndex        =   82
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razão social"
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
            Left            =   5460
            TabIndex        =   81
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   14678
            TabIndex        =   80
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   75
         Top             =   9090
         Width           =   15195
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
            TabIndex        =   21
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
            Left            =   9540
            TabIndex        =   22
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   26
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmvendas_tele_clientes.frx":2717
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
            TabIndex        =   25
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmvendas_tele_clientes.frx":5EBB
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
            TabIndex        =   23
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
            TabIndex        =   24
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmvendas_tele_clientes.frx":99C4
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
            TabIndex        =   27
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmvendas_tele_clientes.frx":DAB3
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
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
            Left            =   3510
            TabIndex        =   104
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
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
            Index           =   48
            Left            =   2190
            TabIndex        =   78
            Top             =   240
            Width           =   645
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
            TabIndex        =   77
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
            Left            =   13050
            TabIndex        =   76
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   6375
         Left            =   -67005
         TabIndex        =   69
         Top             =   2760
         Width           =   7275
         Begin VB.TextBox Txt_telefonel_contato 
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
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Telefone."
            Top             =   5910
            Width           =   1335
         End
         Begin VB.TextBox Txt_email_contato 
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
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "E-mail."
            Top             =   5910
            Width           =   2670
         End
         Begin VB.TextBox Txt_departamento_contato 
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
            Left            =   4080
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Departamento."
            Top             =   5340
            Width           =   3000
         End
         Begin VB.CommandButton imgemail1 
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
            Left            =   4215
            Picture         =   "frmvendas_tele_clientes.frx":1133F
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Enviar e-mail para o contato."
            Top             =   5910
            Width           =   315
         End
         Begin VB.TextBox Txt_status 
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
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox Txt_contato 
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
            MaxLength       =   40
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Nome do contato."
            Top             =   5340
            Width           =   3480
         End
         Begin VB.CommandButton Cmd_contato 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3690
            Picture         =   "frmvendas_tele_clientes.frx":11749
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Localizar contatos."
            Top             =   5340
            Width           =   315
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3855
         End
         Begin VB.CheckBox chkVisita 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Com visita?"
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
            Left            =   6000
            TabIndex        =   51
            Top             =   5970
            Width           =   1245
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
            Height          =   4365
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            ToolTipText     =   "Descrição."
            Top             =   750
            Width           =   6885
         End
         Begin MSComCtl2.DTPicker txtproximo 
            Height          =   315
            Left            =   4620
            TabIndex        =   50
            ToolTipText     =   "Data do próximo contato."
            Top             =   5910
            Width           =   1245
            _ExtentX        =   2196
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
            Format          =   199229441
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtData 
            Height          =   315
            Left            =   180
            TabIndex        =   40
            ToolTipText     =   "Data do último histórico."
            Top             =   390
            Width           =   1350
            _ExtentX        =   2381
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
            Format          =   199229441
            CurrentDate     =   39057
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone"
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
            Index           =   8
            Left            =   532
            TabIndex        =   103
            Top             =   5700
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
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
            Index           =   7
            Left            =   2655
            TabIndex        =   102
            Top             =   5700
            Width           =   420
         End
         Begin VB.Label lblcontato 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
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
            Index           =   6
            Left            =   5070
            TabIndex        =   101
            Top             =   5130
            Width           =   1035
         End
         Begin VB.Label lblcontato 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   6000
            TabIndex        =   95
            Top             =   180
            Width           =   465
         End
         Begin VB.Label lblcontato 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   4
            Left            =   3000
            TabIndex        =   94
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lblcontato 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   3
            Left            =   683
            TabIndex        =   93
            Top             =   180
            Width           =   345
         End
         Begin VB.Label lblcontato 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   71
            Top             =   5130
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Próx. contato"
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
            Left            =   4747
            TabIndex        =   70
            Top             =   5700
            Width           =   990
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo da visita"
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
         Height          =   825
         Left            =   -74925
         TabIndex        =   68
         Top             =   9150
         Width           =   2670
         Begin VB.CheckBox OPTvisitatecnica 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Técnica"
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
            TabIndex        =   52
            Top             =   420
            Width           =   855
         End
         Begin VB.CheckBox OPTVisita 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Apresentação"
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
            Left            =   1200
            TabIndex        =   53
            Top             =   420
            Width           =   1305
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   72
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
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
         ButtonCaption2  =   "Atualizar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Atualizar data do próximo contato (F7)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   50
         ButtonHeight2   =   21
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
         ButtonLeft3     =   92
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   96
         ButtonTop4      =   2
         ButtonWidth4    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   134
         ButtonTop5      =   2
         ButtonWidth5    =   26
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
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
         ButtonLeft6     =   162
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12210
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmvendas_tele_clientes.frx":1184B
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   73
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   16
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
         ButtonCaption2  =   "Salvar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Salvar (F3)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   38
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   77
         ButtonTop3      =   2
         ButtonWidth3    =   39
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Proposta"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Abrir módulo de proposta (F10)"
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
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   51
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Bloquear"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Bloquear históricos do cliente."
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   224
         ButtonTop6      =   2
         ButtonWidth6    =   50
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Desbloquear"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Desbloquear históricos do cliente."
         ButtonKey7      =   "7"
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
         ButtonLeft7     =   276
         ButtonTop7      =   2
         ButtonWidth7    =   68
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Primeiro"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Primeiro contato."
         ButtonKey8      =   "8"
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
         ButtonLeft8     =   346
         ButtonTop8      =   2
         ButtonWidth8    =   46
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Anterior"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Contato anterior."
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   394
         ButtonTop9      =   2
         ButtonWidth9    =   47
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Próximo"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Próximo contato."
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   443
         ButtonTop10     =   2
         ButtonWidth10   =   46
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Ultimo"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Ultimo contato."
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   491
         ButtonTop11     =   2
         ButtonWidth11   =   37
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Atualizar"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Utilizado pelo administrador do sistema."
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   530
         ButtonTop12     =   2
         ButtonWidth12   =   50
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonAlignment13=   2
         ButtonType13    =   1
         ButtonStyle13   =   -1
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   -1
         ButtonLeft13    =   582
         ButtonTop13     =   4
         ButtonWidth13   =   2
         ButtonHeight13  =   54
         ButtonCaption14 =   "Ajuda"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Ajuda (F1)"
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft14    =   586
         ButtonTop14     =   2
         ButtonWidth14   =   36
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonCaption15 =   "Sair"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Sair (Esc)"
         ButtonKey15     =   "15"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   624
         ButtonTop15     =   2
         ButtonWidth15   =   26
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonEnabled16 =   0   'False
         ButtonIconSize16=   32
         ButtonKey16     =   "16"
         ButtonAlignment16=   2
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState16   =   5
         ButtonLeft16    =   652
         ButtonTop16     =   2
         ButtonWidth16   =   24
         ButtonHeight16  =   24
         ButtonUseMaskColor16=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   12210
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmvendas_tele_clientes.frx":14099
            Count           =   1
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   75
         TabIndex        =   74
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
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   735
         Left            =   9405
         TabIndex        =   97
         Top             =   2160
         Width           =   5865
         Begin MSComCtl2.DTPicker msk_fltFim 
            Height          =   315
            Left            =   4380
            TabIndex        =   19
            ToolTipText     =   "Data final."
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   197328897
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker msk_fltInicio 
            Height          =   315
            Left            =   2550
            TabIndex        =   18
            ToolTipText     =   "Data início."
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   197328897
            CurrentDate     =   39057
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Até :"
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
            Left            =   3945
            TabIndex        =   99
            Top             =   270
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "De :"
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
            TabIndex        =   98
            Top             =   270
            Width           =   300
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   4545
         TabIndex        =   100
         Top             =   2160
         Width           =   4845
         Begin VB.OptionButton Opt_com_historico 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Com histórico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton Opt_sem_historico 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sem histórico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1770
            TabIndex        =   14
            Top             =   300
            Width           =   1485
         End
         Begin VB.CheckBox Chk_bloqueado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bloqueados"
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
            Left            =   3450
            TabIndex        =   15
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resonsável pelo contato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   75
         TabIndex        =   106
         Top             =   2160
         Width           =   4455
         Begin VB.ComboBox Cmb_responsavel_contato 
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
            ItemData        =   "frmvendas_tele_clientes.frx":1CF56
            Left            =   180
            List            =   "frmvendas_tele_clientes.frx":1CF63
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Responsável pelo contato."
            Top             =   270
            Width           =   4095
         End
      End
   End
End
Attribute VB_Name = "frmVendas_Tele_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Telemarketing As Boolean 'OK
Public StrSqlTeleLocaliza As String 'OK
Dim DataFiltro As String 'OK
Dim BloqueadoFiltro As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=eCcgBnOPx84&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=28&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_bloqueado_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_responsavel_contato_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
With Opt_sem_historico
    If Cmb_responsavel_contato <> "" Then
        Opt_com_historico.Value = True
        .Enabled = False
    Else
        .Enabled = True
    End If
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_com_historico_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If Opt_com_historico.Value = True Then
    Chk_ultimo.Enabled = True
    Chk_proximo.Enabled = True
    Chk_bloqueado.Enabled = True
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_proximo_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If Chk_proximo.Value = 1 Then
    Chk_ultimo.Value = 0
    Frame9.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame9.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_sem_historico_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If Opt_sem_historico.Value = True Then
    With Chk_bloqueado
        .Enabled = False
        .Value = 0
    End With
    With Chk_ultimo
        .Enabled = False
        .Value = 0
    End With
    With Chk_proximo
        .Enabled = False
        .Value = 0
    End With
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVisita_Click()
On Error GoTo tratar_erro

If chkVisita.Value = 1 Then
    Frame5.Enabled = True
Else
    OPTVisita.Value = 0
    OPTvisitatecnica.Value = 0
    Frame5.Enabled = False
    Txt_contato_visita = ""
    Txt_departamento_contato_visita = ""
    txtDataVisita.Value = Date
    mskhora_inicio.Value = "00:00:00"
    mskhora_fim.Value = "00:00:00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_contato_Click()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" Then
    Vendas_PI = False
    Vendas_Proposta = False
    Analise_critica = False
    Telemarketing = True
    Qualidade_PPAP_PSW = False
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    Sit_REG = 1
    frmVendas_propostaII_contato.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procPrimeiro()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from Vendas_tele where idcliente = " & txtIDcliente.Text & " order by Ultimo, Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.BOF = False Then
        TBLISTA.MoveFirst
        If TBLISTA.BOF = False Then
            txtCodigo.Text = TBLISTA!CODIGO
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from Vendas_tele where codigo = " & txtCodigo.Text, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                ProclimpaDescricao
                Proclimpacontato
                ProcPuxaDados
            End If
        Else
            USMsgBox ("Fim dos cadastros de históricos."), vbInformation, "CAPRIND v5.0"
        End If
    End If
End If
Novo_Telemarketing = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from Vendas_tele where idcliente = " & txtIDcliente.Text & " order by Ultimo, Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.BOF = False Then
        TBLISTA.Find (" codigo =" & txtCodigo.Text)
        TBLISTA.MovePrevious
        If TBLISTA.BOF = False Then
            txtCodigo.Text = TBLISTA!CODIGO
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from Vendas_tele where codigo = " & txtCodigo.Text, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                ProclimpaDescricao
                Proclimpacontato
                ProcPuxaDados
            End If
        Else
            USMsgBox ("Fim dos cadastros de históricos."), vbInformation, "CAPRIND v5.0"
        End If
    End If
End If
Novo_Telemarketing = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from Vendas_tele where idcliente = " & txtIDcliente.Text & " order by Ultimo, Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.BOF = False Then
        TBLISTA.Find ("codigo = " & txtCodigo.Text)
        TBLISTA.MoveNext
        If TBLISTA.EOF = False Then
            txtCodigo.Text = TBLISTA!CODIGO
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from Vendas_tele where codigo = " & txtCodigo.Text, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                ProclimpaDescricao
                Proclimpacontato
                ProcPuxaDados
            End If
        Else
            USMsgBox ("Fim dos cadastros de históricos."), vbInformation, "CAPRIND v5.0"
        End If
    End If
End If
Novo_Telemarketing = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procUltimo()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from Vendas_tele where idcliente = " & txtIDcliente.Text & " order by Ultimo, Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.BOF = False Then
        TBLISTA.MoveLast
        If TBLISTA.EOF = False Then
            txtCodigo.Text = TBLISTA!CODIGO
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from Vendas_tele where codigo = " & txtCodigo.Text, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                ProclimpaDescricao
                Proclimpacontato
                ProcPuxaDados
            End If
        Else
            USMsgBox ("Fim dos cadastros de históricos."), vbInformation, "CAPRIND v5.0"
        End If
    End If
End If
Novo_Telemarketing = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloquear()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente bloquear o(s) histórico(s) deste cliente?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "Update Vendas_tele Set Bloqueado = 'True' where idcliente = " & txtIDcliente.Text
    USMsgBox ("Histórico(s) do cliente " & txtnomerazao.Text & " bloqueado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Vendas/Telemarketing"
    Evento = "Bloquear"
    ID_documento = txtIDcliente
    Documento = "Cliente: " & txtnomerazao & " - Cidade: " & txtCidade
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from Vendas_tele where codigo = " & txtCodigo.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        ProclimpaDescricao
        Proclimpacontato
        ProcPuxaDados
    End If
    TBClientes.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcontato_visita_Click()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" Then
    Analise_critica = False
    Vendas_Proposta = False
    Vendas_PI = False
    Telemarketing = True
    Qualidade_PPAP_PSW = False
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    Sit_REG = 2
    frmVendas_propostaII_contato.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesbloquear()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente desbloquear o(s) histórico(s) deste cliente?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "Update Vendas_tele Set Bloqueado = 'False' where idcliente = " & txtIDcliente.Text
    USMsgBox ("Histórico(s) do cliente " & txtnomerazao.Text & " desbloqueado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Vendas/Telemarketing"
    Evento = "Desbloquear"
    ID_documento = txtIDcliente
    Documento = "Cliente: " & txtnomerazao & " - Cidade: " & txtCidade
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procProposta()
On Error GoTo tratar_erro

Vendas_PI = False
Vendas_Proposta = True
frmVendas_proposta.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente atualizar data do próximo contato menor que hoje para data de hoje?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select idcliente, NomeRazao from Clientes order by idcliente", Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        Do While TBClientes.EOF = False
            If IsNull(TBClientes!IDCliente) = False Then
                Set TBVendas = CreateObject("adodb.recordset")
                TBVendas.Open "Select * from Vendas_Tele where idcliente = " & TBClientes!IDCliente & " and bloqueado = 'False' order by Ultimo, Codigo", Conexao, adOpenKeyset, adLockOptimistic
                If TBVendas.EOF = False Then
                    TBVendas.MoveLast
                    If TBVendas!Proximo < Date Then
                        If USMsgBox("O cliente " & TBClientes!NomeRazao & " tem a data para próximo contato menor que hoje, deseja atualizar esta data para hoje " & Format(Date, "dd/mm/yy") & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                            TBVendas!Proximo = Format(Date, "dd/mm/yy")
                            TBVendas.Update
                        End If
                    End If
                End If
                TBVendas.Close
            End If
            TBClientes.MoveNext
        Loop
    End If
    TBClientes.Close
    USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Vendas/Telemarketing"
    Evento = "Atualizar"
    ID_documento = 0
    Documento = ""
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgemail1_Click()
On Error GoTo tratar_erro

If Txt_email_contato = "" Then Exit Sub
SendMail (Txt_email_contato.Text)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_historico_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Vendas_tele where codigo = " & Lista_historico.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.BOF = False Then
    txtCodigo.Text = TBClientes!CODIGO
    ProclimpaDescricao
    Proclimpacontato
    ProcPuxaDados
End If
Novo_Telemarketing = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Frame4.Enabled = False
Frame5.Enabled = False
txtIDcliente = ""
Novo_Telemarketing = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_historico_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_historico, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ultimo_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If Chk_ultimo.Value = 1 Then
    Chk_proximo.Value = 0
    Frame9.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame9.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPTVisita_Click()
On Error GoTo tratar_erro

If OPTVisita.Value = 1 Then
    Frame3.Enabled = True
    OPTvisitatecnica.Value = 0
Else
    If OPTvisitatecnica.Value = 0 Then
        Proclimpacontato
        Frame3.Enabled = False
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPTvisitatecnica_Click()
On Error GoTo tratar_erro

If OPTvisitatecnica.Value = 1 Then
    Frame3.Enabled = True
    OPTVisita.Value = 0
Else
    If OPTVisita.Value = 0 Then
        Proclimpacontato
        Frame3.Enabled = False
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgemail_Click()
On Error GoTo tratar_erro

If txtEmail = "" Then Exit Sub
SendMail (txtEmail.Text)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub SendMail(Optional Address As String, Optional Subject As String, Optional Body As String, Optional CC As String, Optional BCC As String)
On Error GoTo tratar_erro
Dim StrCommand As String 'OK

'constroi a string do email
If Len(Subject) Then StrCommand = "&Subject=" & Subject
If Len(Body) Then StrCommand = StrCommand & "&Body=" & Body
If Len(CC) Then StrCommand = StrCommand & "&CC=" & CC
If Len(BCC) Then StrCommand = StrCommand & "&BCC=" & BCC

'substitui o primeiro &
'com interrogacao
If Len(StrCommand) Then
   Mid(StrCommand, 1, 1) = "?"
End If

'Inclui o comando mailto: e o endereço de e-mail
StrCommand = "mailto: " & Address & StrCommand

'executa o comando via API
Call ShellExecute(Me.hWnd, "open", StrCommand, vbNullString, vbNullString, SW_SHOWNORMAL)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaCliente()
On Error GoTo tratar_erro

txtIDcliente.Text = TBClientes!IDCli
txtnomerazao.Text = TBClientes!NomeRazao
txtuf = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
txttel01 = IIf(IsNull(TBClientes!Tel01), "", Format(TBClientes!Tel01, "(##)####-####"))
txttel02 = IIf(IsNull(TBClientes!tel02), "", Format(TBClientes!tel02, "(##)####-####"))
txttel03 = IIf(IsNull(TBClientes!tel03), "", Format(TBClientes!tel03, "(##)####-####"))
txtFax = IIf(IsNull(TBClientes!Fax), "", Format(TBClientes!Fax, "(##)####-####"))
txtEmail = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
txtCidade = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Acao = "excluir"
If txtCodigo = 0 Then
    NomeCampo = "o histórico"
    ProcVerificaAcao
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir este histórico?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    'Verifica se é o último registro
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from Vendas_tele where IdCliente = " & txtIDcliente.Text & " order by codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        TBVendas.MoveLast
        If TBVendas!CODIGO <> txtCodigo Then
            USMsgBox ("Só é permitido excluir o último histórico deste cliente."), vbExclamation, "CAPRIND v5.0"
            TBVendas.Close
            Exit Sub
        End If
    End If
    TBVendas.Close
    Conexao.Execute "DELETE from vendas_tele where Codigo = " & txtCodigo
    USMsgBox ("Histórico excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Vendas/Telemarketing"
    Evento = "Excluir"
    ID_documento = txtCodigo
    Documento = "Cliente: " & txtnomerazao & " - Cidade: " & txtCidade
    Documento1 = "Data: " & Txtproximo & " - Histórico: " & txtdescricao
    ProcGravaEvento
    '==================================
    ProclimpaDescricao
    Proclimpacontato
    Frame4.Enabled = False
    Frame5.Enabled = False
    procUltimo
    procCarregalista_hist
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
'Verifica se o cliente está bloqueado
Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_tele where idcliente = " & txtIDcliente.Text & " and Bloqueado = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    USMsgBox ("Não é permitido criar um novo registro para este cliente, pois os históricos estão bloqueados."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
TBVendas.Close
Novo_Telemarketing = True
Frame4.Enabled = True
Frame5.Enabled = True
ProclimpaDescricao
Proclimpacontato
txtData.SetFocus

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
Acao = "salvar"
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtdescricao.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
'If Txt_contato = "" Then
'    NomeCampo = "o contato"
'    ProcVerificaAcao
'    Cmd_contato_Click
'    Exit Sub
'End If
If chkVisita.Value = 1 And OPTVisita.Value = 0 And OPTvisitatecnica.Value = 0 Then
    NomeCampo = "o tipo da visita"
    ProcVerificaAcao
    Exit Sub
End If
If OPTVisita.Value = 1 Or OPTvisitatecnica.Value = 1 Then
    If Txt_contato_visita = "" Then
        NomeCampo = "o nome do contato"
        ProcVerificaAcao
        cmdcontato_visita_Click
        Exit Sub
    End If
    If mskhora_inicio.Value = "00:00:00" Then
        NomeCampo = "a hora início da visita"
        ProcVerificaAcao
        mskhora_inicio.SetFocus
        Exit Sub
    End If
    If mskhora_fim.Value = "00:00:00" Then
        NomeCampo = "a hora final da visita"
        ProcVerificaAcao
        mskhora_fim.SetFocus
        Exit Sub
    End If
End If
Dataini = txtData
DataFim = Txtproximo
If Dataini = DataFim Then
    If USMsgBox("A data do próximo contato está igual a data do último contato, deseja prosseguir mesmo assim?.", vbYesNo, "CAPRIND v5.0") = vbNo Then
        Txtproximo.SetFocus
        Exit Sub
    End If
End If

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Vendas_tele where Codigo = " & txtCodigo, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    TBClientes.AddNew
    TBClientes!Responsavel = pubUsuario
    TBClientes!Bloqueado = False
End If
ProcEnviaDados
TBClientes.Update
txtCodigo = TBClientes!CODIGO
TBClientes.Close

'Grava dados do penúltimo histórico
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Vendas_tele where IDcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.Find ("codigo = " & txtCodigo.Text)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Vendas_tele where codigo = " & TBLISTA!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            TBGravar!Proximo = txtData
            TBGravar.Update
        End If
        TBGravar.Close
    End If
End If
TBLISTA.Close

If Novo_Telemarketing = True Then
    USMsgBox ("Novo histórico cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
'==================================
Modulo = "Vendas/Telemarketing"
ID_documento = txtCodigo
Documento = "Cliente: " & txtnomerazao & " - Cidade: " & txtCidade
Documento1 = "Data: " & Txtproximo & " - Histórico: " & txtdescricao
ProcGravaEvento
'==================================
procCarregalista_hist
Novo_Telemarketing = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados()
On Error GoTo tratar_erro

If OPTVisita.Value = 0 Then TBClientes!visita = False
If OPTvisitatecnica.Value = 0 Then TBClientes!visita_tecnica = False
If OPTVisita.Value = 1 Or OPTvisitatecnica.Value = 1 Then
    TBClientes!Contato_visita = Txt_contato_visita.Text
    TBClientes!data_visita = txtDataVisita
    
    Diasemana = Weekday(DateValue(txtDataVisita))
    Diasemana = WeekdayName(Diasemana)
    TBClientes!dia_semana = DiaSemanaFunValorExtenso
    
    If mskhora_inicio.Value <> "00:00:00" Then TBClientes!hora_inicio = mskhora_inicio.Value
    If mskhora_fim.Value <> "00:00:00" Then TBClientes!hora_fim = mskhora_fim.Value
    If OPTVisita.Value = 1 Then
        TBClientes!visita = True
        TBClientes!visita_tecnica = False
    Else
        TBClientes!visita_tecnica = True
        TBClientes!visita = False
    End If
ElseIf OPTVisita.Value = 0 And OPTvisitatecnica.Value = 0 Then
        TBClientes!Contato_visita = Null
        TBClientes!data_visita = Null
        TBClientes!dia_semana = Null
        TBClientes!hora_inicio = Null
        TBClientes!hora_fim = Null
End If
TBClientes!IDCliente = txtIDcliente
TBClientes!Cliente = txtnomerazao
TBClientes!Ultimo = txtData
TBClientes!Descricao = txtdescricao
TBClientes!contato = Txt_contato
TBClientes!Proximo = Txtproximo

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
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF7: ProcAtualizarData
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCliente()
On Error GoTo tratar_erro

txtnomerazao.Text = ""
txttel01.Text = ""
txttel02.Text = ""
txttel03.Text = ""
txttel04.Text = ""
txtFax.Text = ""
txtCidade.Text = ""
txtEmail.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProclimpaDescricao()
On Error GoTo tratar_erro

txtCodigo.Text = 0
txtData.Value = Date
txtResponsavel = pubUsuario
Txt_status = "Liberado"
txtdescricao = ""
Txt_contato = ""
Txt_departamento_contato = ""
Txt_telefonel_contato = ""
Txt_email_contato = ""
chkVisita.Value = 0
Txtproximo.Value = Date
OPTvisitatecnica.Value = 0
OPTVisita.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Proclimpacontato()
On Error GoTo tratar_erro

Txt_contato_visita.Text = ""
Txt_departamento_contato_visita.Text = ""
txtDataVisita = Date
mskhora_inicio.Value = "00:00:00"
mskhora_fim.Value = "00:00:00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
ProcCarregaToolBar2 Me, 15195, 16, True
Formulario = "Vendas/Telemarketing"
Direitos
SSTab1.Tab = 0
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
Txtproximo.Value = Date
txtData.Value = Date
txtDataVisita.Value = Date
cmbfiltrarpor = "Razão social"

ProcCarregaComboUsuario Cmb_responsavel_contato, "U.usuario IS NOT NULL", True
ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

Acao = "visualizar impressão"
If txtIDcliente.Text = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    Exit Sub
End If
If txtCodigo = 0 Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Exit Sub
End If
If OPTvisitatecnica.Value = 0 And OPTVisita.Value = 0 Then
    USMsgBox ("Não há visitas cadastradas para este histórico."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If OPTVisita.Value = 1 Then NomeRel = "Telemarketing_Visita.rpt" Else NomeRel = "Telemarketing_Visita tecnica.rpt"
ProcImprimirRel "{clientes.idcliente}= " & txtIDcliente.Text & " and {vendas_tele.codigo}= " & txtCodigo.Text, ""
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Telemarketing = True Then
    If USMsgBox("O histórico ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Telemarketing = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Telemarketing = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtCodigo = TBClientes!CODIGO
txtData = IIf(IsNull(TBClientes!Ultimo), Date, TBClientes!Ultimo)
txtResponsavel = IIf(IsNull(TBClientes!Responsavel) = False, TBClientes!Responsavel, pubUsuario)
If TBClientes!Bloqueado = True Then Txt_status = "Bloqueado" Else Txt_status = "Liberado"
txtdescricao = IIf(IsNull(TBClientes!Descricao) = False, TBClientes!Descricao, "")

Txt_contato = IIf(IsNull(TBClientes!contato) = False, TBClientes!contato, "")
If Txt_contato <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Departamento, Telefone, Email from Clientes_Contatos where IDcliente = " & txtIDcliente & " and NomeContato = '" & Txt_contato & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_departamento_contato = IIf(IsNull(TBAbrir!Departamento), "", TBAbrir!Departamento)
        Txt_telefonel_contato = IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone)
        Txt_email_contato = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
    End If
    TBAbrir.Close
End If

If TBClientes!visita Or TBClientes!visita_tecnica = True Then chkVisita.Value = 1 Else chkVisita.Value = 0
Txtproximo = IIf(IsNull(TBClientes!Proximo), Date, TBClientes!Proximo)

If TBClientes!visita Then OPTVisita.Value = 1 Else OPTVisita.Value = 0
If TBClientes!visita_tecnica Then OPTvisitatecnica.Value = 1 Else OPTvisitatecnica.Value = 0

Txt_contato_visita.Text = IIf(IsNull(TBClientes!Contato_visita), "", TBClientes!Contato_visita)
If Txt_contato_visita <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Departamento from Clientes_Contatos where IDcliente = " & txtIDcliente & " and NomeContato = '" & Txt_contato_visita & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_departamento_contato_visita = IIf(IsNull(TBAbrir!Departamento), "", TBAbrir!Departamento)
    End If
    TBAbrir.Close
End If

txtDataVisita.Value = IIf(IsNull(TBClientes!data_visita), Date, TBClientes!data_visita)
mskhora_inicio.Value = IIf(IsNull(TBClientes!hora_inicio), "00:00:00", Left(TBClientes!hora_inicio, 8))
mskhora_fim.Value = IIf(IsNull(TBClientes!hora_fim), "00:00:00", Left(TBClientes!hora_fim, 8))
Frame4.Enabled = True
Frame5.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        ListView1.Visible = True
        Lista_historico.Visible = False
        If ListView1.Visible = True Then ListView1.SetFocus
    Case 1:
        ListView1.Visible = False
        Lista_historico.Visible = True
        If ListView1.ListItems.Count = 0 Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        txtIDcliente = ListView1.SelectedItem
        If ListView1.SelectedItem.ListSubItems(7) <> "" Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from Vendas_tele where Codigo = " & ListView1.SelectedItem.ListSubItems(7), Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                ProcPuxaDados
            End If
            TBClientes.Close
        End If
        procCarregalista_hist
        If Lista_historico.Visible = True Then Lista_historico.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDcliente_Change()
On Error GoTo tratar_erro

Frame4.Enabled = False
Frame5.Enabled = False
ProcLimpaCliente
ProclimpaDescricao
Proclimpacontato
If txtIDcliente <> "" Then
    VerifNumero = txtIDcliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDcliente = ""
        txtIDcliente.SetFocus
        Exit Sub
    End If
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select IDCliente as IDCli, NomeRazao, UF, Tel01, tel02, tel03, Fax, Email, Cidade from clientes where IDCliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        ProcCarregaCliente
        procUltimo
        Novo_Telemarketing = False
    End If
    TBClientes.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtproximo_LostFocus()
On Error GoTo tratar_erro

Dataini = txtData.Value
DataFim = Txtproximo.Value
If DataFim < Dataini Then
    USMsgBox ("A data do próximo histórico não pode ser menor que a data do último histórico."), vbExclamation, "CAPRIND v5.0"
    VerifData = False
Else
    VerifData = True
End If
If VerifData = False Then
    Txtproximo.Value = Date
    Txtproximo.SetFocus
    Exit Sub
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcAtualizarData
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_cliente_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_cliente_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_cliente_padrao.PageCount - 1)
    Else
        TBLocalizar_cliente_padrao.AbsolutePage = TBLocalizar_cliente_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)
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
    TBLocalizar_cliente_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_cliente_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_cliente_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_cliente_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_cliente_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_cliente_padrao.AbsolutePage = TBLocalizar_cliente_padrao.PageCount
ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListView1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlTeleLocaliza = "" Then Exit Sub
Set TBLocalizar_cliente_padrao = CreateObject("adodb.recordset")
TBLocalizar_cliente_padrao.Open StrSqlTeleLocaliza, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_cliente_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_cliente_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_cliente_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_cliente_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
If Opt_com_historico.Value = True Then
    PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_cliente_padrao!TotalRecords - IIf(Pagina > 1, (TBLocalizar_cliente_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_cliente_padrao.PageSize)
Else
    PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_cliente_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_cliente_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_cliente_padrao.PageSize)
End If
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_cliente_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_cliente_padrao!IDCliente
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_cliente_padrao!NomeRazao), "", TBLocalizar_cliente_padrao!NomeRazao)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_cliente_padrao!Cidade), "", TBLocalizar_cliente_padrao!Cidade)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_cliente_padrao!UF), "", TBLocalizar_cliente_padrao!UF)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_cliente_padrao!Ultimo), "", Format(TBLocalizar_cliente_padrao!Ultimo, "dd/mm/yy"))
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_cliente_padrao!Proximo), "", Format(TBLocalizar_cliente_padrao!Proximo, "dd/mm/yy"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_cliente_padrao!Descricao), "", TBLocalizar_cliente_padrao!Descricao)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLocalizar_cliente_padrao!CODIGO), "", TBLocalizar_cliente_padrao!CODIGO)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLocalizar_cliente_padrao!Responsavel), "", TBLocalizar_cliente_padrao!Responsavel)
    End With
    TBLocalizar_cliente_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLocalizar_cliente_padrao.RecordCount
If TBLocalizar_cliente_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_cliente_padrao.PageCount
ElseIf TBLocalizar_cliente_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_cliente_padrao.PageCount & " de: " & TBLocalizar_cliente_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_cliente_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_cliente_padrao.PageCount
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

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia.Text <> "" Then
    txtTexto.Text = ""
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Razão social" Or cmbfiltrarpor = "Nome fantasia" Or cmbfiltrarpor = "Cidade" Or cmbfiltrarpor = "Código do cliente" Then
    txtTexto.Visible = True
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
    cmbfamilia.Clear
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbfamilia, "familia IS NOT NULL and vendas = 'True'", False
    Else
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select Texto from Clientes_grupos where Texto IS NOT NULL group by Texto", Conexao, adOpenKeyset, adLockOptimistic
        If TBFamilia.EOF = False Then
            Do While TBFamilia.EOF = False
                cmbfamilia.AddItem TBFamilia!Texto
                TBFamilia.MoveNext
            Loop
        End If
    End If
    TBFamilia.Close
End If
If cmbfiltrarpor = "Status" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = True
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optJuridica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = True
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optFisica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbstatus_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbStatus.Text <> "" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

If optFisica.Value = True Then
    TipoPessoa = "(C.tipo = 'FP' or C.tipo = 'FR')"
    TipoPessoaRel = "{C.tipo} = 'FP' or {C.tipo} = 'FR'"
    CpfCnpj = "C.cpf_cnpj = '" & txtCpf.Text & "'"
Else
    TipoPessoa = "(C.tipo = 'JP' or C.tipo = 'JR')"
    TipoPessoaRel = "({C.tipo} = 'JP' or {C.tipo} = 'JR')"
    CpfCnpj = "C.cpf_cnpj = '" & txtcnpj.Text & "'"
End If

DataFiltro = ""
If Chk_ultimo.Value = 1 Or Chk_proximo = 1 Then
    If Chk_ultimo.Value = 1 Then DataTexto = "VT.Ultimo" Else DataTexto = "VT.Proximo"
    DataFiltro = " and " & DataTexto & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
End If

BloqueadoFiltro = ""
CamposFiltro = "C.IDCliente, C.NomeRazao, C.Cidade, C.UF, VT.codigo, VT.Proximo, VT.Ultimo, VT.Descricao, VT.Responsavel"
If Opt_com_historico.Value = True Then
    If Chk_bloqueado.Value = 1 Then BloqueadoFiltro = " and VT.bloqueado = 'True'" Else BloqueadoFiltro = " and VT.bloqueado <> 'True'"
    TextoFiltroHist = " and VT.Codigo IS NOT NULL"
    INNERJOINTEXTO = "Select DISTINCT COUNT(C.IDCliente) OVER () AS TotalRecords, " & CamposFiltro & " from ((clientes C INNER JOIN vendas_tele VT ON C.idcliente = VT.IDCliente) LEFT JOIN compras_fornecedores_familia CFF ON C.IDCliente = CFF.IDCliente) LEFT JOIN Clientes_grupos CG ON C.IDGrupo = CG.ID"
    TextoFiltroHist1 = "VT.Codigo = (Select MAX(VT1.Codigo) from vendas_tele VT1 where VT1.IDcliente = VT.IDCliente " & Replace(BloqueadoFiltro, "VT.", "VT1.") & ") and "
Else
    TextoFiltroHist = " and VT.Codigo IS NULL"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from ((clientes C LEFT JOIN vendas_tele VT ON C.idcliente = VT.IDCliente) LEFT JOIN compras_fornecedores_familia CFF ON C.IDCliente = CFF.IDCliente) LEFT JOIN Clientes_grupos CG ON C.IDGrupo = CG.ID"
    TextoFiltroHist1 = ""
End If

If Cmb_responsavel_contato <> "" Then TextoFiltroResp = " and VT.Responsavel = '" & Cmb_responsavel_contato & "'" Else TextoFiltroResp = ""

TextoFiltroPadrao = TipoPessoa & TextoFiltroHist & BloqueadoFiltro & DataFiltro & TextoFiltroResp & " and C.DtValidacao IS NOT NULL group by " & CamposFiltro & " order by VT.Proximo, C.nomerazao"

If txtTexto <> "" Or cmbfamilia <> "" Or cmbStatus <> "" Or txtcnpj <> "__.___.___/____-__" Or txtCpf <> "___.___.___-__" Then
    If cmbfiltrarpor = "Status" Then
        StrSqlTeleLocaliza = INNERJOINTEXTO & " where " & TextoFiltroHist1 & "C.status = '" & cmbStatus.Text & "'" & " and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor = "Família" Then
            StrSqlTeleLocaliza = INNERJOINTEXTO & " where " & TextoFiltroHist1 & " CFF.Familia = '" & cmbfamilia & "' and CFF.tipo = 'C' and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Grupo" Then
                StrSqlTeleLocaliza = INNERJOINTEXTO & " where " & TextoFiltroHist1 & " CG.Texto = '" & cmbfamilia & "' and " & TextoFiltroPadrao
            ElseIf cmbfiltrarpor = "CNPJ/CPF" Then
                    StrSqlTeleLocaliza = INNERJOINTEXTO & " where " & TextoFiltroHist1 & CpfCnpj & " and " & TextoFiltroPadrao
                Else
                    Select Case cmbfiltrarpor
                        Case "Razão social": TextoFiltro = "C.nomerazao"
                        Case "Nome fantasia": TextoFiltro = "C.nomefantasia"
                        Case "Cidade": TextoFiltro = "C.cidade"
                        Case "Código do cliente": TextoFiltro = "C.IDCliente"
                    End Select
                    StrSqlTeleLocaliza = INNERJOINTEXTO & " where " & TextoFiltroHist1 & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlTeleLocaliza = INNERJOINTEXTO & " where " & TextoFiltroHist1 & TextoFiltroPadrao
End If
'Debug.print StrSqlTeleLocaliza

ProcCarregaLista
'
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optFisica_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If optFisica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = False
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optJuridica_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If optJuridica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = True
    txtCpf.Visible = False
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcnpj_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtcnpj.Text <> "__.___.___/____-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCpf_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtCpf.Text <> "___.___.___-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: ProcImprimir
    Case 5: procProposta
    Case 6: ProcBloquear
    Case 7: ProcDesbloquear
    Case 8: procPrimeiro
    Case 9: ProcAnterior
    Case 10: ProcProximo
    Case 11: procUltimo
    Case 12: procAtualiza
    Case 14: ProcAjuda
    Case 15: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCarregalista_hist()
On Error GoTo tratar_erro

Lista_historico.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CODIGO, Ultimo, Responsavel, Contato from Vendas_tele where IDcliente = " & txtIDcliente & " order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBlista_hist.Min = 0
    PBlista_hist.Max = TBLISTA.RecordCount
    PBlista_hist.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_historico.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Ultimo), "", Format(TBLISTA!Ultimo, "DD/MM/YY"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!contato), "", TBLISTA!contato)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBlista_hist.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizarData()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If USMsgBox("Deseja realmente alterar a data do próximo contato desse(s) cliente(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                Exit Sub
            Else
                Permitido = True
                GoTo Prosseguir
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) cliente(s) antes de alterar a data do próximo contato."), vbExclamation, "CAPRIND v5.0"
Else
Prosseguir:
    frmvendas_tele_clientes_data.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
