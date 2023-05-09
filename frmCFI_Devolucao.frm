VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCFI_Devolucao 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estoque - Almoxarifado - Devolver"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9435
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   7035
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   12409
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
      TabCaption(0)   =   "Dados da devolução"
      TabPicture(0)   =   "frmCFI_Devolucao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Destino/aplicação"
      TabPicture(1)   =   "frmCFI_Devolucao.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista"
      Tab(1).Control(1)=   "PBLista"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).ControlCount=   3
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
         Height          =   4395
         Left            =   75
         TabIndex        =   31
         Top             =   1320
         Width           =   9270
         Begin VB.TextBox Txt_RE 
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
            Left            =   5370
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Número da rastreabilidade de estoque."
            Top             =   370
            Width           =   1155
         End
         Begin VB.TextBox txtMaquina 
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
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Posto de trabalho."
            Top             =   1650
            Width           =   2535
         End
         Begin VB.TextBox txtDescricao_maquina 
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
            Left            =   2730
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do posto de trabalho."
            Top             =   1644
            Width           =   6345
         End
         Begin VB.CheckBox Optdevolucao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dev. c/ probl."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7650
            TabIndex        =   15
            Top             =   2985
            Width           =   1425
         End
         Begin VB.TextBox Txt_cod_ref 
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
            Left            =   7560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   375
            Width           =   1515
         End
         Begin VB.TextBox txtlote 
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
            Left            =   6540
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Número do lote."
            Top             =   370
            Width           =   1005
         End
         Begin VB.TextBox txtfamilia 
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
            MaxLength       =   100
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Família."
            Top             =   2918
            Width           =   6135
         End
         Begin VB.TextBox txtObservacao 
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
            Height          =   675
            Left            =   180
            MousePointer    =   99  'Custom
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            ToolTipText     =   "Observações."
            Top             =   3555
            Width           =   8895
         End
         Begin VB.TextBox txtDiasatraso 
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
            Left            =   3870
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Dias em atraso."
            Top             =   370
            Width           =   1485
         End
         Begin VB.TextBox txtID 
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
            Left            =   180
            TabIndex        =   32
            Text            =   "0"
            Top             =   370
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtQuantRetirada 
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
            Left            =   6330
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade retirada."
            Top             =   2925
            Width           =   1215
         End
         Begin VB.TextBox txtDataprevdevolucao 
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
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data de prevista da devolução."
            Top             =   370
            Width           =   1155
         End
         Begin VB.TextBox txtCodInterno 
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
            MaxLength       =   50
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   2281
            Width           =   1395
         End
         Begin VB.TextBox txtFuncionario 
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
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Funcionário."
            Top             =   1007
            Width           =   8895
         End
         Begin VB.TextBox txtDataDevolucao 
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
            Left            =   2700
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Data da devolução."
            Top             =   370
            Width           =   1155
         End
         Begin VB.TextBox txtdescricao 
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
            Left            =   1590
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   2281
            Width           =   7485
         End
         Begin VB.TextBox txtData_retirada 
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
            MaxLength       =   50
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Data de retirada."
            Top             =   370
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° RE"
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
            Index           =   19
            Left            =   5737
            TabIndex        =   50
            Top             =   180
            Width           =   420
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posto de trabalho"
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
            Index           =   18
            Left            =   810
            TabIndex        =   49
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição posto de trabalho"
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
            Index           =   17
            Left            =   4897
            TabIndex        =   48
            Top             =   1440
            Width           =   2010
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Index           =   13
            Left            =   7642
            TabIndex        =   44
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº lote"
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
            Index           =   12
            Left            =   6795
            TabIndex        =   43
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   11
            Left            =   4980
            TabIndex        =   42
            Top             =   2070
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   10
            Left            =   3007
            TabIndex        =   41
            Top             =   2730
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   4155
            TabIndex        =   40
            Top             =   3330
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. retirada"
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
            Index           =   9
            Left            =   6420
            TabIndex        =   39
            Top             =   2730
            Width           =   1035
         End
         Begin VB.Label lblDiasAtraso 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dias em atraso"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   3975
            TabIndex        =   38
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prev. devol."
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
            Left            =   1665
            TabIndex        =   37
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   420
            TabIndex        =   36
            Top             =   2070
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. devol."
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
            Left            =   2917
            TabIndex        =   35
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. retirada"
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
            Left            =   427
            TabIndex        =   34
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionário"
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
            Left            =   4215
            TabIndex        =   33
            Top             =   810
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   1095
         Left            =   75
         TabIndex        =   24
         Top             =   5700
         Width           =   9270
         Begin VB.TextBox txtqtdedevolvidaprobl 
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
            Left            =   7830
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade devolvida com problemas."
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox txtqtdedevolverprobl 
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
            Left            =   7830
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade a devolver com problemas."
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox txtqtdedevolver 
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
            Left            =   4365
            TabIndex        =   20
            Text            =   "0,000"
            ToolTipText     =   "Quantidade a devolver."
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox txtquantdevolvida 
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
            Left            =   4365
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade devolvida."
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox txtestoque_atualizado 
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
            Left            =   1590
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque atualizado."
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox txtEstoque 
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
            Left            =   1590
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade em estoque."
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. devolvida c/ probl. :"
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
            Left            =   5835
            TabIndex        =   30
            Top             =   270
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. a devolver c/ probl. :"
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
            Index           =   15
            Left            =   5760
            TabIndex        =   29
            Top             =   630
            Width           =   1980
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. a devolver :"
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
            Index           =   14
            Left            =   2970
            TabIndex        =   28
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. devolvida :"
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
            Left            =   3045
            TabIndex        =   27
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estoque real :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   375
            TabIndex        =   26
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. atualizado :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   195
            TabIndex        =   25
            Top             =   630
            Width           =   1320
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   45
         Top             =   330
         Width           =   9270
         _ExtentX        =   16351
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
         ButtonCaption1  =   "Devolver"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Devolver (F3)"
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
         ButtonWidth1    =   51
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonAlignment2=   2
         ButtonType2     =   1
         ButtonStyle2    =   -1
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState2    =   -1
         ButtonLeft2     =   55
         ButtonTop2      =   4
         ButtonWidth2    =   2
         ButtonHeight2   =   54
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
         ButtonLeft3     =   59
         ButtonTop3      =   2
         ButtonWidth3    =   36
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
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
         ButtonLeft4     =   97
         ButtonTop4      =   2
         ButtonWidth4    =   26
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonKey5      =   "5"
         ButtonAlignment5=   2
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   5
         ButtonLeft5     =   125
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
         ButtonUseMaskColor5=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   7650
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCFI_Devolucao.frx":0038
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   46
         Top             =   330
         Width           =   9270
         _ExtentX        =   16351
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
         ButtonCaption1  =   "Novo"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Salvar (F3)"
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
         ButtonCaption2  =   "Excluir"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Excluir (F4)"
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
         ButtonWidth2    =   39
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
         ButtonLeft3     =   78
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
         ButtonLeft4     =   82
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
         ButtonLeft5     =   120
         ButtonTop5      =   2
         ButtonWidth5    =   26
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   148
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   7530
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCFI_Devolucao.frx":22A5
            Count           =   1
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   -74925
         TabIndex        =   47
         Top             =   6510
         Width           =   9270
         _ExtentX        =   16351
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
      Begin MSComctlLib.ListView Lista 
         Height          =   5160
         Left            =   -74925
         TabIndex        =   22
         Top             =   1335
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   9102
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   512
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cod produto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   8431
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   1058
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCFI_Devolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtData_retirada.Text = ""
txtDataprevdevolucao.Text = ""
txtDataDevolucao.Text = ""
txtDiasatraso.Text = ""
txtmaquina.Text = ""
txtDescricao_maquina.Text = ""
txtFuncionario.Text = ""
txtCodinterno.Text = ""
txtfamilia.Text = ""
txtdescricao.Text = ""
txtQuantRetirada.Text = "0,0000"
txt_RE = ""
txtLote.Text = ""
Txt_cod_ref = ""
txtObservacao.Text = ""
optDevolucao.Value = 0
txtEstoque.Text = "0,0000"
txtquantdevolvida.Text = "0,0000"
txtqtdedevolvidaprobl.Text = "0,0000"
txtestoque_atualizado.Text = "0,0000"
txtqtdedevolver.Text = "0,0000"
txtqtdedevolverprobl.Text = "0,0000"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyEscape: Unload Me
            Case vbKeyF3: ProcDevolver
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyEscape: Unload Me
            Case vbKeyF4: ProcExcluir
            Case vbKeyInsert: ProcNovo
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim Dataretirada        As Date 'OK
Dim Datadevolucao       As Date 'OK
Dim Dataprevdevolucao   As Date 'OK

ProcCarregaToolBar1 Me, 9270, 5, True
ProcCarregaToolBar2 Me, 9270, 6, True
ProcLimpaVariaveisPrincipais

If Qualidade_Almox = True Then Caption = "Qualidade - Almoxarifado - Devolver"
SSTab.Tab = 0
txtId.Text = frmCFI.txtId
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "SELECT * from CFI WHERE IDCFI = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtData_retirada = Format(TBLISTA!Dataretirada, "dd/mm/yy")
    txtDataprevdevolucao = Format(TBLISTA!dataprevisao, "dd/mm/yy")
    txtDataDevolucao = Format(Date, "dd/mm/yy")
    Datadevolucao = txtDataDevolucao
    Dataprevdevolucao = txtDataprevdevolucao
    If Datadevolucao > Dataprevdevolucao Then
        txtDiasatraso = Datadevolucao - Dataprevdevolucao
        txtDiasatraso.ForeColor = &HC0&
    Else
        txtDiasatraso.Text = 0
    End If
    txtFuncionario = TBLISTA!Funcionario
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select maquina, descricao from Cadmaquinas where idmaquina = " & IIf(IsNull(TBLISTA!ID_Maquina), 0, TBLISTA!ID_Maquina), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtmaquina = TBAbrir!maquina
        txtDescricao_maquina = TBAbrir!Descricao
    End If
    TBAbrir.Close
    
    txtCodinterno = TBLISTA!Codigo_produto
    txtfamilia.Text = TBLISTA!Familia
    txtdescricao.Text = TBLISTA!Descricao
    txtQuantRetirada = Format(TBLISTA!Quantretirada, "###,##0.0000")
    txt_RE = TBLISTA!IDEstoque
    txtLote.Text = TBLISTA!LOTE
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Ref from Estoque_controle where IDEstoque = " & IIf(IsNull(TBLISTA!IDEstoque), 0, TBLISTA!IDEstoque), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_cod_ref = IIf(IsNull(TBAbrir!Ref), "", TBAbrir!Ref)
    End If
    TBAbrir.Close
    
    txtObservacao.Text = IIf(IsNull(TBLISTA!Observacao), "", TBLISTA!Observacao)
    txtEstoque = Format(frmCFI.txtquantestoque, "###,##0.0000")
    txtquantdevolvida.Text = Format(TBLISTA!Quantdevolvido, "###,##0.0000")
    txtqtdedevolvidaprobl.Text = Format(TBLISTA!Quantdevolvidoprobl, "###,##0.0000")
    If TBLISTA!restricao = True Then
        optDevolucao.Value = 1
        optDevolucao.Enabled = False
    Else
        optDevolucao.Value = 0
        optDevolucao.Enabled = True
    End If
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lst_NatOp, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optdevolucao_Click()
On Error GoTo tratar_erro

If optDevolucao.Value = 1 Then
    txtObservacao.Text = ""
    txtqtdedevolverprobl.Text = ""
    txtqtdedevolverprobl.Locked = False
    txtqtdedevolverprobl.TabStop = True
Else
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "SELECT * from CFI WHERE IDCFI = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        txtObservacao.Text = TBLISTA!Observacao
    End If
    TBLISTA.Close
    txtqtdedevolverprobl.Text = ""
    txtqtdedevolverprobl.Locked = True
    txtqtdedevolverprobl.TabStop = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab.Tab
    Case 0: If txtqtdedevolver.Visible = True Then txtqtdedevolver.SetFocus
    Case 1:
        If Lista.Visible = True Then Lista.SetFocus
        ProcCarregaLista
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtdedevolver_Change()
On Error GoTo tratar_erro

If txtqtdedevolver.Text <> "" Then
    VerifNumero = txtqtdedevolver.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtqtdedevolver.Text = ""
        txtqtdedevolver.SetFocus
        Exit Sub
    End If
End If
Estoque = IIf(txtEstoque.Text = "", "0", txtEstoque.Text)
QuantSolicitado = IIf(txtqtdedevolver.Text = "", "0", txtqtdedevolver.Text)
txtestoque_atualizado = Format(Estoque + QuantSolicitado, "###,##0.0000")
Estoque = 0
QuantSolicitado = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtdedevolver_GotFocus()
On Error GoTo tratar_erro

If txtqtdedevolver = "0,0000" Then txtqtdedevolver = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtdedevolver_LostFocus()
On Error GoTo tratar_erro

txtqtdedevolver.Text = Format(txtqtdedevolver.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtdedevolverprobl_Change()
On Error GoTo tratar_erro

If txtqtdedevolverprobl.Text <> "" Then
    VerifNumero = txtqtdedevolverprobl.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtqtdedevolverprobl.Text = ""
        txtqtdedevolverprobl.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtdedevolverprobl_GotFocus()
On Error GoTo tratar_erro

If txtqtdedevolverprobl = "0,0000" Then txtqtdedevolverprobl = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtdedevolverprobl_LostFocus()
On Error GoTo tratar_erro

txtqtdedevolverprobl.Text = Format(txtqtdedevolverprobl.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcDevolver()
On Error GoTo tratar_erro

quantestoque = IIf(txtquantdevolvida.Text = "", 0, txtquantdevolvida.Text)
ValorTotal = IIf(txtqtdedevolvidaprobl.Text = "", 0, txtqtdedevolvidaprobl.Text)
quantnovo = IIf(txtqtdedevolver.Text = "", 0, txtqtdedevolver.Text)
qtdeliberar = IIf(txtqtdedevolverprobl = "", 0, txtqtdedevolverprobl)
qtdeliberada = IIf(txtQuantRetirada.Text = "", 0, txtQuantRetirada.Text)

Acao = "devolver"
If optDevolucao.Value = 0 Then
    If quantnovo = 0 Then
        NomeCampo = "a quantidade a devolver"
        ProcVerificaAcao
        txtqtdedevolver.SetFocus
        Exit Sub
    End If
End If
If optDevolucao.Value = 1 Then
    If txtObservacao.Text = "" Then
        USMsgBox ("Informe uma observação referente a devolução do material com problema."), vbExclamation, "CAPRIND v5.0"
        txtObservacao.SetFocus
        Exit Sub
    End If
    If txtqtdedevolver = "" Then
        NomeCampo = "a quantidade a devolver"
        ProcVerificaAcao
        txtqtdedevolver.SetFocus
        Exit Sub
    End If
    If qtdeliberar = 0 Then
        NomeCampo = "a quantidade a devolver c/ problema"
        ProcVerificaAcao
        txtqtdedevolverprobl.SetFocus
        Exit Sub
    End If
End If

If ((quantestoque + ValorTotal) + (quantnovo + qtdeliberar)) > qtdeliberada Then
    USMsgBox ("A quantidade a devolver é maior que a quantidade retirada."), vbExclamation, "CAPRIND v5.0"
    txtqtdedevolver.Text = ""
    txtqtdedevolver.SetFocus
    Exit Sub
End If

Set TBCorretiva = CreateObject("adodb.recordset")
TBCorretiva.Open "SELECT * from CFI where idcfi = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If ((quantestoque + ValorTotal) + (quantnovo + qtdeliberar)) < qtdeliberada Then
    TBCorretiva!status = "DEVOLVIDO PARCIAL"
Else
    TBCorretiva!status = "DEVOLVIDO"
End If
TBCorretiva!Datadevolucao = txtDataDevolucao.Text
If optDevolucao.Value = 1 Then TBCorretiva!restricao = True Else TBCorretiva!restricao = False
TBCorretiva!Quantdevolvido = quantestoque + quantnovo
TBCorretiva!Quantdevolvidoprobl = ValorTotal + qtdeliberar

TBCorretiva!Observacao = txtObservacao.Text
TBCorretiva.Update

Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select EC.* from estoque_controle EC INNER JOIN Estoque_movimentacao EM ON EC.IDestoque = EM.IDestoque where EM.Id_cfi = " & txtId & " and EM.Operacao = 'SAIDA_ALMOXARIFADO'", Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    If optDevolucao.Value = 1 Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
        TBAbrir.AddNew
        TBAbrir!Destino = "Interno"
        TBAbrir!Terceiros = False
        TBAbrir!Id_cfi = txtId
        TBAbrir!IDEstoque = TBEstoque!IDEstoque
        TBAbrir!Operacao = "DEVOLUCAO_ALMOXARIFADO C/ PROB."
        TBAbrir!Desenho = TBEstoque!Desenho
        TBAbrir!Descricao = TBEstoque!Descricao
        TBAbrir!Data = txtDataDevolucao
        TBAbrir!Entrada = qtdeliberar
        TBAbrir!Entrada_PC = qtdeliberar
        TBAbrir!Responsavel = pubUsuario
        TBAbrir!LOTE = TBEstoque!LOTE
        TBAbrir!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
        TBAbrir!vlrTotal = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * qtdeliberar, "###,##0.00")
        TBAbrir!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
        TBAbrir!Destino = "Interno"
        TBAbrir.Update
    End If
    
    If quantnovo > 0 Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
        TBAbrir.AddNew
        TBAbrir!Destino = "Interno"
        TBAbrir!Terceiros = False
        TBAbrir!Id_cfi = txtId
        TBAbrir!IDEstoque = TBEstoque!IDEstoque
        TBAbrir!Operacao = "DEVOLUCAO_ALMOXARIFADO"
        TBAbrir!Desenho = TBEstoque!Desenho
        TBAbrir!Descricao = TBEstoque!Descricao
        TBAbrir!Data = txtDataDevolucao
        TBAbrir!Entrada = quantnovo
        TBAbrir!Entrada_PC = quantnovo
        TBAbrir!Responsavel = pubUsuario
        TBAbrir!LOTE = txtLote.Text
        TBAbrir!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
        TBAbrir!vlrTotal = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * quantnovo, "###,##0.00")
        TBAbrir!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
        
           Set TBFamilia = CreateObject("adodb.recordset")
           TBFamilia.Open "select * from ProjFamilia where Familia = '" & TBEstoque!Classe & "'", Conexao, adOpenKeyset, adLockOptimistic
           If TBFamilia.EOF = False Then
           TBAbrir!Grupo = TBFamilia!Grupo
           End If
           TBFamilia.Close
        
        TBAbrir!Destino = "Interno"
        TBAbrir.Update
        TBAbrir.Close
        
        'Atualiza valor do material no estoque
        'Estoque_controle
        TBEstoque!estoque_real = Format(TBEstoque!estoque_real + quantnovo, "###,##0.00")
        TBEstoque!estoque_real_PC = TBEstoque!estoque_real
        TBEstoque!estoque_venda = Format(TBEstoque!estoque_venda + quantnovo, "###,##0.00")
        quantestoque = IIf(IsNull(TBEstoque!estoque_real), 0, TBEstoque!estoque_real)
        TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * quantestoque, "###,##0.00")
        TBEstoque.Update
    End If
End If
TBEstoque.Close

USMsgBox ("Devolução efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = Formulario
Evento = "Devolver"
ID_documento = txtId
Documento = "Cód. interno: " & txtCodinterno.Text & " - RE: " & txt_RE & " - Lote: " & txtLote.Text
Documento1 = ""
ProcGravaEvento
'==================================
ProcLimpaCampos
With frmCFI
    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    .ProcLimpaCampos
    .Frame2.Enabled = True
End With
Unload Me

qtdeliberada = 0
quantestoque = 0
quantidade = 0
quantnovo = 0
qtdeliberada = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from CFI_itens WHERE id = " & .ListItems.Item(InitFor)
            '==================================
            Modulo = Formulario
            Evento = "Excluir produto de destino/aplicação"
            ID_documento = Lista.SelectedItem
            Documento = "Cód. interno: " & .ListItems.Item(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

frmCFI_locprod.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcDevolver
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from CFI_Itens where ID_CFI = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codproduto), "", TBLISTA!Codproduto)
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "select * from Projproduto where codproduto = " & TBLISTA!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBItem!Descricao), "", Trim(TBItem!Descricao))
                .Item(.Count).SubItems(4) = IIf(IsNull(TBItem!Classe), "", Trim(TBItem!Classe))
                .Item(.Count).SubItems(5) = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
            End If
            TBItem.Close
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
