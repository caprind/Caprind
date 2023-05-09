VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpcoesGeral 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Configurações do sistema - Opções gerais"
   ClientHeight    =   10035
   ClientLeft      =   180
   ClientTop       =   450
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   17648
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Dados empresa"
      TabPicture(0)   =   "frmOpcoesGeral.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSTabEmpresa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CommonDialog1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Cadastro de moedas"
      TabPicture(1)   =   "frmOpcoesGeral.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(33)"
      Tab(1).Control(1)=   "txtidmoeda"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(3)=   "ListaMoeda"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Cadastro de unidades"
      TabPicture(2)   =   "frmOpcoesGeral.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(1)=   "SSTab2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Condições. de pgto./receb."
      TabPicture(3)   =   "frmOpcoesGeral.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(36)"
      Tab(3).Control(1)=   "USToolBar6"
      Tab(3).Control(2)=   "SSTab3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Feriados"
      TabPicture(4)   =   "frmOpcoesGeral.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(38)"
      Tab(4).Control(1)=   "Txt_ID_feriado"
      Tab(4).Control(2)=   "Frame1(37)"
      Tab(4).Control(3)=   "USToolBar7"
      Tab(4).Control(4)=   "Lista_feriado"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Configurações do sistema"
      TabPicture(5)   =   "frmOpcoesGeral.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "PBLista"
      Tab(5).Control(1)=   "ListaBancos"
      Tab(5).Control(2)=   "USToolBar1"
      Tab(5).Control(3)=   "Frame1(1)"
      Tab(5).Control(4)=   "Frame1(2)"
      Tab(5).ControlCount=   5
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   1425
         Index           =   2
         Left            =   -74970
         TabIndex        =   353
         Top             =   2220
         Width           =   15195
         Begin VB.CommandButton cmdLocalnovo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   14670
            Picture         =   "frmOpcoesGeral.frx":00A8
            TabIndex        =   357
            ToolTipText     =   "Localizar."
            Top             =   960
            Width           =   315
         End
         Begin VB.CommandButton cmdLocalantigo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   14670
            Picture         =   "frmOpcoesGeral.frx":01AA
            TabIndex        =   356
            ToolTipText     =   "Localizar."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtlocalnovo 
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   355
            TabStop         =   0   'False
            ToolTipText     =   "Local dos arquivos atualizados."
            Top             =   960
            Width           =   14475
         End
         Begin VB.TextBox txtlocalantigo 
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   354
            TabStop         =   0   'False
            ToolTipText     =   "Local dos arquivos antigos."
            Top             =   390
            Width           =   14475
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local dos arquivos atualizados*"
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
            Left            =   6285
            TabIndex        =   359
            Top             =   765
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local dos arquivos antigos*"
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
            Left            =   6427
            TabIndex        =   358
            Top             =   180
            Width           =   1980
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Index           =   1
         Left            =   -74970
         TabIndex        =   345
         Top             =   1335
         Width           =   15195
         Begin VB.TextBox txtLocalrel 
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
            Left            =   180
            MaxLength       =   100
            TabIndex        =   349
            ToolTipText     =   "Local onde está armazenado os relatórios."
            Top             =   390
            Width           =   5205
         End
         Begin VB.ComboBox Cmb_servidor 
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
            Left            =   5730
            Sorted          =   -1  'True
            TabIndex        =   348
            ToolTipText     =   "Nome da instância SQL."
            Top             =   390
            Width           =   4875
         End
         Begin VB.ComboBox Cmb_nome_banco 
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
            Left            =   10620
            Sorted          =   -1  'True
            TabIndex        =   347
            ToolTipText     =   "Nome do banco de dados."
            Top             =   390
            Width           =   3650
         End
         Begin VB.CommandButton Cmd_localizar_rel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   5400
            Picture         =   "frmOpcoesGeral.frx":02AC
            TabIndex        =   346
            ToolTipText     =   "Localizar."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome da instância  SQL*                                                     Nome do banco de dados*"
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
            Left            =   7470
            TabIndex        =   351
            Top             =   210
            Width           =   6060
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local dos relatórios do sistema*"
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
            Left            =   1890
            TabIndex        =   350
            Top             =   210
            Width           =   2280
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar"
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
         Index           =   38
         Left            =   -60930
         TabIndex        =   342
         Top             =   1290
         Width           =   1155
         Begin VB.ComboBox Cmb_ano_feriado 
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
            ItemData        =   "frmOpcoesGeral.frx":03AE
            Left            =   180
            List            =   "frmOpcoesGeral.frx":03B0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   343
            ToolTipText     =   "Ano."
            Top             =   390
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano"
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
            Index           =   26
            Left            =   435
            TabIndex        =   344
            Top             =   180
            Width           =   285
         End
      End
      Begin VB.TextBox Txt_ID_feriado 
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
         Left            =   -74220
         TabIndex        =   340
         Text            =   "0"
         Top             =   3930
         Visible         =   0   'False
         Width           =   855
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
         Index           =   37
         Left            =   -74970
         TabIndex        =   331
         Top             =   1290
         Width           =   14025
         Begin VB.TextBox Txt_data_feriado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   334
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   875
         End
         Begin VB.TextBox Txt_descricao_feriado 
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
            Left            =   7020
            MaxLength       =   50
            TabIndex        =   333
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   6810
         End
         Begin VB.TextBox Txt_responsavel_feriado 
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
            Left            =   1070
            Locked          =   -1  'True
            TabIndex        =   332
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   4740
         End
         Begin MSComCtl2.DTPicker Cmb_data_feriado 
            Height          =   315
            Left            =   5820
            TabIndex        =   335
            ToolTipText     =   "Data do feriado."
            Top             =   390
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
            Format          =   885653507
            CurrentDate     =   39057
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   450
            TabIndex        =   339
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
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
            Index           =   63
            Left            =   2983
            TabIndex        =   338
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. do feriado*"
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
            Index           =   64
            Left            =   5865
            TabIndex        =   337
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   " Descrição*"
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
            Index           =   65
            Left            =   10013
            TabIndex        =   336
            Top             =   180
            Width           =   825
         End
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
         Index           =   36
         Left            =   -74970
         TabIndex        =   312
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_data_cond 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   316
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   875
         End
         Begin VB.TextBox Txt_texto_cond 
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
            Left            =   7770
            TabIndex        =   315
            ToolTipText     =   "Condição de pagamento/recebimento."
            Top             =   390
            Width           =   6165
         End
         Begin VB.TextBox Txt_responsavel_cond 
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
            Left            =   1070
            Locked          =   -1  'True
            TabIndex        =   314
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   6690
         End
         Begin VB.ComboBox Cmb_aplicacao_cond 
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
            ItemData        =   "frmOpcoesGeral.frx":03B2
            Left            =   13950
            List            =   "frmOpcoesGeral.frx":03BC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   313
            ToolTipText     =   "Tipo."
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   445
            TabIndex        =   320
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aplicação*"
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
            Left            =   14100
            TabIndex        =   319
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label1 
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
            Index           =   61
            Left            =   3958
            TabIndex        =   318
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Condição de pagamento/recebimento*"
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
            Index           =   62
            Left            =   9465
            TabIndex        =   317
            Top             =   180
            Width           =   2775
         End
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
         Index           =   33
         Left            =   -74970
         TabIndex        =   274
         Top             =   1290
         Width           =   15195
         Begin VB.TextBox Txt_responsavel_moeda 
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
            Left            =   1070
            Locked          =   -1  'True
            TabIndex        =   278
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   10950
         End
         Begin VB.TextBox txtSimbolo 
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
            Left            =   13995
            MaxLength       =   3
            TabIndex        =   277
            ToolTipText     =   "Símbolo."
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox txtMoeda 
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
            Left            =   12030
            MaxLength       =   10
            TabIndex        =   276
            ToolTipText     =   "Moeda."
            Top             =   390
            Width           =   1955
         End
         Begin VB.TextBox Txt_data_moeda 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   275
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   875
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   282
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moeda*"
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
            Left            =   12722
            TabIndex        =   281
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label1 
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
            Index           =   52
            Left            =   6088
            TabIndex        =   280
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Símbolo*"
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
            Index           =   54
            Left            =   14182
            TabIndex        =   279
            Top             =   180
            Width           =   630
         End
      End
      Begin VB.TextBox txtidmoeda 
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
         Left            =   -74250
         TabIndex        =   272
         Text            =   "0"
         ToolTipText     =   "Unidade."
         Top             =   4950
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9210
         Top             =   4530
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TabDlg.SSTab SSTabEmpresa 
         Height          =   8670
         Left            =   0
         TabIndex        =   1
         Top             =   1320
         Width           =   15300
         _ExtentX        =   26988
         _ExtentY        =   15293
         _Version        =   393216
         Tabs            =   6
         Tab             =   2
         TabsPerRow      =   6
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
         TabCaption(0)   =   "Dados gerais"
         TabPicture(0)   =   "frmOpcoesGeral.frx":03D1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Lista_empresas"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtidempresa"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Regime tributário/Impostos"
         TabPicture(1)   =   "frmOpcoesGeral.frx":03ED
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(4)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame1(5)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Dados adicionais"
         TabPicture(2)   =   "frmOpcoesGeral.frx":0409
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame1(39)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame1(40)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "E-mail"
         TabPicture(3)   =   "frmOpcoesGeral.frx":0425
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Lista_email"
         Tab(3).Control(1)=   "Frame1(30)"
         Tab(3).Control(2)=   "Txt_ID_email"
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "Filtros"
         TabPicture(4)   =   "frmOpcoesGeral.frx":0441
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Lista_filtros"
         Tab(4).Control(1)=   "txtID_Filtros"
         Tab(4).Control(2)=   "Frame1(31)"
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "Armazenamento (.PDF)"
         TabPicture(5)   =   "frmOpcoesGeral.frx":045D
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Lista_armaz"
         Tab(5).Control(1)=   "Frame1(32)"
         Tab(5).Control(2)=   "Txt_ID_armaz"
         Tab(5).ControlCount=   3
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Impostos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7605
            Index           =   5
            Left            =   -74945
            TabIndex        =   188
            Top             =   1020
            Width           =   15195
            Begin VB.TextBox txtID_imposto 
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
               Left            =   4860
               TabIndex        =   189
               Text            =   "0"
               Top             =   3060
               Visible         =   0   'False
               Width           =   795
            End
            Begin TabDlg.SSTab SSTab5 
               Height          =   7305
               Left            =   0
               TabIndex        =   190
               Top             =   270
               Width           =   15195
               _ExtentX        =   26802
               _ExtentY        =   12885
               _Version        =   393216
               Tabs            =   2
               Tab             =   1
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
               TabCaption(0)   =   "Lucro presumido | Lucro real | Simples nacional (excesso de sublimite de receita bruta)"
               TabPicture(0)   =   "frmOpcoesGeral.frx":0479
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "Frame1(43)"
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Simples nacional"
               TabPicture(1)   =   "frmOpcoesGeral.frx":0495
               Tab(1).ControlEnabled=   -1  'True
               Tab(1).Control(0)=   "USToolBar8"
               Tab(1).Control(0).Enabled=   0   'False
               Tab(1).Control(1)=   "Lista_TBSN"
               Tab(1).Control(1).Enabled=   0   'False
               Tab(1).Control(2)=   "Frame1(45)"
               Tab(1).Control(2).Enabled=   0   'False
               Tab(1).Control(3)=   "Txt_ID_TBSN"
               Tab(1).Control(3).Enabled=   0   'False
               Tab(1).Control(4)=   "Frame1(47)"
               Tab(1).Control(4).Enabled=   0   'False
               Tab(1).Control(5)=   "Frame1(19)"
               Tab(1).Control(5).Enabled=   0   'False
               Tab(1).ControlCount=   6
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1095
                  Index           =   19
                  Left            =   30
                  TabIndex        =   262
                  Top             =   330
                  Width           =   15105
                  Begin VB.TextBox Txt_ICMS_ind 
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
                     Left            =   10170
                     MaxLength       =   50
                     TabIndex        =   373
                     ToolTipText     =   "Alíquota de ICMS para industrialização."
                     Top             =   360
                     Width           =   705
                  End
                  Begin VB.CheckBox chkDuplicata 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Não reter PIS/Cofins no desconto de duplicatas"
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
                     Left            =   11280
                     TabIndex        =   372
                     Top             =   450
                     Width           =   3705
                  End
                  Begin VB.TextBox Txt_valor_total_faturado 
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
                     Left            =   210
                     Locked          =   -1  'True
                     MaxLength       =   50
                     TabIndex        =   370
                     TabStop         =   0   'False
                     ToolTipText     =   "Valor total faturado antes da emissão da primeira nota fiscal no Caprind."
                     Top             =   360
                     Width           =   1605
                  End
                  Begin DrawSuite2022.USButton Cmd_valor_faturado_mes 
                     Height          =   330
                     Left            =   210
                     TabIndex        =   369
                     ToolTipText     =   "Cadastrar valor total faturado nos ultimos doze meses."
                     Top             =   690
                     Width           =   1605
                     _ExtentX        =   2831
                     _ExtentY        =   582
                     DibPicture      =   "frmOpcoesGeral.frx":04B1
                     Caption         =   "Tabela 12 meses"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
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
                     PicSize         =   1
                     Theme           =   5
                  End
                  Begin VB.ComboBox Cmb_tipo_TBSN 
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
                     ForeColor       =   &H00000000&
                     Height          =   330
                     ItemData        =   "frmOpcoesGeral.frx":A5D4
                     Left            =   2940
                     List            =   "frmOpcoesGeral.frx":A5E7
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   264
                     ToolTipText     =   "Tipo da tabela do simples nacional."
                     Top             =   360
                     Width           =   5400
                  End
                  Begin VB.TextBox txtCNAE_TBSN 
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
                     Left            =   8370
                     MaxLength       =   7
                     TabIndex        =   263
                     ToolTipText     =   "CNAE."
                     Top             =   360
                     Width           =   975
                  End
                  Begin DrawSuite2022.USButton Cmd_ativar_tabelaSN 
                     Height          =   330
                     Left            =   2940
                     TabIndex        =   265
                     ToolTipText     =   "Ativar/Desativar tabela"
                     Top             =   720
                     Width           =   5370
                     _ExtentX        =   9472
                     _ExtentY        =   582
                     Caption         =   "Ativar/Desativar tabela"
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
                     HandPointer     =   0   'False
                     PicAlign        =   8
                     PicSize         =   4
                     PicSizeH        =   48
                     PicSizeW        =   48
                     Theme           =   4
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Aliq.ICMS"
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
                     Index           =   28
                     Left            =   10185
                     TabIndex        =   374
                     Top             =   180
                     Width           =   690
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Total faturado"
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
                     Left            =   480
                     TabIndex        =   371
                     Top             =   180
                     Width           =   1035
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Tipo da tabela*"
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
                     Index           =   84
                     Left            =   2985
                     TabIndex        =   268
                     Top             =   180
                     Width           =   1110
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "CNAE"
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
                     Index           =   86
                     Left            =   8655
                     TabIndex        =   267
                     Top             =   180
                     Width           =   405
                  End
                  Begin VB.Label Lbl_status 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Status: Ativada"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   -1  'True
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000080&
                     Height          =   270
                     Left            =   4470
                     TabIndex        =   266
                     Top             =   120
                     Width           =   2460
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   5655
                  Index           =   44
                  Left            =   -74970
                  TabIndex        =   256
                  Top             =   330
                  Width           =   15105
                  Begin VB.CheckBox Check37 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear apontamento sem baixar toda a lista de requisição da ordem"
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
                     Left            =   180
                     TabIndex        =   261
                     Top             =   792
                     Width           =   7215
                  End
                  Begin VB.CheckBox Check36 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Apontamento por código"
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
                     Left            =   180
                     TabIndex        =   260
                     Top             =   300
                     Width           =   2415
                  End
                  Begin VB.CheckBox Check35 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear apontamento sem baixar matéria-prima da lista de requisição da ordem"
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
                     Left            =   180
                     TabIndex        =   259
                     Top             =   546
                     Width           =   7215
                  End
                  Begin VB.CheckBox Check34 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Desbloquear primeiro apontamento de OS com processo controlado"
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
                     Left            =   180
                     TabIndex        =   258
                     Top             =   1038
                     Width           =   6045
                  End
                  Begin VB.CheckBox Check33 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Carregar posto de trabalho por grupo no Gerprod"
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
                     Left            =   180
                     TabIndex        =   257
                     Top             =   1284
                     Width           =   5355
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   6915
                  Index           =   43
                  Left            =   -74970
                  TabIndex        =   217
                  Top             =   360
                  Width           =   15105
                  Begin VB.Frame Frame1 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Sobre produtos"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1065
                     Index           =   18
                     Left            =   0
                     TabIndex        =   247
                     Top             =   1830
                     Width           =   4155
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "IRPJ (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   24
                        Left            =   3078
                        TabIndex        =   254
                        Top             =   240
                        Width           =   915
                        Begin VB.TextBox txtIRPJ1 
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
                           Left            =   90
                           MaxLength       =   50
                           TabIndex        =   255
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "CSLL (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   23
                        Left            =   2112
                        TabIndex        =   252
                        Top             =   240
                        Width           =   915
                        Begin VB.TextBox txtCSLL1 
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
                           Left            =   90
                           MaxLength       =   50
                           TabIndex        =   253
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "Cofins (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   22
                        Left            =   1146
                        TabIndex        =   250
                        Top             =   240
                        Width           =   915
                        Begin VB.TextBox txtCofins1 
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
                           Left            =   90
                           MaxLength       =   50
                           TabIndex        =   251
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "PIS (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   21
                        Left            =   180
                        TabIndex        =   248
                        Top             =   240
                        Width           =   915
                        Begin VB.TextBox txtPIS1 
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
                           Left            =   90
                           MaxLength       =   50
                           TabIndex        =   249
                           Top             =   270
                           Width           =   705
                        End
                     End
                  End
                  Begin VB.Frame Frame1 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Sobre serviços"
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
                     Height          =   1815
                     Index           =   6
                     Left            =   0
                     TabIndex        =   218
                     Top             =   0
                     Width           =   9165
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "Valor acima de"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   13
                        Left            =   7290
                        TabIndex        =   245
                        Top             =   240
                        Width           =   1725
                        Begin VB.TextBox txtVLR 
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
                           Left            =   90
                           MaxLength       =   50
                           TabIndex        =   246
                           Top             =   270
                           Width           =   1455
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "Valor acima de"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   15
                        Left            =   1602
                        TabIndex        =   243
                        Top             =   990
                        Width           =   2235
                        Begin VB.TextBox txtVLR1 
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
                           Left            =   90
                           MaxLength       =   50
                           TabIndex        =   244
                           Top             =   270
                           Width           =   2025
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "IRRF (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   12
                        Left            =   5868
                        TabIndex        =   240
                        Top             =   240
                        Width           =   1395
                        Begin VB.CommandButton cmdIRRF 
                           Appearance      =   0  'Flat
                           BackColor       =   &H00C0C0C0&
                           Height          =   315
                           Left            =   900
                           Picture         =   "frmOpcoesGeral.frx":A782
                           Style           =   1  'Graphical
                           TabIndex        =   242
                           ToolTipText     =   "Localizar dados para criar contas a pagar."
                           Top             =   270
                           Width           =   315
                        End
                        Begin VB.TextBox txtIRRF 
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
                           MaxLength       =   50
                           TabIndex        =   241
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "IRPJ (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   16
                        Left            =   3900
                        TabIndex        =   238
                        Top             =   990
                        Width           =   915
                        Begin VB.TextBox txtIRPJ 
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
                           Left            =   90
                           MaxLength       =   50
                           TabIndex        =   239
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "CSLL (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   10
                        Left            =   3024
                        TabIndex        =   235
                        Top             =   240
                        Width           =   1395
                        Begin VB.CommandButton cmdCSLL 
                           Appearance      =   0  'Flat
                           BackColor       =   &H00C0C0C0&
                           Height          =   315
                           Left            =   900
                           Picture         =   "frmOpcoesGeral.frx":A884
                           Style           =   1  'Graphical
                           TabIndex        =   237
                           ToolTipText     =   "Localizar dados para criar contas a pagar."
                           Top             =   270
                           Width           =   315
                        End
                        Begin VB.TextBox txtCSLL 
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
                           MaxLength       =   50
                           TabIndex        =   236
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "Cofins (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   9
                        Left            =   1602
                        TabIndex        =   232
                        Top             =   240
                        Width           =   1395
                        Begin VB.CommandButton cmdCofins 
                           Appearance      =   0  'Flat
                           BackColor       =   &H00C0C0C0&
                           Height          =   315
                           Left            =   900
                           Picture         =   "frmOpcoesGeral.frx":A986
                           Style           =   1  'Graphical
                           TabIndex        =   234
                           ToolTipText     =   "Localizar dados para criar contas a pagar."
                           Top             =   270
                           Width           =   315
                        End
                        Begin VB.TextBox txtCofins 
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
                           MaxLength       =   50
                           TabIndex        =   233
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "PIS (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   8
                        Left            =   180
                        TabIndex        =   229
                        Top             =   240
                        Width           =   1395
                        Begin VB.CommandButton cmdPIS 
                           Appearance      =   0  'Flat
                           BackColor       =   &H00C0C0C0&
                           Height          =   315
                           Left            =   900
                           Picture         =   "frmOpcoesGeral.frx":AA88
                           Style           =   1  'Graphical
                           TabIndex        =   231
                           ToolTipText     =   "Localizar dados para criar contas a pagar."
                           Top             =   270
                           Width           =   315
                        End
                        Begin VB.TextBox txtPIS 
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
                           MaxLength       =   50
                           TabIndex        =   230
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "INSS (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   14
                        Left            =   180
                        TabIndex        =   226
                        Top             =   990
                        Width           =   1395
                        Begin VB.CommandButton cmdINSS 
                           Appearance      =   0  'Flat
                           BackColor       =   &H00C0C0C0&
                           Height          =   315
                           Left            =   900
                           Picture         =   "frmOpcoesGeral.frx":AB8A
                           Style           =   1  'Graphical
                           TabIndex        =   228
                           ToolTipText     =   "Localizar dados para criar contas a pagar."
                           Top             =   270
                           Width           =   315
                        End
                        Begin VB.TextBox txtINSS 
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
                           MaxLength       =   50
                           TabIndex        =   227
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "ISSQN (%)"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   11
                        Left            =   4446
                        TabIndex        =   223
                        Top             =   240
                        Width           =   1395
                        Begin VB.CommandButton cmdISSQN 
                           Appearance      =   0  'Flat
                           BackColor       =   &H00C0C0C0&
                           Height          =   315
                           Left            =   900
                           Picture         =   "frmOpcoesGeral.frx":AC8C
                           Style           =   1  'Graphical
                           TabIndex        =   225
                           ToolTipText     =   "Localizar dados para criar contas a pagar."
                           Top             =   270
                           Width           =   315
                        End
                        Begin VB.TextBox txtISS 
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
                           MaxLength       =   50
                           TabIndex        =   224
                           Top             =   270
                           Width           =   705
                        End
                     End
                     Begin VB.Frame Frame1 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "Alíquota IRPJ sobre faturamento anual"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   705
                        Index           =   17
                        Left            =   4830
                        TabIndex        =   219
                        Top             =   990
                        Width           =   3615
                        Begin VB.TextBox txtIRPJ_serv_maior 
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
                           Left            =   660
                           MaxLength       =   50
                           TabIndex        =   221
                           Top             =   270
                           Width           =   2115
                        End
                        Begin VB.TextBox txtIRPJ_serv 
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
                           Left            =   2790
                           MaxLength       =   50
                           TabIndex        =   220
                           Top             =   270
                           Width           =   705
                        End
                        Begin VB.Label Label2 
                           AutoSize        =   -1  'True
                           BackColor       =   &H00C0C0C0&
                           BackStyle       =   0  'Transparent
                           Caption         =   "maior :"
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
                           Left            =   120
                           TabIndex        =   222
                           Top             =   270
                           Width           =   495
                        End
                     End
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Alíquotas (%)"
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
                  Height          =   885
                  Index           =   47
                  Left            =   3360
                  TabIndex        =   197
                  Top             =   2460
                  Width           =   11775
                  Begin VB.TextBox Txt_Aliquota_TBSN 
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
                     Left            =   180
                     MaxLength       =   255
                     TabIndex        =   206
                     ToolTipText     =   "Alíquota."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.TextBox Txt_ICMS_TBSN 
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
                     Left            =   7110
                     MaxLength       =   255
                     TabIndex        =   205
                     ToolTipText     =   "Alíquota do ICMS."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.TextBox Txt_valor_deduzir_TBSN 
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
                     Left            =   8100
                     MaxLength       =   255
                     TabIndex        =   204
                     ToolTipText     =   "Valor até."
                     Top             =   390
                     Width           =   1455
                  End
                  Begin VB.TextBox Txt_Cofins_TBSN 
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
                     Left            =   3150
                     MaxLength       =   255
                     TabIndex        =   203
                     ToolTipText     =   "Alíquota do Cofins."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.TextBox Txt_CSLL_TBSN 
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
                     Left            =   2160
                     MaxLength       =   255
                     TabIndex        =   202
                     ToolTipText     =   "Alíquota do CSLL."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.TextBox Txt_IRPJ_TBSN 
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
                     Left            =   1170
                     MaxLength       =   255
                     TabIndex        =   201
                     ToolTipText     =   "Alíquota do IRPJ."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.TextBox Txt_PIS_TBSN 
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
                     Left            =   4140
                     MaxLength       =   255
                     TabIndex        =   200
                     ToolTipText     =   "Alíquota do PIS."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.TextBox Txt_IPI_TBSN 
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
                     Left            =   6120
                     MaxLength       =   255
                     TabIndex        =   199
                     ToolTipText     =   "Alíquota do IPI."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.TextBox Txt_CPP_TBSN 
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
                     Left            =   5130
                     MaxLength       =   255
                     TabIndex        =   198
                     ToolTipText     =   "Alíquota do CPP."
                     Top             =   390
                     Width           =   975
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Alíquota*"
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
                     Index           =   89
                     Left            =   330
                     TabIndex        =   216
                     Top             =   210
                     Width           =   675
                  End
                  Begin VB.Label Lbl_ICMS_TBSN 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ICMS*"
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
                     Left            =   7365
                     TabIndex        =   215
                     Top             =   210
                     Width           =   465
                  End
                  Begin VB.Label Lbl_ISS_TBSN 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ISS*"
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
                     Left            =   7425
                     TabIndex        =   214
                     Top             =   210
                     Visible         =   0   'False
                     Width           =   330
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Valor a deduzir*"
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
                     Left            =   8250
                     TabIndex        =   213
                     Top             =   210
                     Width           =   1155
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "IPI*"
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
                     Index           =   92
                     Left            =   6450
                     TabIndex        =   212
                     Top             =   210
                     Width           =   300
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "PIS*"
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
                     Index           =   91
                     Left            =   4462
                     TabIndex        =   211
                     Top             =   210
                     Width           =   330
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cofins*"
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
                     Index           =   90
                     Left            =   3367
                     TabIndex        =   210
                     Top             =   210
                     Width           =   540
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "IRPJ*"
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
                     Index           =   33
                     Left            =   1447
                     TabIndex        =   209
                     Top             =   210
                     Width           =   420
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "CSLL*"
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
                     Index           =   34
                     Left            =   2430
                     TabIndex        =   208
                     Top             =   210
                     Width           =   435
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "CPP*"
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
                     Index           =   93
                     Left            =   5430
                     TabIndex        =   207
                     Top             =   210
                     Width           =   375
                  End
               End
               Begin VB.TextBox Txt_ID_TBSN 
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
                  Left            =   1410
                  MaxLength       =   60
                  TabIndex        =   196
                  Text            =   "0"
                  Top             =   5310
                  Visible         =   0   'False
                  Width           =   950
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Receita bruta em 12 meses (em R$)"
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
                  Height          =   885
                  Index           =   45
                  Left            =   30
                  TabIndex        =   191
                  Top             =   2460
                  Width           =   3315
                  Begin VB.TextBox Txt_de_TBSN 
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
                     Left            =   180
                     MaxLength       =   255
                     TabIndex        =   193
                     ToolTipText     =   "Valor de."
                     Top             =   390
                     Width           =   1485
                  End
                  Begin VB.TextBox Txt_ate_TBSN 
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
                     Left            =   1680
                     MaxLength       =   255
                     TabIndex        =   192
                     ToolTipText     =   "Valor até."
                     Top             =   390
                     Width           =   1455
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "De*"
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
                     Index           =   87
                     Left            =   780
                     TabIndex        =   195
                     Top             =   210
                     Width           =   285
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00C0C0C0&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Até*"
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
                     Index           =   88
                     Left            =   2235
                     TabIndex        =   194
                     Top             =   210
                     Width           =   345
                  End
               End
               Begin MSComctlLib.ListView Lista_TBSN 
                  Height          =   3735
                  Left            =   30
                  TabIndex        =   269
                  Top             =   3360
                  Width           =   15105
                  _ExtentX        =   26644
                  _ExtentY        =   6588
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  Checkboxes      =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
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
                  NumItems        =   12
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Object.Tag             =   "N"
                     Object.Width           =   512
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   1
                     Object.Tag             =   "N"
                     Text            =   "De"
                     Object.Width           =   2615
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   2
                     Object.Tag             =   "N"
                     Text            =   "Até"
                     Object.Width           =   2615
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   3
                     Object.Tag             =   "N"
                     Text            =   "Alíquota (%)"
                     Object.Width           =   2293
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   4
                     Object.Tag             =   "N"
                     Text            =   "IRPJ (%)"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   5
                     Object.Tag             =   "N"
                     Text            =   "CSLL (%)"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   6
                     Object.Tag             =   "N"
                     Text            =   "Cofins (%)"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   7
                     Object.Tag             =   "N"
                     Text            =   "PIS (%)"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   8
                     Object.Tag             =   "N"
                     Text            =   "CPP (%)"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   9
                     Object.Tag             =   "N"
                     Text            =   "IPI (%)"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   10
                     Object.Tag             =   "N"
                     Text            =   "ICMS (%)"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   11
                     Object.Tag             =   "N"
                     Text            =   "Vlr. a deduzir"
                     Object.Width           =   2117
                  EndProperty
               End
               Begin DrawSuite2022.USToolBar USToolBar8 
                  Height          =   975
                  Left            =   30
                  TabIndex        =   270
                  Top             =   1440
                  Width           =   15105
                  _ExtentX        =   26644
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
                  ButtonToolTipText2=   "Salvar (F7)"
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
                  ButtonCaption4  =   "Atualizar"
                  ButtonEnabled4  =   0   'False
                  ButtonIconSize4 =   32
                  ButtonToolTipText4=   "Atualizar tabela do simples nacional nos registros."
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
                  ButtonWidth4    =   50
                  ButtonHeight4   =   21
                  ButtonUseMaskColor4=   0   'False
                  ButtonEnabled5  =   0   'False
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
                  ButtonLeft5     =   170
                  ButtonTop5      =   2
                  ButtonWidth5    =   24
                  ButtonHeight5   =   24
                  ButtonUseMaskColor5=   0   'False
                  Begin DrawSuite2022.USImageList USImageList8 
                     Left            =   6180
                     Top             =   180
                     _ExtentX        =   900
                     _ExtentY        =   767
                     Img1            =   "frmOpcoesGeral.frx":AD8E
                     Count           =   1
                  End
               End
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Optante pelo regime tributário"
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
            Height          =   675
            Index           =   4
            Left            =   -74945
            TabIndex        =   180
            Top             =   330
            Width           =   15195
            Begin VB.OptionButton optSimples 
               BackColor       =   &H00E0E0E0&
               Caption         =   "1 - Simples nacional "
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
               Left            =   210
               TabIndex        =   184
               Top             =   330
               Width           =   1815
            End
            Begin VB.OptionButton optPresumido 
               BackColor       =   &H00E0E0E0&
               Caption         =   "2 - Lucro presumido"
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
               Left            =   3405
               TabIndex        =   183
               Top             =   330
               Width           =   2325
            End
            Begin VB.OptionButton optReal 
               BackColor       =   &H00E0E0E0&
               Caption         =   "3 - Lucro real"
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
               Left            =   6165
               TabIndex        =   182
               Top             =   330
               Width           =   1545
            End
            Begin VB.OptionButton optSimples1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "4 - Simples nacional (excesso de sublimite de receita bruta)"
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
               Left            =   8190
               TabIndex        =   181
               Top             =   330
               Width           =   4575
            End
         End
         Begin VB.TextBox txtidempresa 
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
            Left            =   -70140
            TabIndex        =   179
            Text            =   "0"
            Top             =   6030
            Visible         =   0   'False
            Width           =   855
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
            Height          =   3855
            Index           =   3
            Left            =   -74940
            TabIndex        =   124
            Top             =   330
            Width           =   15195
            Begin DrawSuite2022.USButton BtnLogotipo 
               Height          =   795
               Left            =   12240
               TabIndex        =   162
               ToolTipText     =   "Localizar logotipo empresa (125 px alt. x 196 px larg. *.jpg ou *.bmp)"
               Top             =   2970
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   1402
               DibPicture      =   "frmOpcoesGeral.frx":D880
               Caption         =   "Logotipo"
               CaptionDistance =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8421504
               BorderColorDown =   7907521
               BorderColorOver =   7907521
               ForeColor       =   0
               ForeColorOver   =   0
               ForeColorDown   =   0
               GradientColor1  =   16777215
               GradientColor2  =   14737632
               GradientColor3  =   12632256
               GradientColor4  =   12632256
               GradientColorOver1=   14417407
               GradientColorOver2=   12317439
               GradientColorOver3=   4838399
               GradientColorOver4=   9627391
               GradientColorDown1=   10802943
               GradientColorDown2=   7979263
               GradientColorDown3=   4370174
               GradientColorDown4=   7395582
               PicAlign        =   7
               PicSize         =   2
               PicSizeH        =   24
               PicSizeW        =   24
               ShowFocusRect   =   0   'False
               Theme           =   1
               ToolTipCentered =   -1  'True
            End
            Begin VB.TextBox Txt_endereco_entrega 
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
               Left            =   180
               MaxLength       =   255
               TabIndex        =   153
               ToolTipText     =   "Endereço de entrega."
               Top             =   3390
               Width           =   7215
            End
            Begin VB.TextBox Txt_cod_SUFRAMA 
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
               Left            =   10635
               MaxLength       =   10
               TabIndex        =   151
               ToolTipText     =   "Código SUFRAMA."
               Top             =   2790
               Width           =   1470
            End
            Begin VB.TextBox txtComplemento 
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
               Left            =   10710
               MaxLength       =   30
               TabIndex        =   134
               ToolTipText     =   "Complemento."
               Top             =   990
               Width           =   1395
            End
            Begin VB.ComboBox Txt_pais 
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
               ItemData        =   "frmOpcoesGeral.frx":10ED0
               Left            =   9570
               List            =   "frmOpcoesGeral.frx":10ED2
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   140
               ToolTipText     =   "País."
               Top             =   1590
               Width           =   2535
            End
            Begin VB.ComboBox cmbTipo_endereco 
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
               ForeColor       =   &H00000000&
               Height          =   330
               ItemData        =   "frmOpcoesGeral.frx":10ED4
               Left            =   4290
               List            =   "frmOpcoesGeral.frx":10F0B
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   131
               ToolTipText     =   "Tipo do endereço."
               Top             =   990
               Width           =   1530
            End
            Begin VB.ComboBox cmbTipo_bairro 
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
               ForeColor       =   &H00000000&
               Height          =   330
               ItemData        =   "frmOpcoesGeral.frx":10F8A
               Left            =   180
               List            =   "frmOpcoesGeral.frx":10FCA
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   135
               ToolTipText     =   "Tipo do bairro."
               Top             =   1590
               Width           =   1650
            End
            Begin VB.TextBox txtCNAE 
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
               Left            =   9420
               MaxLength       =   7
               TabIndex        =   149
               ToolTipText     =   "CNAE fiscal."
               Top             =   2790
               Width           =   840
            End
            Begin VB.TextBox txtNumero 
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
               Left            =   9690
               MaxLength       =   60
               TabIndex        =   133
               ToolTipText     =   "Número."
               Top             =   990
               Width           =   1005
            End
            Begin VB.TextBox Txt_site 
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
               Left            =   7935
               MaxLength       =   150
               TabIndex        =   144
               ToolTipText     =   "Site."
               Top             =   2190
               Width           =   4170
            End
            Begin VB.ComboBox Cmb_uf 
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
               ForeColor       =   &H00000000&
               Height          =   330
               ItemData        =   "frmOpcoesGeral.frx":1107A
               Left            =   4800
               List            =   "frmOpcoesGeral.frx":110CF
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   137
               ToolTipText     =   "UF."
               Top             =   1590
               Width           =   750
            End
            Begin VB.TextBox Txt_bairro 
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
               Left            =   1830
               MaxLength       =   150
               TabIndex        =   136
               ToolTipText     =   "Bairro."
               Top             =   1590
               Width           =   2955
            End
            Begin VB.TextBox Txt_IM 
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
               Left            =   10380
               MaxLength       =   15
               TabIndex        =   129
               ToolTipText     =   "Número da inscrição municipal."
               Top             =   390
               Width           =   1725
            End
            Begin VB.TextBox txtRamo 
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
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
               Left            =   6600
               MaxLength       =   150
               TabIndex        =   147
               ToolTipText     =   "Ramo de atividade."
               Top             =   2790
               Width           =   2805
            End
            Begin VB.TextBox txtEndereco 
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
               Left            =   5830
               MaxLength       =   150
               TabIndex        =   132
               ToolTipText     =   "Endereço."
               Top             =   990
               Width           =   3845
            End
            Begin VB.TextBox txtEndereco_cob 
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
               Left            =   180
               MaxLength       =   255
               TabIndex        =   145
               ToolTipText     =   "Endereço de cobrança."
               Top             =   2790
               Width           =   6405
            End
            Begin VB.TextBox txtRazao 
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
               Left            =   180
               MaxLength       =   60
               TabIndex        =   126
               ToolTipText     =   "Razão social."
               Top             =   390
               Width           =   6855
            End
            Begin VB.TextBox txtEmpresa 
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
               Left            =   180
               MaxLength       =   50
               TabIndex        =   130
               ToolTipText     =   "Nome fantasia."
               Top             =   990
               Width           =   4095
            End
            Begin VB.TextBox txtRG_IE 
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
               Left            =   8700
               MaxLength       =   15
               TabIndex        =   128
               ToolTipText     =   "Número da inscrição estadual."
               Top             =   390
               Width           =   1665
            End
            Begin VB.ComboBox Cmb_cidade 
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
               ForeColor       =   &H00000000&
               Height          =   330
               ItemData        =   "frmOpcoesGeral.frx":1113F
               Left            =   5550
               List            =   "frmOpcoesGeral.frx":11141
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   138
               ToolTipText     =   "Cidade."
               Top             =   1590
               Width           =   2955
            End
            Begin VB.CheckBox Chk_atualizacao_autom 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Atualização automática"
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
               Left            =   7470
               TabIndex        =   157
               Top             =   3510
               Width           =   2415
            End
            Begin VB.TextBox Txt_email 
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
               Left            =   3690
               MaxLength       =   150
               TabIndex        =   143
               ToolTipText     =   "E-mail."
               Top             =   2190
               Width           =   4230
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Incentivador"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Index           =   26
               Left            =   9930
               TabIndex        =   125
               Top             =   3135
               Width           =   2175
               Begin VB.CheckBox chkCultural 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cultural"
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
                  TabIndex        =   159
                  Top             =   270
                  Width           =   915
               End
               Begin VB.CheckBox chkFiscal 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Fiscal"
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
                  Left            =   1110
                  TabIndex        =   161
                  Top             =   270
                  Width           =   855
               End
            End
            Begin VB.CheckBox chkPrincipal 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Empresa principal"
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
               Left            =   7470
               TabIndex        =   155
               Top             =   3270
               Width           =   2415
            End
            Begin MSMask.MaskEdBox txtcnpj 
               Height          =   315
               Left            =   7050
               TabIndex        =   127
               ToolTipText     =   "Número do CNPJ."
               Top             =   390
               Width           =   1635
               _ExtentX        =   2884
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
            Begin MSMask.MaskEdBox Txt_CEP 
               Height          =   315
               Left            =   8520
               TabIndex        =   139
               ToolTipText     =   "CEP."
               Top             =   1590
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "#####-###"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Txt_telefones 
               Height          =   315
               Left            =   180
               TabIndex        =   141
               ToolTipText     =   "Número do(s) telefone(s)."
               Top             =   2190
               Width           =   1740
               _ExtentX        =   3069
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
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Txt_fax 
               Height          =   315
               Left            =   1935
               TabIndex        =   142
               ToolTipText     =   "Número do fax."
               Top             =   2190
               Width           =   1740
               _ExtentX        =   3069
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
               PromptChar      =   "_"
            End
            Begin DrawSuite2022.USButton cmdConsultar 
               Height          =   795
               Left            =   13680
               TabIndex        =   364
               ToolTipText     =   "Consultar contrato de locação do sistema"
               Top             =   2970
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   1402
               DibPicture      =   "frmOpcoesGeral.frx":11143
               Caption         =   "Contrato"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8421504
               BorderColorDown =   7907521
               BorderColorOver =   7907521
               ForeColor       =   0
               ForeColorOver   =   0
               ForeColorDown   =   0
               GradientColor1  =   16777215
               GradientColor2  =   14737632
               GradientColor3  =   12632256
               GradientColor4  =   12632256
               GradientColorOver1=   14417407
               GradientColorOver2=   12317439
               GradientColorOver3=   4838399
               GradientColorOver4=   9627391
               GradientColorDown1=   10802943
               GradientColorDown2=   7979263
               GradientColorDown3=   4370174
               GradientColorDown4=   7395582
               PicAlign        =   7
               PicSize         =   2
               PicSizeH        =   24
               PicSizeW        =   24
               Theme           =   1
               ToolTipTitle    =   "CAPRIND v5.0"
            End
            Begin DrawSuite2022.USButton cmdCnae 
               Height          =   315
               Left            =   10260
               TabIndex        =   365
               ToolTipText     =   "Consultar cadastro na receita federal."
               Top             =   2790
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmOpcoesGeral.frx":1ABF0
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
               BorderColor     =   8421504
               BorderColorDown =   7907521
               BorderColorOver =   7907521
               ForeColor       =   0
               ForeColorOver   =   0
               ForeColorDown   =   0
               GradientColor1  =   16777215
               GradientColor2  =   14737632
               GradientColor3  =   12632256
               GradientColor4  =   12632256
               GradientColorOver1=   14417407
               GradientColorOver2=   12317439
               GradientColorOver3=   4838399
               GradientColorOver4=   9627391
               GradientColorDown1=   10802943
               GradientColorDown2=   7979263
               GradientColorDown3=   4370174
               GradientColorDown4=   7395582
               PicAlign        =   0
               Theme           =   1
               ToolTipTitle    =   "CAPRIND v5.0"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Endereço de entrega"
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
               Left            =   3030
               TabIndex        =   178
               Top             =   3180
               Width           =   1515
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo*"
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
               Index           =   41
               Left            =   810
               TabIndex        =   177
               Top             =   1380
               Width           =   390
            End
            Begin VB.Image picimagem 
               Appearance      =   0  'Flat
               Height          =   2505
               Left            =   12240
               Stretch         =   -1  'True
               ToolTipText     =   "Clique aqui para localizar imagem do logotipo da empresa."
               Top             =   390
               Width           =   2745
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Endereço de cobrança*"
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
               Left            =   2535
               TabIndex        =   176
               Top             =   2580
               Width           =   1695
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Telefone(s)"
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
               Left            =   638
               TabIndex        =   175
               Top             =   1980
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nome fantasia*"
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
               Left            =   1665
               TabIndex        =   174
               Top             =   780
               Width           =   1125
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Razão social*"
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
               Left            =   3075
               TabIndex        =   173
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CNPJ*"
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
               Index           =   66
               Left            =   7680
               TabIndex        =   172
               Top             =   180
               Width           =   465
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "IE*"
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
               Index           =   67
               Left            =   9412
               TabIndex        =   171
               Top             =   180
               Width           =   240
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "IM"
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
               Index           =   68
               Left            =   11152
               TabIndex        =   170
               Top             =   180
               Width           =   180
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo*"
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
               Index           =   69
               Left            =   4860
               TabIndex        =   169
               Top             =   780
               Width           =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Endereço*"
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
               Index           =   70
               Left            =   7370
               TabIndex        =   168
               Top             =   780
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número*"
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
               Index           =   71
               Left            =   9870
               TabIndex        =   167
               Top             =   780
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Complemento"
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
               Index           =   72
               Left            =   10920
               TabIndex        =   166
               Top             =   780
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bairro*"
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
               Index           =   73
               Left            =   3052
               TabIndex        =   165
               Top             =   1380
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UF*"
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
               Index           =   74
               Left            =   5033
               TabIndex        =   164
               Top             =   1380
               Width           =   285
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cidade*"
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
               Index           =   75
               Left            =   6735
               TabIndex        =   163
               Top             =   1380
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CEP*"
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
               Index           =   76
               Left            =   8850
               TabIndex        =   160
               Top             =   1380
               Width           =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "País*"
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
               Index           =   77
               Left            =   10650
               TabIndex        =   158
               Top             =   1380
               Width           =   375
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
               Index           =   78
               Left            =   2670
               TabIndex        =   156
               Top             =   1980
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
               Index           =   79
               Left            =   5595
               TabIndex        =   154
               Top             =   1980
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Site"
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
               Index           =   80
               Left            =   9885
               TabIndex        =   152
               Top             =   1980
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ramo de atividade"
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
               Index           =   81
               Left            =   7335
               TabIndex        =   150
               Top             =   2580
               Width           =   1335
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CNAE"
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
               Index           =   82
               Left            =   9638
               TabIndex        =   148
               Top             =   2580
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. SUFRAMA"
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
               Index           =   83
               Left            =   10842
               TabIndex        =   146
               Top             =   2580
               Width           =   1110
            End
         End
         Begin VB.TextBox Txt_ID_email 
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
            Left            =   -73020
            TabIndex        =   123
            Text            =   "0"
            Top             =   3990
            Visible         =   0   'False
            Width           =   855
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
            Height          =   1455
            Index           =   30
            Left            =   -74945
            TabIndex        =   100
            Top             =   330
            Width           =   15195
            Begin VB.TextBox Txt_senha_email 
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
               IMEMode         =   3  'DISABLE
               Left            =   13410
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   111
               ToolTipText     =   "Senha."
               Top             =   990
               Width           =   1605
            End
            Begin VB.TextBox Txt_usuario_email 
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
               Left            =   9660
               MaxLength       =   100
               TabIndex        =   110
               ToolTipText     =   "Usuário."
               Top             =   990
               Width           =   3735
            End
            Begin VB.TextBox Txt_nome_email 
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
               Left            =   2310
               MaxLength       =   100
               TabIndex        =   109
               ToolTipText     =   "Nome."
               Top             =   990
               Width           =   3675
            End
            Begin VB.ComboBox Cmb_seguranca_email 
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
               ItemData        =   "frmOpcoesGeral.frx":1E240
               Left            =   1060
               List            =   "frmOpcoesGeral.frx":1E24A
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   108
               ToolTipText     =   "Segurança."
               Top             =   990
               Width           =   1245
            End
            Begin VB.TextBox txt_porta_email 
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
               Left            =   180
               TabIndex        =   107
               ToolTipText     =   "Porta."
               Top             =   990
               Width           =   875
            End
            Begin VB.TextBox Txt_data_email 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """R$ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
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
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   106
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   875
            End
            Begin VB.TextBox Txt_servidor_SMTP_email 
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
               Left            =   12150
               MaxLength       =   50
               TabIndex        =   105
               ToolTipText     =   "Servidor SMTP."
               Top             =   390
               Width           =   2865
            End
            Begin VB.TextBox Txt_responsavel_email 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   1070
               Locked          =   -1  'True
               TabIndex        =   104
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   6510
            End
            Begin VB.ComboBox Cmb_aplicacao_email 
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
               ItemData        =   "frmOpcoesGeral.frx":1E263
               Left            =   7590
               List            =   "frmOpcoesGeral.frx":1E276
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   103
               ToolTipText     =   "Aplicação."
               Top             =   390
               Width           =   1245
            End
            Begin VB.TextBox Txt_email_email 
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
               Left            =   6000
               MaxLength       =   150
               TabIndex        =   102
               ToolTipText     =   "E-mail."
               Top             =   990
               Width           =   3645
            End
            Begin VB.ComboBox Cmb_usuario_caprind_email 
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
               ItemData        =   "frmOpcoesGeral.frx":1E2AC
               Left            =   8850
               List            =   "frmOpcoesGeral.frx":1E2AE
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   101
               ToolTipText     =   "Usuário caprind."
               Top             =   390
               Width           =   3285
            End
            Begin VB.Label Label1 
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
               Index           =   17
               Left            =   445
               TabIndex        =   122
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Porta*"
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
               Left            =   377
               TabIndex        =   121
               Top             =   780
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Senha*"
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
               Left            =   13942
               TabIndex        =   120
               Top             =   780
               Width           =   540
            End
            Begin VB.Label Label1 
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
               Index           =   36
               Left            =   3868
               TabIndex        =   119
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Aplicação*"
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
               Index           =   37
               Left            =   7830
               TabIndex        =   118
               Top             =   180
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Servidor SMTP*"
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
               Index           =   38
               Left            =   13020
               TabIndex        =   117
               Top             =   180
               Width           =   1125
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Segurança*"
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
               Index           =   39
               Left            =   1255
               TabIndex        =   116
               Top             =   780
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Nome*"
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
               Index           =   40
               Left            =   3900
               TabIndex        =   115
               Top             =   780
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "E-mail*"
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
               Index           =   42
               Left            =   7567
               TabIndex        =   114
               Top             =   780
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Usuário*"
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
               Index           =   43
               Left            =   11212
               TabIndex        =   113
               Top             =   780
               Width           =   630
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Usuário caprind*"
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
               Index           =   44
               Left            =   9892
               TabIndex        =   112
               Top             =   180
               Width           =   1200
            End
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
            Index           =   31
            Left            =   -74940
            TabIndex        =   88
            Top             =   330
            Width           =   15195
            Begin VB.ComboBox cmbAplicacao_Filtros 
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
               ItemData        =   "frmOpcoesGeral.frx":1E2B0
               Left            =   3900
               List            =   "frmOpcoesGeral.frx":1E2CF
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   94
               ToolTipText     =   "Aplicação."
               Top             =   390
               Width           =   1905
            End
            Begin VB.TextBox txtResponsavel_Filtros 
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
               Left            =   1070
               Locked          =   -1  'True
               TabIndex        =   93
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   2820
            End
            Begin VB.TextBox txtData_Filtros 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """R$ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
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
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   92
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   875
            End
            Begin VB.ComboBox cmbFrase_Filtros 
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
               ItemData        =   "frmOpcoesGeral.frx":1E32A
               Left            =   13410
               List            =   "frmOpcoesGeral.frx":1E33A
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   91
               ToolTipText     =   "Frase."
               Top             =   390
               Width           =   1605
            End
            Begin VB.ComboBox cmbfiltrarpor_Filtros 
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
               ItemData        =   "frmOpcoesGeral.frx":1E358
               Left            =   7740
               List            =   "frmOpcoesGeral.frx":1E36B
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   90
               ToolTipText     =   "Filtrar por."
               Top             =   390
               Width           =   5655
            End
            Begin VB.ComboBox cmbTipo_Filtros 
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
               ItemData        =   "frmOpcoesGeral.frx":1E3BE
               Left            =   5820
               List            =   "frmOpcoesGeral.frx":1E3C5
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   89
               ToolTipText     =   "Tipo."
               Top             =   390
               Width           =   1905
            End
            Begin VB.Label Label1 
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
               Index           =   29
               Left            =   480
               TabIndex        =   99
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Frase*"
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
               Index           =   27
               Left            =   13965
               TabIndex        =   98
               Top             =   180
               Width           =   495
            End
            Begin VB.Label Label1 
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
               Index           =   45
               Left            =   2023
               TabIndex        =   97
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Filtrar por*"
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
               Index           =   46
               Left            =   10170
               TabIndex        =   96
               Top             =   180
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Aplicação*                              Tipo*"
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
               Left            =   4470
               TabIndex        =   95
               Top             =   180
               Width           =   2505
            End
         End
         Begin VB.TextBox txtID_Filtros 
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
            Left            =   -73080
            TabIndex        =   87
            Text            =   "0"
            Top             =   2205
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Txt_ID_armaz 
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
            Left            =   -73110
            TabIndex        =   86
            Text            =   "0"
            Top             =   2205
            Visible         =   0   'False
            Width           =   855
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
            Index           =   32
            Left            =   -74945
            TabIndex        =   76
            Top             =   330
            Width           =   15195
            Begin VB.TextBox Txt_data_armaz 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """R$ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
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
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   81
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   875
            End
            Begin VB.TextBox Txt_responsavel_armaz 
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
               Left            =   1070
               Locked          =   -1  'True
               TabIndex        =   80
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   2820
            End
            Begin VB.ComboBox Cmb_relatorio_armaz 
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
               ItemData        =   "frmOpcoesGeral.frx":1E3DC
               Left            =   3900
               List            =   "frmOpcoesGeral.frx":1E3EC
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   79
               ToolTipText     =   "Relatório."
               Top             =   390
               Width           =   2595
            End
            Begin VB.TextBox Txt_local_armaz 
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
               Left            =   6510
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   78
               TabStop         =   0   'False
               ToolTipText     =   "Local de armazenamento."
               Top             =   390
               Width           =   8145
            End
            Begin VB.CommandButton Cmd_localizar_armaz 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "..."
               Height          =   315
               Left            =   14670
               MaskColor       =   &H00000000&
               Picture         =   "frmOpcoesGeral.frx":1E43E
               TabIndex        =   77
               ToolTipText     =   "Localizar."
               Top             =   390
               Width           =   315
            End
            Begin VB.Label Label1 
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
               Index           =   30
               Left            =   445
               TabIndex        =   85
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
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
               Index           =   49
               Left            =   2023
               TabIndex        =   84
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Relatório"
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
               Index           =   50
               Left            =   4875
               TabIndex        =   83
               Top             =   180
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   " Local de armazenamento"
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
               Index           =   51
               Left            =   9667
               TabIndex        =   82
               Top             =   180
               Width           =   1830
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
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
            Height          =   8295
            Index           =   40
            Left            =   55
            TabIndex        =   7
            Top             =   330
            Width           =   15195
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Configurações Nota fiscal Eletrônica (SEFAZ)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               Index           =   25
               Left            =   0
               TabIndex        =   66
               Top             =   30
               Width           =   4245
               Begin VB.TextBox Txt_local_armaz_NFe 
                  Alignment       =   2  'Center
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
                  Height          =   315
                  Left            =   180
                  Locked          =   -1  'True
                  MaxLength       =   255
                  TabIndex        =   69
                  TabStop         =   0   'False
                  ToolTipText     =   "Diretório do arquivo para envio da nota fiscal."
                  Top             =   1020
                  Width           =   3585
               End
               Begin VB.TextBox txtCaminhoXMLDanfe 
                  Alignment       =   2  'Center
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
                  Height          =   315
                  Left            =   180
                  Locked          =   -1  'True
                  MaxLength       =   255
                  TabIndex        =   68
                  TabStop         =   0   'False
                  ToolTipText     =   "Diretório dos arquivos XML e Danfe."
                  Top             =   450
                  Width           =   3585
               End
               Begin VB.TextBox txtRetornoNF 
                  Alignment       =   2  'Center
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
                  Height          =   315
                  Left            =   180
                  Locked          =   -1  'True
                  MaxLength       =   255
                  TabIndex        =   67
                  TabStop         =   0   'False
                  ToolTipText     =   "Diretório dos arquivos de retorno da nota fiscal."
                  Top             =   1590
                  Width           =   3585
               End
               Begin DrawSuite2022.USButton Cmd_localizar_NFe 
                  Height          =   315
                  Left            =   3780
                  TabIndex        =   70
                  ToolTipText     =   "Abrir diretório de envio..."
                  Top             =   1020
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   556
                  DibPicture      =   "frmOpcoesGeral.frx":1E540
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
                  BorderColor     =   8421504
                  BorderColorDisabled=   13160660
                  BorderColorDown =   7907521
                  BorderColorOver =   7907521
                  GradientColor2  =   14737632
                  GradientColor3  =   12632256
                  GradientColor4  =   12632256
                  GradientColorDisabled1=   14215660
                  GradientColorDisabled2=   14215660
                  GradientColorDisabled3=   14215660
                  GradientColorDisabled4=   14215660
                  GradientColorOver1=   14417407
                  GradientColorOver2=   12317439
                  GradientColorOver3=   4838399
                  GradientColorOver4=   9627391
                  GradientColorDown1=   10802943
                  GradientColorDown2=   7979263
                  GradientColorDown3=   4370174
                  GradientColorDown4=   7395582
                  GradientColors  =   1
                  PicAlign        =   0
                  ShowFocusRect   =   0   'False
                  Theme           =   1
                  ToolTipTitle    =   "CAPRIND v5.0"
               End
               Begin DrawSuite2022.USButton cmdLocalizarXMLDanfe 
                  Height          =   315
                  Left            =   3780
                  TabIndex        =   71
                  ToolTipText     =   "Abrir diretório DANFE-XML..."
                  Top             =   450
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   556
                  DibPicture      =   "frmOpcoesGeral.frx":3C645
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
                  BorderColor     =   8421504
                  BorderColorDisabled=   13160660
                  BorderColorDown =   7907521
                  BorderColorOver =   7907521
                  GradientColor2  =   14737632
                  GradientColor3  =   12632256
                  GradientColor4  =   12632256
                  GradientColorDisabled1=   14215660
                  GradientColorDisabled2=   14215660
                  GradientColorDisabled3=   14215660
                  GradientColorDisabled4=   14215660
                  GradientColorOver1=   14417407
                  GradientColorOver2=   12317439
                  GradientColorOver3=   4838399
                  GradientColorOver4=   9627391
                  GradientColorDown1=   10802943
                  GradientColorDown2=   7979263
                  GradientColorDown3=   4370174
                  GradientColorDown4=   7395582
                  GradientColors  =   1
                  PicAlign        =   0
                  ShowFocusRect   =   0   'False
                  Theme           =   1
                  ToolTipTitle    =   "CAPRIND v5.0"
               End
               Begin DrawSuite2022.USButton cmdLocalizarRetorno 
                  Height          =   315
                  Left            =   3780
                  TabIndex        =   72
                  ToolTipText     =   "Abrir diretório de retorno"
                  Top             =   1590
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   556
                  DibPicture      =   "frmOpcoesGeral.frx":5A74A
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
                  BorderColor     =   8421504
                  BorderColorDisabled=   13160660
                  BorderColorDown =   7907521
                  BorderColorOver =   7907521
                  GradientColor2  =   14737632
                  GradientColor3  =   12632256
                  GradientColor4  =   12632256
                  GradientColorDisabled1=   14215660
                  GradientColorDisabled2=   14215660
                  GradientColorDisabled3=   14215660
                  GradientColorDisabled4=   14215660
                  GradientColorOver1=   14417407
                  GradientColorOver2=   12317439
                  GradientColorOver3=   4838399
                  GradientColorOver4=   9627391
                  GradientColorDown1=   10802943
                  GradientColorDown2=   7979263
                  GradientColorDown3=   4370174
                  GradientColorDown4=   7395582
                  GradientColors  =   1
                  PicAlign        =   0
                  ShowFocusRect   =   0   'False
                  Theme           =   1
                  ToolTipTitle    =   "CAPRIND v5.0"
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Diretório dos arquivos XML e Danfe"
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
                  Left            =   705
                  TabIndex        =   75
                  Top             =   270
                  Width           =   2520
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Diretório de envio"
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
                  Left            =   1335
                  TabIndex        =   74
                  Top             =   840
                  Width           =   1275
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Diretório de retorno"
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
                  Left            =   1260
                  TabIndex        =   73
                  Top             =   1410
                  Width           =   1425
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1155
               Index           =   27
               Left            =   9480
               TabIndex        =   19
               Top             =   930
               Width           =   5715
               Begin VB.TextBox txtSerie_Nf 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   4980
                  TabIndex        =   376
                  Text            =   "1"
                  ToolTipText     =   "Digite á ser utlizada na emissão de nota fiscal eletrônica."
                  Top             =   690
                  Width           =   525
               End
               Begin VB.TextBox Txt_apelido_contimatic 
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
                  Left            =   3720
                  MaxLength       =   30
                  TabIndex        =   21
                  ToolTipText     =   "Apelido cadastrado no sistema Contimatic."
                  Top             =   270
                  Width           =   1815
               End
               Begin VB.TextBox Txt_registro_boleto 
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
                  Height          =   315
                  Left            =   1380
                  MaxLength       =   50
                  TabIndex        =   20
                  ToolTipText     =   "Registro do boleto."
                  Top             =   285
                  Width           =   1395
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Série para emissão de Nota fiscal eletrônica"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Index           =   31
                  Left            =   1080
                  TabIndex        =   375
                  Top             =   750
                  Width           =   3780
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contimatic"
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
                  Left            =   2910
                  TabIndex        =   23
                  Top             =   300
                  Width           =   750
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Registro boleto"
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
                  Left            =   240
                  TabIndex        =   22
                  Top             =   300
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Dados Certificado digital"
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
               Height          =   885
               Index           =   46
               Left            =   9480
               TabIndex        =   14
               Top             =   30
               Width           =   5715
               Begin VB.TextBox txtCertificadodigital 
                  Alignment       =   2  'Center
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
                  Height          =   315
                  Left            =   900
                  Locked          =   -1  'True
                  MaxLength       =   255
                  TabIndex        =   16
                  TabStop         =   0   'False
                  ToolTipText     =   "Diretório do arquivo para envio da nota fiscal."
                  Top             =   420
                  Width           =   4245
               End
               Begin VB.TextBox txttpEmissor 
                  Alignment       =   2  'Center
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
                  Height          =   315
                  Left            =   120
                  Locked          =   -1  'True
                  MaxLength       =   255
                  TabIndex        =   15
                  TabStop         =   0   'False
                  ToolTipText     =   "Tipo do certificado digital"
                  Top             =   420
                  Width           =   765
               End
               Begin DrawSuite2022.USButton cmdCertificado 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   17
                  ToolTipText     =   "Configurar certificado digital a ser utilizado."
                  Top             =   420
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   556
                  DibPicture      =   "frmOpcoesGeral.frx":7884F
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
                  BorderColor     =   8421504
                  BorderColorDisabled=   13160660
                  BorderColorDown =   7907521
                  BorderColorOver =   7907521
                  GradientColor2  =   14737632
                  GradientColor3  =   12632256
                  GradientColor4  =   12632256
                  GradientColorDisabled1=   14215660
                  GradientColorDisabled2=   14215660
                  GradientColorDisabled3=   14215660
                  GradientColorDisabled4=   14215660
                  GradientColorOver1=   14417407
                  GradientColorOver2=   12317439
                  GradientColorOver3=   4838399
                  GradientColorOver4=   9627391
                  GradientColorDown1=   10802943
                  GradientColorDown2=   7979263
                  GradientColorDown3=   4370174
                  GradientColorDown4=   7395582
                  GradientColors  =   1
                  PicAlign        =   0
                  ShowFocusRect   =   0   'False
                  Theme           =   1
                  ToolTipTitle    =   "CAPRIND v5.0"
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo                                    Serial do certificado"
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
                  Left            =   360
                  TabIndex        =   18
                  Top             =   210
                  Width           =   3375
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Opções emissão Nota fiscal eletrônica"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               Left            =   4260
               TabIndex        =   8
               Top             =   30
               Width           =   5205
               Begin VB.CheckBox Chk_calcular_IPI 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Ativar calculo do IPI sobre valor sem desconto"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   150
                  TabIndex        =   368
                  Top             =   1470
                  Width           =   4005
               End
               Begin VB.CheckBox chkSemEstoque 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Bloquear emissão da NF de produtos sem estoque"
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
                  Left            =   150
                  TabIndex        =   366
                  Top             =   420
                  Width           =   4875
               End
               Begin VB.CheckBox chk_TPAmb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Ativar emissão de NFe em ambiente de testes"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   345
                  Left            =   150
                  TabIndex        =   13
                  Top             =   1650
                  Width           =   5025
               End
               Begin VB.CheckBox chk_Baixa_Auto_Estoque_NF 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Ativar baixa automática do estoque por nota fiscal"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   150
                  TabIndex        =   12
                  Top             =   840
                  Width           =   4875
               End
               Begin VB.CheckBox Chk_bloquear_NF_prod_serv_sem_cad 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Bloquear emissão da NF de produtos/serviços sem cadastro"
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
                  Left            =   150
                  TabIndex        =   11
                  Top             =   630
                  Width           =   4875
               End
               Begin VB.CheckBox Chk_codigo_ref_DANFE 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Ativar utilização do código de referência na DANFE"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   150
                  TabIndex        =   10
                  Top             =   1050
                  Width           =   4905
               End
               Begin VB.CheckBox Chk_codigo_ref_desc_DANFE 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Ativar código referência junto com a descrição na DANFE"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   150
                  TabIndex        =   9
                  Top             =   1260
                  Width           =   4815
               End
            End
            Begin TabDlg.SSTab SSTab4 
               Height          =   6195
               Left            =   0
               TabIndex        =   24
               Top             =   2100
               Width           =   15195
               _ExtentX        =   26802
               _ExtentY        =   10927
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
               TabCaption(0)   =   "Configurações do Caprind"
               TabPicture(0)   =   "frmOpcoesGeral.frx":96954
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Frame1(28)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Configurações do Gerprod"
               TabPicture(1)   =   "frmOpcoesGeral.frx":96970
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Frame1(29)"
               Tab(1).ControlCount=   1
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   5835
                  Index           =   29
                  Left            =   -74970
                  TabIndex        =   57
                  Top             =   330
                  Width           =   15105
                  Begin VB.CheckBox Chk_bloquear_apontamento_simultaneo 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear apontamento ""ABRIR VÁRIAS OS's"""
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
                     TabIndex        =   65
                     Top             =   1950
                     Width           =   7215
                  End
                  Begin VB.CheckBox Chk_bloquear_apontamento_sem_baixa_total 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear apontamento se não baixar toda a lista de requisição da ordem que não movimenta estoque automaticamente"
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
                     Height          =   405
                     Left            =   180
                     TabIndex        =   64
                     Top             =   1005
                     Width           =   10695
                  End
                  Begin VB.CheckBox Chk_ap_codigo 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Apontamento por código"
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
                     TabIndex        =   63
                     Top             =   300
                     Width           =   7215
                  End
                  Begin VB.CheckBox Chk_bloquear_apontamento_sem_baixa 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear apontamento se não baixar matéria-prima da lista de requisição da ordem que não movimenta estoque automaticamente"
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
                     Height          =   405
                     Left            =   180
                     TabIndex        =   62
                     Top             =   546
                     Width           =   12075
                  End
                  Begin VB.CheckBox Chk_desbloquear_primeiro_apontamento_OS_proc_controlado 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Desbloquear primeiro apontamento de OS com processo controlado"
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
                     TabIndex        =   61
                     Top             =   1455
                     Width           =   7215
                  End
                  Begin VB.CheckBox chk_Grupo_Gerprod 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Carregar posto de trabalho por grupo"
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
                     TabIndex        =   60
                     Top             =   1710
                     Width           =   7215
                  End
                  Begin VB.CheckBox Chk_NC_parecer 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Criar quantidade NC com parecer ""Rejeitar"" ao apontar"
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
                     TabIndex        =   59
                     Top             =   2430
                     Width           =   7215
                  End
                  Begin VB.CheckBox Chk_apontar_NC_descricao 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Apontar quantidade NC por descrição da não conformidade"
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
                     TabIndex        =   58
                     Top             =   2190
                     Width           =   7215
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   5805
                  Index           =   28
                  Left            =   30
                  TabIndex        =   25
                  Top             =   330
                  Width           =   15105
                  Begin VB.CheckBox chkClienteVendedor 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear cliente por vendedor interno"
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
                     Left            =   6240
                     TabIndex        =   367
                     Top             =   570
                     Width           =   5565
                  End
                  Begin VB.Frame Frame1 
                     BackColor       =   &H00E0E0E0&
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Index           =   20
                     Left            =   6240
                     TabIndex        =   54
                     Top             =   270
                     Width           =   5685
                     Begin VB.TextBox Txt_minutos_desconectar 
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
                        Left            =   3300
                        Locked          =   -1  'True
                        MaxLength       =   5
                        TabIndex        =   55
                        TabStop         =   0   'False
                        ToolTipText     =   "Minutos."
                        Top             =   0
                        Width           =   645
                     End
                     Begin VB.CheckBox Chk_verificar_desconectar_usuario 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "Verificar e desconectar usuário a cada                         (minutos)"
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
                        Left            =   0
                        TabIndex        =   56
                        Top             =   30
                        Width           =   5295
                     End
                  End
                  Begin VB.CheckBox Chk_bloquear_forn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear utilização do fornecedores sem homologação"
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
                     Left            =   6240
                     TabIndex        =   53
                     Top             =   1080
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_CC_obrigatorio 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar centro de custo obrigatório"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   52
                     Top             =   975
                     Width           =   6015
                  End
                  Begin VB.CheckBox Chk_bloquear_prod_cliente 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear produtos/serviços por cliente"
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
                     Left            =   6240
                     TabIndex        =   51
                     Top             =   825
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_liberar_qtde_MRP 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar alteração da quantidade ao gerar o MRP"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   50
                     Top             =   1470
                     Width           =   6015
                  End
                  Begin VB.CheckBox Chk_bloquear_cli_forn_regime 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear utilização do cliente/fornecedor sem regime tributário"
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
                     Left            =   6240
                     TabIndex        =   49
                     Top             =   1335
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_ativar_empenho_aut 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar empenho automático do estoque"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   48
                     Top             =   240
                     Width           =   6015
                  End
                  Begin VB.CheckBox Chk_carregar_CFOP_ST 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Carregar CFOP do produto conforme configurações da ST"
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
                     TabIndex        =   47
                     Top             =   4020
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_agregar_ordem_valor_PC 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Agregar na ordem o valor do produto do pedido de compra"
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
                     TabIndex        =   46
                     Top             =   4275
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_gerar_RM_ordem_PC 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Gerar RM da ordem pelo pedido de compra"
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
                     TabIndex        =   45
                     Top             =   4515
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_liberar_campos_estrutura 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar ""Kg/unidade"" e ""Un/Kg"" no cadastro da estrutura"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   44
                     Top             =   1230
                     Width           =   4665
                  End
                  Begin VB.CheckBox chk_Esconder_ValorOF 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar valores ocultos na ordem de faturamento"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   255
                     Left            =   180
                     TabIndex        =   43
                     Top             =   1680
                     Width           =   4185
                  End
                  Begin VB.CheckBox Chk_movimentar_estoque_pc 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar movimentação de estoque por peça"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   42
                     Top             =   3150
                     Width           =   3765
                  End
                  Begin VB.CheckBox Chk_ativar_produtos_similares 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar recurso para produtos similares"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   41
                     Top             =   2910
                     Width           =   3495
                  End
                  Begin VB.CheckBox Chk_validar_proposta_pi_autom 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar validação da proposta e do pedido automaticamente"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   40
                     Top             =   1950
                     Width           =   6015
                  End
                  Begin VB.CheckBox Chk_codigo_ref_SPED_forn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar utilização do código de referência no SPED (fornecedor)"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   39
                     Top             =   3405
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkLiberar_LoteMinimo 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar emissão da ordem com qtde. menor que o lote mínimo"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   38
                     Top             =   735
                     Width           =   5025
                  End
                  Begin VB.CheckBox Chk_carregar_LA_entrada 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Carregar local de armazen. com estoque zerado (entrada)"
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
                     TabIndex        =   37
                     Top             =   4770
                     Width           =   5715
                  End
                  Begin VB.CheckBox chkNao_inspecionar 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Não inspecionar quando fornecedor for certificado ou avaliado"
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
                     TabIndex        =   36
                     Top             =   5025
                     Width           =   5715
                  End
                  Begin VB.CheckBox ChkBloc_CC_Previsao 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear utilização de centro de custo sem previsão"
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
                     Left            =   6240
                     TabIndex        =   35
                     Top             =   1590
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_bloq_OP_estrutura 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear emissão da ordem sem validação da estrutura"
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
                     Left            =   6240
                     TabIndex        =   34
                     Top             =   1845
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_bloq_OP_processo 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear emissão da ordem sem validação do processo"
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
                     Left            =   6240
                     TabIndex        =   33
                     Top             =   2100
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_bloq_OP_plano 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear emissão da ordem sem validação do plano de insp."
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
                     Left            =   6240
                     TabIndex        =   32
                     Top             =   2355
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_bloq_compra_cot_valida 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Bloquear compra sem cotação válida"
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
                     Left            =   6240
                     TabIndex        =   31
                     Top             =   2625
                     Width           =   5715
                  End
                  Begin VB.CheckBox Chk_ativar_empenho_aut_prod 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar empenho automático da produção"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   30
                     Top             =   480
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkCodigo_sequencial 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar a geração de código sequencial ao criar produto final"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   29
                     Top             =   2670
                     Width           =   6015
                  End
                  Begin VB.CheckBox Chk_salvar_status_aprovado_PC 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar salvamento do status do pedido de compra como ""APROVADO"""
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   28
                     Top             =   2190
                     Width           =   6015
                  End
                  Begin VB.CheckBox Chk_enviar_email_outlook 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Ativar envio de email pelo outlook"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   180
                     TabIndex        =   27
                     Top             =   2430
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkMargemAnalise 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Usar margem no cálculo de recíproca/fator na análise crítica"
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
                     TabIndex        =   26
                     Top             =   5295
                     Width           =   5715
                  End
               End
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Configuração Nota Fiscal de Serviços eletrônica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Index           =   39
            Left            =   11040
            TabIndex        =   2
            Top             =   2430
            Width           =   4215
            Begin VB.TextBox txtSenhaPref 
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
               Left            =   2070
               MaxLength       =   100
               TabIndex        =   4
               ToolTipText     =   "Senha para envio de NFS-e (Obrigatório em algumas cidadas para emissão da nota de serviço)."
               Top             =   450
               Width           =   1980
            End
            Begin VB.TextBox txtUsuarioPref 
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
               Left            =   180
               MaxLength       =   100
               TabIndex        =   3
               ToolTipText     =   "Usúario para envio da NFS-e (Obrigatório em algumas cidadas para emissão da nota de serviço)."
               Top             =   450
               Width           =   1800
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Senha prefeitura"
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
               Index           =   35
               Left            =   2460
               TabIndex        =   6
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Usuário prefeitura"
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
               Index           =   85
               Left            =   480
               TabIndex        =   5
               Top             =   240
               Width           =   1305
            End
         End
         Begin MSComctlLib.ListView Lista_email 
            Height          =   6825
            Left            =   -74940
            TabIndex        =   185
            Top             =   1800
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   12039
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
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
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Aplicação"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Usuário caprind"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "E-mail"
               Object.Width           =   9534
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_filtros 
            Height          =   7395
            Left            =   -74940
            TabIndex        =   186
            Top             =   1200
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   13044
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
               Object.Width           =   529
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
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Aplicação"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Tipo"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Filtrar por"
               Object.Width           =   12532
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Frase"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_armaz 
            Height          =   7425
            Left            =   -74940
            TabIndex        =   187
            Top             =   1200
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   13097
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
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
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Relatório"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Local de armazenamento"
               Object.Width           =   14647
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_empresas 
            Height          =   4455
            Left            =   -74940
            TabIndex        =   362
            Top             =   4170
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   7858
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Razão social"
               Object.Width           =   25585
            EndProperty
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   0
         TabIndex        =   271
         Top             =   360
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   9
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDownRight1=   10802943
         GradientColorDownRight2=   7979263
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOverRight1=   14417407
         GradientColorOverRight2=   12317439
         IsStrech        =   -1  'True
         RightColor1     =   14737632
         RightColor2     =   16777215
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
         ButtonCaption4  =   "Conf. relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Configurar dados da empresa nos relatórios"
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
         ButtonWidth4    =   78
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Atualizar"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Atualizar regime tributário dos impostos e ID das empresas na tabela."
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
         ButtonLeft5     =   198
         ButtonTop5      =   2
         ButtonWidth5    =   50
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
         ButtonLeft6     =   250
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   54
         ButtonCaption7  =   "Ajuda"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Ajuda (F1)"
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
         ButtonLeft7     =   254
         ButtonTop7      =   2
         ButtonWidth7    =   36
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Sair"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Sair (Esc)"
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
         ButtonLeft8     =   292
         ButtonTop8      =   2
         ButtonWidth8    =   26
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
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
         ButtonState9    =   5
         ButtonLeft9     =   320
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   6180
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOpcoesGeral.frx":9698C
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74970
         TabIndex        =   273
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   7
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
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonAlignment4=   2
         ButtonType4     =   1
         ButtonStyle4    =   -1
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState4    =   -1
         ButtonLeft4     =   118
         ButtonTop4      =   4
         ButtonWidth4    =   2
         ButtonHeight4   =   54
         ButtonCaption5  =   "Ajuda"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Ajuda (F1)"
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
         ButtonLeft5     =   122
         ButtonTop5      =   2
         ButtonWidth5    =   36
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Sair"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Sair (Esc)"
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
         ButtonLeft6     =   160
         ButtonTop6      =   2
         ButtonWidth6    =   26
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
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
         ButtonState7    =   5
         ButtonLeft7     =   188
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   4920
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOpcoesGeral.frx":9B8C7
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView ListaMoeda 
         Height          =   7515
         Left            =   -74970
         TabIndex        =   283
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13256
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
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
            Object.Width           =   17295
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Moeda"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "SÍmbolo"
            Object.Width           =   2646
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74970
         TabIndex        =   284
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   7
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
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonAlignment4=   2
         ButtonType4     =   1
         ButtonStyle4    =   -1
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState4    =   -1
         ButtonLeft4     =   118
         ButtonTop4      =   4
         ButtonWidth4    =   2
         ButtonHeight4   =   54
         ButtonCaption5  =   "Ajuda"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Ajuda (F1)"
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
         ButtonLeft5     =   122
         ButtonTop5      =   2
         ButtonWidth5    =   36
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Sair"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Sair (Esc)"
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
         ButtonLeft6     =   160
         ButtonTop6      =   2
         ButtonWidth6    =   26
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
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
         ButtonState7    =   5
         ButtonLeft7     =   188
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   7740
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOpcoesGeral.frx":9EC9F
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   10035
         Left            =   -74970
         TabIndex        =   285
         Top             =   1290
         Width           =   15300
         _ExtentX        =   26988
         _ExtentY        =   17701
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
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
         TabCaption(0)   =   "Unidades"
         TabPicture(0)   =   "frmOpcoesGeral.frx":A2077
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Lista_unidade"
         Tab(0).Control(1)=   "Frame1(34)"
         Tab(0).Control(2)=   "txtidunidade"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Tabela de conversão"
         TabPicture(1)   =   "frmOpcoesGeral.frx":A2093
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Lista_conversao"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame1(35)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Txt_ID_conversao"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.TextBox txtidunidade 
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
            Left            =   -73855
            TabIndex        =   309
            Text            =   "0"
            ToolTipText     =   "Unidade."
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
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
            Index           =   34
            Left            =   -74945
            TabIndex        =   300
            Top             =   330
            Width           =   15195
            Begin VB.TextBox Txt_descricao_unidade 
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
               Left            =   7920
               MaxLength       =   50
               TabIndex        =   304
               ToolTipText     =   "Descrição."
               Top             =   390
               Width           =   7080
            End
            Begin VB.TextBox Txt_unidade 
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
               Left            =   7050
               MaxLength       =   6
               TabIndex        =   303
               ToolTipText     =   "Unidade."
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox Txt_responsavel_unidade 
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
               Left            =   1070
               Locked          =   -1  'True
               TabIndex        =   302
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   5970
            End
            Begin VB.TextBox Txt_data_unidade 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """R$ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
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
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   301
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   875
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   20
               Left            =   450
               TabIndex        =   308
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
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
               Index           =   53
               Left            =   3598
               TabIndex        =   307
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Unidade*"
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
               Index           =   55
               Left            =   7140
               TabIndex        =   306
               Top             =   180
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição*"
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
               Index           =   56
               Left            =   11070
               TabIndex        =   305
               Top             =   180
               Width           =   780
            End
         End
         Begin VB.TextBox Txt_ID_conversao 
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
            Left            =   1090
            TabIndex        =   299
            Text            =   "0"
            ToolTipText     =   "Unidade."
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
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
            Index           =   35
            Left            =   60
            TabIndex        =   286
            Top             =   330
            Width           =   15195
            Begin VB.TextBox Txt_data_conversao 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """R$ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
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
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   292
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   390
               Width           =   875
            End
            Begin VB.TextBox Txt_responsavel_conversao 
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
               Left            =   1070
               Locked          =   -1  'True
               TabIndex        =   291
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   390
               Width           =   3240
            End
            Begin VB.TextBox Txt_qtde_de_conversao 
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
               Height          =   315
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   290
               TabStop         =   0   'False
               Text            =   "1,000"
               ToolTipText     =   "Quantidade."
               Top             =   390
               Width           =   1125
            End
            Begin VB.ComboBox Cmb_unidade_de_conversao 
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
               Height          =   330
               ItemData        =   "frmOpcoesGeral.frx":A20AF
               Left            =   8160
               List            =   "frmOpcoesGeral.frx":A20B1
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   289
               ToolTipText     =   "Unidade de estoque."
               Top             =   390
               Width           =   1035
            End
            Begin VB.TextBox Txt_qtde_para_conversao 
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
               Height          =   315
               Left            =   11190
               TabIndex        =   288
               ToolTipText     =   "Quantidade."
               Top             =   390
               Width           =   1365
            End
            Begin VB.ComboBox Cmb_unidade_para_conversao 
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
               Height          =   330
               ItemData        =   "frmOpcoesGeral.frx":A20B3
               Left            =   13650
               List            =   "frmOpcoesGeral.frx":A20B5
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   287
               ToolTipText     =   "Unidade comercial."
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "equivale a quantidade de"
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
               Left            =   9255
               TabIndex        =   298
               Top             =   450
               Width           =   1815
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   21
               Left            =   450
               TabIndex        =   297
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
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
               Index           =   57
               Left            =   2233
               TabIndex        =   296
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "A quantidade de "
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
               Index           =   58
               Left            =   4800
               TabIndex        =   295
               Top             =   450
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "da unidade"
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
               Index           =   59
               Left            =   7290
               TabIndex        =   294
               Top             =   480
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "da unidade"
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
               Index           =   60
               Left            =   12720
               TabIndex        =   293
               Top             =   480
               Width           =   795
            End
         End
         Begin MSComctlLib.ListView Lista_unidade 
            Height          =   7185
            Left            =   -74940
            TabIndex        =   310
            Top             =   1200
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   12674
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
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
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Un."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   16413
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_conversao 
            Height          =   7185
            Left            =   60
            TabIndex        =   311
            Top             =   1200
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   12674
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
               Object.Width           =   529
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
               Object.Width           =   16237
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Un. est."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Un. com."
               Object.Width           =   1499
            EndProperty
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar6 
         Height          =   975
         Left            =   -74970
         TabIndex        =   321
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   7
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
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonAlignment4=   2
         ButtonType4     =   1
         ButtonStyle4    =   -1
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState4    =   -1
         ButtonLeft4     =   118
         ButtonTop4      =   4
         ButtonWidth4    =   2
         ButtonHeight4   =   54
         ButtonCaption5  =   "Ajuda"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Ajuda (F1)"
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
         ButtonLeft5     =   122
         ButtonTop5      =   2
         ButtonWidth5    =   36
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Sair"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Sair (Esc)"
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
         ButtonLeft6     =   160
         ButtonTop6      =   2
         ButtonWidth6    =   26
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
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
         ButtonState7    =   5
         ButtonLeft7     =   188
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         Begin DrawSuite2022.USImageList USImageList6 
            Left            =   4920
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOpcoesGeral.frx":A20B7
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   7845
         Left            =   -74970
         TabIndex        =   322
         Top             =   2160
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   13838
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
         TabCaption(0)   =   "Compras"
         TabPicture(0)   =   "frmOpcoesGeral.frx":A548F
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Lista_cond"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Txt_ID_cond"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Vendas"
         TabPicture(1)   =   "frmOpcoesGeral.frx":A54AB
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lista_cond1"
         Tab(1).ControlCount=   1
         Begin VB.TextBox Txt_ID_cond 
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
            Left            =   660
            TabIndex        =   329
            Text            =   "0"
            Top             =   1140
            Visible         =   0   'False
            Width           =   855
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
            Index           =   41
            Left            =   -74945
            TabIndex        =   323
            Top             =   2
            Width           =   15195
            Begin VB.ComboBox Combo2 
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
               ItemData        =   "frmOpcoesGeral.frx":A54C7
               Left            =   12090
               List            =   "frmOpcoesGeral.frx":A54C9
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   325
               ToolTipText     =   "Unidade de estoque."
               Top             =   390
               Width           =   765
            End
            Begin VB.ComboBox Combo1 
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
               ItemData        =   "frmOpcoesGeral.frx":A54CB
               Left            =   14250
               List            =   "frmOpcoesGeral.frx":A54CD
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   324
               ToolTipText     =   "Unidade comercial."
               Top             =   390
               Width           =   765
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
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
               Left            =   14310
               TabIndex        =   326
               Top             =   180
               Width           =   645
            End
         End
         Begin MSComctlLib.ListView Lista_cond 
            Height          =   7165
            Left            =   30
            TabIndex        =   327
            Top             =   345
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   12647
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
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
               Object.Width           =   15180
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Texto"
               Object.Width           =   8291
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_cond1 
            Height          =   7165
            Left            =   -74970
            TabIndex        =   328
            Top             =   345
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   12647
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
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
               Object.Width           =   15180
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Texto"
               Object.Width           =   8291
            EndProperty
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar7 
         Height          =   975
         Left            =   -74940
         TabIndex        =   330
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   7
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
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonAlignment4=   2
         ButtonType4     =   1
         ButtonStyle4    =   -1
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState4    =   -1
         ButtonLeft4     =   118
         ButtonTop4      =   4
         ButtonWidth4    =   2
         ButtonHeight4   =   54
         ButtonCaption5  =   "Ajuda"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Ajuda (F1)"
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
         ButtonLeft5     =   122
         ButtonTop5      =   2
         ButtonWidth5    =   36
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Sair"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Sair (Esc)"
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
         ButtonLeft6     =   160
         ButtonTop6      =   2
         ButtonWidth6    =   26
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
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
         ButtonState7    =   5
         ButtonLeft7     =   188
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         Begin DrawSuite2022.USImageList USImageList7 
            Left            =   4920
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOpcoesGeral.frx":A54CF
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_feriado 
         Height          =   7515
         Left            =   -74970
         TabIndex        =   341
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13256
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
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
            Object.Width           =   15178
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Dt. do feriado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   6174
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74940
         TabIndex        =   352
         Top             =   360
         Width           =   15195
         _ExtentX        =   26802
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
         ButtonCaption4  =   "Base de dados"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Alterar caminho do acesso (F6)"
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
         ButtonWidth4    =   78
         ButtonHeight4   =   21
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
         ButtonLeft5     =   198
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   202
         ButtonTop6      =   2
         ButtonWidth6    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   240
         ButtonTop7      =   2
         ButtonWidth7    =   26
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
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
         ButtonState8    =   5
         ButtonLeft8     =   268
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   6600
            Top             =   120
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOpcoesGeral.frx":A88A7
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView ListaBancos 
         Height          =   6105
         Left            =   -74970
         TabIndex        =   360
         Top             =   3660
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10769
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Local dos relatórios"
            Object.Width           =   13238
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nome da instância SQL"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Nome do banco de dados"
            Object.Width           =   6350
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   -74970
         TabIndex        =   361
         Top             =   9750
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
      Begin DrawSuite2022.USToolBar USToolBar5 
         Height          =   975
         Left            =   30
         TabIndex        =   363
         Top             =   360
         Width           =   15195
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
         ButtonCaption1  =   "Salvar"
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
         ButtonWidth1    =   38
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
         ButtonLeft2     =   42
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
         ButtonLeft3     =   46
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
         ButtonLeft4     =   84
         ButtonTop4      =   2
         ButtonWidth4    =   26
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
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
         ButtonState5    =   5
         ButtonLeft5     =   112
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
         Begin DrawSuite2022.USImageList USImageList5 
            Left            =   9450
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmOpcoesGeral.frx":AC6B1
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frmOpcoesGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Local    As Boolean 'OK
Dim Novo_geral    As Boolean 'OK
Dim Novo_geral1   As Boolean 'OK
Dim Novo_geral2   As Boolean 'OK
Dim Novo_geral3   As Boolean 'OK
Dim Novo_geral4   As Boolean 'OK
Dim Novo_geral5   As Boolean 'OK
Dim Novo_geral6   As Boolean 'OK
Dim Novo_geral7   As Boolean 'OK
Dim Novo_geral8   As Boolean 'OK
Dim Novo_geral9   As Boolean 'OK
Public PC_PIS     As Boolean
Public PC_Cofins  As Boolean
Public PC_CSLL    As Boolean
Public PC_ISSQN   As Boolean
Public PC_IRRF    As Boolean
Public PC_INSS    As Boolean
Public Regime     As Integer 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=o9dhuFSCS5I&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=53&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

If Cmb_tipo_TBSN <> "" Then
    Select Case Mid(Cmb_tipo_TBSN, 8, 3)
        Case "I -": IDConta = 1
        Case "II ": IDConta = 2
        Case "III": IDConta = 3
        Case "IV ": IDConta = 4
        Case "V -": IDConta = 5
    End Select
    If IDConta <> 2 Then ProcCorrigeFormTabelaSN IDConta, True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub



Private Sub BtnLogotipo_Click()
On Error GoTo tratar_erro

'If Frame1(3).Enabled = False Then Exit Sub
ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "Arquivos jpg (*.jpg) | *.jpg| Arquivos bmp (*.bmp) | *.bmp"
If caminho <> "" Then
    picimagem.Picture = LoadPicture(caminho)
    'If fotopadrao = Localrel & "\imagens\caprind.bmp" Then Label8.Visible = True Else Label8.Visible = False
Else
    picimagem.Picture = LoadPicture("")
    'Label8.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_bloquear_apontamento_sem_baixa_Click()
On Error GoTo tratar_erro

If Chk_bloquear_apontamento_sem_baixa.Value = 1 Then Chk_bloquear_apontamento_sem_baixa_total.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_bloquear_apontamento_sem_baixa_total_Click()
On Error GoTo tratar_erro

If Chk_bloquear_apontamento_sem_baixa_total.Value = 1 Then Chk_bloquear_apontamento_sem_baixa.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_empresa()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_empresas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) empresa(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Empresa where codigo = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Impostos where ID_empresa = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir empresa"
            ID_documento = .ListItems(InitFor)
            Documento = "Empresa: " & .ListItems(InitFor).ListSubItems(1)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) empresa(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empresa(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposEmpresa
    ProcCarregaListaEmpresa
    Frame1(3).Enabled = False
    Novo_geral = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirTabelaSN()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_TBSN
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) registro(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            
            Permitido = True
            Conexao.Execute "DELETE from Impostos_TabelaDAS where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir registro da tabela do DAS"
            ID_documento = .ListItems(InitFor)
            Documento = "Empresa: " & frmOpcoesGeral.txtRazao
            Documento1 = "De: " & .ListItems(InitFor).SubItems(1) & " - Até: " & .ListItems(InitFor).SubItems(2) & " - DAS: " & .ListItems(InitFor).SubItems(3) & " - ICMS: " & .ListItems(InitFor).SubItems(4)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) registro(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Registro(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposTBSN
    Frame1(45).Enabled = False
    Frame1(47).Enabled = False
    ProcCarregaLista_TBSN Cmb_tipo_TBSN
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_email()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_email
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) e-mail(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Empresa_email where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir e-mail"
            ID_documento = .ListItems(InitFor)
            Documento = "Empresa: " & txtRazao
            Documento1 = "E-mail: " & .ListItems(InitFor).ListSubItems(4)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) e-mail(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("E-mail(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposEmail
    ProcCarregaListaEmail
    Frame1(30).Enabled = False
    Novo_geral6 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Filtros()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_filtros
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) filtro(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Empresa_filtros where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir filtro"
            ID_documento = .ListItems(InitFor)
            Documento = "Empresa: " & txtRazao
            Documento1 = "Filtro: " & .ListItems(InitFor).ListSubItems(4)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) filtro(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Filtro(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposFiltros
    ProcCarregaListaFiltros
    Frame1(31).Enabled = False
    Novo_geral7 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Armaz()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_armaz
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) local(ais) de armazenamento?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Empresa_armazenamento_PDF where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir local de armazenamento"
            ID_documento = .ListItems(InitFor)
            Documento = "Empresa: " & txtRazao
            Documento1 = "Relatório: " & .ListItems(InitFor).ListSubItems(3) & " - Local de armazenamento: " & .ListItems(InitFor).ListSubItems(4)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) local(ais) de armazenamento antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Local(ais) de armazenamento excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposArmaz
    ProcCarregaListaArmaz
    Frame1(32).Enabled = False
    Novo_geral8 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_unidade()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_unidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) unidade(s) de medida?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from unidade_medida where codigo = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir unidade de medida"
            ID_documento = .ListItems(InitFor)
            Documento = "Unidade: " & .ListItems(InitFor).ListSubItems(3) & " - Descrição: " & .ListItems(InitFor).ListSubItems(4)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) unidade(s) de medida antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Unidade(s) de medida excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposUnidade
    ProcCarregaListaUnidade
    Frame1(34).Enabled = False
    Novo_geral2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_conversao()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_conversao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) regra(s) de conversão?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Tabela_conversao_unidade where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir regra de conversão"
            ID_documento = .ListItems(InitFor)
            Documento = "Qtde. de: " & .ListItems(InitFor).ListSubItems(1) & " - Unidade de: " & .ListItems(InitFor).ListSubItems(2) & " - Qtde. para: " & .ListItems(InitFor).ListSubItems(3) & " - Unidade para: " & .ListItems(InitFor).ListSubItems(4)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) regra(s) de conversão antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Regra(s) de conversão excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposConversao
    ProcCarregaListaConversao
    Frame1(35).Enabled = False
    Novo_geral3 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_condicao()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If SSTab3.Tab = 0 Then ProcExcluir_condicao1 Lista_cond, "compras" Else ProcExcluir_condicao1 Lista_cond1, "vendas"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_condicao1(Lista As ListView, Texto As String)
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) condição(ões) de pagamento/recebimento de " & Texto & "?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_proposta_dadoscomerciais_padrao where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir condições de pagamento/recebimento de " & Texto
            ID_documento = .ListItems(InitFor)
            Documento = "Condições de pagamento/recebimento: " & .ListItems(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) condição(ões) de pagamento/recebimento de " & Texto & " antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Condição(ões) de pagamento/recebimento de " & Texto & " excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCondicoes
    ProcCarregaListaCondicoes
    Frame1(36).Enabled = False
    Novo_geral4 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_feriado()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_feriado
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) feriado(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Feriados where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir feriado"
            ID_documento = .ListItems(InitFor)
            Documento = "Data do feriado: " & .ListItems(InitFor).ListSubItems(3) & " - Descrição: " & .ListItems(InitFor).ListSubItems(4)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) feriado(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Feriado(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposFeriados
    ProcCarregaListaFeriados
    Frame1(37).Enabled = False
    Novo_geral5 = False
    ProcCarregaComboAnoFeriado
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_verificar_desconectar_usuario_Click()
On Error GoTo tratar_erro

With Txt_minutos_desconectar
    If Chk_verificar_desconectar_usuario.Value = 1 Then
        .Locked = False
        .TabStop = True
    Else
        .Text = ""
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Cmb_ano_feriado_Click()
On Error GoTo tratar_erro

ProcCarregaListaFeriados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_servidor_Change()
On Error GoTo tratar_erro

Cmb_nome_banco.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_servidor_LostFocus()
On Error GoTo tratar_erro

If Cmb_servidor <> "" Then
    With Cmb_nome_banco
        .Clear
        For Each vDb In EnumSqlDbAdo(Cmb_servidor.Text, "Procam", "PRO0902loc$?")
            .AddItem vDb
        Next
    End With
End If

Exit Sub
tratar_erro:
    If Err.Number = 13 Then
        USMsgBox ("Não foi encontrado nenhum banco de dados ao efetuar a conexão nessa instância."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_TBSN_Click()
On Error GoTo tratar_erro

ProcLimpaCamposTBSN
Select Case Mid(Cmb_tipo_TBSN, 8, 3)
    Case "I -": IDConta = 1
    Case "II ": IDConta = 2
    Case "III": IDConta = 3
    Case "IV ": IDConta = 4
    Case "V -": IDConta = 5
End Select
ProcCorrigeFormTabelaSN IDConta, False

txtCNAE_TBSN = ""
Lbl_status = "Status: Desativada"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CNAE, Ativado from Impostos_TabelaDAS where ID_empresa = " & txtIDEmpresa & " and Tabela = " & IDConta, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtCNAE_TBSN = IIf(IsNull(TBLISTA!CNAE), "", TBLISTA!CNAE)
    If TBLISTA!Ativado = True Then Lbl_status = "Status: Ativada"
    Cmd_ativar_tabelaSN.Enabled = True
Else
    Cmd_ativar_tabelaSN.Enabled = False
End If
ProcCarregaLista_TBSN Cmb_tipo_TBSN

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeFormTabelaSN(Tabela As Integer, Resize As Boolean)
On Error GoTo tratar_erro

'valor = Txt_IPI_TBSN.Width
'If Tabela = 2 Then 'Tem IPI
'    If Txt_IPI_TBSN.Visible = False And Resize = False Then
'        Lbl_ICMS_TBSN.Left = Lbl_ICMS_TBSN.Left + valor
'        Lbl_ISS_TBSN.Left = Lbl_ISS_TBSN.Left + valor
'        Txt_ICMS_TBSN.Left = Txt_ICMS_TBSN.Left + valor
'        Label3(4).Left = Label3(4).Left + valor
'        Txt_valor_deduzir_TBSN.Left = Txt_valor_deduzir_TBSN.Left + valor
'    End If
'    Label1(92).Visible = True
'    Txt_IPI_TBSN.Visible = True
'    Lista_TBSN.ColumnHeaders(10).Width = 1000
'Else
'    If Txt_IPI_TBSN.Visible = True Or Resize = True Then
'        Lbl_ICMS_TBSN.Left = Lbl_ICMS_TBSN.Left - valor
'        Lbl_ISS_TBSN.Left = Lbl_ISS_TBSN.Left - valor
'        Txt_ICMS_TBSN.Left = Txt_ICMS_TBSN.Left - valor
'        Label3(4).Left = Label3(4).Left - valor
'        Txt_valor_deduzir_TBSN.Left = Txt_valor_deduzir_TBSN.Left - valor
'    End If
'    Label1(92).Visible = False
'    Txt_IPI_TBSN.Visible = False
'    Lista_TBSN.ColumnHeaders(10).Width = 0
'End If
'
'Valor1 = Txt_CPP_TBSN.Width
'If Tabela = 4 Then 'Não tem CPP
'    If Txt_CPP_TBSN.Visible = True Or Resize = True Then
'        Lbl_ICMS_TBSN.Left = Lbl_ICMS_TBSN.Left - Valor1
'        Lbl_ISS_TBSN.Left = Lbl_ISS_TBSN.Left - Valor1
'        Txt_ICMS_TBSN.Left = Txt_ICMS_TBSN.Left - Valor1
'        Label3(4).Left = Label3(4).Left - Valor1
'        Txt_valor_deduzir_TBSN.Left = Txt_valor_deduzir_TBSN.Left - Valor1
'    End If
'    Label1(93).Visible = False
'    Txt_CPP_TBSN.Visible = False
'    Lista_TBSN.ColumnHeaders(9).Width = 0
'Else
'    If Txt_CPP_TBSN.Visible = False And Resize = False Then
'        Lbl_ICMS_TBSN.Left = Lbl_ICMS_TBSN.Left + Valor1
'        Lbl_ISS_TBSN.Left = Lbl_ISS_TBSN.Left + Valor1
'        Txt_ICMS_TBSN.Left = Txt_ICMS_TBSN.Left + Valor1
'        Label3(4).Left = Label3(4).Left + Valor1
'        Txt_valor_deduzir_TBSN.Left = Txt_valor_deduzir_TBSN.Left + Valor1
'    End If
'    Label1(93).Visible = True
'    Txt_CPP_TBSN.Visible = True
'    Lista_TBSN.ColumnHeaders(9).Width = 1000
'End If
'
'If Tabela = 3 Or Tabela = 4 Or Tabela = 5 Then
'    Lbl_ICMS_TBSN.Visible = False
'    Lbl_ISS_TBSN.Visible = True
'    Txt_ICMS_TBSN.ToolTipText = "Alíquota do ISS."
'    Lista_TBSN.ColumnHeaders(11).Text = "ISS (%)"
'    Txt_ICMS_TBSN.ToolTipText = "Alíquota do ISS."
'Else
'    Lbl_ICMS_TBSN.Visible = True
'    Lbl_ISS_TBSN.Visible = False
'    Txt_ICMS_TBSN.ToolTipText = "Alíquota do ICMS."
'    Lista_TBSN.ColumnHeaders(11).Text = "ICMS (%)"
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_uf_Click()
On Error GoTo tratar_erro

If Cmb_uf <> "" Then ProcCarregaComboCidade Cmb_cidade, "Sigla_UF = '" & Cmb_uf & "'", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_usuario_caprind_email_Click()
On Error GoTo tratar_erro

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select Email from Usuarios where Usuario = '" & Cmb_usuario_caprind_email & "' and Email IS NULL and Email <> N''", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    Txt_email_email = TBUsuarios!Email
End If
TBUsuarios.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_ativar_tabelaSN_Click()
On Error GoTo tratar_erro

If Left(Cmb_tipo_TBSN, 1) = "T" Then
    Select Case Mid(Cmb_tipo_TBSN, 8, 3)
        Case "I -": Tabela = 1
        Case "II ": Tabela = 2
        Case "III": Tabela = 3
        Case "IV ": Tabela = 4
        Case "V -": Tabela = 5
    End Select
Else
    Tabela = 6
End If
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Ativado from Impostos_TabelaDAS where ID_empresa = " & txtIDEmpresa & " and Tabela = " & Tabela, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    If TBLISTA!Ativado = True Then
        IDConta = 0
        Evento = "Desativar tabela do simples nacional"
    Else
        IDConta = 1
        Evento = "Ativar tabela do simples nacional"
    End If
    Conexao.Execute "UPDATE Impostos_TabelaDAS Set Ativado = " & IDConta & " where ID_empresa = " & txtIDEmpresa & " and Tabela = " & Tabela
    If Left(Evento, 1) = "D" Then StatusTexto = "desativada" Else StatusTexto = "ativada"
    USMsgBox ("Tabela " & StatusTexto & " com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Configuração do sistema/Opções gerais"
    ID_documento = txtIDEmpresa.Text
    Documento = "Empresa: " & txtEmpresa
    Documento1 = "Tabela: " & Cmb_tipo_TBSN
    ProcGravaEvento
    '==================================
    If Left(Evento, 1) = "D" Then Lbl_status = "Satus: Desativada" Else Lbl_status = "Satus: Ativada"
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_localizar_NFe_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Diretório dos arquivos de envio da nota fiscal"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Txt_local_armaz_NFe.Text = sBuffer & "\"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdCertificado_Click()
On Error GoTo tratar_erro
Dim Stor As New Store
Dim Cert As Certificate
Dim Certs As New Certificates
Dim CForNext As Integer

'Abrir o store
Stor.Open

Certs.Clear
For CForNext = 1 To Stor.Certificates.Count
Certs.Add Stor.Certificates.Item(CForNext)
Next CForNext

Set Certs = Certs.Select("LaRoche", "Selecione o Certificado Digital.", False)

'Exibir mensagem com data de validade do certificado
'cert.
For Each Cert In Certs
'=============================================
'Verifica se é A1 ou A3
'=============================================
'Var1 = Year(Cert.ValidToDate)
'Var2 = Year(Cert.ValidFromDate)

  If Int(Year(Cert.ValidToDate)) - Int(Year(Cert.ValidFromDate)) = 1 Then
   txttpEmissor.Text = "A1"
  Else
  txttpEmissor.Text = "A3"
  End If
'=============================================
    USMsgBox "Nome razão: " & Cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME) & vbCrLf & "Certificado válido até: " & Cert.ValidToDate, vbInformation, "CAPRIND v5.0" '(CAPICOM_CHECK_TIME_VALIDITY)
    txtCertificadodigital.Text = Cert.SerialNumber
    'txtval.Text = IIf(Cert.IsValid = True, "Sim", "Não")
Next


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdConsultar_Click()
On Error GoTo tratar_erro
  
frmEmpresa_Contrato.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdLocalizarRetorno_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Diretório dos arquivos de retorno da nota fiscal"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtRetornoNF.Text = sBuffer & "\"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdLocalizarXMLDanfe_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Diretório dos arquivos XML e Danfe"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtCaminhoXMLDanfe.Text = sBuffer & "\"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_localizar_armaz_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Localizar local de armazenamento padrão (.PDF)"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Txt_local_armaz = sBuffer
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_servidor_Click()
On Error GoTo tratar_erro

Cmb_nome_banco.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_valor_faturado_mes_Click()
On Error GoTo tratar_erro

frmOpcoesGeral_TotalFat.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCnae_Click()
On Error GoTo tratar_erro

If txtIDEmpresa <> "" Then frmOpcoesGeral_CNAE.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_armaz_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_armaz
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_armaz, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_armaz_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_armaz.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Empresa_armazenamento_PDF where ID = " & Lista_armaz.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosArmaz
    CodigoLista8 = Lista_armaz.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_cond1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_cond1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_cond1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_email_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_email
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_email, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_email_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_email.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Empresa_email where ID = " & Lista_email.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimpaCamposEmail
    ProcCarregaDadosEmail
    CodigoLista6 = Lista_email.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_filtros_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_filtros
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_filtros, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_filtros_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_filtros.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Empresa_filtros where ID = " & Lista_filtros.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosFiltros
    CodigoLista7 = Lista_filtros.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_TBSN_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_TBSN
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_TBSN, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_TBSN_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_TBSN.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Impostos_TabelaDAS where ID = " & Lista_TBSN.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposTBSN
    ProcCarregaDadosTBSN
    CodigoLista = Lista_TBSN.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab5_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab5.Tab = 1 Then ProcCarregaLista_TBSN Cmb_tipo_TBSN

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Cofins_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_Cofins_TBSN <> "" Then
    VerifNumero = Txt_Cofins_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_Cofins_TBSN = ""
        Txt_Cofins_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Cofins_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_Cofins_TBSN = Format(Txt_Cofins_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_CPP_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_CPP_TBSN <> "" Then
    VerifNumero = Txt_CPP_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_CPP_TBSN = ""
        Txt_CPP_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_CPP_TBSN_LostFocus()
On Error GoTo tratar_erro

 Txt_CPP_TBSN = Format(Txt_CPP_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_CSLL_TBSN_Change()
On Error GoTo tratar_erro
    
If CSLL_TBSN <> "" Then
    VerifNumero = CSLL_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        CSLL_TBSN = ""
        CSLL_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_CSLL_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_CSLL_TBSN = Format(Txt_CSLL_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_email_email_LostFocus()
On Error GoTo tratar_erro

Txt_email_email = LCase(Txt_email_email)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_email_LostFocus()
On Error GoTo tratar_erro

Txt_email = LCase(Txt_email)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_IPI_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_IPI_TBSN <> "" Then
    VerifNumero = Txt_IPI_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_IPI_TBSN = ""
        Txt_IPI_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_IPI_TBSN_LostFocus()
On Error GoTo tratar_erro

 Txt_IPI_TBSN = Format(Txt_IPI_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_IRPJ_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_IRPJ_TBSN <> "" Then
    VerifNumero = Txt_IRPJ_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_IRPJ_TBSN = ""
        Txt_IRPJ_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_IRPJ_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_IRPJ_TBSN = Format(Txt_IRPJ_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_PIS_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_PIS_TBSN <> "" Then
    VerifNumero = Txt_PIS_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_PIS_TBSN = ""
        Txt_PIS_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_PIS_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_PIS_TBSN = Format(Txt_PIS_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txt_porta_email_Change()
On Error GoTo tratar_erro

If txt_porta_email <> "" Then
    VerifNumero = txt_porta_email
    ProcVerificaNumero
    If VerifNumero = False Then
        txt_porta_email = ""
        txt_porta_email.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_senha_Change()
On Error GoTo tratar_erro

Cmb_nome_banco.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_site_LostFocus()
On Error GoTo tratar_erro

Txt_site = LCase(Txt_site)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'Private Sub Txt_usuario_Change()
'On Error GoTo tratar_erro
'
'Cmb_nome_banco.Clear
'
'Exit Sub
'tratar_erro:
'    usMsgbox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Private Sub Cmd_localizar_rel_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Localizar local dos relatórios"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtLocalrel.Text = sBuffer
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_empresa()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposEmpresa
Novo_geral = True
Frame1(3).Enabled = True
txtRazao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoTabelaSN()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposTBSN
Novo_geral9 = True
Frame1(47).Enabled = True
Frame1(45).Enabled = True
Txt_de_TBSN.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_email()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposEmail
Novo_geral6 = True
Frame1(30).Enabled = True
Cmb_aplicacao_email.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_filtros()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposFiltros
Novo_geral7 = True
Frame1(31).Enabled = True
cmbAplicacao_Filtros.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_armaz()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposArmaz
Novo_geral8 = True
Frame1(32).Enabled = True
Cmb_relatorio_armaz.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_unidade()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposUnidade
Novo_geral2 = True
Frame1(34).Enabled = True
Txt_unidade.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_conversao()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposConversao
Novo_geral3 = True
Frame1(35).Enabled = True
Cmb_unidade_de_conversao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_condicao()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposCondicoes
Novo_geral4 = True
Frame1(36).Enabled = True
Txt_texto_cond.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_feriado()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposFeriados
Novo_geral5 = True
Frame1(37).Enabled = True
Cmb_data_feriado.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlterarBD()
On Error GoTo tratar_erro

frmOpcoesGeral2_Subs.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirBD()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With listaBancos
    If .ListItems.Count = 0 Then
        USMsgBox ("Informe o local do banco de dados que deseja excluir."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If USMsgBox("Deseja realmente excluir esta configuração? " & vbCrLf & "Nome da instância SQL: " & .SelectedItem.ListSubItems(2) & vbCrLf & "Nome do banco de dados: " & .SelectedItem.ListSubItems(3), vbYesNo) = vbYes Then
        If .SelectedItem.ListSubItems(1) = Localrel And .SelectedItem.ListSubItems(2) = NomeServidor And .SelectedItem.ListSubItems(3) = Nome_banco Then
            USMsgBox ("Não é permitido excluir, pois essa configuração está em uso."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        ElseIf .SelectedItem.ListSubItems(1) = Localrel1 And .SelectedItem.ListSubItems(2) = NomeServidor1 And .SelectedItem.ListSubItems(3) = Nome_banco1 Then
                DeleteSetting "Procam", "CaprindSQL", "NomeServidor1"
                DeleteSetting "Procam", "CaprindSQL", "LocalRel1"
                DeleteSetting "Procam", "CaprindSQL", "Nome_banco1"
                If Usuario_banco1 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Usuario_banco1"
                If Senha_banco1 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Senha_banco1"
                Nome_banco1 = ""
                Localrel1 = ""
            ElseIf .SelectedItem.ListSubItems(1) = Localrel2 And .SelectedItem.ListSubItems(2) = NomeServidor2 And .SelectedItem.ListSubItems(3) = Nome_banco2 Then
                    DeleteSetting "Procam", "CaprindSQL", "NomeServidor2"
                    DeleteSetting "Procam", "CaprindSQL", "LocalRel2"
                    DeleteSetting "Procam", "CaprindSQL", "Nome_banco2"
                    If Usuario_banco2 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Usuario_banco2"
                    If Senha_banco2 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Senha_banco2"
                    Nome_banco2 = ""
                    Localrel2 = ""
                ElseIf .SelectedItem.ListSubItems(1) = Localrel3 And .SelectedItem.ListSubItems(2) = NomeServidor3 And .SelectedItem.ListSubItems(3) = Nome_banco3 Then
                        DeleteSetting "Procam", "CaprindSQL", "NomeServidor3"
                        DeleteSetting "Procam", "CaprindSQL", "LocalRel3"
                        DeleteSetting "Procam", "CaprindSQL", "Nome_banco3"
                        If Usuario_banco3 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Usuario_banco3"
                        If Senha_banco3 <> "" Then DeleteSetting "Procam", "CaprindSQL", "Senha_banco3"
                        Nome_banco3 = ""
                        Localrel3 = ""
                    Else
                        USMsgBox ("Configuração não encontrada nos registros do windows."), vbExclamation, "CAPRIND v5.0"
            End If
        USMsgBox ("Configuração excluída com sucesso."), vbInformation, "CAPRIND v5.0"
        ProcLimpaCamposBanco
        ProcCarregaBancoDados
        ProcCarregaListaBancos
        Novo_Local = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_moeda()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaMoeda
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) moeda(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from moeda where codigo = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir moeda"
            ID_documento = .ListItems(InitFor)
            Documento = "Moeda: " & .ListItems(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) moeda(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Moeda(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposMoeda
    ProcCarregaListaMoeda
    Frame1(33).Enabled = False
    Novo_geral1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_geral = True Then
    If USMsgBox("A empresa ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_empresa
        If Novo_geral = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral1 = True Then
    If USMsgBox("A moeda ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_Moeda
        If Novo_geral1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral2 = True Then
    If USMsgBox("A unidade ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_unidade
        If Novo_geral2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral3 = True Then
    If USMsgBox("A regra de conversão ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_conversao
        If Novo_geral3 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral4 = True Then
    If USMsgBox("A condição de pagamento/recebimento ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_condicao
        If Novo_geral4 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral5 = True Then
    If USMsgBox("O feriado ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_feriado
        If Novo_geral5 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral6 = True Then
    If USMsgBox("O e-mail ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_email
        If Novo_geral6 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral7 = True Then
    If USMsgBox("O filtro ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_Filtros
        If Novo_geral7 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral8 = True Then
    If USMsgBox("O local de armazenamento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_Armaz
        If Novo_geral8 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_geral9 = True Then
    If USMsgBox("A tabela do simples nacional ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarTabelaSN
        If Novo_geral9 = True Then Exit Sub Else Unload Me
    End If
End If
Novo_geral = False
Novo_geral1 = False
Novo_geral2 = False
Novo_geral3 = False
Novo_geral4 = False
Novo_geral5 = False
Novo_geral6 = False
Novo_geral7 = False
Novo_geral8 = False
Novo_geral9 = False
Unload Me

frmMDI.ProcVerificaLogoffAutomatico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarBD()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Local = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtLocalrel.Text = "" Then
    NomeCampo = "o local dos relatórios"
    ProcVerificaAcao
    Cmd_localizar_rel.SetFocus
    Exit Sub
End If
If Cmb_servidor = "" Then
    NomeCampo = "o nome da instância SQL"
    ProcVerificaAcao
    Cmb_servidor.SetFocus
    Exit Sub
End If
If Cmb_nome_banco = "" Then
    NomeCampo = "o nome do banco de dados"
    ProcVerificaAcao
    Cmb_nome_banco.SetFocus
    Exit Sub
End If
If txtlocalantigo.Text = "" Then
    NomeCampo = "o local onde esta armazenado os arquivos antigos"
    ProcVerificaAcao
    cmdLocalantigo.SetFocus
    Exit Sub
End If
If txtlocalnovo.Text = "" Then
    NomeCampo = "o local onde esta armazenado os novos arquivos"
    ProcVerificaAcao
    cmdLocalnovo.SetFocus
    Exit Sub
End If
Caprind = "\Caprind.exe"
Gerprod = "\Gerprod.exe"
If Cmb_servidor = NomeServidor And Cmb_nome_banco = Nome_banco Or Cmb_servidor = NomeServidor1 And Cmb_nome_banco = Nome_banco1 Or Cmb_servidor = NomeServidor2 And Cmb_nome_banco = Nome_banco2 Then
    USMsgBox ("Essa configuração de instância SQL e banco de dados já foi cadastrada, favor alterar."), vbExclamation, "CAPRIND v5.0"
    Cmb_servidor.SetFocus
    Exit Sub
End If

Permitido = True
ProcFunAbreBD_Configuracao Cmb_servidor, Cmb_nome_banco, "Procam", "PRO0902loc$?"
If Permitido = False Then
    USMsgBox "Não foi possivel salvar pois não foi econtrado essa instância e banco de dados.", vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Localrel = "" Then
    NomeServidor = Cmb_servidor
    SaveSetting "Procam", "CaprindSQL", "NomeServidor", NomeServidor
    
    Localrel = txtLocalrel.Text
    SaveSetting "Procam", "CaprindSQL", "LocalRel", Localrel
    
    Nome_banco = Cmb_nome_banco
    SaveSetting "Procam", "CaprindSQL", "Nome_banco", Nome_banco
    
    'Usuario_banco = Txt_usuario
    'SaveSetting "Procam", "CaprindSQL", "Usuario_banco", Usuario_banco
    
    'S'enha_banco = Txt_senha
    'SaveSetting "Procam", "CaprindSQL", "Senha_banco", Senha_banco
    
    LocalAntigoCaprind = txtlocalantigo.Text & Caprind
    LocalAntigoGerprod = txtlocalantigo.Text & Gerprod
    LocalNovoCaprind = txtlocalnovo.Text & Caprind
    LocalNovoGerprod = txtlocalnovo.Text & Gerprod
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind", LocalAntigoCaprind
    SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind", LocalNovoCaprind
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod", LocalAntigoGerprod
    SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod", LocalNovoGerprod
    USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
ElseIf Localrel1 = "" Then
        NomeServidor1 = Cmb_servidor
        SaveSetting "Procam", "CaprindSQL", "NomeServidor1", NomeServidor1
        
        Localrel1 = txtLocalrel.Text
        SaveSetting "Procam", "CaprindSQL", "LocalRel1", Localrel1
        
        Nome_banco1 = Cmb_nome_banco
        SaveSetting "Procam", "CaprindSQL", "Nome_banco1", Nome_banco1
        
        'Usuario_banco1 = Txt_usuario
        'SaveSetting "Procam", "CaprindSQL", "Usuario_banco1", Usuario_banco1
        '
        'Senha_banco1 = Txt_senha
        'SaveSetting "Procam", "CaprindSQL", "Senha_banco1", Senha_banco1
        
        LocalAntigoCaprind1 = txtlocalantigo.Text & Caprind
        LocalAntigoGerprod1 = txtlocalantigo.Text & Gerprod
        LocalNovoCaprind1 = txtlocalnovo.Text & Caprind
        LocalNovoGerprod1 = txtlocalnovo.Text & Gerprod
        SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind1", LocalAntigoCaprind1
        SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind1", LocalNovoCaprind1
        SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod1", LocalAntigoGerprod1
        SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod1", LocalNovoGerprod1
        USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
    ElseIf Localrel2 = "" Then
            NomeServidor2 = Cmb_servidor
            SaveSetting "Procam", "CaprindSQL", "NomeServidor2", NomeServidor2
            
            Localrel2 = txtLocalrel.Text
            SaveSetting "Procam", "CaprindSQL", "LocalRel2", Localrel2
            
            Nome_banco2 = Cmb_nome_banco
            SaveSetting "Procam", "CaprindSQL", "Nome_banco2", Nome_banco2
            
            'Usuario_banco2 = Txt_usuario
            'SaveSetting "Procam", "CaprindSQL", "Usuario_banco2", Usuario_banco2
            
            'Senha_banco2 = Txt_senha
            'SaveSetting "Procam", "CaprindSQL", "Senha_banco2", Senha_banco2
            
            LocalAntigoCaprind2 = txtlocalantigo.Text & Caprind
            LocalAntigoGerprod2 = txtlocalantigo.Text & Gerprod
            LocalNovoCaprind2 = txtlocalnovo.Text & Caprind
            LocalNovoGerprod2 = txtlocalnovo.Text & Gerprod
            SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind2", LocalAntigoCaprind2
            SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind2", LocalNovoCaprind2
            SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod2", LocalAntigoGerprod2
            SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod2", LocalNovoGerprod2
            USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
        ElseIf Localrel3 = "" Then
                NomeServidor3 = Cmb_servidor
                SaveSetting "Procam", "CaprindSQL", "NomeServidor3", NomeServidor3
                
                Localrel3 = txtLocalrel.Text
                SaveSetting "Procam", "CaprindSQL", "LocalRel3", Localrel3
                
                Nome_banco3 = Cmb_nome_banco
                SaveSetting "Procam", "CaprindSQL", "Nome_banco3", Nome_banco3
                
'                Usuario_banco3 = Txt_usuario
'                SaveSetting "Procam", "CaprindSQL", "Usuario_banco3", Usuario_banco3
'
'                Senha_banco3 = Txt_senha
'                SaveSetting "Procam", "CaprindSQL", "Senha_banco3", Senha_banco3
                
                LocalAntigoCaprind3 = txtlocalantigo.Text & Caprind
                LocalAntigoGerprod3 = txtlocalantigo.Text & Gerprod
                LocalNovoCaprind3 = txtlocalnovo.Text & Caprind
                LocalNovoGerprod3 = txtlocalnovo.Text & Gerprod
                SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind3", LocalAntigoCaprind3
                SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind3", LocalNovoCaprind3
                SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod3", LocalAntigoGerprod3
                SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod3", LocalNovoGerprod3
                USMsgBox "Cadastro realizado com sucesso.", vbInformation, "CAPRIND v5.0"
            Else
                 USMsgBox ("Você só pode armazenar quatro configurações diferentes."), vbExclamation, "CAPRIND v5.0"
End If

ProcBloqueiaCampos
Salvarrel = True
Main
Salvarrel = False
Novo_Local = False
ProcCarregaListaBancos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_empresa()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(3).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If txtRazao.Text = "" Then
    NomeCampo = "a razão social"
    ProcVerificaAcao
    txtRazao.SetFocus
    Exit Sub
End If
If txtcnpj.Text = "__.___.___/____-__" Then
    NomeCampo = "o CNPJ da empresa"
    ProcVerificaAcao
    txtcnpj.SetFocus
    Exit Sub
End If
If txtRG_IE.Text = "" Then
    NomeCampo = "a IE"
    ProcVerificaAcao
    txtRG_IE.SetFocus
    Exit Sub
End If
If txtEmpresa.Text = "" Then
    NomeCampo = "o nome fantasia"
    ProcVerificaAcao
    txtEmpresa.SetFocus
    Exit Sub
End If
If txtendereco <> "" And cmbTipo_endereco = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_endereco.SetFocus
    Exit Sub
End If
If txtendereco = "" Then
    NomeCampo = "o endereço"
    ProcVerificaAcao
    txtendereco.SetFocus
    Exit Sub
End If
If txtNumero.Text = "" Then
    NomeCampo = "o número"
    ProcVerificaAcao
    txtNumero.SetFocus
    Exit Sub
End If
If txt_Bairro <> "" And cmbTipo_bairro = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_bairro.SetFocus
    Exit Sub
End If
If txt_Bairro = "" Then
    NomeCampo = "o bairro"
    ProcVerificaAcao
    txt_Bairro.SetFocus
    Exit Sub
End If
If Cmb_cidade = "" Then
    NomeCampo = "a cidade"
    ProcVerificaAcao
    Cmb_cidade.SetFocus
    Exit Sub
End If
If Cmb_uf = "" Then
    NomeCampo = "o estado"
    ProcVerificaAcao
    Cmb_uf.SetFocus
    Exit Sub
End If
If Txt_pais = "" Then
    NomeCampo = "o país"
    ProcVerificaAcao
    Txt_pais.SetFocus
    Exit Sub
End If
If txtEndereco_cob.Text = "" Then
    NomeCampo = "o endereço de cobrança"
    ProcVerificaAcao
    txtEndereco_cob.SetFocus
    Exit Sub
End If
'If FunVerificaCidade(Txt_cidade, Cmb_uf) = False Then Exit Sub

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from empresa where codigo = " & txtIDEmpresa.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!Empresa <> txtEmpresa Then
        'Transportadora
        Conexao.Execute "Update Compras_fornecedores Set Transportadora = '" & txtEmpresa & "' where IDTransp = " & IIf(txtIDEmpresa = "", 0, txtIDEmpresa) & " and Tipo_transp = 'E'"
        Conexao.Execute "Update vendas_comercial Set Transportadora = '" & txtEmpresa & "' where IDInttransp = " & IIf(txtIDEmpresa = "", 0, txtIDEmpresa) & " and Tipo_transp = 'E'"
        Conexao.Execute "Update tbl_Dados_Transp Set txt_Razao = '" & txtEmpresa & "' where IdIntTransp = " & IIf(txtIDEmpresa = "", 0, txtIDEmpresa) & " and Tipo_transp = 'E'"
    End If
Else
    TBProduto.AddNew
End If
ProcEnviaDadosEmpresa
TBProduto.Update
txtIDEmpresa = TBProduto!CODIGO
TBProduto.Close
ProcCarregaListaEmpresa
If Novo_geral = True Then
    USMsgBox ("Nova empresa cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova empresa"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar empresa"
    If CodigoLista <> 0 And Lista_empresas.ListItems.Count <> 0 Then
        Lista_empresas.SelectedItem = Lista_empresas.ListItems(CodigoLista)
        Lista_empresas.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = txtIDEmpresa.Text
Documento = "Empresa: " & txtEmpresa
Documento1 = ""
ProcGravaEvento
'==================================
Novo_geral = False

If TemInternet = True And ErroDriverMYSQL = False Then ProcSalvarEmpresaSite

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarTabelaSN()
On Error GoTo tratar_erro

If Frame1(45).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
valor = IIf(Txt_de_TBSN = "", 0, Txt_de_TBSN)
If Txt_de_TBSN = "" Or valor < 0 Then
    NomeCampo = "o valor de"
    ProcVerificaAcao
    Txt_de_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_ate_TBSN = "", 0, Txt_ate_TBSN)
If Txt_ate_TBSN = "" Or valor < 0 Then
    NomeCampo = "o valor até"
    ProcVerificaAcao
    Txt_ate_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_Aliquota_TBSN = "", 0, Txt_Aliquota_TBSN)
If valor < 0 Then
    NomeCampo = "a alíquota"
    ProcVerificaAcao
    Txt_Aliquota_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_IRPJ_TBSN = "", 0, Txt_IRPJ_TBSN)
If valor < 0 Then
    NomeCampo = "a alíquota do IRPJ"
    ProcVerificaAcao
    Txt_IRPJ_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_CSLL_TBSN = "", 0, Txt_CSLL_TBSN)
If valor < 0 Then
    NomeCampo = "a alíquota do CSLL"
    ProcVerificaAcao
    Txt_CSLL_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_Cofins_TBSN = "", 0, Txt_Cofins_TBSN)
If valor < 0 Then
    NomeCampo = "a alíquota do Cofins"
    ProcVerificaAcao
    Txt_Cofins_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_PIS_TBSN = "", 0, Txt_PIS_TBSN)
If valor < 0 Then
    NomeCampo = "a alíquota do PIS"
    ProcVerificaAcao
    Txt_PIS_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_CPP_TBSN = "", 0, Txt_CPP_TBSN)
If valor < 0 Then
    NomeCampo = "a alíquota do CPP"
    ProcVerificaAcao
    Txt_CPP_TBSN.SetFocus
    Exit Sub
End If
If Txt_IPI_TBSN.Visible = True Then
    valor = IIf(Txt_IPI_TBSN = "", 0, Txt_IPI_TBSN)
    If valor < 0 Then
        NomeCampo = "a alíquota do IPI"
        ProcVerificaAcao
        Txt_IPI_TBSN.SetFocus
        Exit Sub
    End If
End If
valor = IIf(Txt_ICMS_TBSN = "", 0, Txt_ICMS_TBSN)
If valor < 0 Then
    NomeCampo = "a alíquota do " & IIf(Lbl_ICMS_TBSN.Visible = True, "ICMS", "ISS")
    ProcVerificaAcao
    Txt_ICMS_TBSN.SetFocus
    Exit Sub
End If
valor = IIf(Txt_valor_deduzir_TBSN = "", 0, Txt_valor_deduzir_TBSN)
If valor < 0 Then
    NomeCampo = "o valor a deduzir"
    ProcVerificaAcao
    Txt_valor_deduzir_TBSN.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Impostos_TabelaDAS where ID = " & Txt_ID_TBSN, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_empresa = txtIDEmpresa
If Left(Cmb_tipo_TBSN, 1) = "T" Then
    Select Case Mid(Cmb_tipo_TBSN, 8, 3)
        Case "I -": TBGravar!Tabela = 1
        Case "II ": TBGravar!Tabela = 2
        Case "III": TBGravar!Tabela = 3
        Case "IV ": TBGravar!Tabela = 4
    End Select
Else
    TBGravar!Tabela = 6
End If
TBGravar!De = Txt_de_TBSN
TBGravar!Ate = Txt_ate_TBSN
TBGravar!DAS = Txt_Aliquota_TBSN
TBGravar!IRPJ = Txt_IRPJ_TBSN
TBGravar!CSLL = Txt_CSLL_TBSN
TBGravar!Cofins = Txt_Cofins_TBSN
TBGravar!PIS = Txt_PIS_TBSN
TBGravar!cpp = Txt_CPP_TBSN
If Txt_IPI_TBSN.Visible = True Then TBGravar!IPI = Txt_IPI_TBSN
If Lbl_ICMS_TBSN.Visible = True Then TBGravar!ICMS = Txt_ICMS_TBSN Else TBGravar!ISS = Txt_ICMS_TBSN
TBGravar!Valor_deduzir = Txt_valor_deduzir_TBSN
TBGravar.Update
Txt_ID_TBSN = TBGravar!ID
Conexao.Execute "Update Impostos_TabelaDAS Set CNAE = '" & txtCNAE_TBSN & "' where ID_empresa = " & txtIDEmpresa & " and Tabela = " & TBGravar!Tabela
TBGravar.Close

ProcCarregaLista_TBSN Cmb_tipo_TBSN
If Novo_geral9 = True Then
    USMsgBox ("Novo registro cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo registro da tabela do DAS"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar registro da tabela do DAS"
    If CodigoLista <> 0 And Lista_TBSN.ListItems.Count <> 0 Then
        Lista_TBSN.SelectedItem = Lista_TBSN.ListItems(CodigoLista)
        Lista_TBSN.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID_TBSN
Documento = "Empresa: " & txtRazao
Documento1 = "De: " & Format(Txt_de_TBSN, "###,##0.00") & " - Até: " & Format(Txt_ate_TBSN, "###,##0.00") & " - DAS: " & Format(Txt_Aliquota_TBSN, "###,##0.00") & " - ICMS: " & Format(Txt_ICMS_TBSN, "###,##0.00")
ProcGravaEvento
'==================================
Novo_geral9 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_email()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(30).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If Cmb_aplicacao_email = "" Then
    NomeCampo = "a aplicação"
    ProcVerificaAcao
    Cmb_aplicacao_email.SetFocus
    Exit Sub
End If
If Cmb_usuario_caprind_email = "" Then
    NomeCampo = "o usuário do caprind"
    ProcVerificaAcao
    Cmb_usuario_caprind_email.SetFocus
    Exit Sub
End If
If Txt_servidor_SMTP_email = "" Then
    NomeCampo = "o servidor SMTP"
    ProcVerificaAcao
    Txt_servidor_SMTP_email.SetFocus
    Exit Sub
End If
If txt_porta_email = "" Then
    NomeCampo = "a porta"
    ProcVerificaAcao
    txt_porta_email.SetFocus
    Exit Sub
End If
If Cmb_seguranca_email = "" Then
    NomeCampo = "a segurança"
    ProcVerificaAcao
    Cmb_seguranca_email.SetFocus
    Exit Sub
End If
If Txt_nome_email = "" Then
    NomeCampo = "o nome"
    ProcVerificaAcao
    Txt_nome_email.SetFocus
    Exit Sub
End If
If Txt_email_email.Text = "" Then
    NomeCampo = "o e-mail"
    ProcVerificaAcao
    Txt_email_email.SetFocus
    Exit Sub
End If
If Txt_usuario_email = "" Then
    NomeCampo = "o usuário"
    ProcVerificaAcao
    Txt_usuario_email.SetFocus
    Exit Sub
End If
If Txt_senha_email = "" Then
    NomeCampo = "a senha"
    ProcVerificaAcao
    Txt_senha_email.SetFocus
    Exit Sub
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Empresa_email where ID = " & Txt_ID_email, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then TBProduto.AddNew
ProcEnviaDadosEmail
TBProduto.Update
Txt_ID_email = TBProduto!ID
TBProduto.Close
ProcCarregaListaEmail
If Novo_geral6 = True Then
    USMsgBox ("Novo e-mail cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo e-mail"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar e-mail"
    If CodigoLista6 <> 0 And Lista_email.ListItems.Count <> 0 Then
        Lista_email.SelectedItem = Lista_email.ListItems(CodigoLista6)
        Lista_email.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID_email
Documento = "Empresa: " & txtEmpresa
Documento1 = "E-mail: " & Txt_email_email.Text
ProcGravaEvento
'==================================
Novo_geral6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Filtros()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(31).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If cmbAplicacao_Filtros = "" Then
    NomeCampo = "a aplicação"
    ProcVerificaAcao
    cmbAplicacao_Filtros.SetFocus
    Exit Sub
End If
If cmbfiltrarpor_Filtros = "" Then
    NomeCampo = "o filtro"
    ProcVerificaAcao
    cmbfiltrarpor_Filtros.SetFocus
    Exit Sub
End If
If cmbTipo_Filtros = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_Filtros.SetFocus
    Exit Sub
End If
If cmbFrase_Filtros = "" Then
    NomeCampo = "a frase"
    ProcVerificaAcao
    cmbFrase_Filtros.SetFocus
    Exit Sub
End If

Select Case cmbAplicacao_Filtros
    Case "Compras": Aplicacao = "C"
    Case "Vendas": Aplicacao = "V"
    Case "Qualidade": Aplicacao = "Q"
    Case "Engenharia": Aplicacao = "E"
    Case "PCP": Aplicacao = "P"
    Case "Estoque": Aplicacao = "T"
    Case "Faturamento": Aplicacao = "F"
    Case "Manutenção": Aplicacao = "M"
    Case "Outros": Aplicacao = "O"
End Select
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select ID from Empresa_filtros where ID <> " & txtID_Filtros & " and Aplicacao = '" & Aplicacao & "' and Tipo = '" & cmbTipo_Filtros & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Não é possivel salvar, pois já existe cadastro para a aplicação: " & cmbAplicacao_Filtros & " e tipo: " & cmbTipo_Filtros & "."), vbExclamation, "CAPRIND v5.0"
    TBProduto.Close
    Exit Sub
End If
TBProduto.Close

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Empresa_filtros where ID = " & txtID_Filtros, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then
    TBProduto.AddNew
    TBProduto!Data = Date
    TBProduto!Responsavel = pubUsuario
End If
ProcEnviaDadosFiltros
TBProduto.Update
txtID_Filtros = TBProduto!ID
TBProduto.Close
ProcCarregaListaFiltros
If Novo_geral7 = True Then
    USMsgBox ("Novo filtro cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo filtro"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar filtro"
    If CodigoLista7 <> 0 And Lista_filtros.ListItems.Count <> 0 Then
        Lista_filtros.SelectedItem = Lista_filtros.ListItems(CodigoLista7)
        Lista_filtros.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = txtID_Filtros
Documento = "Empresa: " & txtEmpresa
Documento1 = "Filtro: " & cmbfiltrarpor_Filtros
ProcGravaEvento
'==================================
Novo_geral7 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Armaz()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(32).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If Cmb_relatorio_armaz = "" Then
    NomeCampo = "o relatório"
    ProcVerificaAcao
    Cmb_relatorio_armaz.SetFocus
    Exit Sub
End If
If Txt_local_armaz = "" Then
    NomeCampo = "o local de armazenamento"
    ProcVerificaAcao
    Txt_local_armaz.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select ID from Empresa_armazenamento_PDF where ID <> " & Txt_ID_armaz & " and Relatorio = '" & Cmb_relatorio_armaz & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Não é possivel salvar, pois já existe cadastro para este relatório."), vbExclamation, "CAPRIND v5.0"
    TBProduto.Close
    Exit Sub
End If
TBProduto.Close

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Empresa_armazenamento_PDF where ID = " & Txt_ID_armaz, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then
    TBProduto.AddNew
    TBProduto!Data = Date
    TBProduto!Responsavel = pubUsuario
End If
ProcEnviaDadosArmaz
TBProduto.Update
Txt_ID_armaz = TBProduto!ID
TBProduto.Close
ProcCarregaListaArmaz
If Novo_geral8 = True Then
    USMsgBox ("Novo local de armazenamento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo local de armazenamento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar local de armazenamento"
    If CodigoLista8 <> 0 And Lista_armaz.ListItems.Count <> 0 Then
        Lista_armaz.SelectedItem = Lista_armaz.ListItems(CodigoLista8)
        Lista_armaz.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID_armaz
Documento = "Empresa: " & txtEmpresa
Documento1 = "Relatório: " & Cmb_relatorio_armaz & " - Local de armazenamento: " & Txt_local_armaz
ProcGravaEvento
'==================================
Novo_geral8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarEmpresaSite()
On Error GoTo tratar_erro

FunAbreBDSite
If ConexaoMySql.State = 1 Then
    Set TBMySQL = New ADODB.Recordset
    TBMySQL.Open "Select * From Clientes Where CNPJ = '" & txtcnpj & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
    With TBMySQL
        If .EOF = True Then
            .AddNew
            .Fields!Data = Date
            .Fields!Responsavel = pubUsuario
            .Fields!Cargo = pubSetor
            .Fields!Liberado = "SIM"
        End If
        .Fields!NomeRazao = txtRazao
        .Fields!CNPJ = txtcnpj
        .Fields!telefone = Txt_telefones
        .Fields!Produto = "Caprind"
        .Fields!Email = IIf(Txt_email = "", Null, LCase(Txt_email))
        .Fields!Atualizacao_automatica = Chk_atualizacao_autom.Value
        .Update
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Moeda()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(33).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtMoeda.Text = "" Then
    NomeCampo = "a descrição da moeda"
    ProcVerificaAcao
    txtMoeda.SetFocus
    Exit Sub
End If
If txtSimbolo.Text = "" Then
    NomeCampo = "o símbolo da moeda"
    ProcVerificaAcao
    txtSimbolo.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from moeda where codigo = " & txtidmoeda.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then
    TBProduto.AddNew
Else
    If txtMoeda <> TBProduto!Moeda Then
        Conexao.Execute "Update Vendas_comercial Set moeda = '" & txtMoeda & "' where moeda = '" & TBProduto!Moeda & "'"
        Conexao.Execute "Update tbl_dados_nota_fiscal Set moeda = '" & txtMoeda & "' where moeda = '" & TBProduto!Moeda & "'"
    End If
End If
ProcEnviadadosMoeda
TBProduto.Update
txtidmoeda = TBProduto!CODIGO
TBProduto.Close
ProcCarregaListaMoeda
If Novo_geral1 = True Then
    USMsgBox ("Nova moeda cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova moeda"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar moeda"
    If CodigoLista1 <> 0 And ListaMoeda.ListItems.Count <> 0 Then
        ListaMoeda.SelectedItem = ListaMoeda.ListItems(CodigoLista1)
        ListaMoeda.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = txtidmoeda.Text
Documento = "Moeda: " & txtMoeda
Documento1 = ""
ProcGravaEvento
'==================================
Novo_geral1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_unidade()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(34).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If Txt_unidade.Text = "" Then
    NomeCampo = "a unidade"
    ProcVerificaAcao
    Txt_unidade.SetFocus
    Exit Sub
End If
If Txt_descricao_unidade = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Txt_descricao_unidade.SetFocus
    Exit Sub
End If

'Verificar se já existe uma unidade cadastrada com a sigla ou descrição
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Unidade_Medida where Unidade = '" & Txt_unidade & "' and codigo <> " & txtidunidade, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Esta unidade " & Txt_unidade & " já esta cadastrada, favor alterar."), vbExclamation, "CAPRIND v5.0"
    Txt_unidade.SetFocus
    TBProduto.Close
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Unidade_Medida where Descricao = '" & Txt_descricao_unidade & "' and codigo <> " & txtidunidade, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Esta descrição está sendo utilizada, favor alterar."), vbExclamation, "CAPRIND v5.0"
    Txt_descricao_unidade.SetFocus
    TBProduto.Close
    Exit Sub
End If
TBProduto.Close

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Unidade_Medida where codigo = " & txtidunidade.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then
    TBProduto.AddNew
Else
    If Txt_unidade <> TBProduto!Unidade Then
        Conexao.Execute "Update Compras_pedido_lista Set UN = '" & Txt_unidade & "' where UN = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Compras_pedido_lista Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Compras_programacao Set UN = '" & Txt_unidade & "' where UN = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Compras_programacao Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Estoque_Controle Set UN = '" & Txt_unidade & "' where UN = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Manutencao_defeito Set Unidade = '" & Txt_unidade & "' where Unidade = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Manutencao_defeito Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Producaomaterial Set Unidade = '" & Txt_unidade & "' where Unidade = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Projconjunto Set Unidade = '" & Txt_unidade & "' where Unidade = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update projproduto Set Unidade = '" & Txt_unidade & "' where Unidade = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update projproduto Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Requisicao_materiais_lista Set UN = '" & Txt_unidade & "' where UN = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Requisicao_materiais_lista Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update tbl_Detalhes_Nota Set Txt_Unid = '" & Txt_unidade & "' where Txt_Unid = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update tbl_Detalhes_Nota Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_analise Set Unidade = '" & Txt_unidade & "' where Unidade = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_analise Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_analise_ProdutosProcessos Set Un = '" & Txt_unidade & "' where Un = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_analise_ProdutosProcessos Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_analise_setores Set Un = '" & Txt_unidade & "' where Un = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_analise_setores Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update vendas_carteira Set Unidade = '" & Txt_unidade & "' where Unidade = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update vendas_carteira Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_programacao Set Un = '" & Txt_unidade & "' where Un = '" & TBProduto!Unidade & "'"
        Conexao.Execute "Update Vendas_programacao Set Unidade_com = '" & Txt_unidade & "' where Unidade_com = '" & TBProduto!Unidade & "'"
    End If
End If
ProcEnviadadosUnidade
TBProduto.Update
txtidunidade = TBProduto!CODIGO
TBProduto.Close
ProcCarregaListaUnidade
If Novo_geral2 = True Then
    USMsgBox ("Nova unidade cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova unidade"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar unidade"
    If CodigoLista2 <> 0 And Lista_unidade.ListItems.Count <> 0 Then
        Lista_unidade.SelectedItem = Lista_unidade.ListItems(CodigoLista2)
        Lista_unidade.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = txtidunidade.Text
Documento = "Unidade: " & Txt_unidade & " - Descrição: " & Txt_descricao_unidade
Documento1 = ""
ProcGravaEvento
'==================================
Novo_geral2 = False
ProcCarregaComboUnidade Cmb_unidade_de_conversao, False
ProcCarregaComboUnidade Cmb_unidade_para_conversao, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_conversao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(35).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
valor = IIf(Txt_qtde_de_conversao = "", 0, Txt_qtde_de_conversao)
If valor <= 0 Then
    NomeCampo = "a quantidade de conversão"
    ProcVerificaAcao
    Txt_qtde_de_conversao.SetFocus
    Exit Sub
End If
If Cmb_unidade_de_conversao = "" Then
    NomeCampo = "a unidade de converão"
    ProcVerificaAcao
    Cmb_unidade_de_conversao.SetFocus
    Exit Sub
End If
valor = IIf(Txt_qtde_para_conversao = "", 0, Txt_qtde_para_conversao)
If valor <= 0 Then
    NomeCampo = "a quantidade para conversão"
    ProcVerificaAcao
    Txt_qtde_para_conversao.SetFocus
    Exit Sub
End If
If Cmb_unidade_para_conversao = "" Then
    NomeCampo = "a unidade para converão"
    ProcVerificaAcao
    Cmb_unidade_para_conversao.SetFocus
    Exit Sub
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Tabela_conversao_unidade where ID = " & Txt_ID_conversao, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then TBProduto.AddNew
ProcEnviadadosConversao
TBProduto.Update
Txt_ID_conversao = TBProduto!ID
TBProduto.Close
ProcCarregaListaConversao
If Novo_geral3 = True Then
    USMsgBox ("Nova regra para conversão cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova regra para conversão"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar regra para conversão"
    If CodigoLista3 <> 0 And Lista_conversao.ListItems.Count <> 0 Then
        Lista_conversao.SelectedItem = Lista_conversao.ListItems(CodigoLista3)
        Lista_conversao.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID_conversao
Documento = "Qtde. de: " & Format(Txt_qtde_de_conversao, "###,##0.0000") & " - Unidade de: " & Cmb_unidade_de_conversao & " - Qtde. para: " & Txt_qtde_para_conversao & " - Unidade para: " & Cmb_unidade_para_conversao
Documento1 = ""
ProcGravaEvento
'==================================
Novo_geral3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_condicao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(36).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If Txt_texto_cond = "" Then
    NomeCampo = "a condição de pagamento/recebimento"
    ProcVerificaAcao
    Txt_texto_cond.SetFocus
    Exit Sub
End If
If Cmb_aplicacao_cond = "" Then
    NomeCampo = "a aplicação"
    ProcVerificaAcao
    Cmb_aplicacao_cond.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from vendas_proposta_dadoscomerciais_padrao where ID = " & Txt_ID_cond, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then TBProduto.AddNew
ProcEnviadadosCondicao
TBProduto.Update
Txt_ID_cond = TBProduto!ID
TBProduto.Close
ProcCarregaListaCondicoes
If Novo_geral4 = True Then
    USMsgBox ("Nova condição de pagamento/recebimento cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova condição de pagamento/recebimento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar condição de pagamento/recebimento"
    If SSTab3.Tab = 0 Then
        If CodigoLista4 <> 0 And Lista_cond.ListItems.Count <> 0 Then
            Lista_cond.SelectedItem = Lista_cond.ListItems(CodigoLista4)
            Lista_cond.SetFocus
        End If
    Else
        If CodigoLista4 <> 0 And Lista_cond1.ListItems.Count <> 0 Then
            Lista_cond1.SelectedItem = Lista_cond1.ListItems(CodigoLista4)
            Lista_cond1.SetFocus
        End If
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID_cond
Documento = "Condições de pagamento/recebimento: " & Txt_texto_cond
Documento1 = ""
ProcGravaEvento
'==================================
Novo_geral4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_feriado()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(37).Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If Txt_descricao_feriado = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Txt_descricao_feriado.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Feriados where ID = " & Txt_ID_feriado, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then TBProduto.AddNew
ProcEnviadadosFeriado
TBProduto.Update
Txt_ID_feriado = TBProduto!ID
TBProduto.Close
ProcCarregaListaFeriados
If Novo_geral5 = True Then
    USMsgBox ("Novo feriado cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo feriado"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar feriado"
    If CodigoLista5 <> 0 And Lista_feriado.ListItems.Count <> 0 Then
        Lista_feriado.SelectedItem = Lista_feriado.ListItems(CodigoLista5)
        Lista_feriado.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID_feriado
Documento = "Data do feriado: " & Format(Cmb_data_feriado, "dd/mm/yy") & " - Descrição: " & Txt_descricao_feriado
Documento1 = ""
ProcGravaEvento
'==================================
Novo_geral5 = False
ProcCarregaComboAnoFeriado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCofins_Click()
On Error GoTo tratar_erro

PC_PIS = False
PC_Cofins = True
PC_CSLL = False
PC_ISSQN = False
PC_IRRF = False
PC_INSS = False
frmOpcoesGeral_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCSLL_Click()
On Error GoTo tratar_erro

PC_PIS = False
PC_Cofins = False
PC_CSLL = True
PC_ISSQN = False
PC_IRRF = False
PC_INSS = False
frmOpcoesGeral_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdINSS_Click()
On Error GoTo tratar_erro

PC_PIS = False
PC_Cofins = False
PC_CSLL = False
PC_ISSQN = False
PC_IRRF = False
PC_INSS = True
frmOpcoesGeral_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdIRRF_Click()
On Error GoTo tratar_erro

PC_PIS = False
PC_Cofins = False
PC_CSLL = False
PC_ISSQN = False
PC_IRRF = True
PC_INSS = False
frmOpcoesGeral_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdISSQN_Click()
On Error GoTo tratar_erro

PC_PIS = False
PC_Cofins = False
PC_CSLL = False
PC_ISSQN = True
PC_IRRF = False
PC_INSS = False
frmOpcoesGeral_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalantigo_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Localizar local arquivos antigos"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtlocalantigo.Text = sBuffer
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalnovo_Click()
On Error GoTo tratar_erro

szTitle = vbCr & vbCr & "Localizar local novos arquivos"
With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    txtlocalnovo.Text = sBuffer
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoBD()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcVerifInstancia Cmb_servidor
Novo_Local = True
ProcLimpaCamposBanco
ProcHabilitaCampos
Cmd_localizar_rel_Click

Exit Sub
tratar_erro:
    If Err.Number <> 32755 Then USMsgBox "Erro. Banco de dados inválido." & Chr(13) & "Se você utilizar este arquivo como banco de dados o sistema pode não funcionar corretamente." & Chr(13) & "Descrição do erro: " & Err.Description, vbCritical, "CAPRIND v5.0", Err.Number
    Exit Sub
End Sub

Private Sub ProcLimpaCamposBanco()
On Error GoTo tratar_erro

txtLocalrel = ""
'Txt_usuario = ""
'Txt_senha = ""
Cmb_servidor = ""
Cmb_nome_banco = ""
txtlocalantigo = App.Path
txtlocalnovo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaCampos()
On Error GoTo tratar_erro

Cmd_localizar_rel.Enabled = True
cmdLocalantigo.Enabled = True
cmdLocalnovo.Enabled = True
'With Txt_usuario
'    .Locked = False
'    .TabStop = True
'End With
'With Txt_senha
'    .Locked = False
'    .TabStop = True
'End With
With Cmb_servidor
    .Locked = False
    .TabStop = True
End With
With Cmb_nome_banco
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

Cmd_localizar_rel.Enabled = False
cmdLocalantigo.Enabled = False
cmdLocalnovo.Enabled = False
With Cmb_servidor
    .Locked = True
    .TabStop = False
End With
With Cmb_nome_banco
    .Locked = True
    .TabStop = False
End With
'With Txt_usuario
'    .Locked = True
'    .TabStop = False
'End With
'With Txt_senha
'    .Locked = True
'    .TabStop = False
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Moeda()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposMoeda
Novo_geral1 = True
Frame1(33).Enabled = True
txtMoeda.Locked = False
txtMoeda.TabStop = True
txtMoeda.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_empresa2()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Impostos where ID_empresa = " & txtIDEmpresa, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar regime tributário"
Else
    TBGravar.AddNew
    USMsgBox ("Regime tributário cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo regime tributário"
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
End If
ProcEnviadadosImpostos
TBGravar.Update
txtID_imposto = TBGravar!ID

'Atualiza dados na empresa
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from empresa where codigo  = " & txtIDEmpresa, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If optSimples.Value = True Then
'====================================================================================
'    TBGravar!AliquotaSN = IIf(txtAliquotaSN.Text <> "", txtAliquotaSN.Text, "0")
'====================================================================================
    TBGravar!Simples = True
    Else
    TBGravar!Simples = False
    End If
    If optSimples1.Value = True Then TBGravar!Simples1 = True Else TBGravar!Simples1 = False
    If optPresumido.Value = True Then TBGravar!Presumido = True Else TBGravar!Presumido = False
    If optReal.Value = True Then TBGravar!Real = True Else TBGravar!Real = False
    TBGravar.Update
End If
TBGravar.Close
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = txtID_imposto
Documento = "Imposto: " & txtID_imposto
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_empresa3()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If Chk_verificar_desconectar_usuario.Value = 1 Then
    valor = IIf(Txt_minutos_desconectar = "", 0, Txt_minutos_desconectar)
    If valor <= 0 Then
        NomeCampo = "os minutos para verificar e desconectar o usuário automaticamente"
        ProcVerificaAcao
        Txt_minutos_desconectar.SetFocus
        Exit Sub
    End If
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from empresa where codigo = " & txtIDEmpresa.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then TBProduto.AddNew
TBProduto!Caminho_Nfe = Txt_local_armaz_NFe
TBProduto!Caminho_RetornoNfe = txtRetornoNF
TBProduto!Caminho_XMLDanfe = txtCaminhoXMLDanfe
TBProduto!UsuarioPref = txtUsuarioPref
TBProduto!SenhaPref = txtSenhaPref
TBProduto!Apelido_contimatic = Txt_apelido_contimatic
TBProduto!Registro_boleto = Txt_registro_boleto
TBProduto!Certificadodigital = txtCertificadodigital.Text

'Caprind
If Chk_bloquear_prod_cliente.Value = 0 Then TBProduto!Bloquear_produtos = False Else TBProduto!Bloquear_produtos = True
If Chk_bloquear_forn.Value = 0 Then TBProduto!Bloquear_fornecedores = False Else TBProduto!Bloquear_fornecedores = True
If Chk_bloquear_cli_forn_regime.Value = 0 Then TBProduto!Bloquear_cli_forn_regime = False Else TBProduto!Bloquear_cli_forn_regime = True
If Chk_CC_obrigatorio.Value = 0 Then TBProduto!CC_obrigatorio = False Else TBProduto!CC_obrigatorio = True
If Chk_codigo_ref_DANFE.Value = 0 Then TBProduto!Codigo_ref_DANFE = False Else TBProduto!Codigo_ref_DANFE = True
If Chk_codigo_ref_desc_DANFE.Value = 0 Then TBProduto!Codigo_ref_desc_DANFE = False Else TBProduto!Codigo_ref_desc_DANFE = True
If Chk_liberar_qtde_MRP.Value = 0 Then TBProduto!Liberar_qtde_MRP = False Else TBProduto!Liberar_qtde_MRP = True
If Chk_calcular_IPI.Value = 0 Then TBProduto!Calcular_IPI_sem_desc = False Else TBProduto!Calcular_IPI_sem_desc = True
If Chk_bloquear_NF_prod_serv_sem_cad.Value = 0 Then TBProduto!Bloquear_NF_prod_serv_sem_cadastro = False Else TBProduto!Bloquear_NF_prod_serv_sem_cadastro = True
If Chk_ativar_empenho_aut.Value = 0 Then TBProduto!Ativar_empenho_autom = False Else TBProduto!Ativar_empenho_autom = True
If Chk_ativar_empenho_aut_prod.Value = 0 Then TBProduto!Ativar_empenho_autom_prod = False Else TBProduto!Ativar_empenho_autom_prod = True
If Chk_carregar_CFOP_ST.Value = 0 Then TBProduto!Carregar_CFOP_ST = False Else TBProduto!Carregar_CFOP_ST = True
If Chk_agregar_ordem_valor_PC.Value = 0 Then TBProduto!Agregar_ordem_valor_PC = False Else TBProduto!Agregar_ordem_valor_PC = True
If Chk_gerar_RM_ordem_PC.Value = 0 Then TBProduto!Gerar_RM_ordem_PC = False Else TBProduto!Gerar_RM_ordem_PC = True
If Chk_liberar_campos_estrutura.Value = 0 Then TBProduto!Liberar_campos_estrutura = False Else TBProduto!Liberar_campos_estrutura = True

If chk_TPAmb.Value = 1 Then TBProduto!tpAmb = "2" Else TBProduto!tpAmb = "1"

If chkSemEstoque.Value = 0 Then TBProduto!SemEstoque = False Else TBProduto!SemEstoque = True
If chkClienteVendedor.Value = 0 Then TBProduto!ClienteVendedor = False Else TBProduto!ClienteVendedor = True

TBProduto!TPCertificado = txttpEmissor

'===========================================================================================================
' Novo
'===========================================================================================================
If chkSemEstoque.Value = 0 Then TBProduto!SemEstoque = False Else TBProduto!SemEstoque = True
If chkClienteVendedor.Value = 0 Then TBProduto!ClienteVendedor = False Else TBProduto!ClienteVendedor = True
TBProduto!NF_Serie = txtSerie_Nf.Text
'============================================================================================================

If Chk_verificar_desconectar_usuario.Value = 0 Then
    TBProduto!Verificar_desconectar_usuario = False
    TBProduto!Minutos_desconectar = Null
Else
    TBProduto!Verificar_desconectar_usuario = True
    TBProduto!Minutos_desconectar = Txt_minutos_desconectar
End If

If chk_Esconder_ValorOF.Value = 0 Then TBProduto!Esconder_ValorOF = False Else TBProduto!Esconder_ValorOF = True
If Chk_movimentar_estoque_pc.Value = 0 Then TBProduto!Movimentar_estoque_pc = False Else TBProduto!Movimentar_estoque_pc = True
If Chk_ativar_produtos_similares.Value = 0 Then TBProduto!Ativar_prod_similares = False Else TBProduto!Ativar_prod_similares = True
If Chk_validar_proposta_pi_autom.Value = 0 Then TBProduto!Validar_prop_pi_autom = False Else TBProduto!Validar_prop_pi_autom = True
If Chk_codigo_ref_SPED_forn.Value = 0 Then TBProduto!Codigo_ref_SPED_forn = False Else TBProduto!Codigo_ref_SPED_forn = True
If chkLiberar_LoteMinimo.Value = 0 Then TBProduto!Liberar_LoteMinimo = False Else TBProduto!Liberar_LoteMinimo = True
If Chk_carregar_LA_entrada.Value = 0 Then TBProduto!Carregar_LAentrada = False Else TBProduto!Carregar_LAentrada = True
If chkNao_inspecionar.Value = 0 Then TBProduto!Nao_inspecionar = False Else TBProduto!Nao_inspecionar = True
If ChkBloc_CC_Previsao.Value = 0 Then TBProduto!Bloc_CC_Previsao = False Else TBProduto!Bloc_CC_Previsao = True
If chk_Baixa_Auto_Estoque_NF.Value = 0 Then TBProduto!Baixa_Auto_Estoque_NF = False Else TBProduto!Baixa_Auto_Estoque_NF = True
If Chk_bloq_OP_estrutura.Value = 0 Then TBProduto!Bloq_OP_estrutura = False Else TBProduto!Bloq_OP_estrutura = True
If Chk_bloq_OP_processo.Value = 0 Then TBProduto!Bloq_OP_processo = False Else TBProduto!Bloq_OP_processo = True
If Chk_bloq_OP_plano.Value = 0 Then TBProduto!Bloq_OP_plano = False Else TBProduto!Bloq_OP_plano = True
If Chk_bloq_compra_cot_valida.Value = 0 Then TBProduto!Bloq_compra_cot_valida = False Else TBProduto!Bloq_compra_cot_valida = True
If chkCodigo_sequencial.Value = 0 Then TBProduto!Codigo_sequencial = False Else TBProduto!Codigo_sequencial = True
If Chk_salvar_status_aprovado_PC.Value = 0 Then TBProduto!Salvar_status_aprovado_PC = False Else TBProduto!Salvar_status_aprovado_PC = True
If Chk_enviar_email_outlook.Value = 0 Then TBProduto!Enviar_email_outlook = False Else TBProduto!Enviar_email_outlook = True
If chkMargemAnalise.Value = 0 Then TBProduto!MargemAnalise = False Else TBProduto!MargemAnalise = True

'Gerprod
If Chk_ap_codigo.Value = 0 Then TBProduto!Apontamento_codigo = False Else TBProduto!Apontamento_codigo = True
If Chk_bloquear_apontamento_sem_baixa.Value = 0 Then TBProduto!Bloquear_apontamento_sem_baixa = False Else TBProduto!Bloquear_apontamento_sem_baixa = True
If Chk_bloquear_apontamento_sem_baixa_total.Value = 0 Then TBProduto!Bloquear_apontamento_sem_baixa_total = False Else TBProduto!Bloquear_apontamento_sem_baixa_total = True
If Chk_desbloquear_primeiro_apontamento_OS_proc_controlado.Value = 0 Then TBProduto!Desbloquear_prim_apont_OS_proc_controlado = False Else TBProduto!Desbloquear_prim_apont_OS_proc_controlado = True
If chk_Grupo_Gerprod.Value = 0 Then TBProduto!Grupo_Gerprod = False Else TBProduto!Grupo_Gerprod = True
If Chk_bloquear_apontamento_simultaneo.Value = 0 Then TBProduto!Bloquear_apontamento_simultaneo = False Else TBProduto!Bloquear_apontamento_simultaneo = True
If Chk_apontar_NC_descricao.Value = 0 Then TBProduto!Apontar_NC_descricao = False Else TBProduto!Apontar_NC_descricao = True
If Chk_NC_parecer.Value = 0 Then TBProduto!NC_parecer_rejeitado = False Else TBProduto!NC_parecer_rejeitado = True

TBProduto.Update
TBProduto.Close
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Configuração do sistema/Opções gerais"
Evento = "Salvar outros"
ID_documento = txtIDEmpresa.Text
Documento = "Empresa: " & txtEmpresa
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosOutros()
On Error GoTo tratar_erro
  
ProcLimpaCamposEmpresa3
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where codigo = " & txtIDEmpresa, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_local_armaz_NFe = IIf(IsNull(TBAbrir!Caminho_Nfe), "", TBAbrir!Caminho_Nfe)
    txtRetornoNF = IIf(IsNull(TBAbrir!Caminho_RetornoNfe), "", TBAbrir!Caminho_RetornoNfe)
    txtCaminhoXMLDanfe = IIf(IsNull(TBAbrir!Caminho_XMLDanfe), "", TBAbrir!Caminho_XMLDanfe)
    Txt_registro_boleto = IIf(IsNull(TBAbrir!Registro_boleto), "", TBAbrir!Registro_boleto)
    Txt_apelido_contimatic = IIf(IsNull(TBAbrir!Apelido_contimatic), "", TBAbrir!Apelido_contimatic)
    
    txtUsuarioPref = IIf(IsNull(TBAbrir!UsuarioPref), "", TBAbrir!UsuarioPref)
    txtSenhaPref = IIf(IsNull(TBAbrir!SenhaPref), "", TBAbrir!SenhaPref)
    
    'Caprind
    If TBAbrir!Bloquear_produtos = True Then Chk_bloquear_prod_cliente.Value = 1 Else Chk_bloquear_prod_cliente.Value = 0
    If TBAbrir!Bloquear_fornecedores = True Then Chk_bloquear_forn.Value = 1 Else Chk_bloquear_forn.Value = 0
    If TBAbrir!Bloquear_cli_forn_regime = True Then Chk_bloquear_cli_forn_regime.Value = 1 Else Chk_bloquear_cli_forn_regime.Value = 0
    If TBAbrir!CC_obrigatorio = True Then Chk_CC_obrigatorio.Value = 1 Else Chk_CC_obrigatorio.Value = 0
    If TBAbrir!Codigo_ref_DANFE = True Then Chk_codigo_ref_DANFE.Value = 1 Else Chk_codigo_ref_DANFE.Value = 0
    If TBAbrir!Codigo_ref_desc_DANFE = True Then Chk_codigo_ref_desc_DANFE.Value = 1 Else Chk_codigo_ref_desc_DANFE.Value = 0
    If TBAbrir!Liberar_qtde_MRP = True Then Chk_liberar_qtde_MRP.Value = 1 Else Chk_liberar_qtde_MRP.Value = 0
    If TBAbrir!Calcular_IPI_sem_desc = True Then Chk_calcular_IPI.Value = 1 Else Chk_calcular_IPI.Value = 0
    If TBAbrir!Bloquear_NF_prod_serv_sem_cadastro = True Then Chk_bloquear_NF_prod_serv_sem_cad.Value = 1 Else Chk_bloquear_NF_prod_serv_sem_cad.Value = 0
    If TBAbrir!Ativar_empenho_autom = True Then Chk_ativar_empenho_aut.Value = 1 Else Chk_ativar_empenho_aut.Value = 0
    If TBAbrir!Ativar_empenho_autom_prod = True Then Chk_ativar_empenho_aut_prod.Value = 1 Else Chk_ativar_empenho_aut_prod.Value = 0
    If TBAbrir!Carregar_CFOP_ST = True Then Chk_carregar_CFOP_ST.Value = 1 Else Chk_carregar_CFOP_ST.Value = 0
    If TBAbrir!Agregar_ordem_valor_PC = True Then Chk_agregar_ordem_valor_PC.Value = 1 Else Chk_agregar_ordem_valor_PC.Value = 0
    If TBAbrir!Gerar_RM_ordem_PC = True Then Chk_gerar_RM_ordem_PC.Value = 1 Else Chk_gerar_RM_ordem_PC.Value = 0
    If TBAbrir!Liberar_campos_estrutura = True Then Chk_liberar_campos_estrutura.Value = 1 Else Chk_liberar_campos_estrutura.Value = 0
    If TBAbrir!Verificar_desconectar_usuario = True Then Chk_verificar_desconectar_usuario.Value = 1 Else Chk_verificar_desconectar_usuario.Value = 0
    If TBAbrir!Esconder_ValorOF = True Then chk_Esconder_ValorOF.Value = 1 Else chk_Esconder_ValorOF.Value = 0
    Txt_minutos_desconectar = IIf(IsNull(TBAbrir!Minutos_desconectar), "", TBAbrir!Minutos_desconectar)
    If TBAbrir!Movimentar_estoque_pc = True Then Chk_movimentar_estoque_pc.Value = 1 Else Chk_movimentar_estoque_pc.Value = 0
    If TBAbrir!Ativar_prod_similares = True Then Chk_ativar_produtos_similares.Value = 1 Else Chk_ativar_produtos_similares.Value = 0
    If TBAbrir!Validar_prop_pi_autom = True Then Chk_validar_proposta_pi_autom.Value = 1 Else Chk_validar_proposta_pi_autom.Value = 0
    If TBAbrir!Codigo_ref_SPED_forn = True Then Chk_codigo_ref_SPED_forn.Value = 1 Else Chk_codigo_ref_SPED_forn.Value = 0
    If TBAbrir!Liberar_LoteMinimo = True Then chkLiberar_LoteMinimo.Value = 1 Else chkLiberar_LoteMinimo.Value = 0
    If TBAbrir!Carregar_LAentrada = True Then Chk_carregar_LA_entrada.Value = 1 Else Chk_carregar_LA_entrada.Value = 0
    If TBAbrir!Nao_inspecionar = True Then chkNao_inspecionar.Value = 1 Else chkNao_inspecionar.Value = 0
    If TBAbrir!Bloc_CC_Previsao = True Then ChkBloc_CC_Previsao.Value = 1 Else ChkBloc_CC_Previsao.Value = 0
    If TBAbrir!Baixa_Auto_Estoque_NF = True Then chk_Baixa_Auto_Estoque_NF.Value = 1 Else chk_Baixa_Auto_Estoque_NF.Value = 0
    If TBAbrir!Bloq_OP_estrutura = True Then Chk_bloq_OP_estrutura.Value = 1 Else Chk_bloq_OP_estrutura.Value = 0
    If TBAbrir!Bloq_OP_processo = True Then Chk_bloq_OP_processo.Value = 1 Else Chk_bloq_OP_processo.Value = 0
    If TBAbrir!Bloq_OP_plano = True Then Chk_bloq_OP_plano.Value = 1 Else Chk_bloq_OP_plano.Value = 0
    If TBAbrir!Bloq_compra_cot_valida = True Then Chk_bloq_compra_cot_valida.Value = 1 Else Chk_bloq_compra_cot_valida.Value = 0
    If TBAbrir!Codigo_sequencial = True Then chkCodigo_sequencial.Value = 1 Else chkCodigo_sequencial.Value = 0
    If TBAbrir!Salvar_status_aprovado_PC = True Then Chk_salvar_status_aprovado_PC.Value = 1 Else Chk_salvar_status_aprovado_PC.Value = 0
    If TBAbrir!Enviar_email_outlook = True Then Chk_enviar_email_outlook.Value = 1 Else Chk_enviar_email_outlook.Value = 0
    If TBAbrir!MargemAnalise = True Then chkMargemAnalise.Value = 1 Else chkMargemAnalise.Value = 0
    If TBAbrir!tpAmb = "1" Then chk_TPAmb.Value = 0 Else chk_TPAmb.Value = 1
'=======================================
'Novo
'=======================================
    If TBAbrir!SemEstoque = True Then chkSemEstoque.Value = 1 Else chkSemEstoque.Value = 0
    If TBAbrir!ClienteVendedor = True Then chkClienteVendedor.Value = 1 Else chkClienteVendedor.Value = 0
'=======================================

    txttpEmissor = IIf(IsNull(TBAbrir!TPCertificado) = False, TBAbrir!TPCertificado, "")

    'Gerprod
    If TBAbrir!Apontamento_codigo = True Then Chk_ap_codigo.Value = 1 Else Chk_ap_codigo.Value = 0
    If TBAbrir!Bloquear_apontamento_sem_baixa = True Then Chk_bloquear_apontamento_sem_baixa.Value = 1 Else Chk_bloquear_apontamento_sem_baixa.Value = 0
    If TBAbrir!Bloquear_apontamento_sem_baixa_total = True Then Chk_bloquear_apontamento_sem_baixa_total.Value = 1 Else Chk_bloquear_apontamento_sem_baixa_total.Value = 0
    If TBAbrir!Desbloquear_prim_apont_OS_proc_controlado = True Then Chk_desbloquear_primeiro_apontamento_OS_proc_controlado.Value = 1 Else Chk_desbloquear_primeiro_apontamento_OS_proc_controlado.Value = 0
    If TBAbrir!Grupo_Gerprod = True Then chk_Grupo_Gerprod.Value = 1 Else chk_Grupo_Gerprod.Value = 0
    If TBAbrir!Bloquear_apontamento_simultaneo = True Then Chk_bloquear_apontamento_simultaneo.Value = 1 Else Chk_bloquear_apontamento_simultaneo.Value = 0
    If TBAbrir!Apontar_NC_descricao = True Then Chk_apontar_NC_descricao.Value = 1 Else Chk_apontar_NC_descricao.Value = 0
    If TBAbrir!NC_parecer_rejeitado = True Then Chk_NC_parecer.Value = 1 Else Chk_NC_parecer.Value = 0
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o certificado."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
    Else
        USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    End If
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362O" Then
    If USMsgBox("Deseja realmente atualizar o regime tributário dos impostos e o ID da empresa em todas as tabelas?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Empresa order by codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TBAbrir!Simples = True Then Regime = 1
                If TBAbrir!Presumido = True Then Regime = 2
                If TBAbrir!Real = True Then Regime = 3
                If TBAbrir!Simples1 = True Then Regime = 4
                Conexao.Execute "Update impostos Set Regime = " & Regime & " where ID_empresa = " & TBAbrir!CODIGO
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Configuração do sistema/Opções gerais"
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

Private Sub ProcAtualizaTabelaSN()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362O" Then
    If USMsgBox("Deseja realmente atualizar a tabela do simples nacional em todas as propostas, pedidos e notas fiscais que não possuem a tabela do simples nacional?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Tabela from Impostos_TabelaDAS where ID_empresa = " & txtIDEmpresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Conexao.Execute "Update vendas_proposta Set TabelaSN = " & TBAbrir!Tabela & " where ID_empresa = " & txtIDEmpresa & " and (TabelaSN IS NULL or TabelaSN = 0)"
            Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set TabelaSN = " & TBAbrir!Tabela & " where ID_empresa = " & txtIDEmpresa & " and (TabelaSN IS NULL or TabelaSN = 0)"
            USMsgBox ("Atualização da tabela do simples nacional em todas as propostas, pedidos e notas fiscais efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Atualizar tabela do simples nacional"
            ID_documento = 0
            Documento = ""
            Documento1 = ""
            ProcGravaEvento
            '==================================
        Else
            USMsgBox ("É necessário cadastrar a tabela antes de atualizar."), vbExclamation, "CAPRIND v5.0"
        End If
        TBAbrir.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPIS_Click()
On Error GoTo tratar_erro

PC_PIS = True
PC_Cofins = False
PC_CSLL = False
PC_ISSQN = False
PC_IRRF = False
PC_INSS = False
frmOpcoesGeral_PC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Label8_Click()
On Error GoTo tratar_erro

If Frame1(3).Enabled = False Then Exit Sub
ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "Arquivos jpg (*.jpg) | *.jpg| Arquivos bmp (*.bmp) | *.bmp"
If caminho <> "" Then
    picimagem.Picture = LoadPicture(caminho)
    If fotopadrao = Localrel & "\imagens\caprind.bmp" Then Label8.Visible = True Else Label8.Visible = False
Else
    picimagem.Picture = LoadPicture("")
    Label8.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_cond_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_cond
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_cond, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_cond_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_cond.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from vendas_proposta_dadoscomerciais_padrao where ID = " & Lista_cond.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosCondicao
    CodigoLista4 = Lista_cond.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_cond1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_cond1.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from vendas_proposta_dadoscomerciais_padrao where ID = " & Lista_cond1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosCondicao
    CodigoLista4 = Lista_cond1.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_conversao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_conversao
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_conversao, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_conversao_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_conversao.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Tabela_conversao_unidade where ID = " & Lista_conversao.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosConversao
    CodigoLista3 = Lista_conversao.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empresas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_empresas
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                ID_empresa = .ListItems.Item(InitFor)
                ProcVerificaRegistroUtilizadoSemMsg "Compras_Cotacao", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "compras_pedido", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "compras_programa", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Compras_requisicao", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "estoque_controle", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Estoque_controle_recebimento", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Funcionarios", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "impostos", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Producao", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_ContasPagar", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_contas_receber", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Instituicoes", "ID_empresa = " & ID_empresa
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_empresas, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empresas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaMoeda
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Mensagem = "Não é permitido excluir esta empresa, pois a mesma está sendo utilizada no módulo"
            ID_empresa = .ListItems.Item(InitFor)
            ProcVerificaRegistroUtilizado "Compras_Cotacao", "ID_empresa = " & ID_empresa, "Compras/Cotação"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "compras_pedido", "ID_empresa = " & ID_empresa, "Compras/Pedido"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "compras_programa", "ID_empresa = " & ID_empresa, "Compras/Programação de compra"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Compras_requisicao", "ID_empresa = " & ID_empresa, "Outros/Solicitação"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "estoque_controle", "ID_empresa = " & ID_empresa, "Estoque/Movimentação"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Estoque_controle_recebimento", "ID_empresa = " & ID_empresa, "Estoque/Recebimento"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Funcionarios", "ID_empresa = " & ID_empresa, "RH/Funcionários"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "impostos", "ID_empresa = " & ID_empresa, "Configuração do sistema/Opções gerais"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Producao", "ID_empresa = " & ID_empresa, "PCP/Gerenciamento de ordem"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_ContasPagar", "ID_empresa = " & ID_empresa, "Financeiro/Contas a pagar"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_contas_receber", "ID_empresa = " & ID_empresa, "Financeiro/Contas a receber"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal", "ID_empresa = " & ID_empresa, "Faturamento/Nota fiscal"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Instituicoes", "ID_empresa = " & ID_empresa, "Financeiro/Instituições"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empresas_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista_empresas.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where codigo = " & Lista_empresas.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcCarregaDadosEmpresa
    CodigoLista = Lista_empresas.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_feriado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_feriado
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_feriado, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_feriado_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_feriado.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Feriados where ID = " & Lista_feriado.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosFeriado
    CodigoLista5 = Lista_feriado.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_unidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_unidade
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido_lista", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Compras_programacao", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Estoque_Controle", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Manutencao_defeito", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Producaomaterial", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Projconjunto", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "projproduto", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Requisicao_materiais_lista", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "Txt_Unid = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_analise", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_analise_ProdutosProcessos", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_analise_setores", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "vendas_carteira", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Vendas_programacao", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_unidade, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_unidade_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_unidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Mensagem = "Não é permitido excluir esta unidade de medida, pois a mesma está sendo utilizada no módulo"
            ProcVerificaRegistroUtilizado "Compras_pedido_lista", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Compras/Pedido"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Compras_programacao", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Compras/Programação"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Estoque_Controle", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Estoque/Movimentação"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Manutencao_defeito", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Manutenção/Equipamentos"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Producaomaterial", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "'", "PCP/Gerenciamento de ordem"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Projconjunto", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Engenharia/Estrutura"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "projproduto", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Engenharia/Produtos e serviços"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Requisicao_materiais_lista", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Estoque/Requisição de materiais"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Detalhes_Nota", "Txt_Unid = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Faturamento/Nota fiscal"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Vendas_analise", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Outros/Análise crítica"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Vendas_analise_ProdutosProcessos", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Outros/Análise crítica"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Vendas_analise_setores", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Outros/Análise crítica"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "vendas_carteira", "Unidade = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Vendas/Proposta comercial"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Vendas_programacao", "UN = '" & .ListItems(InitFor).ListSubItems(3) & "' or Unidade_com = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Vendas/Programação"
            If Permitido = False Then
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

Private Sub Lista_unidade_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_unidade.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Unidade_Medida where codigo = " & Lista_unidade.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosUnidade
    CodigoLista2 = Lista_unidade.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listaBancos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listaBancos.ListItems.Count = 0 Then Exit Sub
Select Case listaBancos.SelectedItem
    Case 1:
        txtLocalrel = Localrel
        'Txt_usuario = Usuario_banco
        'Txt_senha = Senha_banco
        NomeCampo = NomeServidor
        Cmb_servidor = NomeServidor
        Cmb_nome_banco = Nome_banco
        If LocalAntigoCaprind <> "" And LocalNovoCaprind <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind, Len(LocalAntigoCaprind) - 12)
            txtlocalnovo = Left(LocalNovoCaprind, Len(LocalNovoCaprind) - 12)
        End If
    Case 2:
        txtLocalrel = Localrel1
        'Txt_usuario = Usuario_banco1
        'Txt_senha = Senha_banco1
        NomeCampo = NomeServidor1
        Cmb_servidor = NomeServidor1
        Cmb_nome_banco = Nome_banco1
        If LocalAntigoCaprind1 <> "" And LocalNovoCaprind1 <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind1, Len(LocalAntigoCaprind1) - 12)
            txtlocalnovo = Left(LocalNovoCaprind1, Len(LocalNovoCaprind1) - 12)
        End If
    Case 3:
        txtLocalrel = Localrel2
       ' Txt_usuario = Usuario_banco2
       ' Txt_senha = Senha_banco2
        NomeCampo = NomeServidor2
        Cmb_servidor = NomeServidor2
        Cmb_nome_banco = Nome_banco2
        If LocalAntigoCaprind2 <> "" And LocalNovoCaprind2 <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind2, Len(LocalAntigoCaprind2) - 12)
            txtlocalnovo = Left(LocalNovoCaprind2, Len(LocalNovoCaprind2) - 12)
        End If
    Case 4:
        txtLocalrel = Localrel3
       ' Txt_usuario = Usuario_banco3
       ' Txt_senha = Senha_banco3
        NomeCampo = NomeServidor3
        Cmb_servidor = NomeServidor3
        Cmb_nome_banco = Nome_banco3
        If LocalAntigoCaprind3 <> "" And LocalNovoCaprind3 <> "" Then
            txtlocalantigo = Left(LocalAntigoCaprind3, Len(LocalAntigoCaprind3) - 12)
            txtlocalnovo = Left(LocalNovoCaprind3, Len(LocalNovoCaprind3) - 12)
        End If
End Select
Novo_Local = False
ProcBloqueiaCampos

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("A instância " & NomeCampo & " não está disponível."), vbExclamation, "CAPRIND v5.0"
        txtLocalrel = ""
      '  Txt_usuario = ""
      '  Txt_senha = ""
        Cmb_servidor = ""
        Cmb_nome_banco = ""
        txtlocalantigo = ""
        txtlocalnovo = ""
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaMoeda_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaMoeda
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                ProcVerificaRegistroUtilizadoSemMsg "Compras_comercial", "Moeda = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal", "Moeda = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "vendas_comercial", "Moeda = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaMoeda, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaMoeda_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaMoeda
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Mensagem = "Não é permitido excluir esta moeda, pois a mesma está sendo utilizada no módulo"
            ProcVerificaRegistroUtilizado "Compras_comercial", "Moeda = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Compras/Pedido"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal", "Moeda = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Faturamento/Nota fiscal"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "vendas_comercial", "Moeda = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Vendas/Proposta comercial"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaMoeda_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaMoeda.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from MOEDA where codigo = " & ListaMoeda.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaDadosMoeda
    CodigoLista1 = ListaMoeda.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaBancos()
On Error GoTo tratar_erro

listaBancos.ListItems.Clear
If Localrel <> "" Then
    With listaBancos.ListItems
        .Add , , 1
        .Item(.Count).SubItems(1) = Localrel
        If NomeServidor <> "" Then .Item(.Count).SubItems(2) = NomeServidor
        If Nome_banco <> "" Then .Item(.Count).SubItems(3) = Nome_banco
    End With
End If
If Localrel1 <> "" Then
    With listaBancos.ListItems
        .Add , , 2
        .Item(.Count).SubItems(1) = Localrel1
        If NomeServidor1 <> "" Then .Item(.Count).SubItems(2) = NomeServidor1
        If Nome_banco1 <> "" Then .Item(.Count).SubItems(3) = Nome_banco1
    End With
End If
If Localrel2 <> "" Then
    With listaBancos.ListItems
        .Add , , 3
        .Item(.Count).SubItems(1) = Localrel2
        If NomeServidor2 <> "" Then .Item(.Count).SubItems(2) = NomeServidor2
        If Nome_banco2 <> "" Then .Item(.Count).SubItems(3) = Nome_banco2
    End With
End If
If Localrel3 <> "" Then
    With listaBancos.ListItems
        .Add , , 4
        .Item(.Count).SubItems(1) = Localrel3
        If NomeServidor3 <> "" Then .Item(.Count).SubItems(2) = NomeServidor3
        If Nome_banco3 <> "" Then .Item(.Count).SubItems(3) = Nome_banco3
    End With
End If

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
            Case vbKeyInsert: ProcNovoBD
            Case vbKeyF3: ProcSalvarBD
            Case vbKeyF4: ProcExcluirBD
            Case vbKeyF6: ProcAlterarBD
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case SSTabEmpresa.Tab
            Case 0:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_empresa
                    Case vbKeyF3: ProcSalvar_empresa
                    Case vbKeyF4: ProcExcluir_empresa
                    Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: If SSTab5.Tab = 1 Then ProcNovoTabelaSN
                    Case vbKeyF3: ProcSalvar_empresa2
                    Case vbKeyF4: If SSTab5.Tab = 1 Then ProcExcluirTabelaSN
                    Case vbKeyF7: If SSTab5.Tab = 1 Then ProcSalvarTabelaSN
                    Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 2:
                Select Case KeyCode
                    Case vbKeyF3: ProcSalvar_empresa3
                    Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 3:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_email
                    Case vbKeyF3: ProcSalvar_email
                    Case vbKeyF4: ProcExcluir_email
                    Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 4:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_filtros
                    Case vbKeyF3: ProcSalvar_Filtros
                    Case vbKeyF4: ProcExcluir_Filtros
                    Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
            Case 5:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_armaz
                    Case vbKeyF3: ProcSalvar_Armaz
                    Case vbKeyF4: ProcExcluir_Armaz
                    Case vbKeyF1: ProcAjuda
                    Case vbKeyEscape: ProcSair
                End Select
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_Moeda
            Case vbKeyF3: ProcSalvar_Moeda
            Case vbKeyF4: ProcExcluir_moeda
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        If SSTab2.Tab = 0 Then
            Select Case KeyCode
                Case vbKeyInsert: ProcNovo_unidade
                Case vbKeyF3: ProcSalvar_unidade
                Case vbKeyF4: ProcExcluir_unidade
                Case vbKeyF1: ProcAjuda
                Case vbKeyEscape: ProcSair
            End Select
        Else
            Select Case KeyCode
                Case vbKeyInsert: ProcNovo_conversao
                Case vbKeyF3: ProcSalvar_conversao
                Case vbKeyF4: ProcExcluir_conversao
                Case vbKeyF1: ProcAjuda
                Case vbKeyEscape: ProcSair
            End Select
        End If
    Case 4:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_condicao
            Case vbKeyF3: ProcSalvar_condicao
            Case vbKeyF4: ProcExcluir_condicao
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 5:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_feriado
            Case vbKeyF3: ProcSalvar_feriado
            Case vbKeyF4: ProcExcluir_feriado
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCriarPastasDFE
SSTab1.Tab = 1
SSTabEmpresa.Tab = 0



ProcCarregaToolBar1 Me, 15195, 8, True
ProcCarregaToolBar2 Me, 15195, 6, True
ProcCarregaToolBar3 Me, 15195, 6, True
ProcCarregaToolBar4 Me, 15195, 8, True
ProcCarregaToolBar5 Me, 15195, 5, True
ProcCarregaToolBar6 Me, 15195, 6, True
ProcCarregaToolBar7 Me, 15195, 6, True
ProcCarregaToolBar8 Me, 15105, 5, True

Contador = 6
ProcVerifLiberacaoTab SSTab1, 0, "Configuração do sistema/Opções gerais/Cadastro de empresa"
ProcVerifLiberacaoTab SSTab1, 1, "Configuração do sistema/Opções gerais/Cadastro de moedas"
ProcVerifLiberacaoTab SSTab1, 2, "Configuração do sistema/Opções gerais/Cadastro de unidades"
ProcVerifLiberacaoTab SSTab1, 3, "Configuração do sistema/Opções gerais/Cadastro de condição de pagamento/recebimento"
ProcVerifLiberacaoTab SSTab1, 4, "Configuração do sistema/Opções gerais/Cadastro de feriados"
ProcVerifLiberacaoTab SSTab1, 5, "Configuração do sistema/Opções gerais/Configuração do sistema"

Formulario = "Configuração do sistema/Opções gerais"
Direitos
ProcLimpaVariaveisPrincipais

If SSTab1.TabVisible(4) = True Then
    SSTab1.Tab = 4
    ProcCarregaListaFeriados
End If
If SSTab1.TabVisible(5) = True Then
    SSTab1.Tab = 5
    SSTab3.Tab = 0
    ProcCarregaListaCondicoes
End If
If SSTab1.TabVisible(2) = True Then
    SSTab1.Tab = 2
    SSTab2.Tab = 0
    ProcCarregaListaUnidade
    ProcCarregaListaConversao
End If
If SSTab1.TabVisible(1) = True Then
    SSTab1.Tab = 1
    ProcCarregaListaMoeda
End If
If SSTab1.TabVisible(0) = True Then
    SSTab1.Tab = 0
    SSTabEmpresa.Tab = 0
    SSTab4.Tab = 0
    SSTab5.Tab = 0
    ProcCarregaListaEmpresa
End If
If SSTab1.TabVisible(5) = True Then
    SSTab1.Tab = 5
    ProcCarregaListaBancos
    
    txtLocalrel = Localrel
   ' Txt_usuario = Usuario_banco
   ' Txt_senha = Senha_banco
    Cmb_servidor = NomeServidor
    Cmb_nome_banco = Nome_banco
    If LocalAntigoCaprind <> "" And LocalNovoCaprind <> "" Then
        txtlocalantigo.Text = Left(LocalAntigoCaprind, Len(LocalAntigoCaprind) - 12)
        txtlocalnovo.Text = Left(LocalNovoCaprind, Len(LocalNovoCaprind) - 12)
    End If
End If

1:
    ProcCarregaComboPais Txt_pais
    ProcCarregaComboUnidade Cmb_unidade_de_conversao, False
    ProcCarregaComboUnidade Cmb_unidade_para_conversao, False
    ProcCarregaComboAnoFeriado
    Cmb_data_feriado.Value = Date
        
    ProcRemoveObjetosResize Me
SSTab1.Tab = 0

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where codigo = '1'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcCarregaDadosEmpresa
End If
TBAbrir.Close


Exit Sub
tratar_erro:
    If Err.Number = 383 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboAnoFeriado()
On Error GoTo tratar_erro

With Cmb_ano_feriado
    Tipo = .Text
    .Clear
    Permitido = False
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Year(Data_feriado) as Ano from Feriados Group by Year(Data_feriado)", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Do While TBFI.EOF = False
            .AddItem TBFI!Ano
            If TBFI!Ano = Tipo Then Permitido = True
            TBFI.MoveNext
        Loop
        TBFI.MoveFirst
        .Text = IIf(Tipo <> "" And Permitido = True, Tipo, TBFI!Ano)
    End If
    TBFI.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Configuração do sistema/Opções gerais"
Direitos
ProcLimpaVariaveisPrincipais

txtLocalrel = Localrel
Cmb_servidor = NomeServidor
Cmb_nome_banco = Nome_banco
'Txt_usuario = Usuario_banco
'Txt_senha = Senha_banco

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposMoeda()
On Error GoTo tratar_erro

txtidmoeda.Text = 0
Txt_data_moeda.Text = Format(Date, "dd/mm/yy")
Txt_responsavel_moeda.Text = pubUsuario
txtMoeda.Text = ""
txtSimbolo.Text = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposUnidade()
On Error GoTo tratar_erro

txtidunidade.Text = 0
Txt_data_unidade.Text = Format(Date, "dd/mm/yy")
Txt_responsavel_unidade.Text = pubUsuario
Txt_unidade.Text = ""
Txt_descricao_unidade = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposConversao()
On Error GoTo tratar_erro

Txt_ID_conversao = 0
Txt_data_conversao = Format(Date, "dd/mm/yy")
Txt_responsavel_conversao = pubUsuario
Txt_qtde_de_conversao = "1,000"
Txt_qtde_para_conversao = ""
Cmb_unidade_de_conversao.ListIndex = -1
Cmb_unidade_para_conversao.ListIndex = -1
CodigoLista3 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposCondicoes()
On Error GoTo tratar_erro

Txt_ID_cond = 0
Txt_data_cond = Format(Date, "dd/mm/yy")
Txt_responsavel_cond = pubUsuario
Txt_texto_cond = ""
Cmb_aplicacao_cond.ListIndex = -1
CodigoLista4 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposFeriados()
On Error GoTo tratar_erro

Txt_ID_feriado = 0
Txt_data_feriado = Format(Date, "dd/mm/yy")
Txt_responsavel_feriado = pubUsuario
Cmb_data_feriado.Value = Date
Txt_descricao_feriado = ""
CodigoLista5 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosMoeda()
On Error GoTo tratar_erro
  
TBProduto!Data = IIf(Txt_data_moeda = "", Date, Txt_data_moeda)
TBProduto!Responsavel = IIf(Txt_responsavel_moeda = "", pubUsuario, Txt_responsavel_moeda)
TBProduto!Moeda = txtMoeda.Text
TBProduto!Simbolo = txtSimbolo.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosUnidade()
On Error GoTo tratar_erro

TBProduto!Data = IIf(Txt_data_unidade = "", Date, Txt_data_unidade)
TBProduto!Responsavel = IIf(Txt_responsavel_unidade = "", pubUsuario, Txt_responsavel_unidade)
TBProduto!Unidade = Txt_unidade.Text
TBProduto!Descricao = Txt_descricao_unidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosConversao()
On Error GoTo tratar_erro
  
TBProduto!Data = IIf(Txt_data_conversao = "", Date, Txt_data_conversao)
TBProduto!Responsavel = IIf(Txt_responsavel_conversao = "", pubUsuario, Txt_responsavel_conversao)
TBProduto!Qtde_de = Txt_qtde_de_conversao
TBProduto!Unidade_de = Cmb_unidade_de_conversao
TBProduto!Qtde_para = Txt_qtde_para_conversao
TBProduto!Unidade_para = Cmb_unidade_para_conversao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosCondicao()
On Error GoTo tratar_erro

TBProduto!Data = IIf(Txt_data_cond = "", Date, Txt_data_cond)
TBProduto!Responsavel = IIf(Txt_responsavel_cond = "", pubUsuario, Txt_responsavel_cond)
TBProduto!Aplic = 1
TBProduto!Tipo = IIf(Cmb_aplicacao_cond = "Vendas", "V", "C")
TBProduto!Texto = Txt_texto_cond

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosFeriado()
On Error GoTo tratar_erro

TBProduto!Data = IIf(Txt_data_feriado = "", Date, Txt_data_feriado)
TBProduto!Responsavel = IIf(Txt_responsavel_feriado = "", pubUsuario, Txt_responsavel_feriado)
TBProduto!Data_feriado = Cmb_data_feriado
TBProduto!Descricao = Txt_descricao_feriado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEmpresa()
On Error GoTo tratar_erro

Lista_empresas.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Empresa order by Razao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_empresas.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = TBLISTA!Razao
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

Private Sub ProcCarregaLista_TBSN(TipoTBSN As String)
On Error GoTo tratar_erro

Tabela = 6
Select Case Mid(TipoTBSN, 8, 3)
    Case "I -": Tabela = 1
    Case "II ": Tabela = 2
    Case "III": Tabela = 3
    Case "IV ": Tabela = 4
    Case "V -": Tabela = 5
End Select
Lista_TBSN.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Impostos_TabelaDAS where ID_empresa = " & frmOpcoesGeral.txtIDEmpresa & " and Tabela = " & Tabela, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_TBSN.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!De), "", Format(TBLISTA!De, "###,##0.00"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ate), "", Format(TBLISTA!Ate, "###,##0.00"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!DAS), "", Format(TBLISTA!DAS, "###,##0.00"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!IRPJ), "", Format(TBLISTA!IRPJ, "###,##0.00"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!CSLL), "", Format(TBLISTA!CSLL, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Cofins), "", Format(TBLISTA!Cofins, "###,##0.00"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!PIS), "", Format(TBLISTA!PIS, "###,##0.00"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!cpp), "", Format(TBLISTA!cpp, "###,##0.00"))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!IPI), "", Format(TBLISTA!IPI, "###,##0.00"))
            .Item(.Count).SubItems(10) = IIf(Lbl_ICMS_TBSN.Visible = True, IIf(IsNull(TBLISTA!ICMS), "", Format(TBLISTA!ICMS, "###,##0.00")), IIf(IsNull(TBLISTA!ISS), "", Format(TBLISTA!ISS, "###,##0.00")))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Valor_deduzir), "", Format(TBLISTA!Valor_deduzir, "###,##0.00"))
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

Private Sub ProcCarregaListaEmail()
On Error GoTo tratar_erro

Lista_email.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Empresa_email where ID_empresa = " & txtIDEmpresa & " order by Aplicacao, Email", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_email.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            Select Case TBLISTA!Aplicacao
                Case "C": Aplicacao = "Compras"
                Case "CU": Aplicacao = "Custos"
                Case "F": Aplicacao = "Financeiro"
                Case "V": Aplicacao = "Vendas"
                Case "FA": Aplicacao = "Faturamento"
            End Select
            .Item(.Count).SubItems(3) = Aplicacao
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Usuario_caprind), "", TBLISTA!Usuario_caprind)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Email), "", TBLISTA!Email)
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

Private Sub ProcCarregaListaFiltros()
On Error GoTo tratar_erro

Lista_filtros.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Empresa_Filtros where ID_empresa = " & txtIDEmpresa & " order by Aplicacao, Filtrarpor", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_filtros.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            Select Case TBLISTA!Aplicacao
                Case "C": Aplicacao = "Compras"
                Case "V": Aplicacao = "Vendas"
                Case "Q": Aplicacao = "Qualidade"
                Case "E": Aplicacao = "Engenharia"
                Case "P": Aplicacao = "PCP"
                Case "F": Aplicacao = "Faturamento"
                Case "T": Aplicacao = "Estoque"
                Case "M": Aplicacao = "Manutenção"
                Case "O": Aplicacao = "Outros"
            End Select
            .Item(.Count).SubItems(3) = Aplicacao
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!filtrarpor), "", TBLISTA!filtrarpor)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Frase), "", TBLISTA!Frase)
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

Private Sub ProcCarregaListaArmaz()
On Error GoTo tratar_erro

Lista_armaz.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Empresa_armazenamento_PDF where ID_empresa = " & txtIDEmpresa & " order by Relatorio", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_armaz.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Relatorio), "", TBLISTA!Relatorio)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!caminho), "", TBLISTA!caminho)
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

Private Sub ProcCarregaListaMoeda()
On Error GoTo tratar_erro

ListaMoeda.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from moeda", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaMoeda.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Moeda), "", TBLISTA!Moeda)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Simbolo), "", TBLISTA!Simbolo)
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

Private Sub ProcCarregaListaUnidade()
On Error GoTo tratar_erro

Lista_unidade.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from unidade_medida order by unidade", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_unidade.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = TBLISTA!Unidade
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
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

Private Sub ProcCarregaListaConversao()
On Error GoTo tratar_erro

Lista_conversao.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Tabela_conversao_unidade", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_conversao.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Qtde_de), "", Format(TBLISTA!Qtde_de, "###,##0.0000"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Unidade_de), "", TBLISTA!Unidade_de)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Qtde_para), "", Format(TBLISTA!Qtde_para, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Unidade_para), "", TBLISTA!Unidade_para)
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

Private Sub ProcCarregaListaCondicoes()
On Error GoTo tratar_erro

Select Case SSTab3.Tab
    Case 0:
        Lista_cond.ListItems.Clear
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from vendas_proposta_dadoscomerciais_padrao where Aplic = 1 and Tipo = 'C' order by Texto", Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBLISTA.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBLISTA.EOF = False
                With Lista_cond.ListItems
                    .Add , , TBLISTA!ID
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                End With
                TBLISTA.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
    Case 1:
        Lista_cond1.ListItems.Clear
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from vendas_proposta_dadoscomerciais_padrao where Aplic = 1 and Tipo = 'V' order by Texto", Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBLISTA.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBLISTA.EOF = False
                With Lista_cond1.ListItems
                    .Add , , TBLISTA!ID
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                End With
                TBLISTA.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBLISTA.Close
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaFeriados()
On Error GoTo tratar_erro

Lista_feriado.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Feriados where Year(Data_feriado) = '" & Cmb_ano_feriado & "' order by Data_feriado", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_feriado.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Data_feriado), "", Format(TBLISTA!Data_feriado, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
1:
    TBLISTA.Close

Exit Sub
tratar_erro:
    If Err.Number = 365 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosEmpresa()
On Error GoTo tratar_erro
  
TBProduto!Razao = txtRazao.Text
TBProduto!Empresa = txtEmpresa.Text
If txtcnpj <> "__.___.___/____-__" Then TBProduto!CNPJ = txtcnpj.Text Else TBProduto!CNPJ = ""
TBProduto!IE = txtRG_IE.Text
TBProduto!IM = Txt_IM
TBProduto!Tipo_endereco = cmbTipo_endereco
TBProduto!Endereco = txtendereco.Text
TBProduto!Numero = txtNumero
TBProduto!complemento = IIf(txtComplemento.Text = "", Null, txtComplemento.Text)
TBProduto!Tipo_bairro = cmbTipo_bairro
TBProduto!Bairro = txt_Bairro
TBProduto!Cidade = Cmb_cidade
TBProduto!UF = Cmb_uf
If Txt_CEP <> "_____-___" Then TBProduto!CEP = Txt_CEP Else TBProduto!CEP = ""
TBProduto!Pais = Txt_pais
TBProduto!Codigo_pais = Txt_pais.ItemData(Txt_pais.ListIndex)
TBProduto!telefone = Txt_telefones
TBProduto!Fax = Txt_fax
TBProduto!Email = IIf(Txt_email.Text = "", Null, LCase(Txt_email))
TBProduto!Site = IIf(Txt_site = "", Null, LCase(Txt_site))
TBProduto!endereco_Cobranca = txtEndereco_cob.Text
TBProduto!Ramo = txtRamo
TBProduto!CNAE = txtCNAE
TBProduto!Codigo_SUFRAMA = IIf(Txt_cod_SUFRAMA = "", Null, Txt_cod_SUFRAMA)
TBProduto!endereco_entrega = Txt_endereco_entrega
If Chk_atualizacao_autom.Value = 1 Then TBProduto!Atualizacao_automatica = True Else TBProduto!Atualizacao_automatica = False
If chkPrincipal.Value = 1 Then TBProduto!Principal = True Else TBProduto!Principal = False
If chkFiscal.Value = 1 Then TBProduto!Fiscal = True Else TBProduto!Fiscal = False
If chkCultural.Value = 1 Then TBProduto!Cultural = True Else TBProduto!Cultural = False
TBProduto!Logotipo = CommonDialog1.filename
TBProduto!NF_Serie = txtSerie_Nf.Text

If TemInternet = True And ErroDriverMYSQL = False Then
    If Chk_atualizacao_autom.Value = 1 Then TextoFiltro = 1 Else TextoFiltro = 0
    FunAbreBDSite
    If ConexaoMySql.State = 1 Then ConexaoMySql.Execute "Update Clientes Set Atualizacao_automatica = '" & TextoFiltro & "' where CNPJ = '" & TBProduto!CNPJ & "'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosFiltros()
On Error GoTo tratar_erro

TBProduto!ID_empresa = txtIDEmpresa
Select Case cmbAplicacao_Filtros
    Case "Compras": TBProduto!Aplicacao = "C"
    Case "Vendas": TBProduto!Aplicacao = "V"
    Case "Qualidade": TBProduto!Aplicacao = "Q"
    Case "Engenharia": TBProduto!Aplicacao = "E"
    Case "PCP": TBProduto!Aplicacao = "P"
    Case "Estoque": TBProduto!Aplicacao = "T"
    Case "Faturamento": TBProduto!Aplicacao = "F"
    Case "Manutenção": TBProduto!Aplicacao = "M"
    Case "Outros": TBProduto!Aplicacao = "O"
End Select
TBProduto!Tipo = cmbTipo_Filtros
TBProduto!filtrarpor = cmbfiltrarpor_Filtros
TBProduto!Frase = cmbFrase_Filtros

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosArmaz()
On Error GoTo tratar_erro

TBProduto!ID_empresa = txtIDEmpresa
TBProduto!Relatorio = Cmb_relatorio_armaz
TBProduto!caminho = Txt_local_armaz

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosEmail()
On Error GoTo tratar_erro

TBProduto!Data = IIf(Txt_data_email = "", Date, Txt_data_email)
TBProduto!Responsavel = IIf(Txt_responsavel_email = "", pubUsuario, Txt_responsavel_email)
TBProduto!ID_empresa = txtIDEmpresa
Select Case Cmb_aplicacao_email
    Case "Compras": TBProduto!Aplicacao = "C"
    Case "Custos": TBProduto!Aplicacao = "CU"
    Case "Financeiro": TBProduto!Aplicacao = "F"
    Case "Vendas": TBProduto!Aplicacao = "V"
    Case "Faturamento": TBProduto!Aplicacao = "FA"
End Select
TBProduto!Usuario_caprind = Cmb_usuario_caprind_email
TBProduto!Servidor_SMTP = Txt_servidor_SMTP_email
TBProduto!Porta = txt_porta_email
Select Case Cmb_seguranca_email
    Case "Não segura": TBProduto!Seguranca = "N"
    Case "SSL/TSL": TBProduto!Seguranca = "S"
End Select
TBProduto!Nome = Txt_nome_email
TBProduto!Email = LCase(Txt_email_email)
TBProduto!Usuario = Txt_usuario_email
TBProduto!Senha = Txt_senha_email

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposEmpresa()
On Error GoTo tratar_erro
  
fotopadrao = Localrel & "\imagens\caprind.bmp"
picimagem.Picture = LoadPicture(fotopadrao)
CommonDialog1.filename = fotopadrao
'Label8.Visible = True
3:
    txtIDEmpresa.Text = 0
'    txtAliquotaSN.Text = "0,00"
    txtRazao.Text = ""
    txtcnpj.Text = "__.___.___/____-__"
    txtEmpresa.Text = ""
    txtRG_IE.Text = ""
    cmbTipo_endereco.ListIndex = -1
    txtendereco.Text = ""
    txtNumero = ""
    txtComplemento = ""
    Txt_IM = ""
    cmbTipo_bairro.ListIndex = -1
    txt_Bairro = ""
    Cmb_cidade.ListIndex = -1
    Cmb_uf.ListIndex = -1
    Txt_CEP = "_____-___"
    Txt_pais.ListIndex = -1
    Txt_telefones = ""
    Txt_fax = ""
    Txt_email.Text = ""
    Txt_site = ""
    txtEndereco_cob.Text = ""
    txtRamo = ""
    txtCNAE = ""
    Txt_cod_SUFRAMA = ""
    Txt_endereco_entrega = ""
    Chk_atualizacao_autom.Value = 1
    chkPrincipal.Value = 0
    chkCultural.Value = 0
    chkFiscal.Value = 0
    CodigoLista = 0
    txtSerie_Nf.Text = ""
    
Exit Sub
tratar_erro:
    If Err.Number = "71" Or Err.Number = "75" Or Err.Number = "76" Then
        GoTo 3
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposTBSN()
On Error GoTo tratar_erro

Txt_ID_TBSN = 0
Txt_de_TBSN = ""
Txt_ate_TBSN = ""
Txt_Aliquota_TBSN = ""
Txt_IRPJ_TBSN = ""
Txt_CSLL_TBSN = ""
Txt_Cofins_TBSN = ""
Txt_PIS_TBSN = ""
Txt_CPP_TBSN = ""
Txt_IPI_TBSN = ""
Txt_ICMS_TBSN = ""
Txt_valor_deduzir_TBSN = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposEmail()
On Error GoTo tratar_erro
  
Txt_ID_email = 0
Txt_data_email = Format(Date, "dd/mm/yy")
Txt_responsavel_email = pubUsuario
Cmb_aplicacao_email.ListIndex = -1
ProcCarregaComboUsuario Cmb_usuario_caprind_email, "A.IDAcesso IS NOT NULL", True
Txt_servidor_SMTP_email = ""
txt_porta_email = ""
Cmb_seguranca_email.ListIndex = -1
Txt_nome_email = ""
Txt_email_email.Text = ""
Txt_usuario_email = ""
Txt_senha_email = ""
CodigoLista6 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposFiltros()
On Error GoTo tratar_erro
  
txtID_Filtros = 0
txtData_Filtros = Format(Date, "dd/mm/yy")
txtResponsavel_Filtros = pubUsuario
cmbAplicacao_Filtros.ListIndex = -1
cmbTipo_Filtros.ListIndex = -1
cmbfiltrarpor_Filtros.ListIndex = -1
cmbFrase_Filtros.ListIndex = -1
CodigoLista7 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposArmaz()
On Error GoTo tratar_erro
  
Txt_ID_armaz = 0
Txt_data_armaz = Format(Date, "dd/mm/yy")
Txt_responsavel_armaz = pubUsuario
Cmb_relatorio_armaz.ListIndex = -1
Txt_local_armaz = ""
CodigoLista8 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPresumido_Click()
On Error GoTo tratar_erro

If optPresumido.Value = True Then
    ProcVerifMostrarEsconderTab 1
    Regime = 2
    ProcCarregaDadosImpostos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optReal_Click()
On Error GoTo tratar_erro

If optReal.Value = True Then
    ProcVerifMostrarEsconderTab 1
    Regime = 3
    ProcCarregaDadosImpostos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSimples_Click()
On Error GoTo tratar_erro

If optSimples.Value = True Then
    ProcVerifMostrarEsconderTab 0
    Regime = 1
    ProcCarregaDadosImpostos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSimples1_Click()
On Error GoTo tratar_erro

If optSimples1.Value = True Then
    ProcVerifMostrarEsconderTab 1
    Regime = 4
    ProcCarregaDadosImpostos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub picimagem_Click()
On Error GoTo tratar_erro

If Frame1(3).Enabled = False Then Exit Sub
ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
If caminho <> "" Then
    picimagem.Picture = LoadPicture(caminho)
    'If fotopadrao = Localrel & "\imagens\caprind.bmp" Then Label8.Visible = True Else Label8.Visible = False
Else
    picimagem.Picture = LoadPicture("")
   ' Label8.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTabEmpresa.Tab <> 2 Then PBLista.Visible = True
Select Case SSTab1.Tab
    Case 0:
        If listaBancos.Visible = True Then listaBancos.SetFocus
        SSTabEmpresa.Tab = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Empresa where codigo = '1'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcCarregaDadosEmpresa
        End If
        TBAbrir.Close

    Case 1:
        USToolBar4.Visible = True
        USToolBar5.Visible = False
        If Lista_empresas.Visible = True Then Lista_empresas.SetFocus
        ProcCarregaListaEmpresa
    Case 2:
        If ListaMoeda.Visible = True Then ListaMoeda.SetFocus
        ProcCarregaListaMoeda
    Case 3:
        If Lista_unidade.Visible = True Then Lista_unidade.SetFocus
        ProcCarregaListaUnidade
    Case 4:
        If Lista_cond.Visible = True Then Lista_cond.SetFocus
    Case 5:
        If Lista_feriado.Visible = True Then Lista_feriado.SetFocus
        ProcCarregaListaFeriados
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

ProcCarregaListaCondicoes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTabEmpresa_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

With USToolBar4
    .Visible = True
    USToolBar5.Visible = False
    .ButtonState(4) = 5
    PBLista.Visible = True
    Select Case SSTabEmpresa.Tab
        Case 0: 'Empresas
            .ButtonState(4) = 0
            If Lista_empresas.Visible = True Then Lista_empresas.SetFocus
        Case 1: 'Regime tributário
            ProcVerifProsseguir
            If Permitido = False Then Exit Sub
            If optSimples.Value = True Then PBLista.Visible = True Else PBLista.Visible = False
            chkDuplicata.SetFocus
            USToolBar4.Visible = False
            USToolBar5.Visible = True
            ProcCarregaDadosImpostos
        Case 2: 'Dados adicionais
            ProcVerifProsseguir
            If Permitido = False Then Exit Sub
            Cmd_localizar_NFe.SetFocus
            PBLista.Visible = False
            USToolBar4.Visible = False
            USToolBar5.Visible = True
            ProcCarregaDadosOutros
        Case 3: 'Email
            ProcVerifProsseguir
            If Permitido = False Then Exit Sub
            Lista_email.SetFocus
            ProcCarregaListaEmail
        Case 4: 'Filtros
            ProcVerifProsseguir
            If Permitido = False Then Exit Sub
            Lista_filtros.SetFocus
            ProcCarregaListaFiltros
        Case 5: 'Armazenamento documentos
            ProcVerifProsseguir
            If Permitido = False Then Exit Sub
            .ButtonState(4) = 5
            Lista_armaz.SetFocus
            ProcCarregaListaArmaz
    End Select
    .Refresh
End With
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifProsseguir()
On Error GoTo tratar_erro

Permitido = True
If txtIDEmpresa = 0 Then
    SSTabEmpresa.Tab = 0
    Permitido = False
    USMsgBox ("Informe a empresa antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosEmpresa()
On Error GoTo tratar_erro

'================================================================================
 '   txtAliquotaSN.Text = IIf(IsNull(TBAbrir!AliquotaSN), "", Format(TBAbrir!AliquotaSN, "###,##0.00"))
'================================================================================

txtIDEmpresa.Text = TBAbrir!CODIGO
txtRazao.Text = IIf(IsNull(TBAbrir!Razao), "", TBAbrir!Razao)
If IsNull(TBAbrir!CNPJ) = False And TBAbrir!CNPJ <> "" Then txtcnpj = TBAbrir!CNPJ
txtEmpresa.Text = TBAbrir!Empresa
txtRG_IE.Text = IIf(IsNull(TBAbrir!IE), "", TBAbrir!IE)
Txt_IM = IIf(IsNull(TBAbrir!IM), "", TBAbrir!IM)
txtendereco.Text = IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco)
txtNumero = IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero)
txtComplemento = IIf(IsNull(TBAbrir!complemento), "", TBAbrir!complemento)
txt_Bairro = IIf(IsNull(TBAbrir!Bairro), "", TBAbrir!Bairro)
If IsNull(TBAbrir!CEP) = False And TBAbrir!CEP <> "" Then Txt_CEP = TBAbrir!CEP
Txt_telefones = IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone)
Txt_fax = IIf(IsNull(TBAbrir!Fax), "", TBAbrir!Fax)
Txt_email.Text = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
Txt_site = IIf(IsNull(TBAbrir!Site), "", TBAbrir!Site)
txtEndereco_cob.Text = IIf(IsNull(TBAbrir!endereco_Cobranca), "", TBAbrir!endereco_Cobranca)
txtRamo = IIf(IsNull(TBAbrir!Ramo), "", TBAbrir!Ramo)
txtCNAE = IIf(IsNull(TBAbrir!CNAE), "", TBAbrir!CNAE)
Txt_cod_SUFRAMA = IIf(IsNull(TBAbrir!Codigo_SUFRAMA), "", TBAbrir!Codigo_SUFRAMA)
Txt_endereco_entrega = IIf(IsNull(TBAbrir!endereco_entrega), "", TBAbrir!endereco_entrega)
If TBAbrir!Atualizacao_automatica = True Then Chk_atualizacao_autom.Value = 1 Else Chk_atualizacao_autom.Value = 0
If TBAbrir!Principal = True Then chkPrincipal.Value = 1 Else chkPrincipal.Value = 0
If TBAbrir!Cultural = True Then chkCultural.Value = 1 Else chkCultural.Value = 0
If TBAbrir!Fiscal = True Then chkFiscal.Value = 1 Else chkFiscal.Value = 0

If TBAbrir!Certificadodigital <> "" Then
txtCertificadodigital.Text = TBAbrir!Certificadodigital
Else
txtCertificadodigital.Text = ""
End If

txtSerie_Nf.Text = IIf(IsNull(TBAbrir!NF_Serie), 0, TBAbrir!NF_Serie)


NomeCampo = "o tipo do endereço"
If IsNull(TBAbrir!Tipo_endereco) = False And TBAbrir!Tipo_endereco <> "" Then cmbTipo_endereco = TBAbrir!Tipo_endereco
NomeCampo = "o tipo do bairro"
If IsNull(TBAbrir!Tipo_bairro) = False And TBAbrir!Tipo_bairro <> "" Then cmbTipo_bairro = TBAbrir!Tipo_bairro
NomeCampo = "o estado"
If IsNull(TBAbrir!UF) = False And TBAbrir!UF <> "" Then Cmb_uf = TBAbrir!UF
NomeCampo = "a cidade"
If IsNull(TBAbrir!Cidade) = False And TBAbrir!Cidade <> "" Then Cmb_cidade = TBAbrir!Cidade
NomeCampo = "o país"
If IsNull(TBAbrir!Pais) = False And TBAbrir!Pais <> "" Then Txt_pais = TBAbrir!Pais

1:
    If TBAbrir!Simples = True Then optSimples.Value = True Else optSimples.Value = False
    If TBAbrir!Simples1 = True Then optSimples1.Value = True Else optSimples1.Value = False
    
    If TBAbrir!Simples = True Then Cmb_tipo_TBSN = "Tabela II - Partilha do Simples Nacional - Indústria"
    
    If TBAbrir!Presumido = True Then optPresumido.Value = True Else optPresumido.Value = False
    'If TBAbrir!Real = True Then optReal.Value = True Else optReal.Value = False
    If TBAbrir!Real = True Then optReal.Value = True Else optReal.Value = False
    Frame1(3).Enabled = True
    Novo_geral1 = False
    
    If IsNull(TBAbrir!Logotipo) = False And TBAbrir!Logotipo <> "" Then
        'Label8.Visible = False
        picimagem.Picture = LoadPicture(TBAbrir!Logotipo)
        CommonDialog1.filename = TBAbrir!Logotipo
    Else
        'Label8.Visible = True
        picimagem.Picture = LoadPicture(fotopadrao)
        CommonDialog1.filename = fotopadrao
    End If
2:
    
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " dessa empresa."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    If Err.Number = "13" Or Err.Number = "53" Or Err.Number = "71" Or Err.Number = "75" Or Err.Number = "76" Then GoTo 2
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosTBSN()
On Error GoTo tratar_erro

Txt_ID_TBSN = TBLISTA!ID
Txt_de_TBSN = IIf(IsNull(TBLISTA!De), "", Format(TBLISTA!De, "###,##0.00"))
Txt_ate_TBSN = IIf(IsNull(TBLISTA!Ate), "", Format(TBLISTA!Ate, "###,##0.00"))
Txt_Aliquota_TBSN = IIf(IsNull(TBLISTA!DAS), "", Format(TBLISTA!DAS, "###,##0.00"))
Txt_IRPJ_TBSN = IIf(IsNull(TBLISTA!IRPJ), "", Format(TBLISTA!IRPJ, "###,##0.00"))
Txt_CSLL_TBSN = IIf(IsNull(TBLISTA!CSLL), "", Format(TBLISTA!CSLL, "###,##0.00"))
Txt_Cofins_TBSN = IIf(IsNull(TBLISTA!Cofins), "", Format(TBLISTA!Cofins, "###,##0.00"))
Txt_PIS_TBSN = IIf(IsNull(TBLISTA!PIS), "", Format(TBLISTA!PIS, "###,##0.00"))
Txt_CPP_TBSN = IIf(IsNull(TBLISTA!cpp), "", Format(TBLISTA!cpp, "###,##0.00"))
If Txt_IPI_TBSN.Visible = True Then Txt_IPI_TBSN = IIf(IsNull(TBLISTA!IPI), "", Format(TBLISTA!IPI, "###,##0.00"))
If Lbl_ICMS_TBSN.Visible = True Then Txt_ICMS_TBSN = IIf(IsNull(TBLISTA!ICMS), "", Format(TBLISTA!ICMS, "###,##0.00")) Else Txt_ICMS_TBSN = IIf(IsNull(TBLISTA!ISS), "", Format(TBLISTA!ISS, "###,##0.00"))
Txt_valor_deduzir_TBSN = IIf(IsNull(TBLISTA!Valor_deduzir), "", Format(TBLISTA!Valor_deduzir, "###,##0.00"))
Frame1(45).Enabled = True
Frame1(47).Enabled = True
Novo_geral9 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosEmail()
On Error GoTo tratar_erro

Txt_ID_email = TBProduto!ID
Txt_data_email = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
Txt_responsavel_email = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
Select Case TBProduto!Aplicacao
    Case "C": Aplicacao = "Compras"
    Case "CU": Aplicacao = "Custos"
    Case "F": Aplicacao = "Financeiro"
    Case "V": Aplicacao = "Vendas"
    Case "FA": Aplicacao = "Faturamento"
End Select
Cmb_aplicacao_email = Aplicacao
If IsNull(TBProduto!Usuario_caprind) = False And TBProduto!Usuario_caprind <> "" Then Cmb_usuario_caprind_email = TBProduto!Usuario_caprind
Txt_servidor_SMTP_email = IIf(IsNull(TBProduto!Servidor_SMTP), "", TBProduto!Servidor_SMTP)
txt_porta_email = IIf(IsNull(TBProduto!Porta), "", TBProduto!Porta)
Select Case TBProduto!Seguranca
    Case "N": Seguranca = "Não segura"
    Case "S": Seguranca = "SSL/TSL"
End Select
Cmb_seguranca_email = Seguranca
Txt_nome_email = IIf(IsNull(TBProduto!Nome), "", TBProduto!Nome)
Txt_email_email.Text = IIf(IsNull(TBProduto!Email), "", TBProduto!Email)
Txt_usuario_email = IIf(IsNull(TBProduto!Usuario), "", TBProduto!Usuario)
Txt_senha_email = IIf(IsNull(TBProduto!Senha), "", TBProduto!Senha)
Frame1(30).Enabled = True
Novo_geral6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosFiltros()
On Error GoTo tratar_erro

txtID_Filtros = TBProduto!ID
txtData_Filtros = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
txtResponsavel_Filtros = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
Select Case TBProduto!Aplicacao
    Case "C": Aplicacao = "Compras"
    Case "V": Aplicacao = "Vendas"
    Case "P": Aplicacao = "PCP"
    Case "E": Aplicacao = "Engenharia"
    Case "Q": Aplicacao = "Qualidade"
    Case "T": Aplicacao = "Estoque"
    Case "F": Aplicacao = "Faturamento"
    Case "M": Aplicacao = "Manutenção"
    Case "O": Aplicacao = "Outros"
End Select
cmbAplicacao_Filtros = Aplicacao

If IsNull(TBProduto!Tipo) = False And TBProduto!Tipo <> "" Then cmbTipo_Filtros = TBProduto!Tipo
If IsNull(TBProduto!filtrarpor) = False And TBProduto!filtrarpor <> "" Then cmbfiltrarpor_Filtros = TBProduto!filtrarpor
If IsNull(TBProduto!Frase) = False And TBProduto!Frase <> "" Then cmbFrase_Filtros = TBProduto!Frase
Frame1(31).Enabled = True
Novo_geral7 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosArmaz()
On Error GoTo tratar_erro

Txt_ID_armaz = TBProduto!ID
Txt_data_armaz = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
Txt_responsavel_armaz = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
If IsNull(TBProduto!Relatorio) = False And TBProduto!Relatorio <> "" Then Cmb_relatorio_armaz = TBProduto!Relatorio
Txt_local_armaz = IIf(IsNull(TBProduto!caminho), "", TBProduto!caminho)
Frame1(32).Enabled = True
Novo_geral8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosMoeda()
On Error GoTo tratar_erro

txtidmoeda.Text = TBProduto!CODIGO
txtMoeda.Text = TBProduto!Moeda
txtSimbolo.Text = TBProduto!Simbolo
Txt_data_moeda.Text = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
Txt_responsavel_moeda.Text = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
If txtMoeda = "DÓLAR" Or txtMoeda = "REAL" Then
    txtMoeda.Locked = True
    txtMoeda.TabStop = False
Else
    txtMoeda.Locked = False
    txtMoeda.TabStop = True
End If
Frame1(33).Enabled = True
Novo_geral1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosUnidade()
On Error GoTo tratar_erro

txtidunidade.Text = TBProduto!CODIGO
Txt_data_unidade = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
Txt_responsavel_unidade = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
Txt_unidade = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
Txt_descricao_unidade = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
Frame1(34).Enabled = True
Novo_geral2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosConversao()
On Error GoTo tratar_erro

Txt_ID_conversao = TBProduto!ID
Txt_data_conversao = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
Txt_responsavel_conversao = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
Txt_qtde_de_conversao = IIf(IsNull(TBProduto!Qtde_de), "", Format(TBProduto!Qtde_de, "###,##0.0000"))
If IsNull(TBProduto!Unidade_de) = False And TBProduto!Unidade_de <> "" Then Cmb_unidade_de_conversao = TBProduto!Unidade_de
Txt_qtde_para_conversao = IIf(IsNull(TBProduto!Qtde_para), "", TBProduto!Qtde_para)
If IsNull(TBProduto!Unidade_para) = False And TBProduto!Unidade_para <> "" Then Cmb_unidade_para_conversao = TBProduto!Unidade_para
Frame1(35).Enabled = True
Novo_geral3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosCondicao()
On Error GoTo tratar_erro

Txt_ID_cond = TBProduto!ID
Txt_data_cond = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
Txt_responsavel_cond = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
Txt_texto_cond = IIf(IsNull(TBProduto!Texto), "", TBProduto!Texto)
Cmb_aplicacao_cond = IIf(TBProduto!Tipo = "C", "Compras", "Vendas")
Frame1(36).Enabled = True
Novo_geral4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosFeriado()
On Error GoTo tratar_erro

Txt_ID_feriado = TBProduto!ID
Txt_data_feriado = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
Txt_responsavel_feriado = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
Cmb_data_feriado = TBProduto!Data_feriado
Txt_descricao_feriado = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
Frame1(37).Enabled = True
Novo_geral5 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ICMS_ind_Change()
On Error GoTo tratar_erro

If Txt_ICMS_ind.Text <> "" Then
    VerifNumero = Txt_ICMS_ind.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ICMS_ind.Text = ""
        Txt_ICMS_ind.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_minutos_desconectar_Change()
On Error GoTo tratar_erro

If Txt_minutos_desconectar <> "" Then
    VerifNumero = Txt_minutos_desconectar
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_minutos_desconectar = ""
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_de_conversao_Change()
On Error GoTo tratar_erro

If Txt_qtde_de_conversao <> "" Then
    VerifNumero = Txt_qtde_de_conversao
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_de_conversao = ""
        Txt_qtde_de_conversao.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_de_conversao_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_de_conversao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_de_conversao_LostFocus()
On Error GoTo tratar_erro

Txt_qtde_de_conversao = Format(Txt_qtde_de_conversao, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_para_conversao_Change()
On Error GoTo tratar_erro

If Txt_qtde_para_conversao <> "" Then
    VerifNumero = Txt_qtde_para_conversao
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_para_conversao = ""
        Txt_qtde_para_conversao.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_para_conversao_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_para_conversao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_para_conversao_LostFocus()
On Error GoTo tratar_erro

Txt_qtde_para_conversao = Txt_qtde_para_conversao
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_deduzir_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_valor_deduzir_TBSN <> "" Then
    VerifNumero = Txt_valor_deduzir_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_deduzir_TBSN = ""
        Txt_valor_deduzir_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_deduzir_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_valor_deduzir_TBSN = Format(Txt_valor_deduzir_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub txtCofins_Change()
On Error GoTo tratar_erro

If txtCofins.Text <> "" Then
    VerifNumero = txtCofins.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCofins.Text = ""
        txtCofins.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCofins1_Change()
On Error GoTo tratar_erro

If txtCofins1.Text <> "" Then
    VerifNumero = txtCofins1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCofins1.Text = ""
        txtCofins1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCSLL_Change()
On Error GoTo tratar_erro

If txtCSLL.Text <> "" Then
    VerifNumero = txtCSLL.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCSLL.Text = ""
        txtCSLL.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCSLL1_Change()
On Error GoTo tratar_erro

If txtCSLL1.Text <> "" Then
    VerifNumero = txtCSLL1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCSLL1.Text = ""
        txtCSLL1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtINSS_Change()
On Error GoTo tratar_erro

If txtINSS.Text <> "" Then
    VerifNumero = txtINSS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtINSS.Text = ""
        txtINSS.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIRPJ_Change()
On Error GoTo tratar_erro

If txtIRPJ.Text <> "" Then
    VerifNumero = txtIRPJ.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIRPJ.Text = ""
        txtIRPJ.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIRPJ1_Change()
On Error GoTo tratar_erro

If txtIRPJ1.Text <> "" Then
    VerifNumero = txtIRPJ1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIRPJ1.Text = ""
        txtIRPJ1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIRPJ_serv_Change()
On Error GoTo tratar_erro

If txtIRPJ_serv.Text <> "" Then
    VerifNumero = txtIRPJ_serv.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIRPJ_serv.Text = ""
        txtIRPJ_serv.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIRPJ_serv_maior_Change()
On Error GoTo tratar_erro

If txtIRPJ_serv_maior.Text <> "" Then
    VerifNumero = txtIRPJ_serv_maior.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIRPJ_serv_maior.Text = ""
        txtIRPJ_serv_maior.SetFocus
        Exit Sub
    End If
End If
txtIRPJ_serv_ate = txtIRPJ_serv_maior

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIRPJ_serv_maior_LostFocus()
On Error GoTo tratar_erro

txtIRPJ_serv_maior = Format(txtIRPJ_serv_maior, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIRRF_Change()
On Error GoTo tratar_erro

If txtIRRF.Text <> "" Then
    VerifNumero = txtIRRF.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIRRF.Text = ""
        txtIRRF.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtiss_Change()
On Error GoTo tratar_erro

If txtiss.Text <> "" Then
    VerifNumero = txtiss.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtiss.Text = ""
        txtiss.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpis_Change()
On Error GoTo tratar_erro

If txtPIS.Text <> "" Then
    VerifNumero = txtPIS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPIS.Text = ""
        txtPIS.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPIS1_Change()
On Error GoTo tratar_erro

If txtPIS1.Text <> "" Then
    VerifNumero = txtPIS1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPIS1.Text = ""
        txtPIS1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposEmpresa2()
On Error GoTo tratar_erro

txtID_imposto = 0
    
'Serviços
txtPIS = ""
txtCofins = ""
txtCSLL = ""
txtiss = ""
txtIRRF = ""
txtVLR = ""
txtINSS = ""
txtVLR1 = ""
txtIRPJ = ""
txtIRPJ_serv_maior = ""
txtIRPJ_serv = ""

'Produtos
txtPIS1 = ""
txtCofins1 = ""
txtCSLL1 = ""
txtIRPJ1 = ""
    
'Simples nacional
Txt_valor_total_faturado = ""
ProcLimpaCamposTBSN

Txt_ICMS_ind = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposEmpresa3()
On Error GoTo tratar_erro

Txt_local_armaz_NFe = ""
txtRetornoNF = ""
'Txt_local_armaz_CCe = ""
txtCaminhoXMLDanfe = ""

txtUsuarioPref = ""
txtSenhaPref = ""

Txt_registro_boleto = ""
Txt_apelido_contimatic = ""
'txtChaveMigrate = ""

'Caprind
Chk_bloquear_prod_cliente.Value = 0
Chk_bloquear_forn.Value = 0
Chk_bloquear_cli_forn_regime.Value = 0
Chk_CC_obrigatorio.Value = 0
Chk_codigo_ref_DANFE.Value = 0
Chk_codigo_ref_desc_DANFE.Value = 0
Chk_liberar_qtde_MRP.Value = 0
Chk_calcular_IPI.Value = 0
Chk_bloquear_NF_prod_serv_sem_cad.Value = 0
Chk_ativar_empenho_aut.Value = 0
Chk_ativar_empenho_aut_prod.Value = 0
Chk_carregar_CFOP_ST.Value = 0
Chk_agregar_ordem_valor_PC.Value = 0
Chk_gerar_RM_ordem_PC.Value = 0
Chk_liberar_campos_estrutura.Value = 0
Chk_verificar_desconectar_usuario.Value = 0
Txt_minutos_desconectar = ""
chk_Esconder_ValorOF.Value = 0
Chk_movimentar_estoque_pc.Value = 0
Chk_ativar_produtos_similares.Value = 0
Chk_validar_proposta_pi_autom.Value = 0
Chk_codigo_ref_SPED_forn.Value = 0
chkLiberar_LoteMinimo.Value = 0
Chk_carregar_LA_entrada.Value = 0
chkNao_inspecionar.Value = 0
ChkBloc_CC_Previsao.Value = 0
chk_Baixa_Auto_Estoque_NF.Value = 0
Chk_bloq_OP_estrutura.Value = 0
Chk_bloq_OP_processo.Value = 0
Chk_bloq_OP_plano.Value = 0
Chk_bloq_compra_cot_valida.Value = 0
chkCodigo_sequencial.Value = 0
Chk_salvar_status_aprovado_PC.Value = 0
Chk_enviar_email_outlook.Value = 0
chkMargemAnalise.Value = 0

'Gerprod
Chk_ap_codigo.Value = 0
Chk_bloquear_apontamento_sem_baixa.Value = 0
Chk_bloquear_apontamento_sem_baixa_total.Value = 0
Chk_desbloquear_primeiro_apontamento_OS_proc_controlado.Value = 0
chk_Grupo_Gerprod.Value = 0
Chk_bloquear_apontamento_simultaneo.Value = 0
Chk_apontar_NC_descricao.Value = 0
Chk_NC_parecer.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVLR_Change()
On Error GoTo tratar_erro

If txtVLR.Text <> "" Then
    VerifNumero = txtVLR.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtVLR.Text = ""
        txtVLR.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVLR_LostFocus()
On Error GoTo tratar_erro

txtVLR = Format(txtVLR, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVLR1_Change()
On Error GoTo tratar_erro

If txtVLR1.Text <> "" Then
    VerifNumero = txtVLR1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtVLR1.Text = ""
        txtVLR1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadadosImpostos()
On Error GoTo tratar_erro

TBGravar!ID_empresa = txtIDEmpresa
TBGravar!Regime = Regime
'Serviços
TBGravar!PIS = IIf(txtPIS = "", 0, txtPIS)
TBGravar!Cofins = IIf(txtCofins = "", 0, txtCofins)
TBGravar!CSLL = IIf(txtCSLL = "", 0, txtCSLL)
TBGravar!ISS = IIf(txtiss = "", 0, txtiss)
TBGravar!IRRF = IIf(txtIRRF = "", 0, txtIRRF)
TBGravar!Acima = IIf(txtVLR = "", 0, txtVLR)
TBGravar!INSS = IIf(txtINSS = "", 0, txtINSS)
TBGravar!INSS_acima = IIf(txtVLR1 = "", 0, txtVLR1)
TBGravar!IRPJ = IIf(txtIRPJ = "", 0, txtIRPJ)
TBGravar!IRPJ_serv_atemaior = IIf(txtIRPJ_serv_maior = "", 0, txtIRPJ_serv_maior)
TBGravar!IRPJ_servicos = IIf(txtIRPJ_serv = "", 0, txtIRPJ_serv)

'Produtos
TBGravar!PIS_produtos = IIf(txtPIS1 = "", 0, txtPIS1)
TBGravar!Cofins_produtos = IIf(txtCofins1 = "", 0, txtCofins1)
TBGravar!CSLL_produtos = IIf(txtCSLL1 = "", 0, txtCSLL1)
TBGravar!IRPJ_produtos = IIf(txtIRPJ1 = "", 0, txtIRPJ1)

'Aliquota ICMS industrialização
TBGravar!ICMS_ind = IIf(Txt_ICMS_ind = "", Null, Txt_ICMS_ind)

If chkDuplicata.Value = 1 Then TBGravar!Duplicata = True Else TBGravar!Duplicata = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosImpostos()
On Error GoTo tratar_erro

ProcLimpaCamposEmpresa2
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Impostos where ID_empresa = " & txtIDEmpresa & " and Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    txtID_imposto = TBFIltro!ID
    
    'Serviços
    txtPIS = IIf(IsNull(TBFIltro!PIS), "", TBFIltro!PIS)
    txtCofins = IIf(IsNull(TBFIltro!Cofins), "", TBFIltro!Cofins)
    txtCSLL = IIf(IsNull(TBFIltro!CSLL), "", TBFIltro!CSLL)
    txtiss = IIf(IsNull(TBFIltro!ISS), "", TBFIltro!ISS)
    txtIRRF = IIf(IsNull(TBFIltro!IRRF), "", TBFIltro!IRRF)
    txtVLR = IIf(IsNull(TBFIltro!Acima), "", Format(TBFIltro!Acima, "###,##0.00"))
    txtINSS = IIf(IsNull(TBFIltro!INSS), "", TBFIltro!INSS)
    txtVLR1 = IIf(IsNull(TBFIltro!INSS_acima), "", Format(TBFIltro!INSS_acima, "###,##0.00"))
    txtIRPJ = IIf(IsNull(TBFIltro!IRPJ), "", TBFIltro!IRPJ)
    txtIRPJ_serv_maior = IIf(IsNull(TBFIltro!IRPJ_serv_atemaior), "", Format(TBFIltro!IRPJ_serv_atemaior, "###,##0.00"))
    txtIRPJ_serv = IIf(IsNull(TBFIltro!IRPJ_servicos), "", TBFIltro!IRPJ_servicos)
    
    'Produtos
    txtPIS1 = IIf(IsNull(TBFIltro!PIS_produtos), "", TBFIltro!PIS_produtos)
    txtCofins1 = IIf(IsNull(TBFIltro!Cofins_produtos), "", TBFIltro!Cofins_produtos)
    txtCSLL1 = IIf(IsNull(TBFIltro!CSLL_produtos), "", TBFIltro!CSLL_produtos)
    txtIRPJ1 = IIf(IsNull(TBFIltro!IRPJ_produtos), "", TBFIltro!IRPJ_produtos)
        
    'Simples nacional
    Txt_valor_total_faturado = IIf(IsNull(TBFIltro!Vlr_total_faturado), "", Format(TBFIltro!Vlr_total_faturado, "###,##0.00"))
    'Aliquota ICMS industrialização
    Txt_ICMS_ind = IIf(IsNull(TBFIltro!ICMS_ind), "", TBFIltro!ICMS_ind)
    
    'Duplicata
    If TBFIltro!Duplicata = False Then chkDuplicata.Value = 0 Else chkDuplicata.Value = 1
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVLR1_LostFocus()
On Error GoTo tratar_erro

txtVLR1 = Format(txtVLR1, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoBD
    Case 2: ProcSalvarBD
    Case 3: ProcExcluirBD
    Case 4: ProcAlterarBD
    Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_Moeda
    Case 2: ProcSalvar_Moeda
    Case 3: ProcExcluir_moeda
    Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

If SSTab2.Tab = 0 Then
    Select Case ButtonIndex
        Case 1: ProcNovo_unidade
        Case 2: ProcSalvar_unidade
        Case 3: ProcExcluir_unidade
        Case 5: ProcAjuda
        Case 6: ProcSair
    End Select
Else
    Select Case ButtonIndex
        Case 1: ProcNovo_conversao
        Case 2: ProcSalvar_conversao
        Case 3: ProcExcluir_conversao
        Case 5: ProcAjuda
        Case 6: ProcSair
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case SSTabEmpresa.Tab
    Case 0:
        Select Case ButtonIndex
            Case 1: ProcNovo_empresa
            Case 2: ProcSalvar_empresa
            Case 3: ProcExcluir_empresa
            Case 4: ProcConf_rel
            Case 5: procAtualiza
            Case 7: ProcAjuda
            Case 8: ProcSair
        End Select
    Case 3:
        Select Case ButtonIndex
            Case 1: ProcNovo_email
            Case 2: ProcSalvar_email
            Case 3: ProcExcluir_email
            Case 5: procAtualiza
            Case 7: ProcAjuda
            Case 8: ProcSair
        End Select
    Case 4:
        Select Case ButtonIndex
            Case 1: ProcNovo_filtros
            Case 2: ProcSalvar_Filtros
            Case 3: ProcExcluir_Filtros
            Case 5: procAtualiza
            Case 7: ProcAjuda
            Case 8: ProcSair
        End Select
    Case 5:
        Select Case ButtonIndex
            Case 1: ProcNovo_armaz
            Case 2: ProcSalvar_Armaz
            Case 3: ProcExcluir_Armaz
            Case 5: procAtualiza
            Case 7: ProcAjuda
            Case 8: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar5_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case SSTabEmpresa.Tab
    Case 1:
        Select Case ButtonIndex
            Case 1: ProcSalvar_empresa2
            Case 3: ProcAjuda
            Case 4: ProcSair
        End Select
    Case 2:
        Select Case ButtonIndex
            Case 1: ProcSalvar_empresa3
            Case 3: ProcAjuda
            Case 4: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar6_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_condicao
    Case 2: ProcSalvar_condicao
    Case 3: ProcExcluir_condicao
    Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar7_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_feriado
    Case 2: ProcSalvar_feriado
    Case 3: ProcExcluir_feriado
    Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar8_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoTabelaSN
    Case 2: ProcSalvarTabelaSN
    Case 3: ProcExcluirTabelaSN
    Case 4: ProcAtualizaTabelaSN
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcConf_rel()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDEmpresa = "" Or txtIDEmpresa = "0" Then
    USMsgBox "Informe a empresa na lista antes de configurar os dados dos relatórios.", vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmOpcoesGeral_ConfRelatorio.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifMostrarEsconderTab(NTab As Integer)
On Error GoTo tratar_erro

With SSTab5
    .TabVisible(NTab) = False
    .TabVisible(IIf(NTab = 0, 1, 0)) = True
    .TabsPerRow = 1
    .Tab = IIf(NTab = 0, 1, 0)
End With
If SSTabEmpresa.Tab = 1 Then
    If NTab = 0 Then PBLista.Visible = True Else PBLista.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Aliquota_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_Aliquota_TBSN <> "" Then
    VerifNumero = Txt_Aliquota_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_Aliquota_TBSN = ""
        Txt_Aliquota_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Aliquota_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_Aliquota_TBSN = Format(Txt_Aliquota_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ate_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_ate_TBSN <> "" Then
    VerifNumero = Txt_ate_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ate_TBSN = ""
        Txt_ate_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ate_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_ate_TBSN = Format(Txt_ate_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_de_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_de_TBSN <> "" Then
    VerifNumero = Txt_de_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_de_TBSN = ""
        Txt_de_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_de_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_de_TBSN = Format(Txt_de_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ICMS_TBSN_Change()
On Error GoTo tratar_erro
    
If Txt_ICMS_TBSN <> "" Then
    VerifNumero = Txt_ICMS_TBSN
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ICMS_TBSN = ""
        Txt_ICMS_TBSN.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ICMS_TBSN_LostFocus()
On Error GoTo tratar_erro

Txt_ICMS_TBSN = Format(Txt_ICMS_TBSN, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
