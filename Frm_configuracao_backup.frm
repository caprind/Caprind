VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_configuracao_backup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações do sistema - Criar backup - Configurações"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   8085
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000007&
      Height          =   825
      Left            =   55
      TabIndex        =   28
      Top             =   990
      Width           =   8415
      Begin VB.TextBox Txt_responsavel 
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   375
         Width           =   7095
      End
      Begin VB.TextBox Txt_data 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   375
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Index           =   9
         Left            =   4230
         TabIndex        =   30
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   480
         TabIndex        =   29
         Top             =   180
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar backups realizados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   7
      Left            =   55
      TabIndex        =   23
      Top             =   3855
      Width           =   8415
      Begin VB.OptionButton Opt_ate_mes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Até o mês"
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
         Left            =   990
         TabIndex        =   26
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton Opt_do_mes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do mês"
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
         Left            =   150
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.ComboBox Cmb_ano 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Frm_configuracao_backup.frx":0000
         Left            =   7440
         List            =   "Frm_configuracao_backup.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Ano."
         Top             =   210
         Width           =   795
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Top             =   240
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         MultiRow        =   -1  'True
         TabMinWidth     =   265
         TabStyle        =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   12
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jan"
               Key             =   "Jan"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fev"
               Key             =   "Fev"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mar"
               Key             =   "Mar"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Abr"
               Key             =   "Abr"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Maio"
               Key             =   "Maio"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jun"
               Key             =   "Jun"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jul"
               Key             =   "Jul"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ago"
               Key             =   "Ago"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Set"
               Key             =   "Set"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Out"
               Key             =   "Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Nov"
               Key             =   "Nov"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dez"
               Key             =   "Dez"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7350
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000007&
      Height          =   2055
      Left            =   55
      TabIndex        =   5
      Top             =   1800
      Width           =   8415
      Begin VB.CommandButton Cmd_excluir_armaz_copia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7890
         Picture         =   "Frm_configuracao_backup.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar caminho do pedido de compra."
         Top             =   1320
         Width           =   315
      End
      Begin VB.CheckBox Chk_excluir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Excluir arquivos do local de armazenamento"
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
         Height          =   195
         Left            =   4200
         TabIndex        =   21
         Top             =   1740
         Width           =   4065
      End
      Begin VB.CommandButton Cmd_localizar_armaz_copia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7560
         Picture         =   "Frm_configuracao_backup.frx":0142
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Localizar."
         Top             =   1320
         Width           =   315
      End
      Begin VB.TextBox Txt_local_armaz_copia 
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
         TabIndex        =   18
         ToolTipText     =   "Local de armazenamento da cópia do backup."
         Top             =   1320
         Width           =   7365
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
         Left            =   1320
         MaxLength       =   255
         TabIndex        =   22
         ToolTipText     =   "Local de armazenamento."
         Top             =   720
         Width           =   6585
      End
      Begin VB.CommandButton Cmd_localizar_armaz 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7890
         Picture         =   "Frm_configuracao_backup.frx":0244
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Localizar."
         Top             =   720
         Width           =   315
      End
      Begin VB.CheckBox Chk_compactar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compactar"
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
         Left            =   6960
         TabIndex        =   14
         Top             =   180
         Width           =   1245
      End
      Begin VB.CheckBox Chk_dom 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Domingo"
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
         Left            =   5730
         TabIndex        =   13
         Top             =   180
         Width           =   915
      End
      Begin VB.CheckBox Chk_sab 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sábado"
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
         Left            =   4760
         TabIndex        =   12
         Top             =   180
         Width           =   855
      End
      Begin VB.CheckBox Chk_qui 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quinta"
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
         Left            =   3000
         TabIndex        =   10
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox Chk_qua 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quarta"
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
         Left            =   2060
         TabIndex        =   9
         Top             =   180
         Width           =   825
      End
      Begin VB.CheckBox Chk_ter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Terça"
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
         Left            =   1240
         TabIndex        =   8
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox Chk_seg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Segunda"
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
         TabIndex        =   7
         Top             =   180
         Width           =   1035
      End
      Begin VB.CheckBox Chk_sex 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sexta"
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
         Left            =   3910
         TabIndex        =   6
         Top             =   180
         Width           =   735
      End
      Begin MSComCtl2.DTPicker Txt_hora 
         Height          =   315
         Left            =   180
         TabIndex        =   15
         ToolTipText     =   "Hora do inicio do evento."
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
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
         Format          =   180813826
         CurrentDate     =   39055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Local de armazenamento para cópia do backup"
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
         Left            =   2182
         TabIndex        =   31
         Top             =   1110
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora*"
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
         Index           =   24
         Left            =   472
         TabIndex        =   16
         Top             =   510
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Local de armazenamento*"
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
         Left            =   3645
         TabIndex        =   11
         Top             =   510
         Width           =   1875
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   7770
      Width           =   8415
      _ExtentX        =   14843
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   3
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   44
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   48
      ButtonTop2      =   2
      ButtonWidth2    =   45
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
      ButtonLeft3     =   95
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
      ButtonLeft4     =   99
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
      ButtonLeft5     =   142
      ButtonTop5      =   2
      ButtonWidth5    =   30
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
      ButtonLeft6     =   174
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   6060
         Top             =   0
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "Frm_configuracao_backup.frx":0346
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3360
      Left            =   60
      TabIndex        =   2
      Top             =   4395
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5927
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
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Hora"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Local de armazenamento"
         Object.Width           =   9393
      EndProperty
   End
End
Attribute VB_Name = "Frm_configuracao_backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) histórico(s) de backup?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            '==================================
            Modulo = "Configuração do sistema/Criar backup/Configurações"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Data: " & .ListItems(InitFor).ListSubItems(1) & " - Hora: " & .ListItems(InitFor).ListSubItems(2) & " - Local de armazenamento: " & .ListItems(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Conexao.Execute "DELETE from Backup_historico where id = " & .ListItems(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) histórico(s) de backup antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Histórico(s) de backup excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
End If

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
If Chk_seg.Value = 0 And Chk_ter.Value = 0 And Chk_qua.Value = 0 And Chk_qui.Value = 0 And Chk_sex.Value = 0 And Chk_sab.Value = 0 And Chk_dom.Value = 0 Then
    NomeCampo = "o dia da semana para efetuar o backup"
    ProcVerificaAcao
    Exit Sub
End If
If Txt_hora = "00:00:00" Then
    NomeCampo = "a hora"
    ProcVerificaAcao
    Txt_hora.SetFocus
    Exit Sub
End If
If Txt_local_armaz = "" Then
    NomeCampo = "o local de armazenamento"
    ProcVerificaAcao
    Txt_local_armaz.SetFocus
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Backup_configuracoes", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Data = Date
TBAbrir!Responsavel = pubUsuario
TBAbrir!Seg = Chk_seg.Value
TBAbrir!Ter = Chk_ter.Value
TBAbrir!Qua = Chk_qua.Value
TBAbrir!Qui = Chk_qui.Value
TBAbrir!Sex = Chk_sex.Value
TBAbrir!Sab = Chk_sab.Value
TBAbrir!dom = Chk_dom.Value
TBAbrir!Compactar = Chk_compactar.Value
TBAbrir!Hora = Format(Txt_hora.Value, "hh:mm:ss")
TBAbrir!Local = Txt_local_armaz
TBAbrir!Local_copia = Txt_local_armaz_copia
If Chk_excluir.Value = 1 Then TBAbrir!Excluir_arquivos = True Else TBAbrir!Excluir_arquivos = False
TBAbrir.Update

USMsgBox ("Configurações de backup cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Configuração do sistema/Criar backup/Configurações"
Evento = "Salvar"
ID_documento = TBAbrir!ID
Documento = ""
Documento1 = ""
ProcGravaEvento
'==================================

Txt_data = Format(TBAbrir!Data, "dd/mm/yy")
Txt_responsavel = TBAbrir!Responsavel

TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_excluir_armaz_copia_Click()
On Error GoTo tratar_erro

Txt_local_armaz_copia = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_armaz_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Localizar local de armazenamento do backup"
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

Private Sub Cmd_localizar_armaz_copia_Click()
On Error GoTo tratar_erro
  
szTitle = vbCr & vbCr & "Localizar local de armazenamento para cópia do backup"
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
    Txt_local_armaz_copia = sBuffer
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
 
Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8415, 6, True
Formulario = "Configuração do sistema/Criar backup/Configurações"
Direitos
ProcLimpaVariaveisPrincipais

TabFiltro.Tabs(Month(Date)).Selected = True
ProcCarregaComboAno Cmb_ano, "2005", 1

Txt_hora = "00:00:00"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Backup_configuracoes", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_data = Format(TBAbrir!Data, "dd/mm/yy")
    Txt_responsavel = TBAbrir!Responsavel
    Chk_seg.Value = IIf(TBAbrir!Seg = True, 1, 0)
    Chk_ter.Value = IIf(TBAbrir!Ter = True, 1, 0)
    Chk_qua.Value = IIf(TBAbrir!Qua = True, 1, 0)
    Chk_qui.Value = IIf(TBAbrir!Qui = True, 1, 0)
    Chk_sex.Value = IIf(TBAbrir!Sex = True, 1, 0)
    Chk_sab.Value = IIf(TBAbrir!Sab = True, 1, 0)
    Chk_dom.Value = IIf(TBAbrir!dom = True, 1, 0)
    Chk_compactar.Value = IIf(TBAbrir!Compactar = True, 1, 0)
    Txt_hora.Value = Format(TBAbrir!Hora, "hh:mm:ss")
    Txt_local_armaz = TBAbrir!Local
    Txt_local_armaz_copia = IIf(IsNull(TBAbrir!Local_copia), "", TBAbrir!Local_copia)
    If TBAbrir!Excluir_arquivos = True Then Chk_excluir.Value = 1 Else Chk_excluir.Value = 0
End If
TBAbrir.Close

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

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
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

Private Sub TabFiltro_Click()
On Error GoTo tratar_erro

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

M = FunVerificaMes(TabFiltro.SelectedItem.key)
If Opt_do_mes.Value = True Then DataFiltroTexto = "Month(Data)= '" & M & "' and Year(Data) = '" & IIf(Cmb_ano = "", Year(Date), Cmb_ano) & "'"
If Opt_ate_mes.Value = True Then DataFiltroTexto = "Month(Data)<= '" & M & "' and Year(Data) = '" & IIf(Cmb_ano = "", Year(Date), Cmb_ano) & "'"
If TabFiltro.SelectedItem.key = "Todas" Then DataFiltroTexto = ""

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Backup_historico where " & DataFiltroTexto & " order by Data desc, Hora desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Hora), "", Left(TBLISTA!Hora, 8))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Local), "", TBLISTA!Local)
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

Private Sub Txt_local_armaz_copia_Change()
On Error GoTo tratar_erro

With Chk_excluir
    If Txt_local_armaz_copia = "" Then
        .Value = 0
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
