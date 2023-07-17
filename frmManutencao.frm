VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmManutencao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Manutenção - Equipamentos"
   ClientHeight    =   10035
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15660
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
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
      Height          =   2715
      Left            =   90
      TabIndex        =   90
      Top             =   1680
      Visible         =   0   'False
      Width           =   15555
      Begin VB.CommandButton Cmd_visualizar_arquivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14670
         Picture         =   "frmManutencao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Visualizar arquivo."
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton Cmd_limpar_caminho 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14340
         Picture         =   "frmManutencao.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar caminho."
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton Cmd_localizar_relatorio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14010
         Picture         =   "frmManutencao.frx":0700
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Localizar relatório."
         Top             =   990
         Width           =   315
      End
      Begin VB.TextBox Txt_doc_ref 
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
         Left            =   7650
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Documento de referência."
         Top             =   990
         Width           =   1245
      End
      Begin VB.TextBox Txt_caminho_relatorio 
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
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Caminho do relatório."
         Top             =   990
         Width           =   5115
      End
      Begin VB.TextBox Txt_fornecedor 
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
         Left            =   1890
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Fornecedor."
         Top             =   990
         Width           =   5745
      End
      Begin VB.TextBox Txt_tecnico_responsavel 
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
         Left            =   4920
         MaxLength       =   100
         TabIndex        =   10
         ToolTipText     =   "Nome do técnico responsável pela execução da manutenção."
         Top             =   390
         Width           =   4755
      End
      Begin VB.TextBox txtIDpedido 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   990
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPedido 
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
         MaxLength       =   100
         TabIndex        =   16
         ToolTipText     =   "Pedido de compra."
         Top             =   990
         Width           =   1300
      End
      Begin VB.CommandButton cmdLocalizarPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1500
         Picture         =   "frmManutencao.frx":0802
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Localizar pedido de compra."
         Top             =   990
         Width           =   315
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   4980
         TabIndex        =   91
         Top             =   390
         Width           =   1305
      End
      Begin VB.TextBox Txt_dias_proxima 
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
         Left            =   2490
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Dias para próxima manutenção."
         Top             =   390
         Width           =   1035
      End
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "frmManutencao.frx":0904
         Left            =   9690
         List            =   "frmManutencao.frx":090E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Status."
         Top             =   390
         Width           =   2580
      End
      Begin VB.TextBox Txt_data_conclusao 
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
         Left            =   12285
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Data da conclusão."
         Top             =   390
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox Txt_Hora_conclusao 
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
         Left            =   13680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Hora da conclusão."
         Top             =   390
         Visible         =   0   'False
         Width           =   1300
      End
      Begin VB.TextBox txtObs 
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
         Height          =   975
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         ToolTipText     =   "Descrição da manutenção."
         Top             =   1590
         Width           =   14805
      End
      Begin MSComCtl2.DTPicker txtData_Manutencao 
         Height          =   315
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "Data da manutenção."
         Top             =   390
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
         Format          =   915013633
         CurrentDate     =   39057
      End
      Begin MSMask.MaskEdBox Txt_data_conclusao1 
         Height          =   315
         Left            =   12285
         TabIndex        =   13
         ToolTipText     =   "Data da conclusão."
         Top             =   390
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker txtHora_Manutencao 
         Height          =   315
         Left            =   1470
         TabIndex        =   8
         ToolTipText     =   "Hora da manutenção."
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
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   915079170
         CurrentDate     =   40504.5572337963
      End
      Begin MSMask.MaskEdBox Txt_Hora_conclusao1 
         Height          =   315
         Left            =   13680
         TabIndex        =   15
         ToolTipText     =   "Hora da conclusão."
         Top             =   390
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker Cmb_data_proxima 
         Height          =   315
         Left            =   3540
         TabIndex        =   116
         ToolTipText     =   "Data para próxima manutenção."
         Top             =   390
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   915079169
         CurrentDate     =   39057
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
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
         Left            =   4350
         TabIndex        =   105
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. referência"
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
         Left            =   7717
         TabIndex        =   104
         Top             =   780
         Width           =   1110
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Caminho do relatório"
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
         Left            =   10725
         TabIndex        =   103
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico responsável"
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
         Left            =   6570
         TabIndex        =   102
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ped. de compra"
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
         Left            =   268
         TabIndex        =   101
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora conclusão"
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
         Left            =   13785
         TabIndex        =   100
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
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
         Left            =   1710
         TabIndex        =   99
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dt. p/ próxima"
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
         Left            =   3675
         TabIndex        =   98
         Top             =   180
         Width           =   1035
      End
      Begin VB.Image Img_data_conclusao 
         Height          =   360
         Left            =   13290
         Picture         =   "frmManutencao.frx":0925
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dt.conclusão"
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
         Left            =   12315
         TabIndex        =   97
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dias próxima"
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
         Left            =   2565
         TabIndex        =   96
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label15 
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
         Left            =   7237
         TabIndex        =   95
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   10748
         TabIndex        =   94
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
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
         Left            =   570
         TabIndex        =   93
         Top             =   180
         Width           =   345
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
      FormHeightDT    =   10500
      FormWidthDT     =   15780
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15660
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
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
      ItemData        =   "frmManutencao.frx":0DA8
      Left            =   1410
      List            =   "frmManutencao.frx":0DB2
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   9570
      Visible         =   0   'False
      Width           =   1965
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   65
      Top             =   9000
      Width           =   15555
      _ExtentX        =   27437
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
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   765
      Left            =   75
      TabIndex        =   55
      Top             =   9270
      Width           =   15555
      Begin VB.TextBox txtTotal 
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
         Left            =   8850
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Subtotal."
         Top             =   330
         Width           =   1815
      End
      Begin VB.TextBox txtTotalCheck 
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
         Left            =   6705
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do desconto."
         Top             =   330
         Width           =   1815
      End
      Begin VB.TextBox txtTotalSub 
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
         Left            =   4605
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   1  'Arrow
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do(s) serviço(s)"
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
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
         Left            =   1567
         TabIndex        =   107
         Top             =   120
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valot total"
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
         Left            =   9315
         TabIndex        =   63
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total check-list"
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
         TabIndex        =   62
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total substituição"
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
         Left            =   4755
         TabIndex        =   61
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "="
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8625
         TabIndex        =   60
         Top             =   390
         Width           =   135
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6510
         TabIndex        =   59
         Top             =   390
         Width           =   105
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4680
      Left            =   75
      TabIndex        =   6
      Top             =   4395
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   8255
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cód. equip."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5653
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "D"
         Text            =   "Dt. solic."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Requisitante"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Aprovado"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_check 
      Height          =   6465
      Left            =   75
      TabIndex        =   44
      Top             =   2490
      Visible         =   0   'False
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   11404
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
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   21352
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Executado"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_Data 
      Height          =   4605
      Left            =   105
      TabIndex        =   25
      Top             =   4350
      Visible         =   0   'False
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   8123
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
      NumItems        =   12
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
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Hora"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Dias p/ próxima"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "D"
         Text            =   "Dt. p/ próxima"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Total substituição"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Total check-list"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2178
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Dt. conclusão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "D"
         Text            =   "Hora conclusão"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Doc. ref."
         Object.Width           =   2241
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_desenho 
      Height          =   4980
      Left            =   105
      TabIndex        =   40
      Top             =   3975
      Visible         =   0   'False
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   8784
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
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   15531
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Quantidade"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Valot total"
         Object.Width           =   2646
      EndProperty
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   10065
      Left            =   60
      TabIndex        =   67
      Top             =   330
      Visible         =   0   'False
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   17754
      _Version        =   393216
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
      TabCaption(0)   =   "Data/Status"
      TabPicture(0)   =   "frmManutencao.frx":0DCB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CommonDialog1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImageList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "USImageList2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtIDData"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "USToolBar2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Produtos para substituição"
      TabPicture(1)   =   "frmManutencao.frx":0DE7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtidProduto"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USImageList3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "USToolBar3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Check-list"
      TabPicture(2)   =   "frmManutencao.frx":0E03
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtID_check"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "USImageList4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "USToolBar4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   45
         TabIndex        =   78
         Top             =   330
         Width           =   15465
         _ExtentX        =   27279
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
         ButtonCaption7  =   "Copiar"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Copiar (F7)"
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
         ButtonLeft7     =   309
         ButtonTop7      =   2
         ButtonWidth7    =   39
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
         ButtonLeft8     =   350
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
         ButtonLeft9     =   354
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
         ButtonLeft10    =   397
         ButtonTop10     =   2
         ButtonWidth10   =   30
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   5
         ButtonLeft11    =   429
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
      End
      Begin VB.Frame Frame8 
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
         Left            =   -74925
         TabIndex        =   71
         Top             =   1320
         Width           =   15495
         Begin VB.TextBox txtDescricao_Check 
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
            MaxLength       =   100
            TabIndex        =   41
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   12405
         End
         Begin VB.TextBox txtValorCheck 
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
            Left            =   12990
            TabIndex        =   43
            ToolTipText     =   "Valor."
            Top             =   390
            Width           =   1995
         End
         Begin VB.CommandButton cmdDescricao 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   12600
            Picture         =   "frmManutencao.frx":0E1F
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Localizar descrição padrão do check-list."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label24 
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
            Left            =   6037
            TabIndex        =   73
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            Left            =   13807
            TabIndex        =   72
            Top             =   180
            Width           =   360
         End
      End
      Begin VB.TextBox txtID_check 
         Height          =   285
         Left            =   -72525
         TabIndex        =   70
         Text            =   "0"
         Top             =   3060
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtidProduto 
         Height          =   285
         Left            =   -71925
         TabIndex        =   69
         Text            =   "0"
         Top             =   4140
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtIDData 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3690
         TabIndex        =   68
         Text            =   "0"
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin DrawSuite2022.USImageList USImageList4 
         Left            =   -67290
         Top             =   540
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmManutencao.frx":0F21
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   -68460
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmManutencao.frx":48FA
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   13140
         Top             =   600
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmManutencao.frx":86B6
         Count           =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10620
         Top             =   10020
         _ExtentX        =   794
         _ExtentY        =   794
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":E693
               Key             =   ""
               Object.Tag             =   "Novo"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":E9AD
               Key             =   ""
               Object.Tag             =   "Abrir"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":EDFF
               Key             =   ""
               Object.Tag             =   "Salvar"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":F119
               Key             =   ""
               Object.Tag             =   "Excluir"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":F56B
               Key             =   ""
               Object.Tag             =   "Sair"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":F9BD
               Key             =   ""
               Object.Tag             =   "Afericao"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":FCD7
               Key             =   ""
               Object.Tag             =   "curso"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":FFF1
               Key             =   ""
               Object.Tag             =   "Matricula"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":10443
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":10763
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":10BB7
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManutencao.frx":10CCB
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74940
         TabIndex        =   79
         Top             =   330
         Width           =   15465
         _ExtentX        =   27279
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
         ButtonCaption4  =   "Solicitação"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Criar solicitação (F6)"
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
         ButtonWidth4    =   69
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
         ButtonLeft5     =   204
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
         ButtonLeft6     =   208
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
         ButtonLeft7     =   251
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
         ButtonLeft8     =   283
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
         ButtonUseMaskColor8=   0   'False
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2490
         Top             =   3090
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame7 
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
         Height          =   1485
         Left            =   -74925
         TabIndex        =   80
         Top             =   2175
         Width           =   15465
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
            ItemData        =   "frmManutencao.frx":10FEB
            Left            =   7230
            List            =   "frmManutencao.frx":10FED
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Unidade comercial."
            Top             =   420
            Width           =   825
         End
         Begin VB.TextBox txtvalor_total 
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
            Left            =   2515
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Valor total."
            Top             =   1050
            Width           =   1215
         End
         Begin VB.TextBox txtvalorunitario 
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
            Left            =   1345
            MaxLength       =   50
            TabIndex        =   37
            ToolTipText     =   "Valor unitário."
            Top             =   1050
            Width           =   1155
         End
         Begin VB.TextBox txtQuantidade 
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
            TabIndex        =   36
            ToolTipText     =   "Quantidade."
            Top             =   1050
            Width           =   1155
         End
         Begin VB.TextBox txtdesc_desenho 
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
            Left            =   3750
            MaxLength       =   255
            TabIndex        =   39
            ToolTipText     =   "Descrição."
            Top             =   1050
            Width           =   11265
         End
         Begin VB.TextBox txtdesenho 
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
            MaxLength       =   50
            TabIndex        =   29
            ToolTipText     =   "Código interno."
            Top             =   420
            Width           =   2295
         End
         Begin VB.CommandButton cmdDesenho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2520
            Picture         =   "frmManutencao.frx":10FEF
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Localizar produtos."
            Top             =   420
            Width           =   315
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Criar novo produto"
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
            Left            =   2940
            TabIndex        =   81
            Top             =   210
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
               TabIndex        =   31
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
               TabIndex        =   32
               Top             =   270
               Width           =   1605
            End
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
            Left            =   8070
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            ToolTipText     =   "Família."
            Top             =   420
            Width           =   6930
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
            ItemData        =   "frmManutencao.frx":110F1
            Left            =   6390
            List            =   "frmManutencao.frx":110F3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   33
            ToolTipText     =   "Unidade estoque."
            Top             =   420
            Width           =   825
         End
         Begin VB.Label Label37 
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
            Left            =   7320
            TabIndex        =   106
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor total"
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
            Left            =   2745
            TabIndex        =   88
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor unitário"
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
            Left            =   1440
            TabIndex        =   87
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label16 
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
            Left            =   315
            TabIndex        =   86
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
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
            Left            =   8970
            TabIndex        =   85
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno"
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
            Left            =   712
            TabIndex        =   84
            Top             =   210
            Width           =   1230
         End
         Begin VB.Label Label20 
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
            Left            =   11295
            TabIndex        =   83
            Top             =   210
            Width           =   480
         End
         Begin VB.Label Label21 
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
            Left            =   6510
            TabIndex        =   82
            Top             =   210
            Width           =   585
         End
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
         Height          =   855
         Left            =   -74925
         TabIndex        =   74
         Top             =   1320
         Width           =   15465
         Begin VB.TextBox txtdescricao2 
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
            Left            =   3030
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do posto."
            Top             =   390
            Width           =   10485
         End
         Begin VB.TextBox txtIDmaquina2 
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
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Posto de trabalho."
            Top             =   390
            Width           =   2835
         End
         Begin VB.TextBox txtSolicitacao 
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
            Left            =   13530
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Solicitação."
            Top             =   390
            Width           =   1485
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
            Left            =   960
            TabIndex        =   77
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label13 
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
            Left            =   7927
            TabIndex        =   76
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitação"
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
            Left            =   13912
            TabIndex        =   75
            Top             =   180
            Width           =   750
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74925
         TabIndex        =   89
         Top             =   330
         Width           =   15465
         _ExtentX        =   27279
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
         ButtonCaption4  =   "Check-list"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Salvar marcação do check-list."
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
         ButtonWidth4    =   64
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
         ButtonLeft5     =   199
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
         ButtonLeft6     =   203
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
         ButtonLeft7     =   246
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
         ButtonLeft8     =   278
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
         ButtonUseMaskColor8=   0   'False
      End
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10065
      Left            =   30
      TabIndex        =   45
      Top             =   0
      Width           =   15660
      _ExtentX        =   27623
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
      TabCaption(0)   =   "Manutenção"
      TabPicture(0)   =   "frmManutencao.frx":110F5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "USImageList1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtid"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Dados da manutenção"
      TabPicture(1)   =   "frmManutencao.frx":11111
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo do serviço"
         Height          =   855
         Left            =   5190
         TabIndex        =   121
         Top             =   1320
         Width           =   3915
         Begin VB.CheckBox chkHidraulica 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hidraulica"
            Height          =   255
            Left            =   2010
            TabIndex        =   125
            Top             =   360
            Width           =   1065
         End
         Begin VB.CheckBox chkOutros 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Outros"
            Height          =   255
            Left            =   3060
            TabIndex        =   124
            Top             =   360
            Width           =   825
         End
         Begin VB.CheckBox chkMecanica 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mecânica"
            Height          =   255
            Left            =   960
            TabIndex        =   123
            Top             =   360
            Width           =   1065
         End
         Begin VB.CheckBox chkeletrica 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Elétrica"
            Height          =   255
            Left            =   90
            TabIndex        =   122
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo"
         Height          =   855
         Left            =   1620
         TabIndex        =   112
         Top             =   1320
         Width           =   3555
         Begin VB.OptionButton optPredial 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Predial"
            Height          =   195
            Left            =   1410
            TabIndex        =   115
            Top             =   390
            Width           =   825
         End
         Begin VB.OptionButton optPosto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Equipamento"
            DisabledPicture =   "frmManutencao.frx":1112D
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   90
            TabIndex        =   114
            Top             =   390
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optProduto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produto final"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2280
            TabIndex        =   113
            Top             =   390
            Width           =   1245
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código"
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
         Height          =   855
         Left            =   60
         TabIndex        =   109
         Top             =   1320
         Width           =   1545
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
            Left            =   90
            MaxLength       =   50
            TabIndex        =   111
            ToolTipText     =   "Código"
            Top             =   330
            Width           =   1335
         End
      End
      Begin VB.TextBox txtid 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11820
         TabIndex        =   66
         Text            =   "0"
         Top             =   750
         Visible         =   0   'False
         Width           =   615
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   10500
         Top             =   540
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmManutencao.frx":25B06F
         Count           =   1
      End
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
         Height          =   855
         Left            =   9105
         TabIndex        =   46
         Top             =   1320
         Width           =   6495
         Begin VB.TextBox txtIDMaquina 
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
            MaxLength       =   50
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Máquina."
            Top             =   390
            Width           =   1305
         End
         Begin VB.TextBox txtDescricao 
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
            Left            =   1740
            MaxLength       =   255
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Descrição da máquina."
            Top             =   390
            Width           =   4665
         End
         Begin DrawSuite2022.USButton cmdMáquina 
            Height          =   315
            Left            =   1410
            TabIndex        =   127
            ToolTipText     =   "Localizar..."
            Top             =   390
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmManutencao.frx":26300A
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
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
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
            Left            =   3547
            TabIndex        =   54
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Código"
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
            Left            =   495
            TabIndex        =   47
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2025
         Left            =   75
         TabIndex        =   48
         Top             =   2180
         Width           =   15525
         Begin VB.ComboBox cmbSetor_Requisitante 
            Height          =   315
            Left            =   3150
            TabIndex        =   137
            Top             =   990
            Width           =   9405
         End
         Begin VB.ComboBox cmbLocalizacao 
            Height          =   315
            Left            =   3240
            TabIndex        =   134
            Top             =   420
            Width           =   5985
         End
         Begin VB.TextBox txtAprovado 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   12690
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   131
            TabStop         =   0   'False
            ToolTipText     =   "Nome do responsável pela aprovação da manutenção."
            Top             =   420
            Width           =   2745
         End
         Begin VB.CheckBox chkControlada 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Controlada pelo gerprod?"
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
            Left            =   12780
            TabIndex        =   130
            Top             =   1710
            Width           =   2145
         End
         Begin VB.TextBox txtData 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9240
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   118
            TabStop         =   0   'False
            ToolTipText     =   "Data."
            Top             =   420
            Width           =   975
         End
         Begin VB.TextBox txtResponsavel 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10230
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   117
            TabStop         =   0   'False
            ToolTipText     =   "Responsável."
            Top             =   420
            Width           =   2325
         End
         Begin VB.TextBox txtLista 
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
            Height          =   375
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            ToolTipText     =   "Descrição da manutenção."
            Top             =   1530
            Width           =   12420
         End
         Begin VB.TextBox txttipo 
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
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Tipo da manutenção."
            Top             =   420
            Width           =   1740
         End
         Begin VB.TextBox txtRequisitante 
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
            Left            =   135
            MaxLength       =   100
            TabIndex        =   3
            ToolTipText     =   "Requisitante."
            Top             =   990
            Width           =   2985
         End
         Begin MSMask.MaskEdBox txtData_Solicitacao 
            Height          =   315
            Left            =   1920
            TabIndex        =   110
            ToolTipText     =   "Data da manutenção."
            Top             =   420
            Width           =   975
            _ExtentX        =   1720
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
         Begin DrawSuite2022.USButton imgSolicitacao 
            Height          =   315
            Left            =   2910
            TabIndex        =   126
            ToolTipText     =   "Abrir calendário"
            Top             =   420
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmManutencao.frx":28110F
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
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton cmdAutoriza 
            Height          =   825
            Left            =   12690
            TabIndex        =   132
            ToolTipText     =   "Aprovar documento de manutenção"
            Top             =   810
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   1455
            DibPicture      =   "frmManutencao.frx":2882A2
            Caption         =   "Aprovar documento"
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
            PicAlign        =   7
            PicSize         =   5
            PicSizeH        =   32
            PicSizeW        =   32
            Theme           =   3
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.TextBox txtSetor_Aprovado 
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
            Left            =   8760
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Setor do responsável pela aprovação da manutenção."
            Top             =   990
            Width           =   3795
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Programação"
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
            Left            =   2010
            TabIndex        =   136
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Localização"
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
            Left            =   5715
            TabIndex        =   135
            Top             =   210
            Width           =   810
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aprovado"
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
            Left            =   13800
            TabIndex        =   133
            Top             =   210
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data emissão"
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
            Left            =   9240
            TabIndex        =   120
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável emissão"
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
            Left            =   10627
            TabIndex        =   119
            Top             =   210
            Width           =   1530
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição problema"
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
            Left            =   5633
            TabIndex        =   53
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Objetivo"
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
            Left            =   810
            TabIndex        =   52
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Requisitante"
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
            Left            =   1170
            TabIndex        =   51
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Setor requisitante"
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
            Left            =   7207
            TabIndex        =   50
            Top             =   780
            Width           =   1290
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. solicitação"
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
            Left            =   -5175
            TabIndex        =   49
            Top             =   -960
            Width           =   1005
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   64
         Top             =   330
         Width           =   15525
         _ExtentX        =   27384
         _ExtentY        =   1720
         ButtonCount     =   14
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
         ButtonCaption8  =   "Filtrar todos"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Filtrar todos os registros."
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
         ButtonWidth8    =   77
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Copiar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Copiar (F7)"
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
         ButtonLeft9     =   432
         ButtonTop9      =   2
         ButtonWidth9    =   44
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Agenda"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Agenda (F8)"
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
         ButtonLeft10    =   478
         ButtonTop10     =   2
         ButtonWidth10   =   51
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonAlignment11=   2
         ButtonType11    =   1
         ButtonStyle11   =   -1
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   -1
         ButtonLeft11    =   531
         ButtonTop11     =   4
         ButtonWidth11   =   2
         ButtonHeight11  =   54
         ButtonCaption12 =   "Ajuda"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Ajuda (F1)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   535
         ButtonTop12     =   2
         ButtonWidth12   =   41
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Sair"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Sair (Esc)"
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
         ButtonLeft13    =   578
         ButtonTop13     =   2
         ButtonWidth13   =   30
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonKey14     =   "15"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   5
         ButtonLeft14    =   610
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         ButtonUseMaskColor14=   0   'False
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Setor para manutenção Predial"
         Height          =   855
         Left            =   9120
         TabIndex        =   128
         Top             =   1320
         Width           =   6495
         Begin VB.ComboBox cmbSetorPredial 
            Height          =   315
            Left            =   120
            TabIndex        =   129
            Top             =   360
            Width           =   6255
         End
      End
   End
End
Attribute VB_Name = "frmManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_manutencao As Boolean 'OK
Dim Novo_manutencao2 As Boolean 'OK
Dim Novo_manutencao3 As Boolean 'OK
Dim Novo_manutencao4 As Boolean 'OK
Dim Data_manutencao As Date 'OK

Public Manutencao_Produto As Boolean
Public Manutencao_Equipamento As Boolean
Public Manutencao_Predial As Boolean

Public Sql_Manutencao_Localizar As String 'OK
Public FormulaRel_Manutencao As String 'OK
Public FormulaRelSubReport_Manutencao As String 'OK
Dim VerificaCampos As Boolean 'OK

Private Sub ProcCarregaComboSetor()
On Error GoTo tratar_erro

With cmbSetorPredial

    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Usuarios_Setor order by Setor", Conexao, adOpenKeyset, adLockReadOnly
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
                    .AddItem TBCarregarCombo!Setor
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboSetorRequisitante()
On Error GoTo tratar_erro

With cmbSetor_Requisitante

    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Usuarios_Setor order by Setor", Conexao, adOpenKeyset, adLockReadOnly
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
                    .AddItem TBCarregarCombo!Setor
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboLocalizacao()
On Error GoTo tratar_erro

With cmbLocalizacao
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Usuarios_Setor order by Setor", Conexao, adOpenKeyset, adLockReadOnly
    If TBCarregarCombo.EOF = False Then
        Do While TBCarregarCombo.EOF = False
                    .AddItem TBCarregarCombo!Setor
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub chkAuto_Click()
On Error GoTo tratar_erro

txtdesenho = ""
If chkAuto.Value = 1 Then
    chkManual.Value = 0
    txtdesenho.Locked = True
    txtdesenho.TabStop = False
Else
    txtdesenho.Locked = False
    txtdesenho.TabStop = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkManual_Click()
On Error GoTo tratar_erro

txtdesenho = ""
If chkManual.Value = 1 Then
    chkAuto.Value = 0
    txtdesenho.Locked = False
    txtdesenho.TabStop = True
    txtdesenho.SetFocus
    USMsgBox ("Informe o código interno do produto."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_data_proxima_Change()
On Error GoTo tratar_erro

Dataini = txtData_Manutencao
DataFim = Cmb_data_proxima
If DataFim < Dataini Then
    USMsgBox ("A data da próxima manutenção não pode ser menor que a data da manutenção."), vbExclamation, "CAPRIND v5.0"
    Cmb_data_proxima.Value = Date
    Txt_dias_proxima = ""
    Exit Sub
End If
Txt_dias_proxima = DataFim - Dataini

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

ProcCarregaLista3
With USToolBar4
    Select Case Cmb_opcao_lista
        Case "Excluir":
            .ButtonState(3) = 0
            .ButtonState(4) = 5
        Case "Check-list":
            .ButtonState(3) = 5
            .ButtonState(4) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbstatus_Click()
On Error GoTo tratar_erro

If cmbStatus = "Concluída" Then
    Txt_data_conclusao1.Enabled = True
    Txt_Hora_conclusao1.Enabled = True
Else
    Txt_data_conclusao1 = "__/__/____"
    Txt_data_conclusao1.Enabled = False
    Txt_Hora_conclusao1 = "__:__:__"
    Txt_Hora_conclusao1.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

Txt_caminho_relatorio = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If Txt_caminho_relatorio <> "" Then ProcAbrirArquivo Txt_caminho_relatorio

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_relatorio_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_caminho_relatorio = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdAutoriza_Click()
On Error GoTo tratar_erro

If txttipo = "Solicitação" Then Exit Sub
frmManutencao_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrarTodos()
On Error GoTo tratar_erro

Sql_Manutencao_Localizar = "select * from manutencao order by idmaquina,tipo,codigo"
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCheck()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a alterar neste formulário."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista_check.ListItems.Count = 0 Then Exit Sub

If USMsgBox("Deseja realmente salvar esta(s) marcação(ões) do check-list?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then
    InitFor = 0
    With Lista_check
        PBLista.Max = .ListItems.Count
        PBLista.Min = 0
        PBLista.Value = 0
        Contador = 0
        For InitFor = 1 To .ListItems.Count
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "SELECT * from manutencao_checklist WHERE id = " & Lista_check.ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                If .ListItems.Item(InitFor).Checked = True Then TBGravar!Check_list = True Else TBGravar!Check_list = False
                '==================================
                Modulo = "Manutenção/Controle de manutenção"
                Evento = "Alterar marcação do check-list"
                ID_documento = Lista_check.ListItems(InitFor)
                Documento = "Check-list: " & IIf(IsNull(TBGravar!Descricao), "", TBGravar!Descricao)
                Documento1 = ""
                ProcGravaEvento
                '==================================
                TBGravar.Update
            End If
            TBGravar.Close
            Contador = Contador + 1
            PBLista.Value = Contador
        Next InitFor
    End With
    USMsgBox ("Marcação do(s) check-list(s) cadastrado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista3
    ProcAtualizaStatusData
    
    If txttipo = "Preventiva" Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select ID from Manutencao_Checklist where ID_Data = " & txtIDData & " and Check_list = 'False'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = True Then
            Conexao.Execute "Update Manutencao_data Set Status = 'Concluída', Data_conclusao = '" & Now & "' where ID = " & txtIDData
'            Set TBFIltro = CreateObject("adodb.recordset")
'            TBFIltro.Open "Select * from Manutencao_data where idManutencao = " & txtid & " order by Data", Conexao, adOpenKeyset, adLockOptimistic
'            If TBFIltro.BOF = False Then
'                TBFIltro.Find ("ID = " & txtIDData)
'                TBFIltro.MoveNext
'                If TBFIltro.EOF = True Then
                    If USMsgBox("Deseja criar uma nova manutenção para o dia " & Format(Cmb_data_proxima, "dd/mm/yy") & "?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then
                        ProcCopiaDadosData False, txtIDData, txtId, True
                        USMsgBox ("Nova data cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
                        '==================================
                        Modulo = "Manutenção/Controle de manutenção"
                        ID_documento = txtIDData
                        Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo
                        Documento1 = "Data da manutenção: " & Cmb_data_proxima & " - Hora da manutenção: " & txtHora_Manutencao
                        ProcGravaEvento
                        '==================================
                    End If
'                Else
'                    Conexao.Execute "Update Manutencao_data Set Data = '" & Cmb_data_proxima & "', Dias_proxima = " & Txt_dias_proxima & " where ID = " & TBFIltro!ID
'                End If
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "select * from manutencao_data where id  = " & txtIDData, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    ProcLimpaCampos_data
                    ProcPuxadados_data
                End If
                TBAbrir.Close
'            End If
        End If
        TBFIltro.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmddescricao_Click()
On Error GoTo tratar_erro

frmManutencao_descricao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenho_Click()
On Error GoTo tratar_erro

frmManutencao_item.Show 1

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
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) manutenção(ões)?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from manutencao where Codigo = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Manutenção/Controle de manutenção"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Equipamento : " & TBFI!IDMaquina & " - Tipo da manutenção: " & .ListItems(InitFor).ListSubItems(3) & " - Data da manutenção: " & .ListItems(InitFor).ListSubItems(4)
                Documento1 = ""
                ProcGravaEvento
                '==================================
                
                Conexao.Execute "DELETE from manutencao where Codigo = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from manutencao_data where idmanutencao = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from manutencao_defeito WHERE idmanutencao = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from manutencao_checklist where id_manutencao = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) manutenção(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Manutenção(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcDesabilitar
    ProcCarregaLista
    Novo_manutencao = False
    ProcLimparTudo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcDesabilitar()
On Error GoTo tratar_erro

With txtRequisitante
    .Locked = True
    .TabStop = False
End With
With cmbSetor_Requisitante
    .Locked = True
    .TabStop = False
End With


'Frame_data_sol.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_data()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_Data
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) data(s)?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from manutencao_data where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Manutenção/Controle de manutenção"
                Evento = "Excluir data"
                ID_documento = .ListItems(InitFor)
                Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo
                Documento1 = "Data da manutenção: " & .ListItems(InitFor).ListSubItems(1) & " - Hora da manutenção: " & .ListItems(InitFor).ListSubItems(2)
                ProcGravaEvento
                '==================================
                
                Conexao.Execute "DELETE from manutencao_data WHERE id = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from manutencao_defeito WHERE id_data = " & .ListItems(InitFor)
                Conexao.Execute "DELETE from manutencao_checklist where id_data = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) data(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Data(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_data
    ProcCarregaLista_data
    Novo_manutencao2 = False
    Frame10.Enabled = False
End If
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir2()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_desenho
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s) para substituição?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from manutencao_defeito where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Manutenção/Controle de manutenção"
                Evento = "Excluir produto para substituição"
                ID_documento = .ListItems(InitFor)
                Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo
                Documento1 = "Código interno: " & .ListItems(InitFor).SubItems(1)
                ProcGravaEvento
                '==================================
                
                Conexao.Execute "DELETE from manutencao_defeito WHERE id = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) para substituição antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) para substituição excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos2
    ProcCarregaLista2
    ProcValortotal
    Novo_manutencao3 = False
    Frame7.Enabled = False
End If
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir3()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_check
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) descrição(ões) do check-list?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from manutencao_checklist where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Manutenção/Controle de manutenção"
                Evento = "Excluir descrição do check-list"
                ID_documento = .ListItems(InitFor)
                Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo
                Documento1 = "Descrição: " & .ListItems(InitFor).SubItems(1)
                ProcGravaEvento
                '==================================
                
                Conexao.Execute "DELETE from manutencao_checklist WHERE id = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) descrição(ões) do check-list antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Descrição(ões) do check-list excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos3
    ProcCarregaLista3
    ProcValortotal
    Novo_manutencao4 = False
    Frame8.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
  
frmManutencao_MenuImpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir_Data()
On Error GoTo tratar_erro
  
If txtIDData = 0 Then
    USMsgBox ("Informe a data da manutenção antes de visualizar impressão"), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "Manutencao.rpt"
ProcImprimirRel "{Manutencao.codigo} = " & txtId & " and {Manutencao_data.ID} = " & txtIDData, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarPedido_Click()
On Error GoTo tratar_erro

If txttipo = "Corretiva" Then frmManutencao_Pedido.Show 1
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdMáquina_Click()
On Error GoTo tratar_erro

If optPosto.Value = True Then
frmManutencao_maquina_Abrir.Show 1
End If

If optProduto.Value = True Then
frmManutencao_item.Show 1
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
ProcDesabilitar
Novo_manutencao = True
Frame2.Enabled = True
Frame6.Enabled = True
txtResponsavel = pubUsuario
txtData = Format(Date, "dd/mm/yy")
ProcLimparTudo
frmManutencaoOpcoes.Show 1
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_data()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos_data
Novo_manutencao2 = True
Frame10.Enabled = True
txtData_Manutencao.SetFocus
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo2()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If cmbStatus = "Concluída" Then
    USMsgBox ("Não é permitido criar novo produto para substituição, pois a manutenção está concluída."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos2
Novo_manutencao3 = True
Frame7.Enabled = True
frmManutencao_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcHabilitarPrevCorr()
On Error GoTo tratar_erro

With txtRequisitante
    .Locked = True
    .TabStop = False
End With
With cmbSetor_Requisitante
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcHabilitarSolicitacao()
On Error GoTo tratar_erro

With txtRequisitante
    .Locked = False
    .TabStop = True
End With
With cmbSetor_Requisitante
    .Locked = False
    .TabStop = True
End With
'Frame_data_sol.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo3()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If cmbStatus = "Concluída" Then
    USMsgBox ("Não é permitido criar nova descrição do check-list, pois a manutenção está concluída."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos3
Novo_manutencao4 = True
Frame8.Enabled = True
txtDescricao_Check.SetFocus
  
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

If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If

If txtIDmaquina.Text = "" And optPosto.Value = True Then
    USMsgBox ("Informe o equipamento antes de salvar."), vbExclamation, "CAPRIND v5.0"
    frmManutencao_maquina_Abrir.Show 1
    Exit Sub
End If

If txtIDmaquina.Text = "" And optProduto.Value = True Then
    USMsgBox ("Informe o produto antes de salvar."), vbExclamation, "CAPRIND v5.0"
    frmManutencao_maquina_Abrir.Show 1
    Exit Sub
End If

If cmbSetorPredial.Text = "" And optPredial.Value = True Then
    USMsgBox ("Informe o setor da localização antes de salvar."), vbExclamation, "CAPRIND v5.0"
   cmbLocalizacao.SetFocus
    Exit Sub
End If


If txttipo.Text = "" Then
    USMsgBox ("Informe o tipo da manutenção antes de salvar."), vbExclamation, "CAPRIND v5.0"
    frmManutencao_menu.Show 1
    Exit Sub
End If

If txttipo.Text = "Solicitação" Then
    If txtRequisitante.Text = "" Then
        USMsgBox ("Informe o requisitante antes de salvar."), vbExclamation, "CAPRIND v5.0"
        txtRequisitante.SetFocus
        Exit Sub
    End If
    If cmbSetor_Requisitante.Text = "" Then
        USMsgBox ("Informe o setor do requisitante antes de salvar."), vbExclamation, "CAPRIND v5.0"
        cmbSetor_Requisitante.SetFocus
        Exit Sub
    End If
    If txtData_Solicitacao = "__/__/____" Then
        USMsgBox ("Informe a data da solicitação antes de salvar."), vbExclamation, "CAPRIND v5.0"
        txtData_Solicitacao.SetFocus
        Exit Sub
    End If
End If
If txttipo = "Preventiva" Or txttipo = "Corretiva" Or txttipo = "Predetiva" Then
    ProcVerificaCampos
    If VerificaCampos = False Then Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from manutencao where Codigo = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    Select Case txttipo
        Case "Solicitação": textmsg = "solicitação de manutenção"
        Case "Preventiva": textmsg = "manutenção preventiva"
        Case "Corretiva": textmsg = "manutenção corretiva"
        Case "Predetiva": textmsg = "manutenção predetiva"
    End Select
    USMsgBox ("Nova " & textmsg & " cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    
    Evento = "Nova"
    TBGravar!Data = txtData
    TBGravar!Responsavel = txtResponsavel
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
Select Case txttipo
    Case "Solicitação": ProcenviadadosSolicitacao
    Case "Preventiva": Procenviadadospreventiva
    Case "Corretiva": ProcenviadadosCorretiva
    Case "Preditiva": ProcenviadadosPreditiva
End Select
If chkControlada.Value = 1 Then TBGravar!Controlada = True Else TBGravar!Controlada = False
If Manutencao_Produto = True Then TBGravar!Produto = True Else TBGravar!Produto = False
TBGravar!Lista = txtLista.Text

If optPredial.Value = False Then
TBGravar!IDMaquina = txtIDmaquina.Text
TBGravar!Descricao = txtdescricao.Text
End If

TBGravar!Eletrica = chkeletrica.Value
TBGravar!Mecanica = chkMecanica.Value
TBGravar!Hidraulica = chkHidraulica.Value
TBGravar!Predial = optPredial.Value
TBGravar!Outros = chkOutros.Value
TBGravar!Produto = optProduto.Value


TBGravar.Update
txtId = TBGravar!CODIGO
If optPredial.Value = False Then
Caption = "Manutenção - Equipamentos - (Posto de trabalho : " & TBGravar!IDMaquina & " - Descrição : " & TBGravar!Descricao & ")"
Else
Caption = "Manutenção - Solicitação - (Predial)"
End If

TBGravar.Close

If Novo_manutencao = True Then
    Sql_Manutencao_Localizar = "Select * from manutencao where Codigo = " & txtId
    FormulaRel_Manutencao = "{manutencao_data.Status} = 'Aberta' and {manutencao.Codigo} = " & txtId
    ProcCarregaLista
Else
    ProcCarregaLista
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Manutenção/Controle de manutenção"
ID_documento = txtId
If optPredial.Value = True Then
txtIDmaquina = "PREDIAL"
End If
Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo & " - Data da manutenção: " & txtData_Manutencao
Documento1 = ""
ProcGravaEvento
'==================================

Novo_manutencao = False
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_data()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame10.Enabled = False And Novo_manutencao2 = True Then
    ProcVerificaSalvar
    Exit Sub
End If

Acao = "salvar"
If txttipo = "Preventiva" And Txt_dias_proxima = "" Then
    NomeCampo = "os dias para próxima manutenção"
    ProcVerificaAcao
    Txt_dias_proxima.SetFocus
    Exit Sub
End If
If txttipo = "Preventiva" Then
    Dataini = txtData_Manutencao
    DataFim = Cmb_data_proxima
    If Dataini > DataFim Then
        USMsgBox ("A data da manutenção não pode ser maior que a data da próxima manutenção."), vbExclamation, "CAPRIND v5.0"
        txtData_Manutencao.SetFocus
        Txt_dias_proxima = ""
        Exit Sub
    End If
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from manutencao_data where idManutencao = " & txtId & " and id <> " & txtIDData & " and data = '" & Format(txtData_Manutencao, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Esta data está sendo utilizada, favor informar outra data."), vbExclamation, "CAPRIND v5.0"
    txtData.SetFocus
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close
If Txt_tecnico_responsavel = "" Then
    NomeCampo = "o técnico responsável"
    ProcVerificaAcao
    Txt_tecnico_responsavel.SetFocus
    Exit Sub
End If
If cmbStatus = "Concluída" And Txt_data_conclusao1.Visible = True And IsDate(Txt_data_conclusao1) = False Then
    NomeCampo = "a data de conclusão"
    ProcVerificaAcao
    Txt_data_conclusao1.SetFocus
    Exit Sub
End If
If cmbStatus = "Concluída" And Txt_Hora_conclusao1.Visible = True And IsDate(Txt_Hora_conclusao1) = False Then
    NomeCampo = "a hora de conclusão"
    ProcVerificaAcao
    Txt_Hora_conclusao1.SetFocus
    Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * FROM manutencao_data WHERE id = " & txtIDData, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    USMsgBox ("Nova data cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova data"
Else
    If TBGravar!IDProducao <> 0 Then
        USMsgBox ("Não é possível alterar esta data de manutenção, pois a mesma está sendo apontada no Gerprod."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    
    If txttipo = "Preventiva" Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from Manutencao_data where idManutencao = " & txtId & " order by Data", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.BOF = False Then
            TBFIltro.Find ("ID = " & txtIDData)
            TBFIltro.MoveNext
            If TBFIltro.EOF = False Then
                If TBFIltro!status = "Concluída" Then
                    USMsgBox ("Não é permitido alterar o status para aberta, pois já existe uma manutenção após a esta que está concluída."), vbExclamation, "CAPRIND v5.0"
                    TBFIltro.Close
                    Exit Sub
                End If
            End If
        End If
        TBFIltro.Close
    End If
    
    'Atualiza check list
    If cmbStatus = "Concluída" Then StatusTexto = "True" Else StatusTexto = "False"
    Conexao.Execute "Update Manutencao_Checklist Set Check_list = '" & StatusTexto & "' where ID_Data = " & txtIDData
        
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar data"
End If
TBGravar!IDmanutencao = txtId
TBGravar!Data = txtData_Manutencao & " " & txtHora_Manutencao
TBGravar!Dias_proxima = IIf(Txt_dias_proxima = "", Null, Txt_dias_proxima)
TBGravar!status = cmbStatus
If IsDate(Txt_data_conclusao1) = True And IsDate(Txt_Hora_conclusao1) = True Then
    TBGravar!data_conclusao = Txt_data_conclusao1 & " " & Txt_Hora_conclusao1
Else
    TBGravar!data_conclusao = Null
End If
TBGravar!Tecnico_responsavel = Txt_tecnico_responsavel
If Txt_fornecedor <> "" Then TBGravar!IDpedido = txtIDPedido Else TBGravar!IDpedido = Null
TBGravar!Doc_referencia = Txt_doc_ref
TBGravar!Caminho_relatorio = Txt_caminho_relatorio
TBGravar!Obs = Trim(txtObs)
TBGravar.Update
txtIDData = TBGravar!ID
TBGravar.Close

ProcCarregaLista_data
If Novo_manutencao2 = False Then
    If CodigoLista1 <> 0 And Lista_Data.ListItems.Count <> 0 Then
        Lista_Data.SelectedItem = Lista_Data.ListItems(CodigoLista1)
        Lista_Data.SetFocus
    End If
End If
Novo_manutencao2 = False
'==================================
Modulo = "Manutenção/Controle de manutenção"
ID_documento = txtIDData
Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo
Documento1 = "Data da manutenção: " & txtData_Manutencao & " - Hora da manutenção: " & txtHora_Manutencao
ProcGravaEvento
'==================================

If txttipo = "Preventiva" Then
    If cmbStatus = "Concluída" Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from Manutencao_data where idManutencao = " & txtId & " and CAST(FLOOR(CAST(Data1 as float)) AS datetime) = '" & Format(Cmb_data_proxima, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = True Then
            If USMsgBox("Deseja criar uma nova manutenção para o dia " & Format(Cmb_data_proxima, "dd/mm/yy") & "?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbYes Then
                ProcCopiaDadosData False, txtIDData, txtId, True
                USMsgBox ("Nova data cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "Manutenção/Controle de manutenção"
                ID_documento = txtIDData
                Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo
                Documento1 = "Data da manutenção: " & Cmb_data_proxima & " - Hora da manutenção: " & txtHora_Manutencao
                ProcGravaEvento
                '==================================
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "select * from manutencao_data where id  = " & IDConta, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    ProcLimpaCampos_data
                    ProcPuxadados_data
                End If
                TBAbrir.Close
            End If
        End If
        TBFIltro.Close
    Else
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from Manutencao_data where idManutencao = " & txtId & " order by Data", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.BOF = False Then
            TBFIltro.Find ("ID = " & txtIDData)
            TBFIltro.MoveNext
            Do While TBFIltro.EOF = False
                If TBFIltro!status = "Aberta" And Format(TBFIltro!Data - Txt_dias_proxima, "dd/mm/yy") = txtData And TBFIltro!Dias_proxima = Txt_dias_proxima Then
                    Conexao.Execute "DELETE from Manutencao_defeito where ID_data = " & TBFIltro!ID
                    Conexao.Execute "DELETE from Manutencao_Checklist where ID_data = " & TBFIltro!ID
                    Conexao.Execute "DELETE from manutencao_data where ID = " & TBFIltro!ID
                End If
                TBFIltro.MoveNext
            Loop
        End If
        TBFIltro.Close
    End If
    ProcCarregaLista_data
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiaDadosData(CopiarMan As Boolean, IDData As Long, IDMan As Long, CriarProxima As Boolean)
On Error GoTo tratar_erro

If CopiarMan = True Then TextoFiltro = "IDmanutencao = " & txtId & "" Else TextoFiltro = "ID = " & IDData & ""
Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "Select * from Manutencao_data where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then
    Do While TBCiclo.EOF = False
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select * from Manutencao_data", Conexao, adOpenKeyset, adLockOptimistic
        TBCarteira.AddNew
        TBCarteira!IDmanutencao = IDMan
        TBCarteira!Data = IIf(CriarProxima = False, TBCiclo!Data, TBCiclo!Data + TBCiclo!Dias_proxima)
        TBCarteira!status = "Aberta"
        TBCarteira!Obs = TBCiclo!Obs
        TBCarteira!IDProducao = TBCiclo!IDProducao
        TBCarteira!IDProducao2 = TBCiclo!IDProducao2
        TBCarteira!Dias_proxima = TBCiclo!Dias_proxima
        TBCarteira!data_conclusao = Null
        TBCarteira!IDpedido = TBCiclo!IDpedido
        TBCarteira!Tecnico_responsavel = TBCiclo!Tecnico_responsavel
        TBCarteira!Doc_referencia = TBCiclo!Doc_referencia
        TBCarteira!Caminho_relatorio = TBCiclo!Caminho_relatorio
        TBCarteira!Solicitacao = TBCiclo!Solicitacao
        TBCarteira.Update
        
        ProcCopiaDadosDefeito IDData
        ProcCopiaDadosCheckList IDData
        TBCiclo.MoveNext
    Loop
End If
TBCiclo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiaDadosDefeito(IDData As Long)
On Error GoTo tratar_erro

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from Manutencao_defeito where ID_data = " & IDData, Conexao, adOpenKeyset, adLockOptimistic
Do While TBAliquota.EOF = False
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select * from Manutencao_defeito", Conexao, adOpenKeyset, adLockOptimistic
    TBCFOP.AddNew
    TBCFOP!ID_data = TBCarteira!ID
    TBCFOP!IDmanutencao = TBCarteira!IDmanutencao
    TBCFOP!Desenho = TBAliquota!Desenho
    TBCFOP!Qtde = TBAliquota!Qtde
    TBCFOP!VlrUnit = TBAliquota!VlrUnit
    TBCFOP!vlrTotal = TBAliquota!vlrTotal
    TBCFOP!Descricao = TBAliquota!Descricao
    TBCFOP!Familia = TBAliquota!Familia
    TBCFOP!Unidade = TBAliquota!Unidade
    TBCFOP!Unidade_com = TBAliquota!Unidade_com
    TBCFOP.Update
    TBCFOP.Close
    TBAliquota.MoveNext
Loop
TBAliquota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiaDadosCheckList(IDData As Long)
On Error GoTo tratar_erro

Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select * from Manutencao_Checklist where ID_Data = " & IDData, Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    Do While TBAliquota.EOF = False
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select * from Manutencao_Checklist", Conexao, adOpenKeyset, adLockOptimistic
        TBCFOP.AddNew
        TBCFOP!ID_data = TBCarteira!ID
        TBCFOP!ID_manutencao = TBCarteira!IDmanutencao
        TBCFOP!Descricao = IIf(IsNull(TBAliquota!Descricao), "", TBAliquota!Descricao)
        TBCFOP!valor = IIf(IsNull(TBAliquota!valor), 0, TBAliquota!valor)
        TBCFOP!Check_list = False
        TBCFOP.Update
        TBCFOP.Close
        TBAliquota.MoveNext
    Loop
End If
TBAliquota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar2()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame7.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If chkAuto.Value = 0 And txtdesenho.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdDesenho_Click
    Exit Sub
End If
If txtQuantidade.Text = "" Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQuantidade.SetFocus
    Exit Sub
End If
If txtvalorunitario.Text = "" Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    txtvalorunitario.SetFocus
    Exit Sub
End If
If txtdesc_desenho.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdesc_desenho.SetFocus
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
If cmbfamilia.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
If chkAuto.Value = 1 Then ProcNovoProdutoAuto
If chkManual.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtdesenho.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtN_Estoque.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoManual
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * FROM manutencao_defeito WHERE id = " & txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    USMsgBox ("Novo produto para subistituição cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto para subistituição"
Else
    If cmbStatus = "Concluída" Then
        USMsgBox ("Não é permitido alterar este produto para substituição, pois a manutenção está concluída."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto para subistituição"
End If
TBGravar!IDmanutencao = txtId
TBGravar!ID_data = txtIDData
TBGravar!Desenho = txtdesenho
TBGravar!Qtde = IIf(txtQuantidade = "", Null, txtQuantidade)
TBGravar!VlrUnit = IIf(txtvalorunitario = "", Null, txtvalorunitario)
TBGravar!vlrTotal = IIf(txtvalor_total = "", Null, txtvalor_total)
TBGravar!Descricao = txtdesc_desenho
TBGravar!Unidade = cmbun
TBGravar!Unidade_com = Cmb_un_com
TBGravar!Familia = cmbfamilia
TBGravar.Update
txtidproduto = TBGravar!ID
TBGravar.Close
ProcCarregaLista2
Novo_manutencao3 = False
ProcValortotal
'==================================
Modulo = "Manutenção/Controle de manutenção"
ID_documento = txtidproduto
Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo & " - Data da manutenção: " & txtData_Manutencao
Documento1 = "Código interno: " & txtdesenho
ProcGravaEvento
'==================================
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar3()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame8.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtDescricao_Check = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescricao_Check.SetFocus
    Exit Sub
End If
If txtValorCheck = "" Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    txtValorCheck.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "SELECT * from manutencao_checklist WHERE id = " & txtID_check, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    USMsgBox ("Nova descrição cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova descrição do check-list"
Else
    If cmbStatus = "Concluída" Then
        USMsgBox ("Não é permitido alterar a descrição do check-list, pois a manutenção está concluída."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar descrição do check-list"
    If Lista_check.SelectedItem.Checked = True Then TBGravar!Check_list = True Else TBGravar!Check_list = False
End If
TBGravar!ID_manutencao = txtId
TBGravar!ID_data = txtIDData
TBGravar!Descricao = txtDescricao_Check
TBGravar!valor = txtValorCheck
TBGravar.Update
txtID_check = TBGravar!ID
TBGravar.Close
ProcCarregaLista3
Novo_manutencao4 = False
ProcValortotal
'==================================
Modulo = "Manutenção/Controle de manutenção"
ID_documento = txtID_check
Documento = "Equipamento : " & txtIDmaquina & " - Tipo da manutenção: " & txttipo & " - Data da manutenção: " & txtData_Manutencao
Documento1 = "Descrição: " & txtDescricao_Check
ProcGravaEvento
'==================================
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSolicitacao()
On Error GoTo tratar_erro

If txtSolicitacao <> "" Then
    With frmCompras_Requisicao
        .ProcLimpaCampos
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from compras_requisicao where Requisicaotexto = '" & txtSolicitacao & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = False Then
            ProcCarregaComboEmpresa .Cmb_empresa, False
            If .PBLista.Value = 0 Then .PBLista = 100
            .ProcAbrir
        End If
    End With
Else
    frmManutencao_empresa.Show 1
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
            Case vbKeyInsert: ProcNovo
            Case vbKeyEscape: ProcSair
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCopiar
            Case vbKeyF8: ProcAgenda
        End Select
    Case 1:
        Select Case SSTab2.Tab
            Case 0:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo_data
                    Case vbKeyEscape: ProcSair
                    Case vbKeyF3: ProcSalvar_data
                    Case vbKeyF4: ProcExcluir_data
                    Case vbKeyF5: ProcImprimir_Data
                    Case vbKeyF7: ProcCopiar2
                End Select
            Case 1:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo2
                    Case vbKeyEscape: ProcSair
                    Case vbKeyF3: procSalvar2
                    Case vbKeyF4: procExcluir2
                    Case vbKeyF6: ProcSolicitacao
                End Select
            Case 2:
                Select Case KeyCode
                    Case vbKeyInsert: ProcNovo3
                    Case vbKeyEscape: ProcSair
                    Case vbKeyF3: procSalvar3
                    Case vbKeyF4: procExcluir3
                End Select
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15500, 14, True
ProcCarregaToolBar2 Me, 15500, 11, True
ProcCarregaToolBar3 Me, 15500, 8, True
ProcCarregaToolBar4 Me, 15500, 8, True
Formulario = "Manutenção/Equipamentos"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
ProcCarregaFamiliaUN
txtData_Manutencao.Value = Date
Cmb_opcao_lista = "Excluir"

Set TBSolicitacao = CreateObject("adodb.recordset")
TBSolicitacao.Open "Select * from manutencao where Data_Solicitacao = '" & Date & "' and tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
If TBSolicitacao.EOF = False Then
    USMsgBox ("Existe(m) solicitação(ões) em aberto, favor consultar a agenda de solicitações."), vbExclamation, "CAPRIND v5.0"
End If
TBSolicitacao.Close

ProcRemoveObjetosResize Me
ProcCarregaComboSetor
ProcCarregaComboLocalizacao
ProcCarregaComboSetorRequisitante


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
cmbfamilia.ListIndex = -1
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from manutencao_defeito where id  = " & txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Unidade) = False And TBAbrir!Unidade <> "" Then cmbun = TBAbrir!Unidade
    If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
    If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Manutenção/Equipamentos"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaFamiliaUN

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmManutencao_abrir.Show 1
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAgenda()
On Error GoTo tratar_erro

frmManutencao_Agenda.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro
'Debug.print Sql_Manutencao_Localizar
Lista.ListItems.Clear
If Sql_Manutencao_Localizar = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open Sql_Manutencao_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        Tipo_manutencao = ""
        With Lista.ListItems
            .Add , , TBLISTA!CODIGO
            If IsNull(TBLISTA!CodSol) = False Or IsNull(TBLISTA!codman) = False Then
            .Item(.Count).SubItems(1) = IIf(TBLISTA!Tipo = "S", TBLISTA!CodSol, TBLISTA!codman)
            End If
            If TBLISTA!Predial = False Then
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!IDMaquina), "", TBLISTA!IDMaquina)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            Else
             .Item(.Count).SubItems(2) = "SMP"
            .Item(.Count).SubItems(3) = "PREDIAL"
            End If
            If IsNull(TBLISTA!Tipo) = False Then
                If TBLISTA!Tipo = "S" Then Tipo_manutencao = "Solicitação"
                If TBLISTA!Tipo = "C" Then Tipo_manutencao = "Corretiva"
                If TBLISTA!Tipo = "P" Then Tipo_manutencao = "Preventiva"
                If TBLISTA!Tipo = "PR" Then Tipo_manutencao = "Preditiva"
            End If
            .Item(.Count).SubItems(4) = Tipo_manutencao
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yyyy"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Data_Solicitacao), "", Format(TBLISTA!Data_Solicitacao, "dd/mm/yyyy"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Requisitante), "", TBLISTA!Requisitante)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!setor_requisitante), "", TBLISTA!setor_requisitante)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Aprovado), "", TBLISTA!Aprovado)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Setor_Aprovado), "", TBLISTA!Setor_Aprovado)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista2()
On Error GoTo tratar_erro

Lista_desenho.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao_defeito where id_data = " & txtIDData, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_desenho.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "select * from projproduto where desenho = '" & TBLISTA!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
            Else
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            End If
            TBItem.Close
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Qtde), "", Format(TBLISTA!Qtde, "###,##0.0000"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!VlrUnit), "", Format(TBLISTA!VlrUnit, "###,##0.0000000000"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!vlrTotal), "", Format(TBLISTA!vlrTotal, "###,##0.00"))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista_data()
On Error GoTo tratar_erro

Lista_Data.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao_data where idManutencao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_Data.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "hh:mm:ss"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Dias_proxima), "", TBLISTA!Dias_proxima)
            If IsNull(TBLISTA!Data) = False And TBLISTA!Data <> "" Then
                If IsNull(TBLISTA!Dias_proxima) = False And TBLISTA!Dias_proxima <> "" Then
                    Data_manutencao = Format(TBLISTA!Data, "dd/mm/yy")
                    Data_manutencao = Data_manutencao + TBLISTA!Dias_proxima
                    .Item(.Count).SubItems(4) = Format(Data_manutencao, "dd/mm/yy")
                End If
            End If
            
            Qtde = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(VlrTotal) as Qtde from manutencao_defeito where id_data = " & TBLISTA!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
            End If
                        
            Qtd = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Valor) as Qtd from Manutencao_Checklist where id_data = " & TBLISTA!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtd = IIf(IsNull(TBAbrir!Qtd), 0, TBAbrir!Qtd)
            End If
            TBAbrir.Close
            
            .Item(.Count).SubItems(5) = Format(Qtde, "###,##0.00")
            .Item(.Count).SubItems(6) = Format(Qtd, "###,##0.00")
            .Item(.Count).SubItems(7) = Format(Qtde + Qtd, "###,##0.00")
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!status), "", TBLISTA!status)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!data_conclusao), "", Format(TBLISTA!data_conclusao, "dd/mm/yy"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!data_conclusao), "", Format(TBLISTA!data_conclusao, "hh:mm:ss"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Doc_referencia), "", TBLISTA!Doc_referencia)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos2()
On Error GoTo tratar_erro

txtidproduto = 0
txtdesenho = ""
txtdesc_desenho = ""
txtQuantidade = ""
txtvalorunitario = ""
txtvalor_total = ""
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
cmbfamilia.ListIndex = -1
chkAuto.Value = 0
chkManual.Value = 0
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Img_data_conclusao_Click()
On Error GoTo tratar_erro

If Txt_data_conclusao1.Enabled = False Then Exit Sub
Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = True
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
Estoque_recebimento = False
Sit_REG = 2
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgSolicitacao_Click()
On Error GoTo tratar_erro

If txttipo <> "Solicitação" Then Exit Sub
Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = True
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
Estoque_recebimento = False
Sit_REG = 1
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_check_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_check
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    If cmbStatus = "Aberta" Then .ListItems.Item(InitFor).Checked = True
                Else
                    .ListItems.Item(InitFor).Checked = True
                End If
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_check, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_check_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_check
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir" And cmbStatus = "Concluída" Then
                USMsgBox ("Não é permitido excluir esta descrição do check-list, pois a manutenção está concluída."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_check_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_check.ListItems.Count = 0 Then Exit Sub
Frame8.Enabled = True
Novo_manutencao4 = False
txtID_check = Lista_check.SelectedItem
txtDescricao_Check = Lista_check.SelectedItem.SubItems(1)
txtValorCheck = Lista_check.SelectedItem.SubItems(2)

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
                ProcVerificaRegistroUtilizadoSemMsg "manutencao_data", "idManutencao = " & .ListItems(InitFor) & " and idproducao <> 0"
                If Permitido = False Then GoTo Proximo
                
                ProcVerificaRegistroUtilizadoSemMsg "manutencao_data", "idManutencao = " & .ListItems(InitFor) & " and status = 'Concluída'"
                If Permitido = False Then GoTo Proximo
                
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

Private Sub Lista_Data_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_Data
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                ProcVerificaRegistroUtilizadoSemMsg "manutencao_data", "id = " & .ListItems(InitFor) & " and idproducao <> 0"
                If Permitido = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_Data, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Data_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Mensagem = "Não é possível excluir esta data, pois a mesma está sendo apontada no"
            ProcVerificaRegistroUtilizado "manutencao_data", "id = " & .ListItems(InitFor) & " and idproducao <> 0", "Gerprod"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Data_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_Data.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from manutencao_data where id  = " & Lista_Data.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_data
    ProcPuxadados_data
    CodigoLista1 = Lista_Data.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_desenho_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_desenho
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If cmbStatus = "Aberta" Then .ListItems.Item(InitFor).Checked = True
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

Private Sub Lista_desenho_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_desenho
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If cmbStatus = "Concluída" Then
                USMsgBox ("Não é permitido excluir este produto para substituição, pois a manutenção está concluída."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_desenho_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_desenho.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from manutencao_defeito where id  = " & Lista_desenho.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos2
    ProcPuxadados2
End If
TBAbrir.Close

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
            Mensagem = "Não é possível excluir esta manutenção, pois a mesma está sendo apontada no"
            ProcVerificaRegistroUtilizado "manutencao_data", "idManutencao = " & .ListItems(InitFor) & " and idproducao <> 0", "Gerprod"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            Mensagem = "Não é possível excluir esta manutenção, pois a mesma possui datas de manutenção"
            ProcVerificaRegistroUtilizado "manutencao_data", "idManutencao = " & .ListItems(InitFor) & " and status = 'Concluída'", "concluídas"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
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
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from manutencao where Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcDesabilitar
    ProcLimpaCampos
    ProcPuxaDados
    Novo_manutencao = False
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPosto_Click()
On Error GoTo tratar_erro

If optPosto.Value = True Then
    Frame2.Visible = True
    Frame12.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPredial_Click()
On Error GoTo tratar_erro

If optPredial.Value = True Then
    Frame2.Visible = False
    Frame12.Visible = True
Else
    Frame2.Visible = True
    Frame12.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optProduto_Click()
On Error GoTo tratar_erro

If optProduto.Value = True Then
    Frame2.Visible = True
    Frame12.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = "0" Or txttipo = "Solicitação" Then
    SSTab1.Tab = 0
    Exit Sub
End If

Cmb_opcao_lista.Visible = False
Label38.Visible = False
Select Case SSTab1.Tab
    Case 0:
        SSTab2.Visible = False
        Lista.Visible = True
        Frame10.Visible = False
        Lista_Data.Visible = False
        Lista_desenho.Visible = False
        Lista_check.Visible = False
        Lista.SetFocus
    Case 1:
        SSTab2.Visible = True
        Lista.Visible = False
        Frame10.Visible = True
        Lista_Data.Visible = True
        Lista_desenho.Visible = False
        Lista_check.Visible = False
        ProcCarregaLista_data
        SSTab2.Tab = 0
        If chkControlada.Value = 1 Then
            Txt_data_conclusao.Visible = True
            Txt_data_conclusao1.Visible = False
            Txt_Hora_conclusao.Visible = True
            Txt_Hora_conclusao1.Visible = False
            Img_data_conclusao.Visible = False
        Else
            Txt_data_conclusao.Visible = False
            Txt_data_conclusao1.Visible = True
            Txt_Hora_conclusao.Visible = False
            Txt_Hora_conclusao1.Visible = True
            Img_data_conclusao.Visible = True
        End If
        Lista_Data.SetFocus
        ProcLimpaCampos_data
        txtObs = ""
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = "0"
txtIDmaquina = ""
txtdescricao = ""
txtData = ""
'txttipo.Text = ""
txtRequisitante.Text = ""
cmbSetor_Requisitante.Text = ""
txtData_Solicitacao.Text = "__/__/____"
txtAprovado.Text = ""
txtSetor_Aprovado.Text = ""
txtLista.Text = ""
chkControlada.Value = 0
txtTotalSub = "0,00"
txtTotalCheck = "0,00"
txtTotal = "0,00"
CodigoLista = 0
Manutencao_Produto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_data()
On Error GoTo tratar_erro

txtIDData = 0
txtData_Manutencao.Value = Date
txtHora_Manutencao = Format(Now, "hh:mm:ss")
Txt_dias_proxima = 0
Cmb_data_proxima.Value = Date
cmbStatus = "Aberta"
Txt_data_conclusao = ""
Txt_data_conclusao1 = "__/__/____"
Txt_Hora_conclusao = ""
Txt_Hora_conclusao1 = "__:__:__"
Txt_tecnico_responsavel = ""
txtIDPedido = 0
txtPedido = ""
Txt_fornecedor = ""
Txt_doc_ref = ""
CodigoLista = 0
txtObs = ""
Txt_caminho_relatorio = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procenviadadospreventiva()
On Error GoTo tratar_erro

TBGravar!Tipo = "P"
TBGravar!Aprovado = txtAprovado.Text
TBGravar!Setor_Aprovado = txtSetor_Aprovado.Text

If txtCodigo.Text = "" Then
ProcCriaCodigoManutencao
End If

TBGravar!codman = txtCodigo.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriaCodigoManutencao()
On Error GoTo tratar_erro
Dim CodigoMan As String
Var = "S"

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao where Tipo <> '" & Var & "' order by CodMan", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
TBLISTA.MoveLast

   If TBLISTA!codman <> "" Then
   CodigoMan = TBLISTA!codman
   CodigoMan = Right(CodigoMan, 9)
   CodigoMan = Left(CodigoMan, 6)
   CodigoMan = Int(CodigoMan) + 1
   Else
   CodigoMan = 1
   End If
    Select Case Len(CodigoMan)
        Case 1: CodigoMan = "00000" & CodigoMan
        Case 2: CodigoMan = "0000" & CodigoMan
        Case 3: CodigoMan = "000" & CodigoMan
        Case 4: CodigoMan = "00" & CodigoMan
        Case 5: CodigoMan = "0" & CodigoMan
    End Select
    Ano = Right(Year(Date), 2)
CodigoMan = "MAN-" & CodigoMan & "/" & Right(Year(Date), 2)
Else
    CodigoMan = "MAN-000001" & "/" & Right(Year(Date), 2)
End If
TBLISTA.Close
frmManutencao.txtCodigo.Text = CodigoMan

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcenviadadosCorretiva()
On Error GoTo tratar_erro

TBGravar!Tipo = "C"
TBGravar!Data_Solicitacao = IIf(txtData_Solicitacao = "__/__/____", Null, txtData_Solicitacao)
TBGravar!Requisitante = txtRequisitante.Text
TBGravar!setor_requisitante = cmbSetor_Requisitante.Text
TBGravar!Setor_Predial = cmbSetorPredial.Text
TBGravar!Alocado = cmbLocalizacao.Text
TBGravar!Aprovado = txtAprovado.Text
TBGravar!Setor_Aprovado = txtSetor_Aprovado.Text

If txtCodigo.Text = "" Then
ProcCriaCodigoManutencao
End If

TBGravar!codman = txtCodigo.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriaCodigoSolicitacao()
On Error GoTo tratar_erro
Dim CodigoSol As String
Var = "S"

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao order by CodSOL", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
TBLISTA.MoveLast

CodSol = TBLISTA!CodSol
CodSol = Right(CodSol, 9)
CodSol = Left(CodSol, 6)

CodSol = Int(CodSol) + 1
    Select Case Len(CodSol)
        Case 1: CodigoSol = "00000" & CodSol
        Case 2: CodigoSol = "0000" & CodSol
        Case 3: CodigoSol = "000" & CodSol
        Case 4: CodigoSol = "00" & CodSol
        Case 5: CodigoSol = "0" & CodSol
    End Select
    Ano = Right(Year(Date), 2)
CodigoSol = "SOL-" & CodigoSol & "/" & Right(Year(Date), 2)
Else
    CodigoSol = "SOL-000001" & "/" & Right(Year(Date), 2)
End If
TBLISTA.Close
frmManutencao.txtCodigo.Text = CodigoSol

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcenviadadosSolicitacao()
On Error GoTo tratar_erro
    
TBGravar!Tipo = "S"
TBGravar!Requisitante = txtRequisitante.Text
TBGravar!setor_requisitante = cmbSetor_Requisitante.Text
TBGravar!Setor_Predial = cmbSetorPredial.Text
TBGravar!Alocado = cmbLocalizacao.Text
TBGravar!Data_Solicitacao = Format(txtData_Solicitacao.Text, "dd/MM/yy")

If txtCodigo.Text = "" Then
ProcCriaCodigoSolicitacao
End If


TBGravar!CodSol = txtCodigo.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcenviadadosPreditiva()
On Error GoTo tratar_erro

TBGravar!Tipo = "PR"
TBGravar!Data_Solicitacao = IIf(txtData_Solicitacao = "__/__/____", Null, txtData_Solicitacao)
TBGravar!Requisitante = txtRequisitante.Text
TBGravar!setor_requisitante = cmbSetor_Requisitante.Text
TBGravar!Aprovado = txtAprovado.Text
TBGravar!Setor_Aprovado = txtSetor_Aprovado.Text

If txtCodigo.Text = "" Then
ProcCriaCodigoManutencao
End If


TBGravar!codman = txtCodigo.Text


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtSolicitacao = ""
Select Case TBAbrir!Tipo
    Case "S":
        txttipo.Text = "Solicitação"
        txtRequisitante.Text = IIf(IsNull(TBAbrir!Requisitante), "", TBAbrir!Requisitante)
        cmbSetor_Requisitante.Text = IIf(IsNull(TBAbrir!setor_requisitante), "", TBAbrir!setor_requisitante)
        cmbSetorPredial.Text = IIf(IsNull(TBAbrir!Setor_Predial), "", TBAbrir!Setor_Predial)
        cmbLocalizacao.Text = IIf(IsNull(TBAbrir!Alocado), "", TBAbrir!Alocado)
        txtData_Solicitacao.Text = IIf(IsNull(TBAbrir!Data_Solicitacao), "__/__/____", Format(TBAbrir!Data_Solicitacao, "dd/mm/yyyy"))
        txtCodigo.Text = IIf(IsNull(TBAbrir!CodSol), "", TBAbrir!CodSol)
        ProcHabilitarSolicitacao
    Case "P":
        txttipo.Text = "Preventiva"
        txtAprovado.Text = IIf(IsNull(TBAbrir!Aprovado), "", TBAbrir!Aprovado)
        txtSetor_Aprovado.Text = IIf(IsNull(TBAbrir!Setor_Aprovado), "", TBAbrir!Setor_Aprovado)
        txtCodigo.Text = IIf(IsNull(TBAbrir!codman), "", TBAbrir!codman)

        ProcHabilitarPrevCorr
    Case "C":
        txttipo.Text = "Corretiva"
        txtRequisitante.Text = IIf(IsNull(TBAbrir!Requisitante), "", TBAbrir!Requisitante)
        cmbSetor_Requisitante.Text = IIf(IsNull(TBAbrir!setor_requisitante), "", TBAbrir!setor_requisitante)
        txtData_Solicitacao.Text = IIf(IsNull(TBAbrir!Data_Solicitacao), "__/__/____", Format(TBAbrir!Data_Solicitacao, "dd/mm/yyyy"))
        txtAprovado.Text = IIf(IsNull(TBAbrir!Aprovado), "", TBAbrir!Aprovado)
        txtSetor_Aprovado.Text = IIf(IsNull(TBAbrir!Setor_Aprovado), "", TBAbrir!Setor_Aprovado)
        txtCodigo.Text = IIf(IsNull(TBAbrir!codman), "", TBAbrir!codman)
        
        ProcHabilitarPrevCorr
    Case "PR"
        txttipo.Text = "Preditiva"
        txtAprovado.Text = IIf(IsNull(TBAbrir!Aprovado), "", TBAbrir!Aprovado)
        txtSetor_Aprovado.Text = IIf(IsNull(TBAbrir!Setor_Aprovado), "", TBAbrir!Setor_Aprovado)
        txtCodigo.Text = IIf(IsNull(TBAbrir!codman), "", TBAbrir!codman)
        
        ProcHabilitarPrevCorr
End Select
ProcVerifTipoManut

txtId = TBAbrir!CODIGO
txtIDmaquina = IIf(IsNull(TBAbrir!IDMaquina), "", TBAbrir!IDMaquina)
txtdescricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)

Caption = "Manutenção - Equipamentos - (Posto de trabalho : " & TBAbrir!IDMaquina & " - Descrição : " & TBAbrir!Descricao & ")"

txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)

txtLista = IIf(IsNull(TBAbrir!Lista), "", TBAbrir!Lista)
If TBAbrir!Controlada = True Then chkControlada.Value = 1 Else chkControlada.Value = 0
If TBAbrir!Produto = True Then Manutencao_Produto = True Else Manutencao_Produto = False

chkeletrica.Value = IIf(TBAbrir!Eletrica = True, "1", "0")
chkMecanica.Value = IIf(TBAbrir!Mecanica = True, "1", "0")
chkHidraulica.Value = IIf(TBAbrir!Hidraulica = True, "1", "0")
optPredial.Value = IIf(TBAbrir!Predial = True, True, False)
optProduto.Value = IIf(TBAbrir!Produto = True, True, False)
If optPredial.Value = False And optProduto.Value = False Then
optPosto.Value = True
End If

chkOutros.Value = IIf(TBAbrir!Outros = True, "1", "0")

'Set TBFIltro = CreateObject("adodb.recordset")
'TBFIltro.Open "select * from Manutencao_data where idManutencao = " & txtid & " and idproducao <> 0", Conexao, adOpenKeyset, adLockOptimistic
'If TBFIltro.EOF = False Then frameControla.Enabled = False Else frameControla.Enabled = True
'TBFIltro.Close

ProcLimparTudo

Frame6.Enabled = True
Frame2.Enabled = True
ProcValortotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados2()
On Error GoTo tratar_erro

txtidproduto = TBAbrir!ID
txtdesenho.Text = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
If IsNull(TBAbrir!Unidade) = False And TBAbrir!Unidade <> "" Then cmbun = TBAbrir!Unidade
If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
txtDescricao2 = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
txtQuantidade.Text = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
txtvalorunitario.Text = IIf(IsNull(TBAbrir!VlrUnit), "", Format(TBAbrir!VlrUnit, "###,##0.0000000000"))
txtvalor_total.Text = IIf(IsNull(TBAbrir!vlrTotal), "", Format(TBAbrir!vlrTotal, "###,##0.00"))
Novo_manutencao3 = False
Frame7.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_data()
On Error GoTo tratar_erro

txtIDData = TBAbrir!ID
txtData_Manutencao = IIf(IsNull(TBAbrir!Data), Date, Format(TBAbrir!Data, "dd/mm/yyyy"))
txtHora_Manutencao = IIf(IsNull(TBAbrir!Data), "__:__:__", Format(TBAbrir!Data, "hh:mm:ss"))
Txt_dias_proxima = IIf(IsNull(TBAbrir!Dias_proxima), "", TBAbrir!Dias_proxima)
If Txt_dias_proxima = "0" Then Cmb_data_proxima = txtData_Manutencao
If IsNull(TBAbrir!status) = False And TBAbrir!status <> "" Then cmbStatus = TBAbrir!status

If chkControlada.Value = 1 Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select data from producaofases where IDProducao = " & IIf(IsNull(TBAbrir!IDProducao2), 0, TBAbrir!IDProducao2), Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        Txt_data_conclusao = Format(TBproducao!Data, "dd/mm/yy")
        Txt_Hora_conclusao = Format(TBproducao!TempoFinal, "hh:mm:ss")
    Else
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select data from producaofases_Backup where IDProducao = " & IIf(IsNull(TBAbrir!IDProducao2), 0, TBAbrir!IDProducao2), Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            Txt_data_conclusao = Format(TBproducao!Data, "dd/mm/yy")
            Txt_Hora_conclusao = Format(TBproducao!TempoFinal, "hh:mm:ss")
        End If
    End If
    TBproducao.Close
Else
    Txt_data_conclusao1 = IIf(IsNull(TBAbrir!data_conclusao), "__/__/____", Format(TBAbrir!data_conclusao, "dd/mm/yyyy"))
    Txt_Hora_conclusao1 = IIf(IsNull(TBAbrir!data_conclusao), "__:__:__", Format(TBAbrir!data_conclusao, "hh:mm:ss"))
End If
Txt_tecnico_responsavel = IIf(IsNull(TBAbrir!Tecnico_responsavel), "", TBAbrir!Tecnico_responsavel)

If IsNull(TBAbrir!IDpedido) = False Then
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "select * from compras_pedido where IDpedido = " & TBAbrir!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        txtIDPedido = TBCFOP!IDpedido
        txtPedido = IIf(IsNull(TBCFOP!Pedido), "", TBCFOP!Pedido)
        Txt_fornecedor = IIf(IsNull(TBCFOP!Fornecedor), "", TBCFOP!Fornecedor)
    End If
    TBCFOP.Close
End If

Txt_doc_ref = IIf(IsNull(TBAbrir!Doc_referencia), "", TBAbrir!Doc_referencia)
Txt_caminho_relatorio = IIf(IsNull(TBAbrir!Caminho_relatorio), "", TBAbrir!Caminho_relatorio)
txtObs = IIf(IsNull(TBAbrir!Obs), "", Trim(TBAbrir!Obs))
txtSolicitacao = IIf(IsNull(TBAbrir!Solicitacao), "", TBAbrir!Solicitacao)
If TBAbrir!IDProducao = 0 Then Frame10.Enabled = True Else Frame10.Enabled = False
Novo_manutencao3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtIDData = 0 Then
    SSTab2.Tab = 0
    Exit Sub
End If

Cmb_opcao_lista.Visible = False
Label38.Visible = False
Select Case SSTab2.Tab
    Case 0:
        Frame10.Visible = True
        Lista_Data.Visible = True
        Lista_desenho.Visible = False
        Lista_check.Visible = False
        Lista_Data.SetFocus
        ProcCarregaLista_data
        Novo_manutencao2 = False
        Frame10.Enabled = False
    Case 1:
        Frame10.Visible = False
        Lista_Data.Visible = False
        Lista_desenho.Visible = True
        Lista_check.Visible = False
        If Novo_manutencao = True Then
            SSTab1.Tab = 0
            USMsgBox ("Salve a manutenção antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        Lista_desenho.SetFocus
        ProcLimpaCampos2
        Novo_manutencao3 = False
        Frame7.Enabled = False
        txtIDmaquina2.Text = txtIDmaquina.Text
        txtDescricao2.Text = txtdescricao.Text
        If txttipo <> "Preventiva" Then
            Txt_dias_proxima = 0
            Txt_dias_proxima.Locked = True
            Txt_dias_proxima.TabStop = False
            Frame11.Enabled = False
        Else
            Txt_dias_proxima.Locked = False
            Txt_dias_proxima.TabStop = True
            Frame11.Enabled = True
        End If
        ProcCarregaLista2
    Case 2:
        Frame10.Visible = False
        Lista_Data.Visible = False
        Lista_desenho.Visible = False
        Lista_check.Visible = True
        Cmb_opcao_lista.Visible = True
        Label38.Visible = True
        If Novo_manutencao = True Then
            SSTab1.Tab = 0
            USMsgBox ("Salve a manutenção antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        Lista_check.SetFocus
        ProcLimpaCampos3
        Novo_manutencao4 = False
        Frame8.Enabled = False
        ProcCarregaLista3
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_dias_proxima_Change()
On Error GoTo tratar_erro

If Txt_dias_proxima <> "" Then
    VerifNumero = Txt_dias_proxima
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_dias_proxima = ""
        Txt_dias_proxima.SetFocus
        Exit Sub
    End If
    Quant = Txt_dias_proxima
    Cmb_data_proxima = txtData_Manutencao + Quant
Else
    Cmb_data_proxima = txtData_Manutencao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtData_Manutencao_Change()
On Error GoTo tratar_erro

If txttipo = "Preventiva" Then
    Dataini = txtData_Manutencao
    DataFim = Cmb_data_proxima
    Txt_dias_proxima = DataFim - Dataini
Else
    Txt_dias_proxima = 0
    Cmb_data_proxima.Value = txtData_Manutencao.Value
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

If chkManual.Value = 1 Or chkAuto.Value = 1 Then Exit Sub
With cmbun
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .ListIndex = -1
    .Locked = False
    .TabStop = True
End With
With txtdesc_desenho
    .Text = ""
    .Locked = False
    .TabStop = True
End With
If txtdesenho <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "select * from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        With cmbun
            If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then
                .Text = TBProduto!Unidade
                .Locked = True
                .TabStop = False
            End If
        End With
        With Cmb_un_com
            If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then
                .Text = TBProduto!Unidade_com
                .Locked = True
                .TabStop = False
            End If
        End With
        With cmbfamilia
            If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then
                .Text = TBProduto!Classe
                .Locked = True
                .TabStop = False
            End If
        End With
        With txtdesc_desenho
            .Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
            .Locked = True
            .TabStop = False
        End With
    End If
    TBProduto.Close
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpedido_Change()
On Error GoTo tratar_erro

txtIDPedido = 0
Txt_fornecedor = ""
If txtPedido = "" Then Exit Sub
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "select * from Compras_pedido where Pedido = '" & txtPedido & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
    txtIDPedido = TBPedido!IDpedido
    Txt_fornecedor = TBPedido!Fornecedor
End If
TBPedido.Close

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
Qtde = IIf(txtQuantidade = "", 0, txtQuantidade)
valor = IIf(txtvalorunitario = "", 0, txtvalorunitario)
txtvalor_total = Format(Qtde * valor, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQuantidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantidade_LostFocus()
On Error GoTo tratar_erro

txtQuantidade = Format(txtQuantidade, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txttipo_Change()
On Error GoTo tratar_erro

ProcVerifTipoManut

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifTipoManut()
On Error GoTo tratar_erro
Dim Tipo As String

If txttipo = "Solicitação" Then
    SSTab1.TabVisible(1) = False
    SSTab1.TabsPerRow = 1
    Tipo = "S"
Else
    SSTab1.TabVisible(1) = True
    SSTab1.TabsPerRow = 2
    Tipo = "M"
    If txttipo = "Corretiva" Then
        With txtPedido
            .Locked = False
            .TabStop = True
        End With
        cmdLocalizarPedido.Enabled = True
    Else
        With txtPedido
            .Locked = True
            .TabStop = False
        End With
        cmdLocalizarPedido.Enabled = False
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorCheck_Change()
On Error GoTo tratar_erro

If txtValorCheck <> "" Then
    VerifNumero = txtValorCheck
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValorCheck = ""
        txtValorCheck.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorCheck_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtValorCheck

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorCheck_LostFocus()
On Error GoTo tratar_erro

txtValorCheck = IIf(txtValorCheck = "", "0,00", Format(txtValorCheck, "###,##0.00"))
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitario_Change()
On Error GoTo tratar_erro

If txtvalorunitario.Text <> "" Then
    VerifNumero = txtvalorunitario.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvalorunitario.Text = ""
        txtvalorunitario.SetFocus
        Exit Sub
    End If
End If
Qtde = IIf(txtQuantidade = "", 0, txtQuantidade)
valor = IIf(txtvalorunitario = "", 0, txtvalorunitario)
txtvalor_total = Format(Qtde * valor, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitario_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalorunitario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitario_LostFocus()
On Error GoTo tratar_erro

txtvalorunitario = Format(txtvalorunitario, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaCampos()
On Error GoTo tratar_erro

VerificaCampos = True
If txtAprovado.Text = "" Then
    USMsgBox ("Informe o responsável pela aprovação antes de salvar."), vbExclamation, "CAPRIND v5.0"
    frmManutencao_aut.Show 1
    VerificaCampos = False
    Exit Sub
End If
If txtSetor_Aprovado.Text = "" Then
    USMsgBox ("Informe o setor do responsável pela aprovação antes de salvar."), vbExclamation, "CAPRIND v5.0"
    frmManutencao_aut.Show 1
    VerificaCampos = False
    Exit Sub
End If
If txtLista.Text = "" Then
    USMsgBox ("Informe a descrição da manutenção antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtLista.SetFocus
    VerificaCampos = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFamiliaUN()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamilia, "Familia <> 'Null'", False
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False
If txtidproduto <> 0 Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoAuto()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtdesenho = FunCriaNovoProdServ(True, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtdesenho, "", 0, txtdesc_desenho, txtdesc_desenho, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtdesenho = FunCriaNovoProdServ(True, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtdesenho, "", 0, txtdesc_desenho, txtdesc_desenho, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoManual()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtdesenho = FunCriaNovoProdServ(True, "", txtdesenho, "", 0, txtdesc_desenho, txtdesc_desenho, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtdesenho = FunCriaNovoProdServ(True, "", txtdesenho, "", 0, txtdesc_desenho, txtdesc_desenho, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos3()
On Error GoTo tratar_erro

txtID_check = "0"
txtDescricao_Check = ""
txtValorCheck = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame10.Enabled = False
Frame3.Enabled = False
Frame8.Enabled = False
ProcLimpaCampos_data
ProcLimpaCampos2
ProcLimpaCampos3
Novo_manutencao2 = False
Novo_manutencao3 = False
Novo_manutencao4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista3()
On Error GoTo tratar_erro

Lista_check.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao_checklist where id_data = " & txtIDData, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_check.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!valor), "0,00", Format(TBLISTA!valor, "###,##0.00"))
            If TBLISTA!Check_list = True Then
                .Item(.Count).SubItems(3) = "SIM"
                If Cmb_opcao_lista = "Check-list" Then .Item(.Count).Checked = True
            Else
                .Item(.Count).SubItems(3) = "NÃO"
                If Cmb_opcao_lista = "Check-list" Then .Item(.Count).Checked = False
            End If
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from manutencao_data where id  = " & txtIDData, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_data
    ProcPuxadados_data
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaStatusData()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao_checklist where id_data = " & txtIDData & " and Check_list = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then StatusTexto = "Status = 'Aberta', Data_conclusao = NULL" Else StatusTexto = "Status = 'Concluída', Data_conclusao = '" & Now & "'"
TBLISTA.Close

Conexao.Execute "Update Manutencao_data Set " & StatusTexto & " where ID = " & txtIDData

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcValortotal()
On Error GoTo tratar_erro

Qtde = 0
Qtd = 0
Set TBProgramas = CreateObject("adodb.recordset")
TBProgramas.Open "select * from manutencao_defeito where idmanutencao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBProgramas.EOF = False Then
    Do While TBProgramas.EOF = False
        Qtde = Qtde + IIf(IsNull(TBProgramas!vlrTotal), "0", TBProgramas!vlrTotal)
        TBProgramas.MoveNext
    Loop
End If
TBProgramas.Close
Set TBProgramas = CreateObject("adodb.recordset")
TBProgramas.Open "select * from Manutencao_Checklist where id_manutencao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBProgramas.EOF = False Then
    Do While TBProgramas.EOF = False
        Qtd = Qtd + IIf(IsNull(TBProgramas!valor), "0", TBProgramas!valor)
        TBProgramas.MoveNext
    Loop
End If
TBProgramas.Close
txtTotalCheck = Format(Qtd, "###,##0.00")
txtTotalSub = Format(Qtde, "###,##0.00")
txtTotal = Format(Qtde + Qtd, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSair()
On Error GoTo tratar_erro

If Novo_manutencao = True Then
    If USMsgBox("A manutenção ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        ProcSalvar
        If Novo_manutencao = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_manutencao2 = True Then
    If USMsgBox("A data ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        SSTab1.Tab = 1
        SSTab2.Tab = 0
        ProcSalvar_data
        If Novo_manutencao2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_manutencao3 = True Then
    If USMsgBox("O defeito ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        SSTab1.Tab = 1
        SSTab2.Tab = 1
        procSalvar2
        If Novo_manutencao2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_manutencao4 = True Then
    If USMsgBox("A descrição ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo + vbQuestion, "CAPRIND v5.0") = vbYes Then
        SSTab1.Tab = 1
        SSTab2.Tab = 2
        procSalvar3
        If Novo_manutencao4 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Unload Me
Novo_manutencao = False
Novo_manutencao2 = False
Novo_manutencao3 = False
Novo_manutencao4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from manutencao order by idmaquina,tipo,Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("codigo = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId.Text = TBLISTA!CODIGO
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from manutencao where Codigo = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCampos_data
        ProcLimpaCampos2
        ProcLimpaCampos3
        ProcPuxaDados
        ProcCarregaLista
        ProcCarregaLista_data
        ProcCarregaLista2
        ProcCarregaLista3
    Else
        USMsgBox ("Fim dos cadastros de manutenção."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_manutencao = False
Novo_manutencao2 = False
Novo_manutencao3 = False
Novo_manutencao4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from manutencao order by idmaquina,tipo,Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("codigo = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId.Text = TBLISTA!CODIGO
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from manutencao where Codigo = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCampos_data
        ProcLimpaCampos2
        ProcLimpaCampos3
        ProcPuxaDados
        ProcCarregaLista
        ProcCarregaLista_data
        ProcCarregaLista2
        ProcCarregaLista3
    Else
        USMsgBox ("Fim dos cadastros de manutenção."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_manutencao = False
Novo_manutencao2 = False
Novo_manutencao3 = False
Novo_manutencao4 = False

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
    Case 8: procFiltrarTodos
    Case 9: ProcCopiar
    Case 10: ProcAgenda
    'Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_data
    Case 2: ProcSalvar_data
    Case 3: ProcExcluir_data
    Case 4: ProcImprimir_Data
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcCopiar2
    'Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo2
    Case 2: procSalvar2
    Case 3: procExcluir2
    Case 4: ProcSolicitacao
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo3
    Case 2: procSalvar3
    Case 3: procExcluir3
    Case 4: ProcCheck
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCopiar()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_manutencao = True Then
    USMsgBox ("Salve o registro antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDmaquina = "" Then
    NomeCampo = "o registro"
    Acao = "copiar"
    ProcVerificaAcao
    Exit Sub
End If
frmManutencao_copiar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCopiar2()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar esta data da manutenção?", vbQuestion + vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
If Novo_manutencao2 = True Then
    USMsgBox ("Salve o registro antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDData = "" Or txtIDData = "0" Then
    NomeCampo = "o registro"
    Acao = "copiar"
    ProcVerificaAcao
    Exit Sub
End If
ProcCopiaDadosData False, txtIDData, txtId, False

'==================================
Modulo = "Manutenção/Controle de manutenção"
ID_documento = txtId
Documento = "Equipamento : " & cmbPostoTrabalho & " - Tipo da manutenção: " & txttipo & " - Data da manutenção: " & Date
Documento1 = ""
ProcGravaEvento
'==================================
USMsgBox ("Registro copiado com sucesso."), vbInformation, "CAPRIND v5.0"
ProcCarregaLista_data

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
