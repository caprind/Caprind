VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPCP_Relatorios_NC 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Relatórios - Não conformidade"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPCP_Relatorios_NC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3390
      Top             =   1020
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmPCP_Relatorios_NC.frx":0442
      Count           =   1
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1325
      Left            =   13135
      TabIndex        =   44
      Top             =   990
      Width           =   2115
      Begin VB.ComboBox cmbPor 
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
         ItemData        =   "frmPCP_Relatorios_NC.frx":3236
         Left            =   630
         List            =   "frmPCP_Relatorios_NC.frx":323D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Por."
         Top             =   210
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_mes_de 
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
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmPCP_Relatorios_NC.frx":3246
         Left            =   630
         List            =   "frmPCP_Relatorios_NC.frx":326E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Mês de."
         Top             =   570
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_de 
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
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmPCP_Relatorios_NC.frx":32AF
         Left            =   1260
         List            =   "frmPCP_Relatorios_NC.frx":32B1
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   675
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   8
         ToolTipText     =   "Data inicio."
         Top             =   570
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
         Format          =   489422849
         CurrentDate     =   39799
      End
      Begin VB.ComboBox Cmb_mes_ate 
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
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmPCP_Relatorios_NC.frx":32B3
         Left            =   630
         List            =   "frmPCP_Relatorios_NC.frx":32DB
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Mês até."
         Top             =   930
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox Cmb_ano_ate 
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
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmPCP_Relatorios_NC.frx":331C
         Left            =   1260
         List            =   "frmPCP_Relatorios_NC.frx":331E
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   675
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   9
         ToolTipText     =   "Data final."
         Top             =   930
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
         Format          =   489422849
         CurrentDate     =   39799
      End
      Begin VB.ComboBox Cmb_ano_de1 
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
         ItemData        =   "frmPCP_Relatorios_NC.frx":3320
         Left            =   630
         List            =   "frmPCP_Relatorios_NC.frx":3322
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Ano de."
         Top             =   570
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_ano_ate1 
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
         ItemData        =   "frmPCP_Relatorios_NC.frx":3324
         Left            =   630
         List            =   "frmPCP_Relatorios_NC.frx":3326
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Ano até."
         Top             =   930
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
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
         Left            =   195
         TabIndex        =   47
         Top             =   278
         Width           =   345
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   46
         Top             =   630
         Width           =   300
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
         Left            =   180
         TabIndex        =   45
         Top             =   990
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   55
      TabIndex        =   32
      Top             =   8880
      Width           =   15195
      Begin VB.TextBox Txt_percentual_NC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2248
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Percentual aprovada."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_qtde_outros 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   13514
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total outros."
         Top             =   360
         Width           =   960
      End
      Begin VB.TextBox Txt_percentual_outros 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   14490
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Percentual outros."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_qtde_prod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         MaxLength       =   20
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total não conforme."
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Txt_horas_retrab 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9349
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Total de horas de retrabalho."
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Txt_percentual_reaproveitada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12990
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Percentual reaproveitada."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_qtde_reaproveitada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11926
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total reaproveitada."
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox Txt_percentual_selecionada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11402
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Percentual selecionada."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_qtde_selecionada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10443
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total selecionada."
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox Txt_percentual_retrabalhada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8825
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Percentual de retrabalho."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_percentual_rejeitada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7267
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Percentual rejeitada."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_percentual_aprovada_desvio 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5709
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Percentual aprovada com desvio."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_qtde_aprovada_desvio 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4315
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total aprovada com desvio."
         Top             =   360
         Width           =   1380
      End
      Begin VB.TextBox Txt_percentual_aprovada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3791
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Percentual aprovada."
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Txt_qtde_NC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1274
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total não conforme."
         Top             =   360
         Width           =   960
      End
      Begin VB.TextBox Txt_qtde_retrabalhada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7791
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total retrabalho."
         Top             =   360
         Width           =   1020
      End
      Begin VB.TextBox Txt_qtde_rejeitada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6233
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total rejeitada."
         Top             =   360
         Width           =   1020
      End
      Begin VB.TextBox Txt_qtde_aprovada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2772
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total aprovada."
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2413
         TabIndex        =   66
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total produz."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   225
         TabIndex        =   64
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   14655
         TabIndex        =   63
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total outros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   13537
         TabIndex        =   62
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horas retrab."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   9379
         TabIndex        =   58
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   13155
         TabIndex        =   57
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total reapro."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   11956
         TabIndex        =   56
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   11567
         TabIndex        =   55
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   8990
         TabIndex        =   54
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   7432
         TabIndex        =   53
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   5874
         TabIndex        =   52
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   3956
         TabIndex        =   51
         Top             =   195
         Width           =   180
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total selec."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   10488
         TabIndex        =   50
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total retrab."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   7821
         TabIndex        =   49
         Top             =   180
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total aprov. desv."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   4330
         TabIndex        =   48
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total aprov."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2817
         TabIndex        =   40
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total rejeit."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6293
         TabIndex        =   39
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total NC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1432
         TabIndex        =   33
         Top             =   195
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6525
      Left            =   55
      TabIndex        =   16
      Top             =   2325
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11509
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
      NumItems        =   29
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
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3043
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Fase"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Grupo/op."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Máquina"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3043
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Qtde. prod."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Qtde. NC"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Qtde. aprov."
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "N"
         Text            =   "Qtde. aprov. c/ desv."
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Object.Tag             =   "N"
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "N"
         Text            =   "Qtde. rej."
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   19
         Object.Tag             =   "N"
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "N"
         Text            =   "Qtde. retrab."
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   21
         Object.Tag             =   "N"
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "Hs. retrab."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Tag             =   "N"
         Text            =   "Qtde. selec."
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   24
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Tag             =   "N"
         Text            =   "Qtde. reapro."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Object.Tag             =   "N"
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Tag             =   "N"
         Text            =   "Qtde. outros"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   28
         Object.Tag             =   "N"
         Text            =   "%"
         Object.Width           =   1147
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6525
      Left            =   55
      TabIndex        =   17
      Top             =   2325
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11509
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
      NumItems        =   0
   End
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
      Height          =   1325
      Left            =   55
      TabIndex        =   34
      Top             =   990
      Width           =   1695
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
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
         TabIndex        =   1
         Top             =   750
         Width           =   1425
      End
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
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
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1325
      Left            =   1785
      TabIndex        =   35
      Top             =   990
      Width           =   1455
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
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
         TabIndex        =   3
         Top             =   750
         Width           =   1155
      End
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
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
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   41
      Top             =   9720
      Width           =   11385
      _ExtentX        =   20082
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   43
      Top             =   0
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   51
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   93
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
      ButtonKey4      =   "3"
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
      ButtonWidth4    =   36
      ButtonHeight4   =   21
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
      ButtonKey5      =   "4"
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
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
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
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
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
      Height          =   1325
      Left            =   3275
      TabIndex        =   36
      Top             =   990
      Width           =   9825
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
         Left            =   2760
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   630
         Width           =   6885
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
         ItemData        =   "frmPCP_Relatorios_NC.frx":3328
         Left            =   180
         List            =   "frmPCP_Relatorios_NC.frx":3344
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   630
         Width           =   2505
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
         Height          =   330
         ItemData        =   "frmPCP_Relatorios_NC.frx":33AC
         Left            =   2760
         List            =   "frmPCP_Relatorios_NC.frx":33AE
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   630
         Visible         =   0   'False
         Width           =   6885
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
         Left            =   5460
         TabIndex        =   38
         Top             =   420
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
         Left            =   1005
         TabIndex        =   37
         Top             =   420
         Width           =   840
      End
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11610
      TabIndex        =   42
      Top             =   9750
      Width           =   3315
   End
End
Attribute VB_Name = "frmPCP_Relatorios_NC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=PBTA2obXT08&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=5&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
 
Private Sub Cmb_ano_ate1_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ano_de_Click()
On Error GoTo tratar_erro

Cmb_ano_ate = Cmb_ano_de
ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ano_de1_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_ate_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_de_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPor_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
ProcMostrarEsconderCombosData

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Change()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 And Lista1.ListItems.Count = 0 Then Exit Sub
Sit_REG = 2
frmQualidade_Relatorios_NC_Menu_Impressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Familiatext = ""
Contador1 = 1
Posicao = 0
Lista.ListItems.Clear
Lista1.ListItems.Clear
If TBLISTA.EOF = False Then
    If optDetalhado.Value = True Then Posicao = TBLISTA.RecordCount
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Desenho, N_referencia, Produto from producao where Ordem = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockReadOnly
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Produto), "", TBAbrir!Produto)
                End If
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!OS), "", TBLISTA!OS)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Fase), "", TBLISTA!Fase)
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Grupo_op from ordemservico where idproducao = " & TBLISTA!OS, Conexao, adOpenKeyset, adLockReadOnly
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!Grupo_op), "", TBAbrir!Grupo_op)
                End If
                TBAbrir.Close
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento)
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select Descricao from Cadmaquinas where Maquina = '" & TBLISTA!DescEvento & "'", Conexao, adOpenKeyset, adLockReadOnly
                If TBFIltro.EOF = False Then
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBFIltro!Descricao), "", TBFIltro!Descricao)
                End If
                TBFIltro.Close
                
                .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!QtdePrev), 0, TBLISTA!QtdePrev)
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC)
                .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Refugo), "", Format(TBLISTA!Refugo, "###,##0.00"))
                .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Terceiros), "", TBLISTA!Terceiros)
                .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Numero2), "", Format(TBLISTA!Numero2, "###,##0.00"))
                .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!impostos), "", TBLISTA!impostos)
                .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA!Numero3), "", Format(TBLISTA!Numero3, "###,##0.00"))
                .Item(.Count).SubItems(18) = IIf(IsNull(TBLISTA!Lucro), "", TBLISTA!Lucro)
                .Item(.Count).SubItems(19) = IIf(IsNull(TBLISTA!Numero4), "", Format(TBLISTA!Numero4, "###,##0.00"))
                .Item(.Count).SubItems(20) = IIf(IsNull(TBLISTA!material), "", TBLISTA!material)
                .Item(.Count).SubItems(21) = IIf(IsNull(TBLISTA!Numero5), "", Format(TBLISTA!Numero5, "###,##0.00"))
                .Item(.Count).SubItems(22) = IIf(IsNull(TBLISTA!Totalhsutil), "00:00:00", Format(TBLISTA!Totalhsutil, "hh:mm:ss"))
                .Item(.Count).SubItems(23) = IIf(IsNull(TBLISTA!Servicos), "", TBLISTA!Servicos)
                .Item(.Count).SubItems(24) = IIf(IsNull(TBLISTA!Numero6), "", Format(TBLISTA!Numero6, "###,##0.00"))
                .Item(.Count).SubItems(24) = IIf(IsNull(TBLISTA!Total), "", TBLISTA!Total)
                .Item(.Count).SubItems(26) = IIf(IsNull(TBLISTA!Numero7), "", Format(TBLISTA!Numero7, "###,##0.00"))
                .Item(.Count).SubItems(27) = IIf(IsNull(TBLISTA!Total_peca), "", TBLISTA!Total_peca)
                .Item(.Count).SubItems(28) = IIf(IsNull(TBLISTA!Numero8), "", Format(TBLISTA!Numero8, "###,##0.00"))
            End With
        Else
            If TBLISTA!maquina <> "" Then
                TextoFiltro = " -  Prod.  |      NC       |      %      |    Aprov.   |      %      | Aprov. c/ des. |      %      |   Rej.   |      %      |  Retrab.  |      %      |   Selec.  |      %      |  Reapro.  |      %      |   Outros  |     %"
                With Lista1.ListItems
                    Contador1 = 1
                    If cmbPor = "Dia" Then
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> Format(TBLISTA!Execucaoprev, "dd/mm/yy") & TextoFiltro
                            Contador1 = Contador1 + 1
                        Loop
                    ElseIf cmbPor = "Mês" Then
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev & TextoFiltro
                                Contador1 = Contador1 + 1
                            Loop
                        Else
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev & TextoFiltro
                                Contador1 = Contador1 + 1
                            Loop
                    End If
                    
                    If TBLISTA!maquina <> Familiatext Then
                        .Add , , TBLISTA!maquina
                        Posicao = Posicao + 1
                    End If
                    
                    FamiliaAntiga = IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC)
                    Select Case cmbPor
                        Case "Dia": SubTipoItem = 15
                        Case "Mês": SubTipoItem = 14
                        Case "Ano": SubTipoItem = 12
                    End Select
                    
                    .Item(.Count).SubItems(Contador1) = FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!QtdePrev), 0, TBLISTA!QtdePrev), SubTipoItem) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!qtdeNC), 0, TBLISTA!qtdeNC), 9) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Refugo), "0,00", Format(TBLISTA!Refugo, "###,##0.00")), 8) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Terceiros), 0, TBLISTA!Terceiros), 9) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero2), "0,00", Format(TBLISTA!Numero2, "###,##0.00")), 8) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!impostos), 0, TBLISTA!impostos), 13) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero3), "0,00", Format(TBLISTA!Numero3, "###,##0.00")), 8) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Lucro), 0, TBLISTA!Lucro), 6) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero4), "0,00", Format(TBLISTA!Numero4, "###,##0.00")), 8) & _
                                                        "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!material), 0, TBLISTA!material), 9) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero5), "0,00", Format(TBLISTA!Numero5, "###,##0.00")), 7) & " | " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Servicos), 0, TBLISTA!Servicos), 7) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero6), "0,00", Format(TBLISTA!Numero6, "###,##0.00")), 8) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Total), 0, TBLISTA!Total), 9) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero7), "0,00", Format(TBLISTA!Numero7, "###,##0.00")), 8) & "|" & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Total_peca), 0, TBLISTA!Total_peca), 8) & "|" & IIf(IsNull(TBLISTA!Numero8), "0,00", Format(TBLISTA!Numero8, "###,##0.00"))
                    
                    'Carrega qtde. total
                    Contador1 = 1
                    Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Totais -  Prod.  |      NC       |      %      |    Aprov.   |      %      | Aprov. c/ des. |      %      |   Rej.   |      %      |  Retrab.  |      %      |   Selec.  |      %      |  Reapro.  |      %      |   Outros  |     %"
                        Contador1 = Contador1 + 1
                    Loop
                    
                    SubTipoItem = 11
                    .Item(.Count).SubItems(Contador1) = FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor9), 0, TBLISTA!Valor9), SubTipoItem) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero9), 0, TBLISTA!Numero9), 7) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Eficiencia), "0,00", Format(TBLISTA!Eficiencia, "###,##0.00")), 6) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero10), 0, TBLISTA!Numero10), 7) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero11), "0,00", Format(TBLISTA!Numero11, "###,##0.00")), 6) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero12), 0, TBLISTA!Numero12), 11) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero13), "0,00", Format(TBLISTA!Numero13, "###,##0.00")), 6) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero14), 0, TBLISTA!Numero14), 5) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Numero15), "0,00", Format(TBLISTA!Numero15, "###,##0.00")), 6) & _
                                                        "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor1), 0, TBLISTA!Valor1), 6) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor2), "0,00", Format(TBLISTA!Valor2, "###,##0.00")), 6) & "   |   " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor3), 0, TBLISTA!Valor3), 5) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor4), "0,00", Format(TBLISTA!Valor4, "###,##0.00")), 5) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor5), 0, TBLISTA!Valor5), 7) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor6), "0,00", Format(TBLISTA!Valor6, "###,##0.00")), 6) & "  |  " & FunTamanhoTextoZeroEsq(IIf(IsNull(TBLISTA!Valor7), 0, TBLISTA!Valor7), 6) & "  |  " & IIf(IsNull(TBLISTA!Valor8), "0,00", Format(TBLISTA!Valor8, "###,##0.00"))
                End With
            End If
        End If
        Familiatext = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_qtde_prod = Format(TBLISTA!QtdeProduzida, "###,##0.00")
    Txt_qtde_NC = Format(TBLISTA!qtdeNC, "###,##0.00")
    Txt_percentual_NC = Format(TBLISTA!Numero6, "###,##0.00")
    Txt_qtde_aprovada = Format(TBLISTA!CustoMat, "###,##0.00")
    Txt_percentual_aprovada = Format(TBLISTA!CustoObra, "###,##0.00")
    Txt_qtde_aprovada_desvio = Format(TBLISTA!Terceros, "###,##0.00")
    Txt_percentual_aprovada_desvio = Format(TBLISTA!Lucro, "###,##0.00")
    Txt_qtde_rejeitada = TBLISTA!Valor1
    Txt_percentual_rejeitada = Format(TBLISTA!Valor2, "###,##0.00")
    Txt_qtde_retrabalhada = TBLISTA!Valor3
    Txt_percentual_retrabalhada = Format(TBLISTA!Total1, "###,##0.00")
    Txt_horas_retrab = TBLISTA!TotalEficiencia
    Txt_qtde_selecionada = TBLISTA!Total2
    Txt_percentual_selecionada = Format(TBLISTA!Numero1, "###,##0.00")
    Txt_qtde_reaproveitada = TBLISTA!Numero2
    Txt_percentual_reaproveitada = Format(TBLISTA!Numero3, "###,##0.00")
    Txt_qtde_outros = TBLISTA!Numero4
    Txt_percentual_outros = Format(TBLISTA!Numero5, "###,##0.00")
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaListaeCampos()
On Error GoTo tratar_erro

Lista.ListItems.Clear
With Lista1
    .ColumnHeaders.Clear
    .ListItems.Clear
End With
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
Txt_qtde_prod = ""
Txt_qtde_NC = ""
Txt_qtde_aprovada = ""
Txt_percentual_NC = ""
Txt_percentual_aprovada = ""
Txt_qtde_aprovada_desvio = ""
Txt_percentual_aprovada_desvio = ""
Txt_qtde_rejeitada = ""
Txt_percentual_rejeitada = ""
Txt_qtde_retrabalhada = ""
Txt_percentual_retrabalhada = ""
Txt_horas_retrab = ""
Txt_qtde_selecionada = ""
Txt_percentual_selecionada = ""
Txt_qtde_reaproveitada = ""
Txt_percentual_reaproveitada = ""
Txt_qtde_outros = ""
Txt_percentual_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
Formulario = "PCP/Relatórios/Não conformidade"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcCarregaComboAno Cmb_ano_ate, "2005", 1
ProcCarregaComboAno Cmb_ano_ate1, "2005", 1
ProcCarregaComboAno Cmb_ano_de, "2005", 1
ProcCarregaComboAno Cmb_ano_de1, "2005", 1
cmbfiltrarpor.Text = "Ordem"
cmbPor = "Dia"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Qualidade_NC = False Then Formulario = "PCP/Relatórios/Não conformidade" Else Formulario = "Qualidade/Relatórios/Não conformidade"
Direitos
ProcLimpaVariaveisPrincipais

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

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTexto()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

txtTexto = ""
txtTexto.Visible = False

With cmbTexto
    .Clear
    .Visible = True
    If Opt_individual.Value = True Then
        Select Case cmbfiltrarpor
            Case "Operador"
                Set TBUsuarios = CreateObject("adodb.recordset")
                TBUsuarios.Open "Select usuario, CODIGO FROM usuarios Group by usuario, CODIGO", Conexao, adOpenKeyset, adLockOptimistic
                If TBUsuarios.EOF = False Then
                    Do While TBUsuarios.EOF = False
                        If IsNull(TBUsuarios!CODIGO) = False And TBUsuarios!CODIGO <> "" Then OperadorTexto = TBUsuarios!Usuario & "-" & TBUsuarios!CODIGO Else OperadorTexto = TBUsuarios!Usuario
                        .AddItem OperadorTexto
                        TBUsuarios.MoveNext
                    Loop
                End If
                TBUsuarios.Close
            Case "Posto de trabalho"
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select maquina from CadMaquinas Group by Maquina", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    Do While TBMaquinas.EOF = False
                        .AddItem TBMaquinas!maquina
                        TBMaquinas.MoveNext
                    Loop
                End If
                TBMaquinas.Close
            Case "Turno"
                .AddItem 0
                .AddItem 1
                .AddItem 2
                .AddItem 3
                .AddItem 4
            Case "Setor"
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select Setor from CadMaquinas Group by Setor", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    Do While TBMaquinas.EOF = False
                        .AddItem TBMaquinas!Setor
                        TBMaquinas.MoveNext
                    Loop
                End If
                TBMaquinas.Close
            Case "Código interno"
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select desenho from Producao Group by desenho", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    Do While TBOrdem.EOF = False
                        .AddItem TBOrdem!Desenho
                        TBOrdem.MoveNext
                    Loop
                End If
                TBOrdem.Close
            Case "Código de referência"
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select N_Referencia from Producao where N_Referencia <> 'Null' Group by N_Referencia", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    Do While TBOrdem.EOF = False
                        If TBOrdem!N_referencia <> "" Then .AddItem TBOrdem!N_referencia
                        TBOrdem.MoveNext
                    Loop
                End If
                TBOrdem.Close
            Case "Disposição"
                .AddItem "Aprovado"
                .AddItem "Aprovado com desvio"
                .AddItem "Rejeitar"
                .AddItem "Retrabalhar"
                .AddItem "Selecionar"
                .AddItem "Reaproveitar para outro produto"
                .AddItem "Outros"
                .AddItem "Nada consta"
            Case "Ordem":
                .Visible = False
                txtTexto.Visible = True
        End Select
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If cmbTexto.Visible = True Then
    If Opt_individual.Value = True And cmbTexto = "" Then
        NomeCampo = "o texto para pesquisa"
        ProcVerificaAcao
        cmbTexto.SetFocus
        Exit Sub
    End If
Else
    If Opt_individual.Value = True And txtTexto = "" Then
        NomeCampo = "o texto para pesquisa"
        ProcVerificaAcao
        txtTexto.SetFocus
        Exit Sub
    End If
End If
If cmbPor = "Dia" Then
    With msk_fltFim
        If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
            .Value = Date
            .SetFocus
            Exit Sub
        End If
    End With
End If
If optResumido.Value = True Then
    If cmbPor = "Mês" Then
        If Cmb_mes_de = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_de.SetFocus
            Exit Sub
        End If
        If Cmb_mes_ate = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_ate.SetFocus
            Exit Sub
        End If
        If Cmb_ano_de = "" Then
            NomeCampo = "o ano"
            ProcVerificaAcao
            Cmb_ano_de.SetFocus
            Exit Sub
        End If
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)
        If Qtd < qt And Cmb_ano_ate <= Cmb_ano_de Then
            USMsgBox ("O mês final não pode ser menor que o mês inicial."), vbExclamation, "CAPRIND v5.0"
            Cmb_mes_ate = Cmb_mes_de
        End If
    ElseIf cmbPor = "Ano" Then
            If Cmb_ano_de1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_de1.SetFocus
                Exit Sub
            End If
            If Cmb_ano_ate1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_ate1.SetFocus
                Exit Sub
            End If
            Qtd = Cmb_ano_de1
            qt = Cmb_ano_ate1
            If qt < Qtd Then
                USMsgBox ("O ano final não pode ser menor que o ano inicial."), vbExclamation, "CAPRIND v5.0"
                Cmb_ano_ate1 = Cmb_ano_de1
            End If
    End If
End If

Inicio = Time
ProcLimpaCamposTotais
ProcAbrirTabelas
If optResumido.Value = True Then
    ProcCriaColunas
    
    'Soma e grava o total geral
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select maquina,SUM(qtdePrev) as qtdePrev, Sum(QtdeNC) as QTNC, Sum(Terceiros) as qtdeliberada, Sum(Impostos) as qtdeliberar, Sum(Lucro) as qt, Sum(material) as Qtd, Sum(Servicos) as Qtd_Prog, Sum(Total) as QTLOTE, Sum(Total_peca) as quantestoque from Producao_relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' Group by Maquina", Conexao, adOpenKeyset, adLockReadOnly
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            QTPC = IIf(IsNull(TBLISTA!QtdePrev), 0, TBLISTA!QtdePrev) 'Qtde. produzida
            QTNC = IIf(IsNull(TBLISTA!QTNC), 0, TBLISTA!QTNC) 'Qtde. NC
            
            'Não conforme
            If QTPC <> 0 Then QuantsolicitadoN9 = (QTNC / QTPC) * 100 Else QuantsolicitadoN9 = 0 'Percentual (%)
            
            'Aprovado
            qtdeliberada = IIf(IsNull(TBLISTA!qtdeliberada), 0, TBLISTA!qtdeliberada)
            If QTPC <> 0 Then QuantsolicitadoN2 = (qtdeliberada / QTPC) * 100 Else QuantsolicitadoN2 = 0 'Percentual (%)
            
            'Aprovado c/ desvio
            qtdeliberar = IIf(IsNull(TBLISTA!qtdeliberar), 0, TBLISTA!qtdeliberar)
            If QTPC <> 0 Then QuantsolicitadoN3 = (qtdeliberar / QTPC) * 100 Else QuantsolicitadoN3 = 0 'Percentual (%)
            
            'Rejeitar
            qt = IIf(IsNull(TBLISTA!qt), 0, TBLISTA!qt)
            If QTPC <> 0 Then QuantsolicitadoN4 = (qt / QTPC) * 100 Else QuantsolicitadoN4 = 0 'Percentual (%)
            
            'Retrabalhar
            Qtd = IIf(IsNull(TBLISTA!Qtd), 0, TBLISTA!Qtd)
            If QTPC <> 0 Then QuantsolicitadoN5 = (Qtd / QTPC) * 100 Else QuantsolicitadoN5 = 0 'Percentual (%)
            
             'Selecionar
            Qtd_Prog = IIf(IsNull(TBLISTA!Qtd_Prog), 0, TBLISTA!Qtd_Prog)
            If QTPC <> 0 Then QuantsolicitadoN6 = (Qtd_Prog / QTPC) * 100 Else QuantsolicitadoN6 = 0 'Percentual (%)
            
            'Reaproveitar
            QTLOTE = IIf(IsNull(TBLISTA!QTLOTE), 0, TBLISTA!QTLOTE)
            If QTPC <> 0 Then QuantsolicitadoN7 = (QTLOTE / QTPC) * 100 Else QuantsolicitadoN7 = 0 'Percentual (%)
            
            'Outros
            quantestoque = IIf(IsNull(TBLISTA!quantestoque), 0, TBLISTA!quantestoque)
            If QTPC <> 0 Then QuantsolicitadoN8 = (quantestoque / QTPC) * 100 Else QuantsolicitadoN8 = 0 'Percentual (%)
                        
            NovoValor1 = Replace(QTNC, ",", ".")
            NovoValor2 = Replace(qtdeliberada, ",", ".")
            NovoValor3 = Replace(QuantsolicitadoN2, ",", ".")
            NovoValor4 = Replace(qtdeliberar, ",", ".")
            NovoValor5 = Replace(QuantsolicitadoN3, ",", ".")
            NovoValor6 = Replace(qt, ",", ".")
            NovoValor7 = Replace(QuantsolicitadoN4, ",", ".")
            NovoValor8 = Replace(Qtd, ",", ".")
            NovoValor9 = Replace(QuantsolicitadoN5, ",", ".")
            NovoValor10 = Replace(Qtd_Prog, ",", ".")
            NovoValor11 = Replace(QuantsolicitadoN6, ",", ".")
            NovoValor12 = Replace(QTLOTE, ",", ".")
            NovoValor13 = Replace(QuantsolicitadoN7, ",", ".")
            NovoValor14 = Replace(quantestoque, ",", ".")
            NovoValor15 = Replace(QuantsolicitadoN8, ",", ".")
            NovoValor16 = Replace(QTPC, ",", ".")
            NovoValor17 = Replace(QuantsolicitadoN9, ",", ".")
            
            Conexao.Execute "Update Producao_relatorios Set Numero9 = " & NovoValor1 & ", Numero10 = " & NovoValor2 & ", Numero11 = " & NovoValor3 & ", Numero12 = " & NovoValor4 & ", Numero13 = " & NovoValor5 & ", Numero14 = " & NovoValor6 & ", Numero15 = " & NovoValor7 & ", Valor1 = " & NovoValor8 & ", Valor2 = " & NovoValor9 & ", Valor3 = " & NovoValor10 & ", Valor4 = " & NovoValor11 & ", Valor5 = " & NovoValor12 & ", Valor6 = " & NovoValor13 & ", Valor7 = " & NovoValor14 & ", Valor8 = " & NovoValor15 & ", Valor9 = " & NovoValor16 & ", Eficiencia = " & NovoValor17 & " where Maquina = '" & TBLISTA!maquina & "'"
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
End If
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True And optDetalhado.Value = True Then
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data, Maquina", Conexao, adOpenKeyset, adLockReadOnly
Else
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Maquina <> 'Null' order by Maquina", Conexao, adOpenKeyset, adLockReadOnly
End If
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

Select Case cmbfiltrarpor
    Case "Operador": Familiatext = "Operador"
    Case "Posto de trabalho": Familiatext = "Maquina"
    Case "Turno": Familiatext = "Turno"
    Case "Setor": Familiatext = "Setor"
    Case "Disposição": Familiatext = "ParecerCQ"
    Case "Ordem": Familiatext = "Ordem"
    Case "Código interno": Familiatext = "Desenho"
    Case "Código de referência": Familiatext = "N_referencia"
End Select

If Opt_individual = True Then
    If cmbfiltrarpor = "Operador" Then
        NumeroCartao = ReturnNumbersOnly(cmbTexto)
        If NumeroCartao = "" Then Filtro = cmbTexto Else Filtro = NumeroCartao & "-" & Left(cmbTexto, Len(cmbTexto) - 9)
    Else
        If cmbfiltrarpor <> "Ordem" Then Filtro = cmbTexto Else Filtro = txtTexto
    End If
End If

Set TBCarteira = CreateObject("adodb.recordset")
If optDetalhado.Value = True Then
    TBCarteira.Open "Select * from Qualidade_relatorio_NC_detalhado where " & Familiatext & " = '" & Filtro & "' and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockReadOnly
Else
    Par1 = ""
    Permitido = False
    Select Case cmbPor
        Case "Dia":
            Dataini = msk_fltInicio
            DataFim = msk_fltFim
            Do While Dataini <= DataFim
                If Permitido = False Then Par1 = "[" & Dataini & "]" Else Par1 = Par1 & " , [" & Dataini & "]"
                Permitido = True
                Dataini = Dataini + 1
            Loop
            Pesquisa = "(Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            Pesquisa1 = "PIVOT (Sum(TTNC) for Data In (" & Par1 & "))"
            Pesquisa2 = "Data"
        Case "Mês":
            qt = FunVerificaMes(Cmb_mes_de)
            Qtd = FunVerificaMes(Cmb_mes_ate)
            MesX = qt
            MesX1 = Qtd
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Month(Data) >= '" & MesX & "' and Year(Data) >= '" & Cmb_ano_de & "' and Month(Data) <= '" & MesX1 & "' and Year(Data) <= '" & Cmb_ano_ate & "'"
            Pesquisa1 = "PIVOT (Sum(TTNC) for Mes In (" & Par1 & "))"
            Pesquisa2 = "Mes"
        Case "Ano":
            qt = Cmb_ano_de1
            Qtd = Cmb_ano_ate1
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Year(Data) >= '" & Cmb_ano_de1 & "' and Year(Data) <= '" & Cmb_ano_ate1 & "'"
            Pesquisa1 = "PIVOT (Sum(TTNC) for Ano In (" & Par1 & "))"
            Pesquisa2 = "Ano"
    End Select
    
    If cmbfiltrarpor = "Disposição" Then
        CamposFiltro = Familiatext & ", Produto, Ordem, OS"
    ElseIf cmbfiltrarpor = "Ordem" Then
            CamposFiltro = Familiatext & ", Produto, ParecerCQ, OS"
        Else
            CamposFiltro = Familiatext & ", Produto, ParecerCQ, Ordem, OS"
    End If
    
    Set TBGravar = CreateObject("adodb.recordset")
    If Opt_individual.Value = True Then
        TBCarteira.Open "SELECT " & CamposFiltro & ", " & Par1 & " From (Select " & CamposFiltro & ", " & Pesquisa2 & ", TTNC from Qualidade_relatorio_NC_detalhado Where " & Familiatext & " = '" & Filtro & "' and " & Pesquisa & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
    Else
        TBCarteira.Open "SELECT " & CamposFiltro & ", " & Par1 & " From (Select " & CamposFiltro & ", " & Pesquisa2 & ", TTNC from Qualidade_relatorio_NC_detalhado Where " & Pesquisa & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
    End If
End If
ProcFiltrar1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

Permitido = False
maquina = ""
Familiatext = ""
Produto = ""
quantidade = 0
If TBCarteira.EOF = False Then
    Permitido = True
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBCarteira.EOF = False
        If optDetalhado.Value = True Then
            Set TBProdutividade = CreateObject("adodb.recordset")
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            ProcCriarResumido
        End If
        TBCarteira.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

Permitido = True

TBProdutividade.AddNew
TBProdutividade!Ordem = TBCarteira!Ordem
TBProdutividade!Fase = TBCarteira!Fase
TBProdutividade!Data = Format(TBCarteira!Data, "dd/mm/yy")
TBProdutividade!OS = TBCarteira!OS
TBProdutividade!DescEvento = TBCarteira!maquina

'Verifica quantidade produzindo
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select ap_backup from producao P INNER JOIN Ordemservico O ON O.ordem = P.ordem where O.IDproducao = " & TBCarteira!OS, Conexao, adOpenKeyset, adLockReadOnly
If TBOrdem.EOF = False Then
    If TBOrdem!AP_backup = True Then
        TabelaBackup = "ProducaoFases_Backup"
    Else
        TabelaBackup = "ProducaoFases"
    End If
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select SUM(quantidade) + SUM(Reprovada) as QtdeProd from " & TabelaBackup & " where OS = " & TBCarteira!OS & " AND Data Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and CodigoDesc = 2", Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then
        quantidade = IIf(IsNull(TBFIltro!QtdeProd), 0, TBFIltro!QtdeProd)
    End If
    TBFIltro.Close
End If
TBOrdem.Close

TBProdutividade!QtdePrev = quantidade
TBProdutividade!qtdeNC = TBCarteira!TTNC
'TBProdutividade!Qtdetotalprod = TBCarteira!TTNC 'ver pra que serve

'Qtde./Percentual por parecer
Select Case TBCarteira!ParecerCQ
    Case "Aprovado":  TBProdutividade!Terceiros = TBCarteira!TTNC
    Case "Aprovado c/ desvio": TBProdutividade!impostos = TBCarteira!TTNC
    Case "Rejeitar": TBProdutividade!Lucro = TBCarteira!TTNC
    Case "Retrabalhar":
        TBProdutividade!material = TBCarteira!TTNC
        TempoTotalProd = 0
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select Producaofases.TempoTotal from Producaofases INNER JOIN ordemservico ON Producaofases.IDFase = ordemservico.IDproducao where ordemservico.Ordem = " & TBCarteira!Ordem & " and ordemservico.Fase = " & TBCarteira!Fase & " and ordemservico.Retrabalho = 'True' and Producaofases.Data Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and Producaofases.CodigoDesc = 2", Conexao, adOpenKeyset, adLockReadOnly
        If TBFIltro.EOF = False Then
            Do While TBFIltro.EOF = False
                TempoTotalProd = TempoTotalProd + IIf(IsNull(TBFIltro!TempoTotal), 0, TBFIltro!TempoTotal)
                TBFIltro.MoveNext
            Loop
        End If
        TBFIltro.Close
        ElapsedTime (TempoTotalProd)
        TBProdutividade!qtdeOK = s 'Total de horas utilizadas para retrabalho (segundos)
        TBProdutividade!Totalhsutil = TempoTotalProd
    Case "Selecionar": TBProdutividade!Servicos = TBCarteira!TTNC
    Case "Reaproveitar": TBProdutividade!Total = TBCarteira!TTNC
    Case "Outros": TBProdutividade!Total_peca = TBCarteira!TTNC
End Select
If quantidade <> 0 Then
    TBProdutividade!Refugo = (TBProdutividade!qtdeNC / quantidade) * 100
    TBProdutividade!Numero2 = (TBProdutividade!Terceiros / quantidade) * 100
    TBProdutividade!Numero3 = (TBProdutividade!impostos / quantidade) * 100
    TBProdutividade!Numero4 = (TBProdutividade!Lucro / quantidade) * 100
    TBProdutividade!Numero5 = (TBProdutividade!material / quantidade) * 100
    TBProdutividade!Numero6 = (TBProdutividade!Servicos / quantidade) * 100
    TBProdutividade!Numero7 = (TBProdutividade!Total / quantidade) * 100
    TBProdutividade!Numero8 = (TBProdutividade!Total_peca / quantidade) * 100
End If

If cmbfiltrarpor <> "Ordem" Then TBProdutividade!maquina = cmbTexto Else TBProdutividade!maquina = txtTexto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarResumido()
On Error GoTo tratar_erro

Permitido = True
Select Case cmbPor
    Case "Dia":
        qt = 0
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            qt = qt + 1
            ProcEnviaDadosResumido
            Dataini = Dataini + 1
        Loop
    Case "Mês":
        qt = MesX
        Qtd = MesX1
        Do While qt <= Qtd
            ProcEnviaDadosResumido
            qt = qt + 1
        Loop
    Case "Ano":
        qt = Cmb_ano_de1
        Qtd = Cmb_ano_ate1
        Do While qt <= Qtd
            ProcEnviaDadosResumido
            qt = qt + 1
        Loop
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

Permitido1 = True
Select Case cmbfiltrarpor
    Case "Operador": Familiatext = TBCarteira!Operador
    Case "Posto de trabalho": Familiatext = TBCarteira!maquina
    Case "Turno": Familiatext = TBCarteira!Turno
    Case "Setor": Familiatext = TBCarteira!Setor
    Case "Disposição": Familiatext = TBCarteira!ParecerCQ
    Case "Ordem": Familiatext = Left(TBCarteira!Ordem & " - " & TBCarteira!Produto, 25)
    Case "Código interno": Familiatext = TBCarteira!Desenho
    Case "Código de referência": Familiatext = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
End Select
Select Case cmbPor
    Case "Dia":
        DataFiltro = "Data = '" & Dataini & "'"
        DataTexto = Dataini
    Case "Mês":
        DataFiltro = "Execucaoprev = '" & Format(qt & "/" & Cmb_ano_de, "mm/yyyy") & "'"
        DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano":
        DataFiltro = "Execucaoprev = '" & qt & "'"
        DataTexto = "01" & "/01/" & qt
End Select

Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Maquina = '" & Familiatext & "' and " & DataFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!Data = Format(DataTexto, "dd/mm/yyyy")

Select Case cmbPor
    Case "Dia":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
        TextoFiltroQTOK = "data = '" & Format(Dataini, "Short Date") & "'"
        TextoFiltroQTOKINNER = "P.data = '" & Format(Dataini, "Short Date") & "'"
    Case "Mês":
        Select Case qt
            Case 1:
                TotalCreditar = IIf(IsNull(TBCarteira![1]), 0, Format(TBCarteira![1], "###,##0.00"))
                MesQTOK = 1
            Case 2:
                TotalCreditar = IIf(IsNull(TBCarteira![2]), 0, Format(TBCarteira![2], "###,##0.00"))
                MesQTOK = 2
            Case 3:
                TotalCreditar = IIf(IsNull(TBCarteira![3]), 0, Format(TBCarteira![3], "###,##0.00"))
                MesQTOK = 3
            Case 4:
                TotalCreditar = IIf(IsNull(TBCarteira![4]), 0, Format(TBCarteira![4], "###,##0.00"))
                MesQTOK = 4
            Case 5:
                TotalCreditar = IIf(IsNull(TBCarteira![5]), 0, Format(TBCarteira![5], "###,##0.00"))
                MesQTOK = 5
            Case 6:
                TotalCreditar = IIf(IsNull(TBCarteira![6]), 0, Format(TBCarteira![6], "###,##0.00"))
                MesQTOK = 6
            Case 7:
                TotalCreditar = IIf(IsNull(TBCarteira![7]), 0, Format(TBCarteira![7], "###,##0.00"))
                MesQTOK = 7
            Case 8:
                TotalCreditar = IIf(IsNull(TBCarteira![8]), 0, Format(TBCarteira![8], "###,##0.00"))
                MesQTOK = 8
            Case 9:
                TotalCreditar = IIf(IsNull(TBCarteira![9]), 0, Format(TBCarteira![9], "###,##0.00"))
                MesQTOK = 9
            Case 10:
                TotalCreditar = IIf(IsNull(TBCarteira![10]), 0, Format(TBCarteira![10], "###,##0.00"))
                MesQTOK = 10
            Case 11:
                TotalCreditar = IIf(IsNull(TBCarteira![11]), 0, Format(TBCarteira![11], "###,##0.00"))
                MesQTOK = 11
            Case 12:
                TotalCreditar = IIf(IsNull(TBCarteira![12]), 0, Format(TBCarteira![12], "###,##0.00"))
                MesQTOK = 12
        End Select
        TBProdutividade!Execucaoprev = Format(qt & "/" & Cmb_ano_de, "mm/yyyy")
        TextoFiltroQTOK = "Month(data) = '" & MesQTOK & "' and Year(Data) = '" & Cmb_ano_de & "'"
        TextoFiltroQTOKINNER = "Month(P.data) = '" & MesQTOK & "' and Year(Data) = '" & Cmb_ano_de & "'"
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!Execucaoprev = qt
        TextoFiltroQTOK = "Year(Data) = '" & qt & "'"
        TextoFiltroQTOKINNER = "Year(P.Data) = '" & qt & "'"
End Select

TBProdutividade!qtdeNC = Format(TBProdutividade!qtdeNC + IIf(IsNull(TotalCreditar), 0, TotalCreditar), "###,##0.00")
TBProdutividade!maquina = Left(Familiatext, 25)

'Qtde./Percentual por parecer
Select Case TBCarteira!ParecerCQ
    Case "Aprovado": TBProdutividade!Terceiros = TBProdutividade!Terceiros + TotalCreditar
    Case "Aprovado c/ desvio": TBProdutividade!impostos = TBProdutividade!impostos + TotalCreditar
    Case "Rejeitar": TBProdutividade!Lucro = TBProdutividade!Lucro + TotalCreditar
    Case "Retrabalhar":
        TBProdutividade!material = TBProdutividade!material + TotalCreditar
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Ordem = " & TBCarteira!Ordem & " and " & DataFiltro & " and QtdeOK <> 0", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = True Then
            TempoTotalProd = 0
            Select Case cmbPor
                Case "Dia": DataFiltro1 = "Producaofases.Data Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
                Case "Mês":
                    Qtde = FunVerificaMes(Cmb_mes_de)
                    Quant = FunVerificaMes(Cmb_mes_ate)
                    DataFiltro1 = "Month(Producaofases.Data) >= '" & qt & "' and Year(Producaofases.Data) >= '" & Cmb_ano_de & "' and Month(Producaofases.Data) <= '" & Qtd & "' and Year(Producaofases.Data) <= '" & Cmb_ano_ate & "'"
                Case "Ano": DataFiltro1 = "Year(Producaofases.Data) >= '" & Cmb_ano_de1 & "' and Year(Producaofases.Data) <= '" & Cmb_ano_ate1 & "'"
            End Select
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Producaofases.TempoTotal from Producaofases INNER JOIN ordemservico ON Producaofases.IDFase = ordemservico.IDproducao where ordemservico.Ordem = " & TBCarteira!Ordem & " and ordemservico.Retrabalho = 'True' and " & DataFiltro1 & " and Producaofases.CodigoDesc = 2", Conexao, adOpenKeyset, adLockReadOnly
            If TBFIltro.EOF = False Then
                Do While TBFIltro.EOF = False
                    TempoTotalProd = TempoTotalProd + IIf(IsNull(TBFIltro!TempoTotal), 0, TBFIltro!TempoTotal)
                    TBFIltro.MoveNext
                Loop
            End If
            TBFIltro.Close
            ElapsedTime (TempoTotalProd)
            TBProdutividade!qtdeOK = s 'Total de horas utilizadas para retrabalho (segundos)
        End If
        TBFI.Close
    Case "Selecionar": TBProdutividade!Servicos = TBProdutividade!Servicos + TotalCreditar
    Case "Reaproveitar": TBProdutividade!Total = TBProdutividade!Total + TotalCreditar
    Case "Outros": TBProdutividade!Total_peca = TBProdutividade!Total_peca + TotalCreditar
End Select

TBProdutividade.Update
TotalGeral = TBProdutividade!ID
TBProdutividade.Close

'Verifica se a OS já foi salva nesta data
Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios where ID = " & TotalGeral, Conexao, adOpenKeyset, adLockOptimistic
If TBProdutividade.EOF = False Then
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select numero, IDProd_Rel from Producao_Relatorios_detalhes where IDProd_Rel = " & TBProdutividade!ID & " and Numero = " & TBCarteira!OS, Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = True Then
        TBCFOP.AddNew
        TBCFOP!IDProd_Rel = TBProdutividade!ID
        TBCFOP!Numero = TBCarteira!OS
        TBCFOP.Update
        'Verifica a quantidade produzida
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select ap_backup from producao P INNER JOIN Ordemservico O ON O.ordem = P.ordem where O.IDproducao = " & TBCarteira!OS, Conexao, adOpenKeyset, adLockReadOnly
        If TBOrdem.EOF = False Then
            If TBOrdem!AP_backup = True Then
                TabelaBackup = "ProducaoFases_Backup"
            Else
                TabelaBackup = "ProducaoFases"
            End If
            Set TBCST = CreateObject("adodb.recordset")
            TBCST.Open "Select SUM(quantidade + Reprovada) as QtdeProd from " & TabelaBackup & " where OS = " & TBCarteira!OS & " and " & TextoFiltroQTOK & " and CodigoDesc = 2", Conexao, adOpenKeyset, adLockReadOnly
            If TBCST.EOF = False Then
                TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + IIf(IsNull(TBCST!QtdeProd), 0, TBCST!QtdeProd)
            End If
            TBCST.Close
        End If
        TBOrdem.Close
    End If
    TBCFOP.Close
    
    If TBProdutividade!qtdeNC <> 0 And TBProdutividade!QtdePrev <> 0 Then
        VlrTotal = IIf(IsNull(TBProdutividade!QtdePrev), 0, TBProdutividade!QtdePrev)
        TBProdutividade!Refugo = (TBProdutividade!qtdeNC / VlrTotal) * 100
        TBProdutividade!Numero2 = (TBProdutividade!Terceiros / VlrTotal) * 100
        TBProdutividade!Numero3 = (TBProdutividade!impostos / VlrTotal) * 100
        TBProdutividade!Numero4 = (TBProdutividade!Lucro / VlrTotal) * 100
        TBProdutividade!Numero5 = (TBProdutividade!material / VlrTotal) * 100
        TBProdutividade!Numero6 = (TBProdutividade!Servicos / VlrTotal) * 100
        TBProdutividade!Numero7 = (TBProdutividade!Total / VlrTotal) * 100
        TBProdutividade!Numero8 = (TBProdutividade!Total_peca / VlrTotal) * 100
    End If
    TBProdutividade.Update
End If
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaColunas()
On Error GoTo tratar_erro

Lista1.ColumnHeaders.Clear
contador = 1
TextoFiltro = " -  Prod.  |      NC       |      %      |    Aprov.   |      %      | Aprov. c/ des. |      %      |   Rej.   |      %      |  Retrab.  |      %      |   Selec.  |      %      |  Reapro.  |      %      |   Outros  |     %"
With Lista1.ColumnHeaders
    .Add
    .Item(contador).Text = cmbfiltrarpor.Text
    .Item(contador).Width = 3500
    If cmbPor.Text = "Dia" Then
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            .Add
            contador = contador + 1
            .Item(contador).Text = Format(Dataini, "dd/mm/yy") & TextoFiltro
            .Item(contador).Alignment = lvwColumnLeft
            .Item(contador).Width = 15000
            Dataini = Dataini + 1
        Loop
    End If
    If cmbPor.Text = "Mês" Then
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)
        Do While qt <= Qtd
            .Add
            contador = contador + 1
            .Item(contador).Text = Format(qt & "/" & Cmb_ano_de, "mm/yyyy") & TextoFiltro
            .Item(contador).Alignment = lvwColumnLeft
            .Item(contador).Width = 15000
            qt = qt + 1
        Loop
    End If
    If cmbPor.Text = "Ano" Then
        qt = Cmb_ano_de1
        Do While qt <= Cmb_ano_ate1
            .Add
            contador = contador + 1
            .Item(contador).Text = qt & TextoFiltro
            .Item(contador).Alignment = lvwColumnLeft
            .Item(contador).Width = 15000
            qt = qt + 1
        Loop
    End If
    .Add
    contador = contador + 1
    .Item(contador).Text = "Totais -  Prod.  |      NC       |      %      |    Aprov.   |      %      | Aprov. c/ des. |      %      |   Rej.   |      %      |  Retrab.  |      %      |   Selec.  |      %      |  Reapro.  |      %      |   Outros  |     %"
    .Item(contador).Alignment = lvwColumnLeft
    .Item(contador).Width = 14000
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro
Dim QTPC As Double
Dim QTNC As Double

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
Select Case cmbPor
    Case "Dia":
        Tipo = "D"
        TBAbrir!Totalutilizada = Format(msk_fltInicio.Value, "dd/mm/yy")
        TBAbrir!Totalprevista = Format(msk_fltFim.Value, "dd/mm/yy")
    Case "Mês":
        Tipo = "M"
        TBAbrir!Totalutilizada = Cmb_mes_de & "/" & Cmb_ano_de
        TBAbrir!Totalprevista = Cmb_mes_ate & "/" & Cmb_ano_ate
    Case "Ano":
        Tipo = "A"
        TBAbrir!Totalutilizada = Cmb_ano_de1
        TBAbrir!Totalprevista = Cmb_ano_ate1
End Select

If Opt_individual.Value = True Then
    If cmbTexto.Visible = True Then TBAbrir!Texto = cmbfiltrarpor & ") : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor & ") : " & txtTexto
Else
    TBAbrir!Texto = cmbfiltrarpor & ")"
End If
TBAbrir!Texto1 = Tipo
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario

QTPC = 0
QTNC = 0
DecimoSegundos = 0
qtdeliberada = 0
QuantsolicitadoN2 = 0
qtdeliberar = 0
QuantsolicitadoN3 = 0
qt = 0
QuantsolicitadoN4 = 0
Qtd = 0
QuantsolicitadoN5 = 0
Qtd_Prog = 0
QuantsolicitadoN6 = 0
QTLOTE = 0
QuantsolicitadoN7 = 0
quantestoque = 0
QuantsolicitadoN8 = 0
QuantsolicitadoN9 = 0
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(Qtdeprev) as QTPC, Sum(QtdeNC) as QTNC, Sum(Terceiros) as qtdeliberada, Sum(Impostos) as qtdeliberar, Sum(Lucro) as qt, Sum(material) as Qtd, Sum(QtdeOK) as DecimoSegundos, Sum(Servicos) as Qtd_Prog, Sum(Total) as QTLOTE, Sum(Total_peca) as quantestoque from Producao_relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    If optDetalhado.Value = True Then
        OSRepet = ""
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select Qtdeprev, OS from Producao_relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by ordem", Conexao, adOpenKeyset, adLockOptimistic
        Do While TBOrdem.EOF = False
            If OSRepet <> TBOrdem!OS Then
                QTPC = QTPC + IIf(IsNull(TBOrdem!QtdePrev), 0, TBOrdem!QtdePrev)   'Qtde. produzida
                OSRepet = TBOrdem!OS
            End If
            TBOrdem.MoveNext
        Loop
        TBOrdem.Close
    Else
        QTPC = IIf(IsNull(TBproducao!QTPC), 0, TBproducao!QTPC) 'Qtde. produzida
    End If
    QTNC = IIf(IsNull(TBproducao!QTNC), 0, TBproducao!QTNC) 'Qtde. NC
    
    'Não Conforme
    If QTPC <> 0 Then QuantsolicitadoN9 = (QTNC / QTPC) * 100 Else QuantsolicitadoN2 = 0 'Percentual (%)
    
    'Aprovado
    qtdeliberada = IIf(IsNull(TBproducao!qtdeliberada), 0, TBproducao!qtdeliberada)
    If QTPC <> 0 Then QuantsolicitadoN2 = (qtdeliberada / QTPC) * 100 Else QuantsolicitadoN2 = 0 'Percentual (%)
    
    'Aprovado c/ desvio
    qtdeliberar = IIf(IsNull(TBproducao!qtdeliberar), 0, TBproducao!qtdeliberar)
    If QTPC <> 0 Then QuantsolicitadoN3 = (qtdeliberar / QTPC) * 100 Else QuantsolicitadoN3 = 0 'Percentual (%)
    
    'Rejeitar
    qt = IIf(IsNull(TBproducao!qt), 0, TBproducao!qt)
    If QTPC <> 0 Then QuantsolicitadoN4 = (qt / QTPC) * 100 Else QuantsolicitadoN4 = 0 'Percentual (%)
    
    'Retrabalhar
    Qtd = IIf(IsNull(TBproducao!Qtd), 0, TBproducao!Qtd)
    If QTPC <> 0 Then QuantsolicitadoN5 = (Qtd / QTPC) * 100 Else QuantsolicitadoN5 = 0 'Percentual (%)
    DecimoSegundos = IIf(IsNull(TBproducao!DecimoSegundos), 0, TBproducao!DecimoSegundos) 'Horas de retrabalho
    
    'Selecionar
    Qtd_Prog = IIf(IsNull(TBproducao!Qtd_Prog), 0, TBproducao!Qtd_Prog)
    If QTPC <> 0 Then QuantsolicitadoN6 = (Qtd_Prog / QTPC) * 100 Else QuantsolicitadoN6 = 0 'Percentual (%)
    
    'Reaproveitar
    QTLOTE = IIf(IsNull(TBproducao!QTLOTE), 0, TBproducao!QTLOTE)
    If QTPC <> 0 Then QuantsolicitadoN7 = (QTLOTE / QTPC) * 100 Else QuantsolicitadoN7 = 0 'Percentual (%)
    
    'Outros
    quantestoque = IIf(IsNull(TBproducao!quantestoque), 0, TBproducao!quantestoque)
    If QTPC <> 0 Then QuantsolicitadoN8 = (quantestoque / QTPC) * 100 Else QuantsolicitadoN8 = 0 'Percentual (%)
End If
TBproducao.Close

TBAbrir!QtdeProduzida = QTPC
TBAbrir!qtdeNC = QTNC
TBAbrir!CustoMat = qtdeliberada
TBAbrir!CustoObra = QuantsolicitadoN2
TBAbrir!Terceros = qtdeliberar
TBAbrir!Lucro = QuantsolicitadoN3
TBAbrir!Valor1 = qt
TBAbrir!Valor2 = QuantsolicitadoN4
TBAbrir!Valor3 = Qtd
TBAbrir!Total1 = QuantsolicitadoN5
TBAbrir!TotalEficiencia = FormataTempo(DecimoSegundos)
TBAbrir!Total2 = Qtd_Prog
TBAbrir!Numero1 = QuantsolicitadoN6
TBAbrir!Numero2 = QTLOTE
TBAbrir!Numero3 = QuantsolicitadoN7
TBAbrir!Numero4 = quantestoque
TBAbrir!Numero5 = QuantsolicitadoN8
TBAbrir!Numero6 = QuantsolicitadoN9
TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ProcLimpaListaeCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
    txtTexto = ""
    txtTexto.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
    txtTexto.Enabled = True
    ProcCarregaComboTexto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    ProcLimpaListaeCampos
    Lista.Visible = True
    Lista1.Visible = False
    With cmbPor
        .Clear
        .AddItem "Dia"
        .Text = "Dia"
    End With
    ProcMostrarEsconderCombosData
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    ProcLimpaListaeCampos
    Lista.Visible = False
    Lista1.Visible = True
    With cmbPor
        .Clear
        .AddItem "Dia"
        .AddItem "Mês"
        .AddItem "Ano"
        .Text = "Dia"
    End With
    ProcMostrarEsconderCombosData
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarEsconderCombosData()
On Error GoTo tratar_erro

If cmbPor = "Dia" Then
    msk_fltInicio.Visible = True
    msk_fltFim.Visible = True
    Cmb_mes_de.Visible = False
    Cmb_mes_ate.Visible = False
    Cmb_ano_de.Visible = False
    Cmb_ano_ate.Visible = False
    Cmb_ano_de1.Visible = False
    Cmb_ano_ate1.Visible = False
ElseIf cmbPor = "Mês" Then
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = True
        Cmb_mes_ate.Visible = True
        Cmb_ano_de.Visible = True
        Cmb_ano_ate.Visible = True
        Cmb_ano_de1.Visible = False
        Cmb_ano_ate1.Visible = False
    Else
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = False
        Cmb_mes_ate.Visible = False
        Cmb_ano_de.Visible = False
        Cmb_ano_ate.Visible = False
        Cmb_ano_de1.Visible = True
        Cmb_ano_ate1.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ProcLimpaListaeCampos
If txtTexto <> "" Then
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
