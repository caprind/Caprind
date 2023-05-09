VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProd_Relatorios_Produtividade 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Relatórios - Produtividade"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
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
   Icon            =   "frmProd_Relatorios_Produtividade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   1110
      Left            =   13470
      TabIndex        =   40
      Top             =   990
      Width           =   1875
      Begin VB.CheckBox Chk_filtrar_backup 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar do backup"
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
         Left            =   90
         TabIndex        =   13
         Top             =   30
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   540
         TabIndex        =   15
         ToolTipText     =   "Data final."
         Top             =   690
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   157024257
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   540
         TabIndex        =   14
         ToolTipText     =   "Data inicio."
         Top             =   330
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   157024257
         CurrentDate     =   39057
      End
      Begin VB.Label Label6 
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
         Left            =   150
         TabIndex        =   42
         Top             =   390
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
         Left            =   90
         TabIndex        =   41
         Top             =   750
         Width           =   360
      End
   End
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
      Height          =   855
      Left            =   60
      TabIndex        =   29
      Top             =   8880
      Width           =   15255
      Begin VB.TextBox Txt_eficiencia 
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
         Left            =   13950
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Eficiência média."
         Top             =   390
         Width           =   1050
      End
      Begin VB.TextBox Txt_eficiencia_exec 
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
         Left            =   12890
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Eficiência de execução."
         Top             =   390
         Width           =   1050
      End
      Begin VB.TextBox Txt_tt_exec_prev 
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
         Left            =   3080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total de execução previsto."
         Top             =   390
         Width           =   1430
      End
      Begin VB.TextBox Txt_tt_prep_util 
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
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total previsto de utilizado."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_tt_exec_util 
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
         Left            =   4509
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total de execução utilizado."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_tt_prep_prev 
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
         MaxLength       =   20
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total previsto de preparação."
         Top             =   390
         Width           =   1430
      End
      Begin VB.TextBox Txt_qtde_produzida 
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
         Left            =   10335
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total produzida."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_horas_utilizadas 
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
         Left            =   8885
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Total de horas utilizadas."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_horas_previstas 
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
         Left            =   5975
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Total de horas previstas."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_qtde_prevista 
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
         Left            =   7425
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total prevista."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_eficiencia_prep 
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
         Left            =   11790
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Eficiência de preparação."
         Top             =   390
         Width           =   1080
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic.média"
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
         Left            =   14175
         TabIndex        =   51
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic. exec."
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
         Left            =   13065
         TabIndex        =   50
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " TT exec. util."
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
         Left            =   4800
         TabIndex        =   49
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " TT prep. util."
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
         Left            =   1920
         TabIndex        =   48
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TT exec. prev."
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
         Left            =   3330
         TabIndex        =   47
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TT prep. prev."
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
         Left            =   435
         TabIndex        =   46
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic. prep."
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
         Left            =   11985
         TabIndex        =   34
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total hs. prev."
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
         Left            =   6225
         TabIndex        =   33
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total prev."
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
         Left            =   7605
         TabIndex        =   32
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total hs. util."
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
         Left            =   9195
         TabIndex        =   31
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total prod."
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
         Left            =   10515
         TabIndex        =   30
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1095
      Left            =   55
      TabIndex        =   35
      Top             =   990
      Width           =   1275
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1095
      Left            =   1350
      TabIndex        =   36
      Top             =   990
      Width           =   1305
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   1095
      Left            =   2670
      TabIndex        =   37
      Top             =   990
      Width           =   10785
      Begin VB.CheckBox Chk_setor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Setor"
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
         Left            =   5415
         TabIndex        =   7
         Top             =   270
         Width           =   705
      End
      Begin VB.ComboBox Cmb_setor 
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
         ItemData        =   "frmProd_Relatorios_Produtividade.frx":0442
         Left            =   5430
         List            =   "frmProd_Relatorios_Produtividade.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Setor."
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox Txt_turno 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9900
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Turno"
         Top             =   480
         Width           =   795
      End
      Begin VB.CheckBox Chk_turno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Turno"
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
         Left            =   9930
         TabIndex        =   11
         Top             =   270
         Width           =   735
      End
      Begin VB.CheckBox Chk_fase 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fase"
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
         Left            =   9075
         TabIndex        =   9
         Top             =   270
         Width           =   705
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
         ItemData        =   "frmProd_Relatorios_Produtividade.frx":0446
         Left            =   180
         List            =   "frmProd_Relatorios_Produtividade.frx":045C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox Txt_fase 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Fase."
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox txtTexto 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2190
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Width           =   3225
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
         ItemData        =   "frmProd_Relatorios_Produtividade.frx":04A6
         Left            =   2190
         List            =   "frmProd_Relatorios_Produtividade.frx":04A8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Visible         =   0   'False
         Width           =   3225
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   750
         TabIndex        =   39
         Top             =   270
         Width           =   840
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
         Left            =   3225
         TabIndex        =   38
         Top             =   270
         Width           =   1500
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   43
      Top             =   9750
      Width           =   11565
      _ExtentX        =   20399
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
      BackColor       =   16777215
      BarColor1       =   16777215
      BarColor2       =   14737632
      BorderColor     =   8421504
      ForeColor2      =   0
      SearchText      =   "Atualizando..."
      Theme           =   1
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   44
      Top             =   30
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   1720
      ButtonCount     =   6
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
      ButtonLeft3     =   108
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
      ButtonLeft4     =   112
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
      ButtonLeft5     =   155
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
      ButtonLeft6     =   187
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7350
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProd_Relatorios_Produtividade.frx":04AA
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6765
      Left            =   30
      TabIndex        =   16
      Top             =   2100
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   11933
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
      NumItems        =   28
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
         Text            =   "Setor"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1776
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   1776
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Fase"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Grupo/op."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Posto de trabalho"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Operador"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Turno"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "D"
         Text            =   "Prep. prev."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Tag             =   "D"
         Text            =   "Prep. util."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "D"
         Text            =   "Exec. prev."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Tag             =   "D"
         Text            =   "Exec. util."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "D"
         Text            =   "Hs. previstas"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Tag             =   "N"
         Text            =   "Qtde. prev."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "D"
         Text            =   "Hs. utilizadas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Tag             =   "N"
         Text            =   "Qtde. apont."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Tag             =   "N"
         Text            =   "Qtde. OK"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Tag             =   "N"
         Text            =   "Qtde. NC"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Object.Tag             =   "N"
         Text            =   "Qtde. acumul. OS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Tag             =   "N"
         Text            =   "Efic. prep."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Object.Tag             =   "N"
         Text            =   "Efic. exec."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Tag             =   "N"
         Text            =   "Efic. média"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6765
      Left            =   30
      TabIndex        =   17
      Top             =   2100
      Visible         =   0   'False
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   11933
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Object.Width           =   5512
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
         Text            =   "Cód. de ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5962
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "D"
         Text            =   "Prep. prev."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "D"
         Text            =   "Prep. util."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "D"
         Text            =   "Exec. prev."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "D"
         Text            =   "Exec. util."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Hs. previstas"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Qtde. prev."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "D"
         Text            =   "Hs. utilizadas"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Qtde. apont."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Efic. prep."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Efic. exec."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "Efic. média"
         Object.Width           =   1587
      EndProperty
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   12240
      TabIndex        =   45
      Top             =   9780
      Width           =   2895
   End
End
Attribute VB_Name = "frmProd_Relatorios_Produtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnLista_Click()
On Error GoTo tratar_erro

frmProd_configurar_listas.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_fase_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
With Txt_fase
    If Chk_fase.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
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

Private Sub Chk_filtrar_backup_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_setor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
With Cmb_setor
    If Chk_setor.Value = 1 Then
        .Clear
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select Setor from CadMaquinas Group by Setor", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            Do While TBMaquinas.EOF = False
                .AddItem TBMaquinas!Setor
                TBMaquinas.MoveNext
            Loop
        End If
        TBMaquinas.Close
        
        .Locked = False
        .TabStop = True
        .SetFocus
    Else
        .ListIndex = -1
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_turno_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
With Txt_turno
    If Chk_turno.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
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

Private Sub Cmb_setor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 And Lista1.ListItems.Count = 0 Then Exit Sub
frmProd_Relatorios_Produtividade_menu_impressao.Show 1

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
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Posicao = 0
Lista.ListItems.Clear
Lista1.ListItems.Clear
Quant = 0
OS = 0
If TBLISTA.EOF = False Then
    
    Posicao = TBLISTA.RecordCount
    OS = TBLISTA!OS
    
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from producao where Ordem = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Produto), "", TBAbrir!Produto)
                End If
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!OS), "", TBLISTA!OS)
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Fase), "", TBLISTA!Fase)
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from ordemservico where idproducao = " & TBLISTA!OS, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!Grupo_op), "", TBAbrir!Grupo_op)
                End If
                TBAbrir.Close
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento)
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from Cadmaquinas where Maquina = '" & TBLISTA!DescEvento & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    .Item(.Count).SubItems(11) = IIf(IsNull(TBFIltro!Descricao), "", TBFIltro!Descricao)
                End If
                TBFIltro.Close
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Data4), "", TBLISTA!Data4)
                .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Turno), "", TBLISTA!Turno)
                .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Data1), "", TBLISTA!Data1)
                .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Data2), "", AcertaHora(TBLISTA!Data2))
                .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!Execucaoprev), "", AcertaHora(TBLISTA!Execucaoprev))
                .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA!Data3), "", TBLISTA!Data3)
                .Item(.Count).SubItems(18) = IIf(IsNull(TBLISTA!Totalhsprev), "", AcertaHora(TBLISTA!Totalhsprev))
                .Item(.Count).SubItems(19) = IIf(IsNull(TBLISTA!QtdePrev), "", Format(TBLISTA!QtdePrev, "###,##0.00"))
                .Item(.Count).SubItems(20) = IIf(IsNull(TBLISTA!Totalhsutil), "", AcertaHora(TBLISTA!Totalhsutil))
                .Item(.Count).SubItems(21) = IIf(IsNull(TBLISTA!Qtdetotalprod), "", Format(TBLISTA!Qtdetotalprod, "###,##0.00"))
                .Item(.Count).SubItems(22) = IIf(IsNull(TBLISTA!qtdeOK), "", Format(TBLISTA!qtdeOK, "###,##0.00"))
                .Item(.Count).ListSubItems(22).ForeColor = vbBlue
                .Item(.Count).SubItems(23) = IIf(IsNull(TBLISTA!qtdeNC), "", Format(TBLISTA!qtdeNC, "###,##0.00"))
                
                If TBLISTA!qtdeNC > 0 Then
                .Item(.Count).ListSubItems(23).ForeColor = vbRed
                End If
                
                If OS = TBLISTA!OS Then
                    Quant = Quant + TBLISTA!Qtdetotalprod
                Else
                    Quant = 0
                    Quant = Quant + TBLISTA!Qtdetotalprod
                End If
                .Item(.Count).SubItems(24) = Format(Quant, "###,##0.00")
                
                .Item(.Count).SubItems(25) = IIf(IsNull(TBLISTA!Numero1), "", Format(TBLISTA!Numero1, "###,##0.00") & "%")
                .Item(.Count).SubItems(26) = IIf(IsNull(TBLISTA!Numero2), "", Format(TBLISTA!Numero2, "###,##0.00") & "%")
                .Item(.Count).SubItems(27) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%")
                If TBLISTA!Eficiencia < 80 Then
                .Item(.Count).ListSubItems(27).ForeColor = vbRed
                End If
            End With
        Else
            With Lista1.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                If cmbfiltrarpor = "Ordem" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from producao where Ordem = " & TBLISTA!maquina, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                        .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                        .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Produto), "", TBAbrir!Produto)
                    End If
                    TBAbrir.Close
                End If

                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data1), "", TBLISTA!Data1)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Data2), "", TBLISTA!Data2)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Execucaoprev), "", AcertaHora(TBLISTA!Execucaoprev))
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Data3), "", AcertaHora(TBLISTA!Data3))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Totalhsprev), "", AcertaHora(TBLISTA!Totalhsprev))
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!QtdePrev), "", Format(TBLISTA!QtdePrev, "###,##0.00"))
                .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Totalhsutil), "", c)
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Qtdetotalprod), "", Format(TBLISTA!Qtdetotalprod, "###,##0.00"))
                .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Numero1), "", Format(TBLISTA!Numero1, "###,##0.00") & "%")
                .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Numero2), "", Format(TBLISTA!Numero2, "###,##0.00") & "%")
                .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Eficiencia), "", Format(TBLISTA!Eficiencia, "###,##0.00") & "%")
            End With
        End If
        
        OS = TBLISTA!OS
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    If optDetalhado.Value = True Then Else
End If

ProcLimpaCamposTotais
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_tt_prep_prev = AcertaHora(TBLISTA!Totalprevista)
    Txt_tt_exec_prev = AcertaHora(TBLISTA!Data1)
    Txt_tt_prep_util = AcertaHora(TBLISTA!Totalutilizada)
    Txt_tt_exec_util = AcertaHora(TBLISTA!Data2)
    Txt_horas_previstas = AcertaHora(TBLISTA!Data3)
    Txt_qtde_prevista = Format(TBLISTA!QtdePrevista, "###,##0.00")
    Txt_horas_utilizadas = AcertaHora(TBLISTA!Data4)
    Txt_qtde_produzida = Format(TBLISTA!QtdeProduzida, "###,##0.00")
    Txt_eficiencia_prep = Format(TBLISTA!Valor1, "###,##0.00") & "%"
    Txt_eficiencia_exec = Format(TBLISTA!Valor2, "###,##0.00") & "%"
    Txt_eficiencia = Format(TBLISTA!TotalEficiencia, "###,##0.00") & "%"
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
Txt_tt_prep_prev = ""
Txt_tt_prep_util = ""
Txt_tt_exec_prev = ""
Txt_tt_exec_util = ""
Txt_qtde_prevista = ""
Txt_horas_previstas = ""
Txt_qtde_produzida = ""
Txt_horas_utilizadas = ""
Txt_eficiencia_prep = ""
Txt_eficiencia_exec = ""
Txt_eficiencia = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 17700, 6, True
Formulario = "PCP/Relatórios/Produtividade"
Direitos
ProcVerifColunas
ProcLimpaVariaveisPrincipais
cmbfiltrarpor.Text = "Ordem"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "PCP/Relatórios/Produtividade"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcVerifColunas()
On Error GoTo tratar_erro

ProcCorrigeColunasForm Lista, Formulario, 28, False, 0, 1200, 1300, 1007, 1200, 1400, 2500, 1006, 600, 1000, 1550, 2500, 1200, 1200, 1200, 1200, 1440, 1200, 1200, 1200, 1500, 700, 1200, 1200, 1500, 1200, 1200, 1200, 1200, 0, 0

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

frmProd_configurar_listas.Show 1

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

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

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

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

With Chk_setor
    .Value = 0
    .Enabled = True
End With
Chk_fase.Value = 0
Txt_fase = ""
With Chk_turno
    .Value = 0
    .Enabled = True
End With
Txt_turno = ""
txtTexto = ""
txtTexto.Visible = False

ProcListaPadrao

With cmbTexto
    .Clear
    .Visible = True
    If Opt_individual.Value = True Then
        Select Case cmbfiltrarpor
            Case "Operador":
                Set TBUsuarios = CreateObject("adodb.recordset")
                TBUsuarios.Open "Select usuario FROM usuarios Group by usuario", Conexao, adOpenKeyset, adLockOptimistic
                If TBUsuarios.EOF = False Then
                    Do While TBUsuarios.EOF = False
                        .AddItem UCase(TBUsuarios!Usuario)
                        TBUsuarios.MoveNext
                    Loop
                End If
                TBUsuarios.Close
            Case "Posto de trabalho":
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select maquina from CadMaquinas Group by Maquina", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    Do While TBMaquinas.EOF = False
                        .AddItem UCase(Trim(TBMaquinas!maquina))
                        TBMaquinas.MoveNext
                    Loop
                End If
                TBMaquinas.Close
                Chk_setor.Enabled = True
            Case "Turno":
                .AddItem 0
                .AddItem 1
                .AddItem 2
                .AddItem 3
                .AddItem 4
            Case "Setor":
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select Setor from CadMaquinas Group by Setor", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    Do While TBMaquinas.EOF = False
                        .AddItem UCase(TBMaquinas!Setor)
                        TBMaquinas.MoveNext
                    Loop
                End If
                TBMaquinas.Close
            Case "Ordem":
                .Visible = False
                txtTexto.Visible = True
            Case "Cód. de referência":
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select N_Referencia from Producao where N_Referencia <> 'Null' Group by N_Referencia", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Do While TBAbrir.EOF = False
                        If TBAbrir!N_referencia <> "" Then .AddItem TBAbrir!N_referencia
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
        End Select
    End If
End With
'Lista.ColumnHeaders(3).Text = cmbfiltrarpor
'Lista1.ColumnHeaders(2).Text = cmbfiltrarpor

Select Case cmbfiltrarpor
    Case "Ordem": ProcListaOrdem
    Case "Turno":
        With Chk_turno
            .Value = 0
            .Enabled = False
        End With
    Case "Setor":
        With Chk_setor
            .Value = 0
            .Enabled = False
        End With
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcListaPadrao()
On Error GoTo tratar_erro

With Lista.ColumnHeaders
    .Item(3).Width = 3100
    If cmbfiltrarpor = "Posto de trabalho" Then
        .Item(11).Width = 0
        .Item(12).Width = 0
    Else
        .Item(11).Width = 1550
        .Item(12).Width = 2500
    End If
    If cmbfiltrarpor = "Operador" Then .Item(13).Width = 0 Else .Item(13).Width = 2500
    If cmbfiltrarpor = "Turno" Then .Item(14).Width = 0 Else .Item(14).Width = 700
End With
With Lista1.ColumnHeaders
    .Item(2).Width = 3100
    .Item(3).Width = 0
    .Item(4).Width = 0
    .Item(5).Width = 0
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcListaOrdem()
On Error GoTo tratar_erro

Lista.ColumnHeaders.Item(3).Width = 0
With Lista1.ColumnHeaders
    .Item(2).Width = 1200
    .Item(3).Width = 1200
    .Item(4).Width = 1400
    .Item(5).Width = 3380
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
End If
If Opt_individual.Value = True Then
    If cmbfiltrarpor = "Ordem" And txtTexto = "" Or cmbfiltrarpor <> "Ordem" And cmbTexto = "" Then
        NomeCampo = "o texto para pesquisa"
        ProcVerificaAcao
        If cmbfiltrarpor = "Ordem" Then txtTexto.SetFocus Else cmbTexto.SetFocus
        Exit Sub
    End If
End If
If Chk_setor.Value = 1 And Cmb_setor = "" Then
    NomeCampo = "o setor"
    ProcVerificaAcao
    Cmb_setor.SetFocus
    Exit Sub
End If
If Chk_fase.Value = 1 And Txt_fase = "" Then
    NomeCampo = "a fase"
    ProcVerificaAcao
    Txt_fase.SetFocus
    Exit Sub
End If
If Chk_turno.Value = 1 And Txt_turno = "" Then
    NomeCampo = "o turno"
    ProcVerificaAcao
    Txt_turno.SetFocus
    Exit Sub
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

Inicio = Time
ProcLimpaCamposTotais
ProcAbrirTabelas
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then
        TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Ordem, OS, Data, Turno", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data", Conexao, adOpenKeyset, adLockOptimistic
    End If
Else
    If cmbfiltrarpor = "Ordem" Then Ordenar = "Ordem" Else Ordenar = "Maquina"
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
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

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

With frmProd_Relatorios_Produtividade
    If TBproducao.EOF = False Then
        Permitido = True
        
        If Opt_individual.Value = True Then
            TBproducao.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBproducao.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBproducao.MoveFirst
        End If
        
        Do While TBproducao.EOF = False
            'Tempo de preparação
            ProcFormataHora (IIf(IsNull(TBproducao!Preparacao), 0, TBproducao!Preparacao))
            TPPSEG = s + DecimoSegundos
            ProcFormataHora (TBproducao!TPUTIL)
            TPUSEG = s + DecimoSegundos
            
            'Tempo de execução
            ProcFormataHora (IIf(IsNull(TBproducao!Execucao), 0, TBproducao!Execucao))
            TEPSEG = s + DecimoSegundos
            ProcFormataHora (TBproducao!TEUTIL)
            TEUSEG = s + DecimoSegundos
            
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select * from ordemservico where idproducao = " & TBproducao!OS, Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = False Then
                Set TBProdutividade = CreateObject("adodb.recordset")
                If Opt_individual.Value = True And optDetalhado.Value = True Then
                    TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
                    ProcEnviaDadosDetalhado
                Else
                    Select Case cmbfiltrarpor
                        Case "Operador": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBproducao!Usuario & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Case "Posto de trabalho": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBproducao!maquina & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Case "Turno": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBproducao!Turno & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Case "Setor": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBproducao!Setor & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Case "Ordem": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBproducao!Ordem & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Case "Cód. de referência": TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TBproducao!N_referencia & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    End Select
                    ProcEnviaDadosResumido
                End If
                TBProdutividade.Update
                TBProdutividade.Close
            End If
            TBOrdem.Close
            TBproducao.MoveNext
            
            If Opt_individual.Value = True Then
                Contador = Contador + 1
                .PBLista.Value = Contador
            End If
        Loop
    End If
    TBproducao.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

TPPSEG = 0 'Tempo preparação previsto
TPUSEG = 0 'Tempo preparação real
TEPSEG = 0 'Tempo execução previsto
TEUSEG = 0 'Tempo execução real
TotalPreparacao = 0 'Tempo total previsto
TTE = 0 'Tempo total real
Qtde = 0
Permitido = False
OF = 0
OS = 0
Dataini = 0
maquina = ""

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

If Chk_filtrar_backup.Value = 1 Then
    NomeTabelaAp = "ProducaoFases_Backup"
    NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao_Backup"
Else
    NomeTabelaAp = "ProducaoFases"
    NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao"
End If

ProcVerifFiltroFaseTurno
If Opt_individual.Value = True Then
    Set TBproducao = CreateObject("adodb.recordset")
    Select Case cmbfiltrarpor
        Case "Operador": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where usuario = '" & cmbTexto & "' and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Posto de trabalho": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where maquina = '" & cmbTexto & "' and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Turno": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where turno = " & cmbTexto & " and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Setor": TBproducao.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where CadMaquinas.Setor = '" & cmbTexto & "' and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Ordem": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & txtTexto & " and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Cód. de referência": TBproducao.Open "Select Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".* FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where Producao.N_Referencia = '" & cmbTexto & "' and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
    End Select
    ProcFiltrar1
Else
    Ordem = 0
    Set TBCarteira = CreateObject("adodb.recordset")
    Select Case cmbfiltrarpor
        Case "Operador":
            TBCarteira.Open "Select * from " & NomeTabelaApTotalizacao & " where Usuario is not null and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                TBCarteira.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCarteira.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCarteira.MoveFirst
                Do While TBCarteira.EOF = False
                    If Ordem <> TBCarteira!Ordem Then
                        Set TBproducao = CreateObject("adodb.recordset")
                        TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and Ordem = " & TBCarteira!Ordem & " order by Usuario, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                        ProcFiltrar1
                    End If
                    Ordem = TBCarteira!Ordem
                    TBCarteira.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        Case "Posto de trabalho":
            TBCarteira.Open "Select * from " & NomeTabelaApTotalizacao & " where Maquina is not null and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                TBCarteira.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCarteira.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCarteira.MoveFirst
                Do While TBCarteira.EOF = False
                    If Ordem <> TBCarteira!Ordem Then
                        Set TBproducao = CreateObject("adodb.recordset")
                        If Chk_setor.Value = 0 Then
                            TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and Ordem = " & TBCarteira!Ordem & " order by Maquina, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                        Else
                            TBproducao.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where CadMaquinas.setor = '" & Cmb_setor & "' and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and " & NomeTabelaApTotalizacao & ".Ordem = " & TBCarteira!Ordem & " order by " & NomeTabelaApTotalizacao & ".Maquina, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".Data, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                        End If
                        ProcFiltrar1
                    End If
                    Ordem = TBCarteira!Ordem
                    TBCarteira.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        Case "Turno":
            TBCarteira.Open "Select * from " & NomeTabelaApTotalizacao & " where Turno is not null and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                TBCarteira.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCarteira.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCarteira.MoveFirst
                Do While TBCarteira.EOF = False
                    If Ordem <> TBCarteira!Ordem Then
                        Set TBproducao = CreateObject("adodb.recordset")
                        TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and Ordem = " & TBCarteira!Ordem & " order by Turno, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                        ProcFiltrar1
                    End If
                    Ordem = TBCarteira!Ordem
                    TBCarteira.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        Case "Setor":
            TBCarteira.Open "Select * from " & NomeTabelaApTotalizacao & " where Maquina is not null and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                TBCarteira.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCarteira.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCarteira.MoveFirst
                Do While TBCarteira.EOF = False
                    If Ordem <> TBCarteira!Ordem Then
                        Set TBproducao = CreateObject("adodb.recordset")
                        TBproducao.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and " & NomeTabelaApTotalizacao & ".Ordem = " & TBCarteira!Ordem & " order by " & NomeTabelaApTotalizacao & ".Maquina, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".Data, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                        ProcFiltrar1
                    End If
                    Ordem = TBCarteira!Ordem
                    TBCarteira.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        Case "Ordem":
            TBCarteira.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                TBCarteira.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCarteira.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCarteira.MoveFirst
                Do While TBCarteira.EOF = False
                    If Ordem <> TBCarteira!Ordem Then
                        Set TBproducao = CreateObject("adodb.recordset")
                        TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and Ordem = " & TBCarteira!Ordem & " order by OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                        ProcFiltrar1
                    End If
                    Ordem = TBCarteira!Ordem
                    TBCarteira.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        Case "Cód. de referência":
            TBCarteira.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Data, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                TBCarteira.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCarteira.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCarteira.MoveFirst
                Do While TBCarteira.EOF = False
                    If Ordem <> TBCarteira!Ordem Then
                        Set TBproducao = CreateObject("adodb.recordset")
                        TBproducao.Open "Select Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".* FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where Producao.N_Referencia is not null and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and " & NomeTabelaApTotalizacao & ".Ordem = " & TBCarteira!Ordem & " order by Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".Data, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                        ProcFiltrar1
                    End If
                    Ordem = TBCarteira!Ordem
                    TBCarteira.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
    End Select
    TBCarteira.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas1()
On Error GoTo tratar_erro

ProcVerifFiltroFaseTurno
Set TBproducao = CreateObject("adodb.recordset")
If Opt_individual.Value = True Then
    Select Case cmbfiltrarpor
        Case "Operador": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where usuario = '" & cmbTexto & "' and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Posto de trabalho": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where maquina = '" & cmbTexto & "' and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Turno": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where turno = " & cmbTexto & " and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Setor": TBproducao.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where CadMaquinas.Setor = '" & cmbTexto & "' and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Ordem": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & txtTexto & " and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Cód. de referência": TBproducao.Open "Select Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".* FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where Producao.N_Referencia = '" & cmbTexto & "' and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
    End Select
Else
    Select Case cmbfiltrarpor
        Case "Operador": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Idproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Posto de trabalho":
            If Chk_setor.Value = 0 Then
                TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            Else
                TBproducao.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where CadMaquinas.Setor = '" & Cmb_setor & "' and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
            End If
        Case "Turno": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Setor": TBproducao.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Ordem": TBproducao.Open "Select * from " & NomeTabelaApTotalizacao & " where " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS, Idproducao", Conexao, adOpenKeyset, adLockOptimistic
        Case "Cód. de referência": TBproducao.Open "Select Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".* FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS, " & NomeTabelaApTotalizacao & ".IDproducao", Conexao, adOpenKeyset, adLockOptimistic
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

TBProdutividade.AddNew
TBProdutividade!Data = Format(TBproducao!Data, "dd/mm/yy")
TBProdutividade!Ordem = TBproducao!Ordem
TBProdutividade!OS = TBproducao!OS
TBProdutividade!Fase = TBproducao!Fase
TBProdutividade!DescEvento = TBproducao!maquina
TBProdutividade!Data4 = TBproducao!Usuario
TBProdutividade!Turno = TBproducao!Turno
TBProdutividade!Numero1 = TBproducao!Eficiencia_prep
TBProdutividade!Numero2 = TBproducao!Eficiencia_exec
TBProdutividade!Eficiencia = TBproducao!Eficiencia


'Qtde. N/C
TTOK = 0
TTNC = 0
QtdeSaida = 0

If OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then
    Select Case cmbfiltrarpor
        Case "Operador":
            Set TBUsuarios = CreateObject("adodb.recordset")
            TBUsuarios.Open "Select * from Usuarios where Usuario = '" & TBproducao!Usuario & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBUsuarios.EOF = False Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Operador = '" & TBUsuarios!CODIGO & "-" & TBproducao!Usuario & "' and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Do While TBAbrir.EOF = False
                        If TBAbrir!ParecerCQ = "Rejeitar" Then
                            If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                        Else
                            If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                        End If
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
                TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
                TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
                TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
            End If
            TBUsuarios.Close
        Case "Posto de trabalho":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Maquina = '" & TBproducao!maquina & "' and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Turno":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Turno = " & TBproducao!Turno & " and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Setor":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Setor = '" & TBproducao!Setor & "' and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Ordem":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Ordem = " & TBproducao!Ordem & " and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Cód. de referência":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Ordem = " & TBproducao!Ordem & " and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
    End Select
End If
Quant = TBproducao!QTOK + TTNC

'Qtde. prevista
'Verif. quantidade de peças previstas x tempo utilizado
If TPPSEG > 0 And TPUSEG > 0 And TEPSEG > 0 And TEUSEG > 0 Then
    Qtde = (TPUSEG / TPPSEG) + (TEUSEG / TEPSEG)
ElseIf TPPSEG > 0 And TPUSEG > 0 Then
        Qtde = TPUSEG / TPPSEG
    ElseIf TEPSEG > 0 And TEUSEG > 0 Then
            Qtde = TEUSEG / TEPSEG
        Else
            Qtde = 0
End If
If Quant <> 0 Then TBProdutividade!QtdePrev = Qtde * Quant Else TBProdutividade!QtdePrev = Qtde
'Tempo preparação previsto
VlttTotal = TPPSEG
TBProdutividade!Data1 = FormataTempo(VlttTotal)
'Tempo preparação utilizado
VlttTotal = TPUSEG
TBProdutividade!Data2 = FormataTempo(VlttTotal)
'Tempo execução previsto
If Quant <> 0 Then VlttTotal = TEPSEG * Quant Else VlttTotal = TEPSEG
TBProdutividade!Execucaoprev = FormataTempo(VlttTotal)
'Tempo execução utilizado
If Quant <> 0 Then VlttTotal = TEUSEG * Quant Else VlttTotal = TEUSEG
TBProdutividade!Data3 = FormataTempo(VlttTotal)
'Tempo total de execução prevista
If Quant <> 0 Then DecimoSegundos = TPPSEG + (TEPSEG * Quant) Else DecimoSegundos = TPPSEG + TPUSEG
VlttTotal = DecimoSegundos
TBProdutividade!Totalhsprev = FormataTempo(VlttTotal)
'Tempo de preparação utilizado + (execução utilizado X qtde. produzida)
If Quant <> 0 Then DecimoSegundos = TPUSEG + (TEUSEG * Quant) Else DecimoSegundos = TPUSEG + TEUSEG
VlttTotal = DecimoSegundos
TBProdutividade!Totalhsutil = FormataTempo(VlttTotal)
'Qtde OK
TBProdutividade!qtdeOK = TBproducao!QTOK
'Qtde NC
TBProdutividade!qtdeNC = TBproducao!QTNC
'Qtde. produzida
TBProdutividade!Qtdetotalprod = Quant
TBProdutividade!maquina = cmbTexto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
OF = TBproducao!Ordem
OS = TBproducao!OS
Dataini = TBproducao!Data

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

If TBProdutividade.EOF = True Then TBProdutividade.AddNew
If Opt_individual.Value = True Then
    If cmbTexto.Visible = True Then Texto = cmbTexto Else Texto = txtTexto
Else
    Select Case cmbfiltrarpor
        Case "Operador": Texto = TBproducao!Usuario
        Case "Posto de trabalho": Texto = TBproducao!maquina
        Case "Turno": Texto = TBproducao!Turno
        Case "Setor": Texto = TBproducao!Setor
        Case "Ordem":
            Texto = TBproducao!Ordem
            TBProdutividade!Ordem = Texto
        Case "Cód. de referência": Texto = TBproducao!N_referencia
    End Select
End If
TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario

If Chk_fase.Value = 1 Then TBProdutividade!Fase = TBproducao!Fase
If Chk_turno.Value = 1 Then TBProdutividade!Turno = TBproducao!Turno

'Tempo preparação previsto
If IsNull(TBProdutividade!Data1) = False And TBProdutividade!Data1 <> "" Then ProcFormataHora (TBProdutividade!Data1) Else ProcFormataHora ("00:00:00")
DecimoSegundos = s + DecimoSegundos
VlttTotal = TPPSEG + DecimoSegundos
TBProdutividade!Data1 = FormataTempo(VlttTotal)
'Tempo preparação utilizado
If IsNull(TBProdutividade!Data2) = False And TBProdutividade!Data2 <> "" Then ProcFormataHora (TBProdutividade!Data2) Else ProcFormataHora ("00:00:00")
DecimoSegundos = s + DecimoSegundos
VlttTotal = TPUSEG + DecimoSegundos
TBProdutividade!Data2 = FormataTempo(VlttTotal)
'Tempo execução previsto
If IsNull(TBProdutividade!Execucaoprev) = False And TBProdutividade!Execucaoprev <> "" Then ProcFormataHora (TBProdutividade!Execucaoprev) Else ProcFormataHora ("00:00:00")
DecimoSegundos = s + DecimoSegundos
If TBproducao!Totalprod <> 0 Then VlttTotal = (TEPSEG * TBproducao!Totalprod) + DecimoSegundos Else VlttTotal = TEPSEG + DecimoSegundos
TBProdutividade!Execucaoprev = FormataTempo(VlttTotal)
'Tempo execução utilizado
If IsNull(TBProdutividade!Data3) = False And TBProdutividade!Data3 <> "" Then ProcFormataHora (TBProdutividade!Data3) Else ProcFormataHora ("00:00:00")
DecimoSegundos = s + DecimoSegundos
If TBproducao!Totalprod <> 0 Then VlttTotal = (TEUSEG * TBproducao!Totalprod) + DecimoSegundos Else VlttTotal = TEUSEG + DecimoSegundos
TBProdutividade!Data3 = FormataTempo(VlttTotal)
'Tempo total de execução prevista
If IsNull(TBProdutividade!Totalhsprev) = False And TBProdutividade!Totalhsprev <> "" Then ProcFormataHora (TBProdutividade!Totalhsprev) Else ProcFormataHora ("00:00:00")
DecimoSegundos = s + DecimoSegundos
If TBproducao!Totalprod <> 0 Then TotalPreparacao = TPPSEG + (TEPSEG * TBproducao!Totalprod) Else TotalPreparacao = TPPSEG + TEPSEG
TotalPreparacao = TotalPreparacao + DecimoSegundos
VlttTotal = TotalPreparacao
TBProdutividade!Totalhsprev = FormataTempo(VlttTotal)
'Tempo de preparação utilizado + (execução utilizado X qtde. produzida)
If IsNull(TBProdutividade!Totalhsutil) = False And TBProdutividade!Totalhsutil <> "" Then ProcFormataHora (TBProdutividade!Totalhsutil) Else ProcFormataHora ("00:00:00")
DecimoSegundos = s + DecimoSegundos
If TBproducao!Totalprod <> 0 Then TTE = TPUSEG + (TEUSEG * TBproducao!Totalprod) Else TTE = TPUSEG + TEUSEG
TTE = TTE + DecimoSegundos
VlttTotal = TTE
TBProdutividade!Totalhsutil = FormataTempo(VlttTotal)

'Qtde. N/C
TTOK = 0
TTNC = 0
QtdeSaida = 0

Gravar = False
If Opt_comparativo.Value = True Then
    Select Case cmbfiltrarpor
        Case "Operador": If maquina <> TBproducao!Usuario Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Posto de trabalho": If maquina <> TBproducao!maquina Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Turno": If maquina <> TBproducao!Turno Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Setor": If maquina <> TBproducao!Setor Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Código interno": If maquina <> TBproducao!Ordem Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Código de referência": If maquina <> TBproducao!Ordem Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Parecer CQ": If maquina <> Texto Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Ordem": If maquina <> TBproducao!Ordem Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
        Case "Cód. de referência": If maquina <> TBproducao!N_referencia Or OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
    End Select
Else
    If OS <> TBproducao!OS Or OS = TBproducao!OS And Dataini <> TBproducao!Data Then Gravar = True
End If

If Gravar = True Then
    Select Case cmbfiltrarpor
        Case "Operador":
            Set TBUsuarios = CreateObject("adodb.recordset")
            TBUsuarios.Open "Select * from Usuarios where Usuario = '" & TBproducao!Usuario & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBUsuarios.EOF = False Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Operador = '" & TBUsuarios!CODIGO & "-" & TBproducao!Usuario & "' and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Do While TBAbrir.EOF = False
                        If TBAbrir!ParecerCQ = "Rejeitar" Then
                            If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                        Else
                            If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                        End If
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
                TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
                TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
                TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
            End If
            TBUsuarios.Close
        Case "Posto de trabalho":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Maquina = '" & TBproducao!maquina & "' and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Turno":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Turno = " & TBproducao!Turno & " and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Setor":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Setor = '" & TBproducao!Setor & "' and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Ordem":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Ordem = " & TBproducao!Ordem & " and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
        Case "Cód. de referência":
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!OS & " and Ordem = " & TBproducao!Ordem & " and Data = '" & Format(TBproducao!Data, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!ParecerCQ = "Rejeitar" Then
                        If TBAbrir!IDProducao = 0 Then QtdeSaida = QtdeSaida + TBAbrir!TTNC Else TTNC = TTNC + TBAbrir!TTNC
                    Else
                        If TBAbrir!IDProducao <> 0 Then TTOK = TTOK + TBAbrir!TTNC
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + TTNC
            TBProdutividade!Refugo = TBProdutividade!Refugo + QtdeSaida
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + TTOK
    End Select
End If

ProcGravarTotalizacoesResumido
TBProdutividade!Qtdetotalprod = TBProdutividade!qtdeOK + TBProdutividade!qtdeNC

'Calcula eficiencia
'Tempo preparação previsto
ProcFormataHora (TBProdutividade!Data1)
Qtd = s + DecimoSegundos
'Tempo preparação utilizado
ProcFormataHora (TBProdutividade!Data2)
Qtd_Prog = s + DecimoSegundos
'Tempo execução previsto por peça
ProcFormataHora (TBProdutividade!Execucaoprev)
qtdeliberada = s + DecimoSegundos
'Tempo execução utilizado por peça
ProcFormataHora (TBProdutividade!Data3)
qtdeliberar = s + DecimoSegundos

ProcCalculaEficiencia Qtd, Qtd_Prog, qtdeliberada, qtdeliberar
TBProdutividade!Numero1 = Eficiencia_prep
TBProdutividade!Numero2 = Eficiencia_exec
TBProdutividade!Eficiencia = Eficiencia

OS = TBproducao!OS
Dataini = TBproducao!Data

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
If Opt_individual.Value = True Then
    If cmbTexto.Visible = True Then TBAbrir!Texto = cmbfiltrarpor & ") : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor & ") : " & txtTexto
Else
    TBAbrir!Texto = cmbfiltrarpor
    Numero = 0
    If Chk_fase.Value = 1 And Chk_turno.Value = 1 Then
        Numero = 3
    ElseIf Chk_fase.Value = 0 And Chk_turno.Value = 1 Then
            Numero = 2
        ElseIf Chk_fase.Value = 1 And Chk_turno.Value = 0 Then
                Numero = 1
    End If
    TBAbrir!qtdeNC = Numero
End If
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario

ProcAtualizaProducao
'Tempo total de preparação previsto
qt = TPPSEG
TBAbrir!Totalprevista = FormataTempo(qt)
'Tempo total de preparação utilizado
qt = TPUSEG
TBAbrir!Totalutilizada = FormataTempo(qt)
'Tempo total de execução previsto
qt = TEPSEG
TBAbrir!Data1 = FormataTempo(qt)
'Tempo total de execução utilizado
qt = TEUSEG
TBAbrir!Data2 = FormataTempo(qt)
'Tempo total de preparação + execução previsto
qt = TPPSEG + TEPSEG
TBAbrir!Data3 = FormataTempo(qt)
'Tempo total de preparação + execução utilizado
qt = TPUSEG + TEUSEG
TBAbrir!Data4 = FormataTempo(qt)
TBAbrir!Valor1 = Eficiencia_prep
TBAbrir!Valor2 = Eficiencia_exec
TBAbrir!TotalEficiencia = Eficiencia

If optDetalhado.Value = True Then
    OF = 0
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Ordem", Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        Do While TBproducao.EOF = False
            
            TBAbrir!QtdePrevista = TBAbrir!QtdePrevista + TBproducao!QtdePrev
            
            If OF <> TBproducao!Ordem Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from Producao_Relatorios where Ordem = " & TBproducao!Ordem & " and Qtdetotalprod <> 0 and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by OS", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBFI.MoveLast
                    quantidade = 0
                    Set TBLISTA = CreateObject("adodb.recordset")
                    TBLISTA.Open "Select Sum(Qtdetotalprod) as quantidade from Producao_Relatorios where OS = " & TBFI!OS & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBLISTA.EOF = False Then
                        quantidade = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade)
                    End If
                    TBLISTA.Close
                    TBAbrir!QtdeProduzida = TBAbrir!QtdeProduzida + quantidade
                End If
                TBFI.Close
            End If
            OF = TBproducao!Ordem
            TBproducao.MoveNext
        Loop
    End If
    TBproducao.Close
End If
TBAbrir.Update
TBAbrir.Close

If Opt_individual.Value = True And optResumido.Value = True Then Conexao.Execute "DELETE from Producao_Relatorios where Maquina = 'Nada consta'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoesResumido()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
Gravar = False
If optResumido.Value = True Then
    Select Case cmbfiltrarpor
        Case "Operador": If maquina <> TBproducao!Usuario Or OF <> TBproducao!Ordem Then Gravar = True
        Case "Posto de trabalho": If maquina <> TBproducao!maquina Or OF <> TBproducao!Ordem Then Gravar = True
        Case "Turno": If maquina <> TBproducao!Turno Or OF <> TBproducao!Ordem Then Gravar = True
        Case "Setor": If maquina <> TBproducao!Setor Or OF <> TBproducao!Ordem Then Gravar = True
        Case "Ordem": If maquina <> TBproducao!Ordem Then Gravar = True
        Case "Cód. de referência": If maquina <> TBproducao!N_referencia Or OF <> TBproducao!Ordem Then Gravar = True
    End Select
End If

If Gravar = True Then
    
    'Qtde total OK por filtro
    Set TBFI = CreateObject("adodb.recordset")
    Select Case cmbfiltrarpor
        Case "Operador": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and Usuario = '" & TBproducao!Usuario & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
        Case "Posto de trabalho": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and Maquina = '" & TBproducao!maquina & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
        Case "Turno": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and Turno = " & TBproducao!Turno & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
        Case "Setor": TBFI.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where " & NomeTabelaApTotalizacao & ".Ordem = " & TBproducao!Ordem & " and CadMaquinas.Setor = '" & TBproducao!Setor & "' and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS", Conexao, adOpenKeyset, adLockOptimistic
        Case "Ordem": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
        Case "Cód. de referência": TBFI.Open "Select Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".* FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where " & NomeTabelaApTotalizacao & ".Ordem = " & TBproducao!Ordem & " and Producao.N_Referencia = '" & TBproducao!N_referencia & "' and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS", Conexao, adOpenKeyset, adLockOptimistic
    End Select
    If TBFI.EOF = False Then
        TBFI.MoveLast
        quantidade = 0
        Set TBLISTA = CreateObject("adodb.recordset")
        Select Case cmbfiltrarpor
            Case "Operador": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Usuario = '" & TBproducao!Usuario & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            Case "Posto de trabalho": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Maquina = '" & TBproducao!maquina & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            Case "Turno": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Turno = " & TBproducao!Turno & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            Case "Setor": TBLISTA.Open "Select Sum(" & NomeTabelaApTotalizacao & ".QTOK) as quantidade FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where CadMaquinas.Setor = '" & TBproducao!Setor & "' and " & NomeTabelaApTotalizacao & ".OS = " & TBFI!OS & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            Case "Ordem": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            Case "Cód. de referência": TBLISTA.Open "Select Sum(" & NomeTabelaApTotalizacao & ".QTOK) as quantidade FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where Producao.N_Referencia = '" & TBproducao!N_referencia & "' and " & NomeTabelaApTotalizacao & ".OS = " & TBFI!OS & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        End Select
        If TBLISTA.EOF = False Then
            quantidade = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade)
        End If
        TBLISTA.Close
        
        'Qtde OK
        TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + quantidade
    End If
    
    'Qtde total OK da ordem
    If OF <> TBproducao!Ordem Then
        quantidade = 0
        Set TBFI = CreateObject("adodb.recordset")
        If Opt_individual.Value = True Then
            Select Case cmbfiltrarpor
                Case "Operador": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and Usuario = '" & TBproducao!Usuario & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Posto de trabalho": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and Maquina = '" & TBproducao!maquina & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Turno": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and Turno = " & TBproducao!Turno & " and  totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Setor": TBFI.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where CadMaquinas.Setor = '" & TBproducao!Setor & "' and " & NomeTabelaApTotalizacao & ".Ordem = " & TBproducao!Ordem & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Ordem": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & FamiliaAntiga & " and data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Cód. de referência": TBFI.Open "Select Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".* FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where Producao.N_Referencia = '" & TBproducao!N_referencia & "' and " & NomeTabelaApTotalizacao & ".Ordem = " & TBproducao!Ordem & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS", Conexao, adOpenKeyset, adLockOptimistic
            End Select
        Else
            Select Case cmbfiltrarpor
                Case "Operador": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Posto de trabalho": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Turno": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Setor": TBFI.Open "Select CadMaquinas.Setor, " & NomeTabelaApTotalizacao & ".* FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where " & NomeTabelaApTotalizacao & ".Ordem = " & TBproducao!Ordem & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Ordem": TBFI.Open "Select * from " & NomeTabelaApTotalizacao & " where Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by Ordem, OS", Conexao, adOpenKeyset, adLockOptimistic
                Case "Cód. de referência": TBFI.Open "Select Producao.N_Referencia, " & NomeTabelaApTotalizacao & ".* FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where " & NomeTabelaApTotalizacao & ".Ordem = " & TBproducao!Ordem & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "' order by " & NomeTabelaApTotalizacao & ".Ordem, " & NomeTabelaApTotalizacao & ".OS", Conexao, adOpenKeyset, adLockOptimistic
            End Select
        End If
        If TBFI.EOF = False Then
            TBFI.MoveLast
            Set TBLISTA = CreateObject("adodb.recordset")
            If Opt_individual.Value = True Then
                Select Case cmbfiltrarpor
                    Case "Operador": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Usuario = '" & TBproducao!Usuario & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Posto de trabalho": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Maquina = '" & TBproducao!maquina & "' and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Turno": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Turno = " & TBproducao!Turno & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Setor": TBLISTA.Open "Select Sum(" & NomeTabelaApTotalizacao & ".QTOK) as quantidade FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where CadMaquinas.Setor = '" & TBproducao!Setor & "' and " & NomeTabelaApTotalizacao & ".OS = " & TBFI!OS & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Ordem": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and " & FamiliaAntiga & " and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Cód. de referência": TBLISTA.Open "Select Sum(" & NomeTabelaApTotalizacao & ".QTOK) as quantidade FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where Producao.N_Referencia = '" & TBproducao!N_referencia & "' and " & NomeTabelaApTotalizacao & ".OS = " & TBFI!OS & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                End Select
            Else
                Select Case cmbfiltrarpor
                    Case "Operador": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and totalprod <> 0 and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & FamiliaAntiga & " and data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Posto de trabalho": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and totalprod <> 0 and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & FamiliaAntiga & " and data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Turno": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and totalprod <> 0 and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & FamiliaAntiga & " and data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Setor": TBLISTA.Open "Select Sum(" & NomeTabelaApTotalizacao & ".QTOK) as quantidade FROM CadMaquinas INNER JOIN " & NomeTabelaApTotalizacao & " ON CadMaquinas.Maquina = " & NomeTabelaApTotalizacao & ".maquina where " & NomeTabelaApTotalizacao & ".OS = " & TBFI!OS & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Ordem": TBLISTA.Open "Select Sum(QTOK) as quantidade from " & NomeTabelaApTotalizacao & " where OS = " & TBFI!OS & " and Ordem = " & TBproducao!Ordem & " and totalprod <> 0 and data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & FamiliaAntiga & " and data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Case "Cód. de referência": TBLISTA.Open "Select Sum(" & NomeTabelaApTotalizacao & ".QTOK) as quantidade FROM Producao INNER JOIN " & NomeTabelaApTotalizacao & " ON Producao.Ordem = " & NomeTabelaApTotalizacao & ".Ordem where " & NomeTabelaApTotalizacao & ".OS = " & TBFI!OS & " and " & NomeTabelaApTotalizacao & ".totalprod <> 0 and " & FamiliaAntiga & " and " & NomeTabelaApTotalizacao & ".data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND " & NomeTabelaApTotalizacao & ".data <= '" & Format(msk_fltFim.Value, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                End Select
            End If
            If TBLISTA.EOF = False Then
                quantidade = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade)
            End If
            TBLISTA.Close
            
            'Qtde OK
            TBAbrir!QtdeProduzida = TBAbrir!QtdeProduzida + quantidade
        End If
    End If
    
    TBFI.Close
End If

'Qtde. prevista
'Verif. quantidade de peças previstas x tempo utilizado
'quantidade = TBAbrir!QtdeProduzida

Quant = TBproducao!QTOK + TTNC
If TPPSEG > 0 And TPUSEG > 0 And TEPSEG > 0 And TEUSEG > 0 Then
    Qtde = (TPUSEG / TPPSEG) + (TEUSEG / TEPSEG)
ElseIf TPPSEG > 0 And TPUSEG > 0 Then
        Qtde = TPUSEG / TPPSEG
    ElseIf TEPSEG > 0 And TEUSEG > 0 Then
            Qtde = TEUSEG / TEPSEG
        Else
            Qtde = 0
End If
If Quant <> 0 Then TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + (Qtde * Quant) Else TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + Qtde

If Quant <> 0 Then TBAbrir!QtdePrevista = TBAbrir!QtdePrevista + (Qtde * Quant) Else TBAbrir!QtdePrevista = TBAbrir!QtdePrevista + Qtde

'Qtde. N/C
TBAbrir!qtdeNC = TBAbrir!qtdeNC + TTNC
TBAbrir!Lucro = TBAbrir!Lucro + QtdeSaida
TBAbrir!QtdeProduzida = TBAbrir!QtdeProduzida + TTNC

TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario

TBAbrir.Update
TBAbrir.Close

OF = TBproducao!Ordem
If optResumido.Value = True Then
    Select Case cmbfiltrarpor
        Case "Operador": maquina = TBproducao!Usuario
        Case "Posto de trabalho": maquina = TBproducao!maquina
        Case "Turno": maquina = TBproducao!Turno
        Case "Setor": maquina = TBproducao!Setor
        Case "Ordem": maquina = TBproducao!Ordem
        Case "Cód. de referência": maquina = TBproducao!N_referencia
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaProducao()
On Error GoTo tratar_erro

TPPSEG = 0 'Tempo preparação previsto
TPUSEG = 0 'Tempo preparação real
TEPSEG = 0 'Tempo execução previsto
TEUSEG = 0 'Tempo execução real
ProcAbrirTabelas1
If TBproducao.EOF = False Then
    Do While TBproducao.EOF = False
        'Tempo de preparação previsto
        ProcFormataHora (IIf(IsNull(TBproducao!Preparacao), 0, TBproducao!Preparacao))
        TPPSEG = TPPSEG + s + DecimoSegundos
        
        'Tempo de execução previsto
        ProcFormataHora (IIf(IsNull(TBproducao!Execucao), 0, TBproducao!Execucao))
        TEPSEG = TEPSEG + IIf(TBproducao!Totalprod > 0, ((s + DecimoSegundos) * TBproducao!Totalprod), s + DecimoSegundos)
                
        'Tempo de preparação real
        ProcFormataHora (TBproducao!TPUTIL)
        TPUSEG = TPUSEG + s + DecimoSegundos
        
        'Tempo de execução real
        ProcFormataHora (TBproducao!TEUTIL)
        TEUSEG = TEUSEG + IIf(TBproducao!Totalprod > 0, ((s + DecimoSegundos) * TBproducao!Totalprod), s + DecimoSegundos)
        
        TBproducao.MoveNext
    Loop
End If
TBproducao.Close
ProcCalculaEficiencia TPPSEG, TPUSEG, TEPSEG, TEUSEG

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaEficiencia(TPPSEG As Double, TPUSEG As Double, TEPSEG As Double, TEUSEG As Double)
On Error GoTo tratar_erro

Eficiencia_prep = 0
Eficiencia_exec = 0
Eficiencia = 0
If TPPSEG > 0 And TPUSEG > 0 Then Eficiencia_prep = Format((TPPSEG / TPUSEG) * 100, "###,##0.00")
If TEPSEG > 0 And TEUSEG > 0 Then Eficiencia_exec = Format((TEPSEG / TEUSEG) * 100, "###,##0.00")
If Eficiencia_prep > 0 And Eficiencia_exec > 0 Then
    Eficiencia = (Eficiencia_prep + Eficiencia_exec) / 2
ElseIf Eficiencia_prep > 0 Then
        Eficiencia = Eficiencia_prep
    ElseIf Eficiencia_exec > 0 Then
            Eficiencia = Eficiencia_exec
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
    txtTexto = ""
    txtTexto.Enabled = False
    If cmbfiltrarpor = "Posto de trabalho" Then Chk_setor.Enabled = True
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
    Chk_setor.Value = 0
    Chk_setor.Enabled = False
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
    Lista.ListItems.Clear
    Lista.Visible = True
    Lista1.ListItems.Clear
    Lista1.Visible = False
    ProcLimpaCamposTotais
    If cmbfiltrarpor = "Ordem" Then ProcListaOrdem Else ProcListaPadrao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    Lista.ListItems.Clear
    Lista.Visible = False
    Lista1.ListItems.Clear
    Lista1.Visible = True
    ProcLimpaCamposTotais
    If cmbfiltrarpor = "Ordem" Then ProcListaOrdem Else ProcListaPadrao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_fase_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Txt_fase <> "" Then
    VerifNumero = Txt_fase
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_fase = ""
        Txt_fase.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_turno_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Txt_turno <> "" Then
    VerifNumero = Txt_turno
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_turno = ""
        Txt_turno.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
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

Private Sub ProcVerifFiltroFaseTurno()
On Error GoTo tratar_erro

If Chk_fase.Value = 1 And Chk_turno.Value = 1 Then
    FamiliaAntiga = NomeTabelaApTotalizacao & ".Fase = " & Txt_fase & " and " & NomeTabelaApTotalizacao & ".Turno = " & Txt_turno
ElseIf Chk_fase.Value = 0 And Chk_turno.Value = 1 Then
        FamiliaAntiga = NomeTabelaApTotalizacao & ".Turno = " & Txt_turno
    ElseIf Chk_fase.Value = 1 And Chk_turno.Value = 0 Then
            FamiliaAntiga = NomeTabelaApTotalizacao & ".Fase = " & Txt_fase
        Else
            FamiliaAntiga = NomeTabelaApTotalizacao & ".Maquina <> 'Null'"
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
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
