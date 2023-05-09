VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProd_Relatorios_Monitor_Eventos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Relatórios - Monitor de eventos"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
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
   Icon            =   "frmProd_Relatorios_Monitor_Eventos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
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
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6945
      Left            =   3240
      TabIndex        =   13
      Top             =   2235
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   12250
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
      NumItems        =   10
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
         Object.Tag             =   "T"
         Text            =   "Operador"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Posto de trab."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Turno"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição do evento"
         Object.Width           =   5480
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "D"
         Text            =   "Início"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "D"
         Text            =   "Final"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Tempo total"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar eventos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6945
      Left            =   60
      TabIndex        =   30
      Top             =   2220
      Width           =   3165
      Begin MSComctlLib.ListView Lista_eventos 
         Height          =   6525
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   11509
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3757
         EndProperty
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
      Left            =   75
      TabIndex        =   17
      Top             =   9180
      Width           =   15195
      Begin VB.TextBox Txt_tempo_total_utilizado 
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
         Left            =   12975
         MaxLength       =   20
         TabIndex        =   16
         ToolTipText     =   "Tempo total utilizado."
         Top             =   390
         Width           =   2010
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   330
         Width           =   9165
         _ExtentX        =   16166
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
         Left            =   9450
         TabIndex        =   28
         Top             =   360
         Width           =   3315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo total utilizado"
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
         Left            =   13080
         TabIndex        =   18
         Top             =   180
         Width           =   1800
      End
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
      Height          =   1260
      Left            =   60
      TabIndex        =   19
      Top             =   960
      Width           =   1695
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
         Top             =   420
         Value           =   -1  'True
         Width           =   1185
      End
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
         Top             =   690
         Width           =   1425
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
      Height          =   1260
      Left            =   1770
      TabIndex        =   20
      Top             =   960
      Width           =   1455
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
         Top             =   420
         Value           =   -1  'True
         Width           =   1185
      End
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
         Top             =   690
         Width           =   1155
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   1260
      Left            =   3240
      TabIndex        =   24
      Top             =   960
      Width           =   9885
      Begin VB.CheckBox Chk_turno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Turno :"
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
         Left            =   210
         TabIndex        =   4
         Top             =   510
         Width           =   825
      End
      Begin VB.ComboBox Cmb_turno 
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
         Height          =   330
         ItemData        =   "frmProd_Relatorios_Monitor_Eventos.frx":0442
         Left            =   1050
         List            =   "frmProd_Relatorios_Monitor_Eventos.frx":0455
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Turno."
         Top             =   510
         Width           =   765
      End
      Begin VB.CheckBox chk_Aberto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Em aberto"
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
         Left            =   8700
         TabIndex        =   8
         Top             =   930
         Width           =   1035
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
         ItemData        =   "frmProd_Relatorios_Monitor_Eventos.frx":0468
         Left            =   4200
         List            =   "frmProd_Relatorios_Monitor_Eventos.frx":046A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Texto para pesquisa."
         Top             =   510
         Width           =   5535
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
         ItemData        =   "frmProd_Relatorios_Monitor_Eventos.frx":046C
         Left            =   1830
         List            =   "frmProd_Relatorios_Monitor_Eventos.frx":047F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Opções para filtro."
         Top             =   510
         Width           =   2355
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
         Left            =   2587
         TabIndex        =   26
         Top             =   300
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
         Left            =   6232
         TabIndex        =   25
         Top             =   300
         Width           =   1470
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1260
      Left            =   13140
      TabIndex        =   21
      Top             =   960
      Width           =   2115
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
         Left            =   300
         TabIndex        =   9
         Top             =   180
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   11
         ToolTipText     =   "Data final."
         Top             =   810
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
         Format          =   488570881
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   10
         ToolTipText     =   "Data inicio."
         Top             =   450
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
         Format          =   488570881
         CurrentDate     =   39057
      End
      Begin VB.Label Label4 
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
         TabIndex        =   23
         Top             =   510
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
         TabIndex        =   22
         Top             =   870
         Width           =   360
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   29
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
         Name            =   "MS Sans Serif"
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
      ButtonLeft4     =   97
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
      ButtonLeft5     =   135
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
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   10470
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProd_Relatorios_Monitor_Eventos.frx":04B5
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista_ordem 
      Height          =   6945
      Left            =   12465
      TabIndex        =   14
      Top             =   2235
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   12250
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6945
      Left            =   3240
      TabIndex        =   15
      Top             =   2235
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   12250
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Object.Width           =   18036
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição do evento"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Turno"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "D"
         Text            =   "Tempo total"
         Object.Width           =   2469
      EndProperty
   End
End
Attribute VB_Name = "frmProd_Relatorios_Monitor_Eventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chk_Aberto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
Lista_ordem.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_filtrar_backup_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
Lista1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_turno_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then
        If Lista.ListItems.Count = 0 Then Exit Sub
    Else
        If Lista1.ListItems.Count = 0 Then Exit Sub
    End If
Else
    If Lista1.ListItems.Count = 0 Then Exit Sub
End If
frmProd_Relatorios_Monitor_Eventos_MenuImpressao.Show 1

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
    Case vbKeyF2: ProcAbrir
    Case vbKeyF5: ProcImprimir
    'Case vbKeyF1: Cmd_ajuda_Click
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
If TBLISTA.EOF = False Then
    Posicao = TBLISTA.RecordCount
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With Lista.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Execucaoprev), "", TBLISTA!Execucaoprev)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Data1), "", TBLISTA!Data1)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Turno), "", TBLISTA!Turno)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento)
                .Item(.Count).SubItems(7) = Format(TBLISTA!Data5, "dd/mm/yy hh:mm:ss")
                .Item(.Count).SubItems(8) = Format(TBLISTA!Data6, "dd/mm/yy hh:mm:ss")
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil)
            End With
        Else
            With Lista1.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Turno), "", TBLISTA!Turno)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil)
            End With
        End If
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If

ProcLimpaCamposTotais
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Txt_tempo_total_utilizado = IIf(IsNull(TBLISTA!Totalutilizada), "", TBLISTA!Totalutilizada)
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
Txt_tempo_total_utilizado = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
Formulario = "PCP/Relatórios/Monitor de eventos"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
cmbfiltrarpor.Text = "Operador"
ProcCarregaListaEventos

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEventos()
On Error GoTo tratar_erro

Lista_eventos.ListItems.Clear
Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select * from CodigoDesc order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBCodigoDesc.EOF = False Then
    Do While TBCodigoDesc.EOF = False
        With Lista_eventos.ListItems
            .Add , , TBCodigoDesc!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBCodigoDesc!Descricao), "", TBCodigoDesc!Descricao)
        End With
        TBCodigoDesc.MoveNext
    Loop
End If
TBCodigoDesc.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "PCP/Relatórios/Monitor de eventos"
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

Private Sub ProcArrumaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
With Lista1
    .ListItems.Clear
    If FunVerifEventoSelecionado = True Then
        If Chk_turno.Value = 1 Then .ColumnHeaders(2).Width = 4125 Else .ColumnHeaders(2).Width = 4725
        .ColumnHeaders(3).Width = 5500
    Else
        If Chk_turno.Value = 1 Then .ColumnHeaders(2).Width = 9625 Else .ColumnHeaders(2).Width = 10225
        .ColumnHeaders(3).Width = 0
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
Lista_ordem.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
With Cmb_turno
    If Chk_turno.Value = 1 Then
        .Enabled = True
        .SetFocus
        If FunVerifEventoSelecionado = True Then Lista1.ColumnHeaders(2).Width = 4125 Else Lista1.ColumnHeaders(2).Width = 9625
        Lista1.ColumnHeaders(4).Width = 600
    Else
        .Enabled = False
        .ListIndex = -1
        If FunVerifEventoSelecionado = True Then Lista1.ColumnHeaders(2).Width = 4725 Else Lista1.ColumnHeaders(2).Width = 10225
        Lista1.ColumnHeaders(4).Width = 0
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_eventos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_eventos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
    ProcArrumaLista
Else
    ProcOrdenaListView Lista_eventos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_eventos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcArrumaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Function FunVerifEventoSelecionado() As Boolean
On Error GoTo tratar_erro

FunVerifEventoSelecionado = False
With Lista_eventos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then FunVerifEventoSelecionado = True
    Next InitFor
End With

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Lista_ordem.ListItems.Clear
If Chk_filtrar_backup.Value = 1 Then NomeTabelaAp = "Producaofases_Backup" Else NomeTabelaAp = "Producaofases"
INNERJOINTEXTO = "ordemservico INNER JOIN " & NomeTabelaAp & " ON ordemservico.Idproducao = " & NomeTabelaAp & ".OS"
Select Case cmbfiltrarpor
    Case "Operador": TextoFiltro = NomeTabelaAp & ".usuario = '" & Lista.SelectedItem.ListSubItems(2) & "'"
    Case "Posto de trabalho": TextoFiltro = NomeTabelaAp & ".maquina = '" & Lista.SelectedItem.ListSubItems(3) & "'"
    Case "Turno": TextoFiltro = NomeTabelaAp & ".turno = " & Lista.SelectedItem.ListSubItems(4)
    Case Else:
        TextoFiltro = "Cadmaquinas.setor = '" & Lista.SelectedItem.ListSubItems(4) & "'"
        INNERJOINTEXTO = "(ordemservico INNER JOIN " & NomeTabelaAp & " ON ordemservico.Idproducao = " & NomeTabelaAp & ".OS) INNER JOIN Cadmaquinas ON Cadmaquinas.Maquina = " & NomeTabelaAp & ".Maquina"
End Select
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ordemservico.* from " & INNERJOINTEXTO & " where " & NomeTabelaAp & ".Data = '" & Format(Lista.SelectedItem.ListSubItems(1), "Short Date") & "' and " & TextoFiltro & " and " & NomeTabelaAp & ".Descricao = '" & Lista.SelectedItem.ListSubItems(6) & "' and " & NomeTabelaAp & ".Tempoinicio = '" & Lista.SelectedItem.ListSubItems(7) & "' order by Ordemservico.Ordem, Ordemservico.IDproducao", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_ordem.ListItems
            .Add , , TBLISTA!Ordem
            .Item(.Count).SubItems(1) = TBLISTA!IDProducao
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
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
Lista_ordem.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

If Opt_individual.Value = True Then
    With cmbTexto
        .Clear
        If Chk_turno.Value = 1 Then
            If FunVerifEventoSelecionado = True Then Lista1.ColumnHeaders(3).Width = 4725 Else Lista1.ColumnHeaders(2).Width = 9625
            Lista1.ColumnHeaders(4).Width = 600
        End If
        Select Case cmbfiltrarpor
            Case "Operador": ProcCarregaComboUsuario cmbTexto, "U.Bloqueado = 'False'", True
            Case "Posto de trabalho": ProcCarregaComboPostoTrab cmbTexto, "Bloqueado = 'False'", True, False
            Case "Turno":
                .AddItem ""
                .AddItem 0
                .AddItem 1
                .AddItem 2
                .AddItem 3
                .AddItem 4
                If FunVerifEventoSelecionado = True Then Lista1.ColumnHeaders(3).Width = 5500 Else Lista1.ColumnHeaders(2).Width = 10225
                Lista1.ColumnHeaders(4).Width = 0
            Case "Setor": ProcCarregaComboSetor cmbTexto, "US.Setor is not null", "", False, True, False, "", False, True
            Case "Grupo": ProcCarregaComboGrupoPT cmbTexto, True
        End Select
    End With
End If

If cmbfiltrarpor = "Turno" Then
    With Chk_turno
        .Value = 0
        .Enabled = False
    End With
    With Cmb_turno
        .ListIndex = -1
        .Enabled = False
    End With
    With Lista1
        If FunVerifEventoSelecionado = True Then
            .ColumnHeaders(2).Width = 4725
            .ColumnHeaders(3).Width = 5500
        Else
            .ColumnHeaders(2).Width = 10225
            .ColumnHeaders(3).Width = 0
        End If
        .ColumnHeaders(4).Width = 0
    End With
Else
    Chk_turno.Enabled = True
End If
Lista1.ColumnHeaders(2).Text = cmbfiltrarpor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

Acao = "filtrar"
If optDetalhado.Value = False And optResumido.Value = False Then
    NomeCampo = "se o filtro é detalhado ou resumido"
    ProcVerificaAcao
    Exit Sub
End If
If Chk_turno.Value = 1 And Cmb_turno = "" Then
    NomeCampo = "o turno"
    ProcVerificaAcao
    Cmb_turno.SetFocus
    Exit Sub
End If
If Opt_individual.Value = True And optResumido.Value = True And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
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
ProcFiltrar
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then Ordenar = "Id" Else Ordenar = "Maquina, Descevento"
Else
    Ordenar = "Maquina, Descevento"
End If
TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Permitido = False
Evento = ""
maquina = ""

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

ProcAbrirTabelas
If TBproducao.EOF = False Then
    Permitido = True
    PBLista.Min = 0
    PBLista.Max = TBproducao.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBproducao.EOF = False
        Set TBProdutividade = CreateObject("adodb.recordset")
        If Opt_individual.Value = True And optDetalhado.Value = True Then
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            If FunVerifEventoSelecionado = True Then TextoFiltro = "and Descevento = '" & TBproducao!Descricao & "'" Else TextoFiltro = ""
            Select Case cmbfiltrarpor
                Case "Operador": TextoFiltro1 = TBproducao!Usuario
                Case "Posto de trabalho": TextoFiltro1 = TBproducao!maquina
                Case "Turno": TextoFiltro1 = TBproducao!Turno
                Case "Setor": TextoFiltro1 = TBproducao!Setor
            End Select
            TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TextoFiltro1 & "' " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosResumido
        End If
        TBproducao.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

If Chk_filtrar_backup.Value = 1 Then
    NomeTabelaAp = "Monitor_eventos_detalhado_Backup"
    NomeTabelaAp1 = "Monitor_eventos_resumido_Backup"
Else
    NomeTabelaAp = "Monitor_eventos_detalhado"
    NomeTabelaAp1 = "Monitor_eventos_resumido"
End If

If optDetalhado.Value = True Then
    Ordenar = "Data, Tempoinicio"
    NomeTabelaFiltro = NomeTabelaAp
Else
    Ordenar = "Data"
    NomeTabelaFiltro = NomeTabelaAp1
End If

Evento = ""
With Lista_eventos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Evento = "" Then Evento = NomeTabelaFiltro & ".descricao = '" & .ListItems.Item(InitFor).ListSubItems(1) & "'" Else Evento = Evento & " or " & NomeTabelaFiltro & ".descricao = '" & .ListItems.Item(InitFor).ListSubItems(1) & "'"
        End If
    Next InitFor
End With
If Evento <> "" Then Evento = "and (" & Evento & ")"

If Chk_turno.Value = 1 Then
    Turno = NomeTabelaFiltro & ".Turno = " & Cmb_turno
Else
    Turno = NomeTabelaFiltro & ".descricao <> 'Null'"
End If

If chk_Aberto.Value = 1 Then EmAberto = NomeTabelaFiltro & ".TempoFinal = '30/12/1899 00:00:00' And " & NomeTabelaFiltro & ".Codigodesc <> 3" Else EmAberto = NomeTabelaFiltro & ".descricao <> 'Null'"

Set TBproducao = CreateObject("adodb.recordset")
If Opt_individual.Value = True Then
    INNERJOINTEXTO = "Select * from " & NomeTabelaFiltro
    TextoFiltroPadrao = "data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "' and " & Turno & " and " & EmAberto & " " & Evento & " order by " & Ordenar
    If cmbTexto <> "" Then
        Select Case cmbfiltrarpor
            Case "Operador": TextoFiltro = "usuario"
            Case "Posto de trabalho": TextoFiltro = "maquina"
            Case "Turno": TextoFiltro = "turno"
            Case "Setor": TextoFiltro = "setor"
            Case "Grupo": TextoFiltro = "grupo"
        End Select
        FiltroPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
    Else
        FiltroPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
    End If
    TBproducao.Open FiltroPadrao, Conexao, adOpenKeyset, adLockReadOnly
Else
    INNERJOINTEXTO = "Select * from " & NomeTabelaAp1
    TextoFiltroData = "data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' AND data <= '" & Format(msk_fltFim.Value, "Short Date") & "'"
    TextoFiltroPadrao = " and " & Turno & " " & Evento
    Select Case cmbfiltrarpor
        Case "Operador": OrdenarTexto = "usuario"
        Case "Posto de trabalho": OrdenarTexto = "maquina"
        Case "Turno":
            TextoFiltroPadrao = " and " & Evento
            OrdenarTexto = "turno"
        Case "Setor": OrdenarTexto = "setor"
        Case "Grupo": OrdenarTexto = "grupo"
    End Select
    TBproducao.Open INNERJOINTEXTO & " where " & TextoFiltroData & TextoFiltroPadrao & " order by " & OrdenarTexto, Conexao, adOpenKeyset, adLockReadOnly
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

TBProdutividade.AddNew

If Chk_filtrar_backup.Value = 0 Then NomeTabela = "ProducaoFases" Else NomeTabela = "ProducaoFases_Backup"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select IDProducao from " & NomeTabela & " where Data = '" & TBproducao!Data & "' and Usuario = '" & TBproducao!Usuario & "' and Maquina = '" & TBproducao!maquina & "' and TempoInicio = '" & TBproducao!TempoInicio & "' and Turno = " & TBproducao!Turno, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBProdutividade!Ordem = TBAbrir!IDProducao
End If

TBProdutividade!Data = TBproducao!Data
TBProdutividade!maquina = TBproducao!Usuario
TBProdutividade!Execucaoprev = TBproducao!maquina
TBProdutividade!Data1 = TBproducao!Setor
TBProdutividade!Turno = TBproducao!Turno
TBProdutividade!Data5 = TBproducao!TempoInicio

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CodigoDesc where Descricao = '" & TBproducao!Descricao & "' and Totalizar_relatorio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBProdutividade!Data6 = TBproducao!TempoFinal
    If TBproducao!TempoFinal <> "00:00:00" Then
        TempoTotal = TBproducao!TempoFinal - TBproducao!TempoInicio
        ElapsedTime (TempoTotal)
        TBProdutividade!Totalhsutil = HoraTotal
        TBProdutividade!Eficiencia = Format(s / 3600, "###,##0.00")
    End If
End If
TBAbrir.Close

TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!DescEvento = TBproducao!Descricao
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

If TBProdutividade.EOF = True Then TBProdutividade.AddNew
    Select Case cmbfiltrarpor
        Case "Operador": Texto = TBproducao!Usuario
        Case "Posto de trabalho": Texto = TBproducao!maquina
        Case "Turno": Texto = TBproducao!Turno
        Case "Setor": Texto = TBproducao!Setor
        Case "Grupo": Texto = TBproducao!Grupo
    End Select

TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
If FunVerifEventoSelecionado = True Or Opt_individual.Value = True And optResumido.Value = True Then TBProdutividade!DescEvento = TBproducao!Descricao
If Chk_turno.Value = 1 Or Opt_individual.Value = True And optResumido.Value = True Then TBProdutividade!Turno = TBproducao!Turno

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CodigoDesc where Descricao = '" & TBproducao!Descricao & "' and Totalizar_relatorio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    'Tempo total do evento
    TempoTotalEvento = FormataTempo(IIf(IsNull(TBproducao!TempoTotalSeg), 0, TBproducao!TempoTotalSeg))
    If Len(TempoTotalEvento) > 8 Or Len(TempoTotalEvento) = 8 And TempoTotalEvento > "23:59:59" Then
        ProcFormataHora (TempoTotalEvento)
        TempoTotalProd = DataResultado
    Else
        TempoTotalProd = TempoTotalEvento
    End If
    
    'Total de horas utilizadas
    Totalhsutil = IIf(IsNull(TBProdutividade!Totalhsutil), 0, TBProdutividade!Totalhsutil)
    If Len(Totalhsutil) > 8 Or Len(Totalhsutil) = 8 And Totalhsutil > "23:59:59" Then
        ProcFormataHora (Totalhsutil)
        TempoTotal = DataResultado
    Else
        TempoTotal = Totalhsutil
    End If
    
    TempoTotal = TempoTotal + TempoTotalProd
    ElapsedTime (TempoTotal)
    TBProdutividade!Totalhsutil = HoraTotal
    
    'Calcula tempo total em horas
    Totalhsutil = IIf(IsNull(TBProdutividade!Totalhsutil), 0, TBProdutividade!Totalhsutil)
    If Len(Totalhsutil) > 8 Or Len(Totalhsutil) = 8 And Totalhsutil > "23:59:59" Then
        ProcFormataHora (Totalhsutil)
        TempoTotal = DataResultado
    Else
        TempoTotal = Totalhsutil
    End If
    ElapsedTime (TempoTotal)
    TBProdutividade!Eficiencia = Format(s / 3600, "###,##0.00")
End If
TBAbrir.Close

TBProdutividade.Update
TBProdutividade.Close

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
If optDetalhado.Value = True Or Opt_comparativo.Value = True Then TBAbrir!Texto = cmbfiltrarpor Else TBAbrir!Texto = cmbfiltrarpor & "): " & cmbTexto
If FunVerifEventoSelecionado = True Then TBAbrir!Texto1 = "S" Else TBAbrir!Texto1 = "N"
If Chk_turno.Value = 1 Or cmbfiltrarpor = "Turno" Then TBAbrir!Turno = True Else TBAbrir!Turno = False
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Do While TBproducao.EOF = False
        'Tempo total do evento
        TempoTotalEvento = IIf(IsNull(TBproducao!Totalhsutil), 0, TBproducao!Totalhsutil)
        If Len(TempoTotalEvento) > 8 Or Len(TempoTotalEvento) = 8 And TempoTotalEvento > "23:59:59" Then
            ProcFormataHora (TempoTotalEvento)
            TempoTotalProd = DataResultado
        Else
            TempoTotalProd = TempoTotalEvento
        End If
        
        'Total de horas utilizadas
        Totalhsutil = IIf(IsNull(TBAbrir!Totalutilizada), 0, TBAbrir!Totalutilizada)
        If Len(Totalhsutil) > 8 Or Len(Totalhsutil) = 8 And Totalhsutil > "23:59:59" Then
            ProcFormataHora (Totalhsutil)
            TempoTotal = DataResultado
        Else
            TempoTotal = Totalhsutil
        End If

        ElapsedTime (TempoTotal + TempoTotalProd)
        TBAbrir!Totalutilizada = HoraTotal
        TBAbrir!TotalEficiencia = Format(s / 3600, "###,##0.00")
        TBproducao.MoveNext
    Loop
End If
TBproducao.Close

TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_ordem.ListItems.Clear
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
Lista_ordem.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
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
    With Lista
        .ListItems.Clear
        .Visible = True
    End With
    With Lista_ordem
        .ListItems.Clear
        .Visible = True
    End With
    With Lista1
        .ListItems.Clear
        .Visible = False
    End With
    chk_Aberto.Enabled = True
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    With Lista
        .ListItems.Clear
        .Visible = False
    End With
    With Lista_ordem
        .ListItems.Clear
        .Visible = False
    End With
    With Lista1
        .ListItems.Clear
        .Visible = True
    End With
    chk_Aberto.Value = 0
    chk_Aberto.Enabled = False
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcAbrir
    Case 2: ProcImprimir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
