VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFluxodecaixa 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Fluxo de caixa"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmFluxodecaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
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
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin MSComctlLib.ListView lst_fluxo 
      Height          =   5130
      Left            =   60
      TabIndex        =   5
      Top             =   3795
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   9049
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "IDFluxo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Lançamento"
         Object.Width           =   11421
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Saldo anterior"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Crédito"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Débito"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSComctlLib.ListView lst_fluxo1 
      Height          =   5160
      Left            =   60
      TabIndex        =   6
      Top             =   3780
      Visible         =   0   'False
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   9102
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "IDFluxo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Lançamento"
         Object.Width           =   11421
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Saldo anterior"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "À creditar"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "À debitar"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   2469
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
      ItemData        =   "frmFluxodecaixa.frx":1042
      Left            =   1170
      List            =   "frmFluxodecaixa.frx":1044
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1200
      Width           =   10170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   55
      TabIndex        =   24
      Top             =   990
      Width           =   15195
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   11820
         TabIndex        =   1
         ToolTipText     =   "Data início para pesquisa."
         Top             =   210
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
         Format          =   155189251
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   13710
         TabIndex        =   2
         ToolTipText     =   "Data final para pesquisa."
         Top             =   210
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
         Format          =   155189251
         CurrentDate     =   39057
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
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
         TabIndex        =   45
         Top             =   210
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Left            =   11430
         TabIndex        =   36
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   13260
         TabIndex        =   25
         Top             =   240
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView lst_fluxo2 
      Height          =   5160
      Left            =   60
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   9102
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
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
         Text            =   "IDFluxo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Lançamento"
         Object.Width           =   16360
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "À creditar"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "À debitar"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView lst_Instituicoes 
      Height          =   1740
      Left            =   60
      TabIndex        =   3
      Top             =   1680
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   3069
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Banco"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Agência"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Conta"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   18706
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   46
      Top             =   0
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
      ButtonKey2      =   "4"
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
      ButtonCaption3  =   "Atualizar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Utilizado pelo administrador do sistema."
      ButtonKey3      =   "7"
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
      ButtonLeft3     =   93
      ButtonTop3      =   2
      ButtonWidth3    =   50
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
      ButtonLeft4     =   145
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "9"
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
      ButtonLeft5     =   149
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "10"
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
      ButtonLeft6     =   187
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "11"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   215
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7290
         Top             =   270
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFluxodecaixa.frx":1046
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   47
      Top             =   8940
      Width           =   15165
      _ExtentX        =   26749
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   55
      TabIndex        =   31
      Top             =   9210
      Width           =   15165
      Begin VB.TextBox txtSaldo_Atual1 
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
         Left            =   13410
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSaldofinalvenc1 
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSaldo_Atual 
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
         Left            =   10125
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Txtsaldo_inicial 
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
         Left            =   5190
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txttotaldebito 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   8475
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txttotalcredito 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(= Saldo final)"
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
         Left            =   10305
         TabIndex        =   43
         Top             =   150
         Width           =   1200
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+ Saldo vencidos)"
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
         Left            =   11760
         TabIndex        =   37
         Top             =   150
         Width           =   1590
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(= Saldo final)"
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
         Left            =   13590
         TabIndex        =   35
         Top             =   150
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Saldo inicial)"
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
         Left            =   5407
         TabIndex        =   34
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(- Total débito)"
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
         Left            =   8617
         TabIndex        =   33
         Top             =   150
         Width           =   1290
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+ Total crédito)"
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
         Left            =   6922
         TabIndex        =   32
         Top             =   150
         Width           =   1410
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   55
      TabIndex        =   26
      Top             =   9210
      Width           =   15165
      Begin VB.TextBox txtSaldo_Atual2 
         Alignment       =   1  'Right Justify
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
         Left            =   10125
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSaldofinalvenc2 
         Alignment       =   1  'Right Justify
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSaldo_Atual3 
         Alignment       =   1  'Right Justify
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
         Left            =   13410
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Txtsaldo_inicial1 
         Alignment       =   1  'Right Justify
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
         Left            =   5190
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtcontasreceber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtcontaspagar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   8475
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(= Saldo final)"
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
         Left            =   10312
         TabIndex        =   44
         Top             =   150
         Width           =   1200
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+ Saldo vencidos)"
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
         Left            =   11752
         TabIndex        =   38
         Top             =   150
         Width           =   1590
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(= Saldo final)"
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
         Left            =   13597
         TabIndex        =   30
         Top             =   150
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Saldo inicial)"
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
         Left            =   5400
         TabIndex        =   29
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+ Total creditar)"
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
         Left            =   6885
         TabIndex        =   28
         Top             =   150
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(- Total debitar)"
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
         Left            =   8580
         TabIndex        =   27
         Top             =   150
         Width           =   1365
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   55
      TabIndex        =   39
      Top             =   9210
      Width           =   15165
      Begin VB.TextBox txttotalcreditovenc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   10125
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txttotaldebitovenc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSaldofinalvenc 
         Alignment       =   1  'Right Justify
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
         Left            =   13410
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+ Total crédito)"
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
         Left            =   10200
         TabIndex        =   42
         Top             =   150
         Width           =   1410
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(- Total débito)"
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
         Left            =   11910
         TabIndex        =   41
         Top             =   150
         Width           =   1290
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(= Saldo final)"
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
         Left            =   13590
         TabIndex        =   40
         Top             =   150
         Width           =   1200
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   60
      TabIndex        =   4
      Top             =   3450
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11615
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
      TabCaption(0)   =   "Realizado"
      TabPicture(0)   =   "frmFluxodecaixa.frx":4B10
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Projetado"
      TabPicture(1)   =   "frmFluxodecaixa.frx":4B2C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Vencidos"
      TabPicture(2)   =   "frmFluxodecaixa.frx":4B48
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "frmFluxodecaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalVencidas As Double 'OK
Dim NomeBanco As String 'OK
Dim NomeBanco1 As String 'OK
Public NomeBancoRel As String 'OK
Dim AbrirRelSaldo As Boolean 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=nHz9FbPO_ow&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=19&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txtsaldo_inicial = "0,00"
txttotalcredito = "0,00"
txttotaldebito = "0,00"
txtSaldo_Atual = "0,00"
txtSaldo_Atual1 = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos1()
On Error GoTo tratar_erro

Txtsaldo_inicial1 = "0,00"
txtcontasreceber = "0,00"
txtcontaspagar = "0,00"
txtSaldo_Atual2 = "0,00"
txtSaldo_Atual3 = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos2()
On Error GoTo tratar_erro

txttotalcreditovenc = "0,00"
txttotaldebitovenc = "0,00"
txtSaldofinalvenc = "0,00"
txtSaldofinalvenc1 = "0,00"
txtSaldofinalvenc2 = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcCarregaListaBancos
ProcLimparTudo

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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 7, True

ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
SSTab1.Tab = 0
ProcCarregaComboEmpresa Cmb_empresa, False

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

ProcLimpaVariaveisPrincipais
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaBancos()
On Error GoTo tratar_erro

lst_Instituicoes.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_instituicoes where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and DtValidacao IS NOT NULL order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    With lst_Instituicoes.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!int_NBanco), "", TBLISTA!int_NBanco)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!txt_Agencia), "", TBLISTA!txt_Agencia)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!txt_Conta), "", TBLISTA!txt_Conta)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End With
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    If lst_fluxo.ListItems.Count = 0 Then Exit Sub
    'Vencidos
    Familiatext = "({tbl_Fluxo_de_caixa.Operacao} = 'À Creditar' or {tbl_Fluxo_de_caixa.Operacao} = 'À Debitar') and {tbl_Fluxo_de_caixa.ID_empresa} = " & _
                            Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_Fluxo_de_caixa.Data} >= Date(" & Year(msk_fltInicio) & "," & Month(msk_fltInicio) & "," & Day(msk_fltInicio) & ") and {tbl_Fluxo_de_caixa.Data} <= Date(" & Year(msk_fltFim) & "," & Month(msk_fltFim) & "," & Day(msk_fltFim) & ") and {tbl_Fluxo_de_caixa.Bloqueado} = False and " & NomeBancoRel & " or ({tbl_Fluxo_de_caixa.Operacao} = 'Crédito' or {tbl_Fluxo_de_caixa.Operacao} = 'Débito') and {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_Fluxo_de_caixa.Data} >= Date(" & Year(msk_fltInicio) & "," & Month(msk_fltInicio) & "," & Day(msk_fltInicio) & ") and {tbl_Fluxo_de_caixa.Data} <= Date(" & Year(msk_fltFim) & "," & Month(msk_fltFim) & "," & Day(msk_fltFim) & ") and {tbl_Fluxo_de_caixa.Bloqueado} = True and Left({tbl_Fluxo_de_caixa.Descricao}, 6) = 'Cheque' and " & NomeBancoRel
    If AbrirRelSaldo = True Then
        NomeRel = "Contas_fluxodecaixa_realizado_saldos.rpt"
        ProcImprimirRel "{tbl_Fluxo_de_caixa_saldos.Responsavel} = '" & pubUsuario & "'", Familiatext
    Else
        NomeRel = "Contas_fluxodecaixa_realizado.rpt"
        ProcImprimirRel "{tbl_Fluxo_de_caixa_saldos.Responsavel} = '" & pubUsuario & "' and {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_Fluxo_de_caixa.Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {tbl_Fluxo_de_caixa.Data}<= Date(" & _
                        Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and {tbl_Fluxo_de_caixa.Bloqueado} = False and ({tbl_Fluxo_de_caixa.Operacao} = 'Crédito' or {tbl_Fluxo_de_caixa.Operacao} = 'Débito') and " & NomeBancoRel, Familiatext
    End If
Else
    If lst_fluxo1.ListItems.Count = 0 Then Exit Sub
    If msk_fltInicio < Date Then Dataini = msk_fltInicio Else Dataini = Date
    'Vencidos
    Familiatext = "({tbl_Fluxo_de_caixa.Operacao} = 'À Creditar' or {tbl_Fluxo_de_caixa.Operacao} = 'À Debitar') and {tbl_Fluxo_de_caixa.ID_empresa} = " & _
                        Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_Fluxo_de_caixa.Data} < Date(" & Year(Dataini) & "," & Month(Dataini) & "," & Day(Dataini) & ") and {tbl_Fluxo_de_caixa.Bloqueado} = False and " & NomeBancoRel & " or ({tbl_Fluxo_de_caixa.Operacao} = 'Crédito' or {tbl_Fluxo_de_caixa.Operacao} = 'Débito') and {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_Fluxo_de_caixa.Data} < Date(" & Year(Dataini) & "," & Month(Dataini) & "," & Day(Dataini) & ") and {tbl_Fluxo_de_caixa.Bloqueado} = True and Left({tbl_Fluxo_de_caixa.Descricao}, 6) = 'Cheque' and " & NomeBancoRel
    If AbrirRelSaldo = True Then
        NomeRel = "Contas_fluxodecaixa_projetado_saldos.rpt"
        ProcImprimirRel "{tbl_Fluxo_de_caixa_saldos.Responsavel} = '" & pubUsuario & "'", Familiatext
    Else
        NomeRel = "Contas_fluxodecaixa_projetado.rpt"
        ProcImprimirRel "{tbl_Fluxo_de_caixa_saldos.Responsavel} = '" & pubUsuario & "' and {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and ({tbl_Fluxo_de_caixa.Operacao} = 'À Creditar' or {tbl_Fluxo_de_caixa.Operacao} = 'À Debitar') and {tbl_Fluxo_de_caixa.bloqueado} = False and {tbl_Fluxo_de_caixa.Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {tbl_Fluxo_de_caixa.Data}<= Date(" & _
                        Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and " & NomeBancoRel & " or {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and ({tbl_Fluxo_de_caixa.Operacao} = 'Crédito' or {tbl_Fluxo_de_caixa.Operacao} = 'Débito') and {tbl_Fluxo_de_caixa.bloqueado} = True and {tbl_Fluxo_de_caixa.Data}>=Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {tbl_Fluxo_de_caixa.Data}<= Date(" & _
                        Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and Left({tbl_Fluxo_de_caixa.Descricao}, 6) = 'Cheque' and " & NomeBancoRel, Familiatext
    End If
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
Select Case SSTab1.Tab
    Case 0:
        If msk_fltInicio.Value > Date Then
            USMsgBox ("A data inicial não pode ser maior que a data de hoje " & Format(Date, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
            msk_fltInicio.SetFocus
            Exit Sub
        End If
        If msk_fltFim.Value > Date Then
            USMsgBox ("A data final não pode ser maior que a data de hoje " & Format(Date, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
            msk_fltFim.SetFocus
            Exit Sub
        End If
    Case 1:
        If msk_fltFim.Value < Date Then
            USMsgBox ("A data final não pode ser menor que a data de hoje " & Format(Date, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
            msk_fltFim.SetFocus
            Exit Sub
        End If
End Select

If ProcVerifProsseguir = False Then Exit Sub

lst_fluxo.ListItems.Clear
lst_fluxo1.ListItems.Clear
Valor_total = 0
'Soma saldo de todos os bancos
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Saldo) as Valor_Total from tbl_Instituicoes where " & NomeBanco, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Valor_total = IIf(IsNull(TBAbrir!Valor_total), 0, TBAbrir!Valor_total)
End If
TBAbrir.Close

AbrirRelSaldo = False
If SSTab1.Tab = 0 Then
    ProcCarregaListaVencidas "(Operacao = 'À Creditar' or Operacao = 'À Debitar') and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and Bloqueado = 'False' and " & NomeBanco1 & " or (Operacao = 'Crédito' or Operacao = 'Débito') and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and Bloqueado = 'True' and Left(Descricao, 6) = 'Cheque' and " & NomeBanco1
    ProcCarregaListaEfetivada
Else
    If msk_fltInicio < Date Then Dataini = msk_fltInicio Else Dataini = Date
    ProcCarregaListaVencidas "(Operacao = 'À Creditar' or Operacao = 'À Debitar') and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Data < '" & Format(Dataini, "Short Date") & "' and Bloqueado = 'False' and " & NomeBanco1 & " or (Operacao = 'Crédito' or Operacao = 'Débito') and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Data < '" & Format(Dataini, "Short Date") & "' and Bloqueado = 'True' and Left(Descricao, 6) = 'Cheque' and " & NomeBanco1
    ProcCarregaListaNEfetivada
End If

'Gravar datas para pesquisa e saldos
Conexao.Execute "DELETE from tbl_Fluxo_de_caixa_saldos where Responsavel IS NULL or Responsavel = '" & pubUsuario & "'"
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Fluxo_de_caixa_saldos", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!Responsavel = pubUsuario
TBGravar!DataInicial = msk_fltInicio.Value
TBGravar!DataFinal = msk_fltFim.Value
If SSTab1.Tab = 0 Then
    TBGravar!SaldoInicial = Txtsaldo_inicial
    TBGravar!SaldoFinal = txtSaldo_Atual
Else
    TBGravar!SaldoInicial = Txtsaldo_inicial1
    TBGravar!SaldoFinal = txtSaldo_Atual2
End If
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEfetivada()
    On Error GoTo tratar_erro

lst_fluxo.ListItems.Clear
ProcLimpaCampos

'Verifica saldo inicial
TotalCredito = 0
TotalDebito = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Valor_credito) as TotalCredito, Sum(Valor_debito) as TotalDebito from Fluxo_de_caixa_resumido where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Data >= '" & Format(msk_fltInicio.Value, "Short Date") & "' and " & NomeBanco1, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TotalCredito = IIf(IsNull(TBAbrir!TotalCredito), 0, TBAbrir!TotalCredito)
    TotalDebito = IIf(IsNull(TBAbrir!TotalDebito), 0, TBAbrir!TotalDebito)
End If
TBAbrir.Close
Saldo_Anterior = Valor_total - TotalCredito
Saldo_Anterior = Saldo_Anterior + TotalDebito
Txtsaldo_inicial = Format(Saldo_Anterior, "###,##0.00")

'Carrega saldo inicial na lista
With lst_fluxo.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = Format(msk_fltInicio.Value, "dd/mm/yy")
    .Item(.Count).SubItems(3) = "SALDO INICIAL"
    .Item(.Count).SubItems(7) = Format(Txtsaldo_inicial, "###,##0.00")
End With

TotalCredito = 0
TotalDebito = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_Fluxo_de_caixa where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and Bloqueado = 'False' and (Operacao = 'Crédito' or Operacao = 'Débito') and " & NomeBanco1 & " order by Data, Hora, IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        TBLISTA!Saldo_Ant = Format(Saldo_Anterior, "###,##0.00")
        With lst_fluxo.ListItems
            .Add , , TBLISTA!IDFluxo
            .Item(.Count).SubItems(1) = Format(TBLISTA!Data, "dd/mm/yy")
            'Documento
            If Left(TBLISTA!Descricao, 3) = "Che" Or Left(TBLISTA!Descricao, 3) = "Ted" Or Left(TBLISTA!Descricao, 3) = "Doc" Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
            Else
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!int_NotaFiscal), "", TBLISTA!int_NotaFiscal)
            End If
            
            If TBLISTA!Operacao = "À Creditar" Then
                TabelaFiltro = "tbl_Contas_receber"
            Else
                TabelaFiltro = "tbl_ContasPagar"
            End If
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from " & TabelaFiltro & " where IDFluxo = " & TBLISTA!IDFluxo & " and (Antecipacao = 'True' or Devolucao = 'True')", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If TBContas!Antecipacao = True Then Texto = " (ANTECIPAÇÃO)" Else Texto = " (DEVOLUÇÃO)"
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs) & Texto
            Else
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
            End If
            TBContas.Close
            
            .Item(.Count).SubItems(4) = Format(Saldo_Anterior, "###,##0.00")
            If TBLISTA!Operacao = "Crédito" Then
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                Saldo_Anterior = Saldo_Anterior + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
                TotalCredito = TotalCredito + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
            Else
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                Saldo_Anterior = Saldo_Anterior - IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
                TotalDebito = TotalDebito + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
            End If
            .Item(.Count).SubItems(7) = Format(Saldo_Anterior, "###,##0.00")
        End With
        TBLISTA!Saldo_Atual = Format(Saldo_Anterior, "###,##0.00")
        TBLISTA.Update
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
Else
    USMsgBox ("Não existem movimentações neste período."), vbInformation, "CAPRIND v5.0"
    AbrirRelSaldo = True
End If
TBLISTA.Close
txttotaldebito = Format(TotalDebito, "###,##0.00")
txttotalcredito = Format(TotalCredito, "###,##0.00")
txtSaldo_Atual = Format(Saldo_Anterior, "###,##0.00")
txtSaldo_Atual1 = Format(Saldo_Anterior + TotalVencidas, "###,##0.00")

'Carrega saldo final na lista
With lst_fluxo.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = Format(msk_fltFim.Value, "dd/mm/yy")
    .Item(.Count).SubItems(3) = "SALDO FINAL"
    .Item(.Count).SubItems(7) = Format(txtSaldo_Atual, "###,##0.00")
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaNEfetivada()
On Error GoTo tratar_erro

lst_fluxo1.ListItems.Clear
ProcLimpaCampos1

'Verifica saldo inicial
TotalCreditar = 0
TotalDebitar = 0
If msk_fltInicio <> Date Then
    If msk_fltInicio < Date Then
        Dataini = Date - 1
        DataFiltro = "(Data) Between '" & Format(msk_fltInicio, "Short Date") & "' And '" & Format(Dataini, "Short Date") & "'"
    Else
        Dataini = msk_fltInicio - 1
        DataFiltro = "(Data) Between '" & Format(Date, "Short Date") & "' And '" & Format(Dataini, "Short Date") & "'"
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(Valor_creditar) as TotalCreditar, Sum(Valor_debitar) as TotalDebitar from Fluxo_de_caixa_resumido where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & " and " & NomeBanco1, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TotalCreditar = IIf(IsNull(TBAbrir!TotalCreditar), 0, TBAbrir!TotalCreditar)
        TotalDebitar = IIf(IsNull(TBAbrir!TotalDebitar), 0, TBAbrir!TotalDebitar)
    End If
    TBAbrir.Close
    If msk_fltInicio < Date Then
        Saldo_Anterior = Valor_total - TotalCreditar
        Saldo_Anterior = Saldo_Anterior + TotalDebitar
    Else
        Saldo_Anterior = Valor_total + TotalCreditar
        Saldo_Anterior = Saldo_Anterior - TotalDebitar
    End If
Else
    Saldo_Anterior = Valor_total
End If
Txtsaldo_inicial1 = Format(Saldo_Anterior, "###,##0.00")

'Carrega saldo inicial na lista
With lst_fluxo1.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = Format(msk_fltInicio.Value, "dd/mm/yy")
    .Item(.Count).SubItems(3) = "SALDO INICIAL"
    .Item(.Count).SubItems(7) = Format(Txtsaldo_inicial1, "###,##0.00")
End With

TotalCreditar = 0
TotalDebitar = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_Fluxo_de_caixa where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Operacao = 'À Creditar' or Operacao = 'À Debitar') and Bloqueado = 'False' and (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and " & NomeBanco1 & " or ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (Operacao = 'Crédito' or Operacao = 'Débito') and Bloqueado = 'True' and (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and " & NomeBanco1 & " and Left(Descricao, 6) = 'Cheque' order by Data, IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        TBLISTA!Saldo_Ant = Format(Saldo_Anterior, "###,##0.00")
        With lst_fluxo1.ListItems
            .Add , , TBLISTA!IDFluxo
            .Item(.Count).SubItems(1) = Format(TBLISTA!Data, "dd/mm/yy")
            'Documento
'            If TBLISTA!Operacao = "À Creditar" Then
'                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Documento), "", TBLISTA!Documento)
'            ElseIf TBLISTA!Operacao = "À Debitar" Then
'                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!int_NotaFiscal), "", TBLISTA!int_NotaFiscal)
'                Else
'                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
'            End If
            
            If Left(TBLISTA!Descricao, 3) = "Che" Or Left(TBLISTA!Descricao, 3) = "Ted" Or Left(TBLISTA!Descricao, 3) = "Doc" Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
            Else
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!int_NotaFiscal), "", TBLISTA!int_NotaFiscal)
            End If
            
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = Format(Saldo_Anterior, "###,##0.00")
            If TBLISTA!Operacao = "À Creditar" Or TBLISTA!Operacao = "Crédito" Then
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                Saldo_Anterior = Saldo_Anterior + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
                TotalCreditar = TotalCreditar + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
            Else
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                Saldo_Anterior = Saldo_Anterior - IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
                TotalDebitar = TotalDebitar + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
            End If
            .Item(.Count).SubItems(7) = Format(Saldo_Anterior, "###,##0.00")
        End With
        TBLISTA("Saldo_Atual") = Format(Saldo_Anterior, "###,##0.00")
        TBLISTA.Update
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
Else
    USMsgBox ("Não existem movimentos neste período."), vbInformation, "CAPRIND v5.0"
    AbrirRelSaldo = True
End If
TBLISTA.Close
txtcontaspagar = Format(TotalDebitar, "###,##0.00")
txtcontasreceber = Format(TotalCreditar, "###,##0.00")
txtSaldo_Atual2 = Format(Saldo_Anterior, "###,##0.00")
txtSaldo_Atual3 = Format(Saldo_Anterior + TotalVencidas, "###,##0.00")

'Carrega saldo final na lista
With lst_fluxo1.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = Format(msk_fltFim.Value, "dd/mm/yy")
    .Item(.Count).SubItems(3) = "SALDO FINAL"
    .Item(.Count).SubItems(7) = Format(txtSaldo_Atual2, "###,##0.00")
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaVencidas(TextoFiltro As String)
On Error GoTo tratar_erro

lst_fluxo2.ListItems.Clear
ProcLimpaCampos2
TotalCreditar = 0
TotalDebitar = 0
TotalVencidas = 0

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_Fluxo_de_caixa where " & TextoFiltro & " order by Data", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With lst_fluxo2.ListItems
            .Add , , TBLISTA!IDFluxo
            .Item(.Count).SubItems(1) = Format(TBLISTA!Data, "dd/mm/yy")
            'Documento
'            If TBLISTA!Operacao = "À Creditar" Then
'                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Documento), "", TBLISTA!Documento)
'            ElseIf TBLISTA!Operacao = "À debitar" Then
'                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!int_NotaFiscal), "", TBLISTA!int_NotaFiscal)
'                Else
'                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
'            End If
            
            If Left(TBLISTA!Descricao, 3) = "Che" Or Left(TBLISTA!Descricao, 3) = "Ted" Or Left(TBLISTA!Descricao, 3) = "Doc" Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
            Else
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!int_NotaFiscal), "", TBLISTA!int_NotaFiscal)
            End If
            
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            If TBLISTA!Operacao = "À Creditar" Or TBLISTA!Operacao = "Crédito" Then
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                TotalCreditar = TotalCreditar + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
            Else
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                TotalDebitar = TotalDebitar + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
            End If
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close
txttotalcreditovenc = Format(TotalCreditar, "###,##0.00")
txttotaldebitovenc = Format(TotalDebitar, "###,##0.00")
TotalVencidas = TotalCreditar - TotalDebitar
txtSaldofinalvenc = Format(TotalVencidas, "###,##0.00")
txtSaldofinalvenc1 = Format(TotalVencidas, "###,##0.00")
txtSaldofinalvenc2 = Format(TotalVencidas, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362F" Then
    Permitido1 = False
    If USMsgBox("Deseja realmente recriar os dados na tabela fluxo de caixa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from tbl_Fluxo_de_caixa", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            Conexao.Execute "DELETE FROM tbl_Fluxo_de_caixa"
        End If
        
        'Todas operações: Contas a Pagar, Contas a receber, Contas Pagasn, Contas Recebidas e Transferencia
        Mes = 1
        Ano = 2003
Proximo:
        Dataini = Format(1, "00") & "/" & Format(Mes, "00") & "/" & Format(Ano, "0000")
        If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
            DataFim = Format(31, "00") & "/" & Format(Mes, "00") & "/" & Format(Ano, "0000")
        ElseIf Mes = 2 Then
                If IsDate(Format(29, "00") & "/" & Format(Mes, "00") & "/" & Format(Ano, "0000")) = True Then
                    DataFim = Format(29, "00") & "/" & Format(Mes, "00") & "/" & Format(Ano, "0000")
                Else
                    DataFim = Format(28, "00") & "/" & Format(Mes, "00") & "/" & Format(Ano, "0000")
                End If
            Else
                DataFim = Format(30, "00") & "/" & Format(Mes, "00") & "/" & Format(Ano, "0000")
        End If
            
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_ContasPagar where LogSit = 'N' and Bloqueado = 'False' and (dt_pagamento) Between '" & Format(Dataini, "Short Date") & "' And '" & Format(DataFim, "Short Date") & "' order by dt_pagamento", Conexao, adOpenKeyset, adLockOptimistic
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from tbl_Contas_receber where LogSit = 'N' and Bloqueado = 'False' and (vencimento) Between '" & Format(Dataini, "Short Date") & "' And '" & Format(DataFim, "Short Date") & "' order by vencimento", Conexao, adOpenKeyset, adLockOptimistic
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_ContasPagar where logsit = 'S' and (databaixa) Between '" & Format(Dataini, "Short Date") & "' And '" & Format(DataFim, "Short Date") & "' order by databaixa", Conexao, adOpenKeyset, adLockOptimistic
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from tbl_Contas_receber where logsit = 'S' and (data_pagamento) Between '" & Format(Dataini, "Short Date") & "' And '" & Format(DataFim, "Short Date") & "' order by data_pagamento", Conexao, adOpenKeyset, adLockOptimistic
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select* from tbl_instituicoes_transf where (data_transf) Between '" & Format(Dataini, "Short Date") & "' And '" & Format(DataFim, "Short Date") & "' order by data_transf", Conexao, adOpenKeyset, adLockOptimistic
        ProcVerifTodos
        TBContas.Close
        TBFIltro.Close
        TBAbrir.Close
        TBItem.Close
        TBLISTA.Close
        
        Mes = Mes + 1
        If Mes = 13 Then
            Mes = 1
            Ano = Ano + 1
            Ano1 = Year(Date) + 20
            If Ano = Ano1 Then
                Conexao.Execute "Update tbl_ContasPagar Set FormaBaixa = 'Dinheiro' where FormaBaixa = 'À vista' or FormaBaixa = 'A vista'"
                Conexao.Execute "Update tbl_contas_receber Set FormaBaixa = 'Dinheiro' where FormaBaixa = 'À vista' or FormaBaixa = 'A vista'"
                
                'Atualiza campo obs do fluxo de caixa com a tabela tbl_Fluxo_de_caixaOLD
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select * from tbl_Fluxo_de_caixaOLD where Operacao = 'Débito' or Operacao = 'Crédito'", Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    Do While TBLISTA.EOF = False
                        If IsNull(TBLISTA!Obs) = True Or TBLISTA!Obs = "" Then
                            Conexao.Execute "Update tbl_Fluxo_de_caixa Set Obs = '" & TBLISTA!Obs & "' where idintconta = " & TBLISTA!IDintconta & " and (Operacao = 'Débito' or Operacao = 'Crédito')"
                        End If
                        TBLISTA.MoveNext
                    Loop
                End If
                
                'Atualiza cheques compensados
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select * from tbl_Fluxo_de_caixaOLD where Left(Descricao,6) = 'Cheque' and Data >= '" & Format("01/01/2010", "Short Date") & "' and (Operacao = 'Débito' or Operacao = 'Crédito') order by Data, Descricao", Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    Do While TBLISTA.EOF = False
                        If TBLISTA!Bloqueado = True Then Bloqueado = "Bloqueado = 'True'" Else Bloqueado = "Bloqueado = 'False'"
                        Conexao.Execute "Update tbl_Fluxo_de_caixa Set " & Bloqueado & ", Data = '" & TBLISTA!Data & "' where Descricao = '" & TBLISTA!Descricao & "' and (Operacao = 'Débito' or Operacao = 'Crédito')"
                        TBLISTA.MoveNext
                    Loop
                End If
                TBLISTA.Close
                GoTo Terminou
            End If
        End If
        GoTo Proximo
    End If
    If USMsgBox("Deseja realmente atualizar os dados na tabela fluxo de caixa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Permitido1 = True
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from tbl_Fluxo_de_caixa where Left(Descricao,6) <> 'Cheque' and Left(Descricao,3) <> 'Doc' and Left(Descricao,3) <> 'Ted' and Descricao <> 'Saque' and Left(Descricao,8) <> 'Depósito' and Left(Descricao,8) <> 'Deposito' and (Operacao = 'À Debitar' or Operacao = 'Débito') order by IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            PBLista.Min = 0
            PBLista.Max = TBproducao.RecordCount
            PBLista.Value = 1
            Contador = 0
            Do While TBproducao.EOF = False
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_ContasPagar where IDFluxo = " & TBproducao!IDFluxo, Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = True Then
                    Set TBContas = CreateObject("adodb.recordset")
                    TBContas.Open "Select * from tbl_instituicoes_transf where IDFluxo = " & TBproducao!IDFluxo, Conexao, adOpenKeyset, adLockOptimistic
                    If TBContas.EOF = True Then
                        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBproducao!IDFluxo
                    End If
                End If
                TBContas.Close
                TBproducao.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBproducao.Close
    End If

Terminou:
    If Permitido1 = True Then
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Financeiro/Fluxo de caixa"
        Evento = "Atualizar"
        Documento = 0
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

'=====================NÃO APAGAR================================
'Corrige o ID da movimentação de várias contas
'Set TBproducao = CreateObject("adodb.recordset")
'TBproducao.Open "Select ID_empresa, IDFluxo, Operacao, Instituicao, ID_varias, Data from tbl_Fluxo_de_caixa where ID_varias > 0 and ID_varias IS NOT NULL order by ID_varias", Conexao, adOpenKeyset, adLockOptimistic
'If TBproducao.EOF = False Then
'    Conexao.Execute "Truncate Table tbl_Contas_Varias"
'    Conexao.Execute "DBCC CheckIdent('tbl_Contas_Varias',Reseed,1)"
'
'    PBLista.Min = 0
'    PBLista.Max = TBproducao.RecordCount
'    PBLista.Value = 1
'    Contador = 0
'    Do While TBproducao.EOF = False
'
'        Set TBAbrir = CreateObject("adodb.recordset")
'        TBAbrir.Open "Select * from tbl_Contas_Varias", Conexao, adOpenKeyset, adLockOptimistic
'        TBAbrir.AddNew
'        TBAbrir.Update
'        IDlista = TBAbrir!ID
'        TBAbrir.Close
'
'        Set TBContas = CreateObject("adodb.recordset")
'        If TBproducao!Operacao = "Crédito" Then
'            Conexao.Execute "Update tbl_contas_receber Set ID_varias = " & IDlista & ", Banco = '" & TBproducao!Instituicao & "' where ID_empresa = " & TBproducao!ID_empresa & " and ID_varias = " & TBproducao!ID_varias & " and Data_pagamento = '" & TBproducao!Data & "'"
'        Else
'            Conexao.Execute "Update tbl_ContasPagar Set ID_varias = " & IDlista & ", Banco = '" & TBproducao!Instituicao & "' where ID_empresa = " & TBproducao!ID_empresa & " and ID_varias = " & TBproducao!ID_varias & " and DataBaixa = '" & TBproducao!Data & "'"
'        End If
'
'        Conexao.Execute "Update tbl_Fluxo_de_caixa Set ID_varias = " & IDlista & " where IDFluxo = " & TBproducao!IDFluxo
'
'        TBproducao.MoveNext
'        Contador = Contador + 1
'        PBLista.Value = Contador
'    Loop
'
'    Set TBproducao = CreateObject("adodb.recordset")
'    TBproducao.Open "Select ID_empresa, Operacao, ID_varias, Data from tbl_Fluxo_de_caixa where ID_varias > 0 and ID_varias IS NOT NULL order by ID_varias", Conexao, adOpenKeyset, adLockOptimistic
'    If TBproducao.EOF = False Then
'        PBLista.Min = 0
'        PBLista.Max = TBproducao.RecordCount
'        PBLista.Value = 1
'        Contador = 0
'        Do While TBproducao.EOF = False
'            If TBproducao!Operacao = "Crédito" Then
'                Conexao.Execute "Update tbl_contas_receber Set ID_varias = 0 where ID_empresa = " & TBproducao!ID_empresa & " and ID_varias = " & TBproducao!ID_varias & " and Data_pagamento <> '" & TBproducao!Data & "'"
'            Else
'                Conexao.Execute "Update tbl_ContasPagar Set ID_varias = 0 where ID_empresa = " & TBproducao!ID_empresa & " and ID_varias = " & TBproducao!ID_varias & " and DataBaixa <> '" & TBproducao!Data & "'"
'            End If
'
'            TBproducao.MoveNext
'            Contador = Contador + 1
'            PBLista.Value = Contador
'        Loop
'    End If
'
'    Conexao.Execute "Update CR set CR.ID_varias = 0 from tbl_contas_receber CR LEFT JOIN tbl_Fluxo_de_caixa FC ON FC.ID_varias = CR.ID_varias where FC.ID_empresa <> CR.ID_empresa"
'    Conexao.Execute "Update CP set CP.ID_varias = 0 from tbl_ContasPagar CP LEFT JOIN tbl_Fluxo_de_caixa FC ON FC.ID_varias = CP.ID_varias where FC.ID_empresa <> CP.ID_empresa"
'    Conexao.Execute "Update CR set CR.ID_varias = 0 from tbl_contas_receber CR LEFT JOIN tbl_Contas_Varias CV ON CR.ID_varias = CV.ID where CV.ID IS NULL"
'    Conexao.Execute "Update CP set CP.ID_varias = 0 from tbl_ContasPagar CP LEFT JOIN tbl_Contas_Varias CV ON CP.ID_varias = CV.ID where CV.ID IS NULL"
'
'    Set TBproducao = CreateObject("adodb.recordset")
'    TBproducao.Open "Select ID_empresa, Operacao, Valor, ID_varias from tbl_Fluxo_de_caixa where ID_varias > 0 and ID_varias IS NOT NULL order by ID_varias", Conexao, adOpenKeyset, adLockOptimistic
'    If TBproducao.EOF = False Then
'        PBLista.Min = 0
'        PBLista.Max = TBproducao.RecordCount
'        PBLista.Value = 1
'        Contador = 0
'        Do While TBproducao.EOF = False
'            Set TBContas = CreateObject("adodb.recordset")
'            If TBproducao!Operacao = "Crédito" Then
'                TBContas.Open "Select Sum(valortitulorecebido) As Valor from tbl_contas_receber where ID_empresa = " & TBproducao!ID_empresa & " and ID_varias = " & TBproducao!ID_varias, Conexao, adOpenKeyset, adLockOptimistic
'            Else
'                TBContas.Open "Select Sum(ValorPago) As Valor from tbl_ContasPagar where ID_empresa = " & TBproducao!ID_empresa & " and ID_varias = " & TBproducao!ID_varias, Conexao, adOpenKeyset, adLockOptimistic
'            End If
'            Valor = IIf(IsNull(TBContas!Valor), 0, TBContas!Valor)
'            TBproducao!Valor = Format(Valor, "###,##0.000")
'            TBproducao.Update
'
'            TBproducao.MoveNext
'            Contador = Contador + 1
'            PBLista.Value = Contador
'        Loop
'    End If
'End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifTodos()
On Error GoTo tratar_erro

'Todos
If TBContas.EOF = False And TBFIltro.EOF = False And TBAbrir.EOF = False And TBItem.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento And TBContas!dt_Pagamento <= TBAbrir!DataBaixa And TBContas!dt_Pagamento <= TBItem!Data_pagamento And TBContas!dt_Pagamento <= TBLISTA!data_transf Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBFIltro!Vencimento <= TBContas!dt_Pagamento And TBFIltro!Vencimento <= TBAbrir!DataBaixa And TBFIltro!Vencimento <= TBItem!Data_pagamento And TBFIltro!Vencimento <= TBLISTA!data_transf Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBAbrir!DataBaixa <= TBFIltro!Vencimento And TBAbrir!DataBaixa <= TBContas!dt_Pagamento And TBAbrir!DataBaixa <= TBItem!Data_pagamento And TBAbrir!DataBaixa <= TBLISTA!data_transf Then
        ProcAtualizaPagas
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBFIltro!Vencimento And TBItem!Data_pagamento <= TBAbrir!DataBaixa And TBItem!Data_pagamento <= TBContas!dt_Pagamento And TBItem!Data_pagamento <= TBLISTA!data_transf Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
    If TBLISTA!data_transf <= TBContas!dt_Pagamento And TBLISTA!data_transf <= TBFIltro!Vencimento And TBLISTA!data_transf <= TBAbrir!DataBaixa And TBLISTA!data_transf <= TBItem!Data_pagamento Then
        ProcAtualizaTransf
        GoTo Calcular
    End If
End If
'contas à pagar, contas à receber, contas pagas e trasferencia
If TBContas.EOF = False And TBFIltro.EOF = False And TBAbrir.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento And TBContas!dt_Pagamento <= TBAbrir!DataBaixa Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBFIltro!Vencimento <= TBContas!dt_Pagamento And TBFIltro!Vencimento <= TBAbrir!DataBaixa Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBAbrir!DataBaixa <= TBFIltro!Vencimento And TBAbrir!DataBaixa <= TBContas!dt_Pagamento Then
        ProcAtualizaPagas
        GoTo Calcular
    End If
    If TBLISTA!data_transf <= TBContas!dt_Pagamento And TBLISTA!data_transf <= TBFIltro!Vencimento And TBLISTA!data_transf <= TBAbrir!DataBaixa Then
        ProcAtualizaTransf
        GoTo Calcular
    End If
End If
'contas à pagar, contas à receber, contas recebidas e trasferencia
If TBContas.EOF = False And TBFIltro.EOF = False And TBItem.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento And TBContas!dt_Pagamento <= TBItem!Data_pagamento Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBFIltro!Vencimento <= TBContas!dt_Pagamento And TBFIltro!Vencimento <= TBItem!Data_pagamento Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBFIltro!Vencimento And TBItem!Data_pagamento <= TBContas!dt_Pagamento Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
    If TBLISTA!data_transf <= TBContas!dt_Pagamento And TBLISTA!data_transf <= TBFIltro!Vencimento And TBLISTA!data_transf <= TBItem!Data_pagamento Then
        ProcAtualizaTransf
        GoTo Calcular
    End If
    Exit Sub
End If
'contas à pagar, contas pagas, contas recebidas e transferencia
If TBContas.EOF = False And TBAbrir.EOF = False And TBItem.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBAbrir!DataBaixa And TBContas!dt_Pagamento <= TBItem!Data_pagamento Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBAbrir!DataBaixa <= TBContas!dt_Pagamento And TBAbrir!DataBaixa <= TBItem!Data_pagamento Then
        ProcAtualizaPagas
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBAbrir!DataBaixa And TBItem!Data_pagamento <= TBContas!dt_Pagamento Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
    If TBLISTA!data_transf <= TBContas!dt_Pagamento And TBLISTA!data_transf <= TBAbrir!DataBaixa And TBLISTA!data_transf <= TBItem!Data_pagamento Then
        ProcAtualizaTransf
        GoTo Calcular
    End If
End If
'contas receber, contas pagas, contas recebidas e transferencia
If TBFIltro.EOF = False And TBAbrir.EOF = False And TBItem.EOF = False And TBLISTA.EOF = False Then
    If TBFIltro!Vencimento <= TBAbrir!DataBaixa And TBFIltro!Vencimento <= TBItem!Data_pagamento Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBAbrir!DataBaixa <= TBFIltro!Vencimento And TBAbrir!DataBaixa <= TBItem!Data_pagamento Then
        ProcAtualizaPagas
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBFIltro!Vencimento And TBItem!Data_pagamento <= TBAbrir!DataBaixa Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
    If TBLISTA!data_transf <= TBFIltro!Vencimento And TBLISTA!data_transf <= TBAbrir!DataBaixa And TBLISTA!data_transf <= TBItem!Data_pagamento Then
        ProcAtualizaTransf
        GoTo Calcular
    End If
End If
'contas à pagar, contas à receber e contas recebidas e transferencia
If TBContas.EOF = False And TBFIltro.EOF = False And TBItem.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento And TBContas!dt_Pagamento <= TBItem!Data_pagamento Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBFIltro!Vencimento <= TBContas!dt_Pagamento And TBFIltro!Vencimento <= TBItem!Data_pagamento Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBFIltro!Vencimento And TBItem!Data_pagamento <= TBContas!dt_Pagamento Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
    If TBLISTA!data_transf <= TBContas!dt_Pagamento And TBLISTA!data_transf <= TBFIltro!Vencimento And TBLISTA!data_transf <= TBItem!Data_pagamento Then
        ProcAtualizaTransf
        GoTo Calcular
    End If
    Exit Sub
End If
'contas à pagar, contas à receber e transferencia / contas à receber, contas à pagar e transferencia
If TBContas.EOF = False And TBFIltro.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento And TBContas!dt_Pagamento <= TBLISTA!data_transf Then
        ProcAtualizaPagar
    ElseIf TBFIltro!Vencimento <= TBContas!dt_Pagamento And TBFIltro!Vencimento <= TBLISTA!data_transf Then
            ProcAtualizaReceber
        Else
            ProcAtualizaTransf
    End If
End If
'contas pagas, contas recebidas e transferencia / contas recebidas, contas pagas e transferencia
If TBAbrir.EOF = False And TBItem.EOF = False And TBLISTA.EOF = False Then
    If TBAbrir!DataBaixa <= TBItem!Data_pagamento And TBAbrir!DataBaixa <= TBLISTA!data_transf Then
        ProcAtualizaPagas
    ElseIf TBItem!Data_pagamento <= TBAbrir!DataBaixa And TBItem!Data_pagamento <= TBLISTA!data_transf Then
            ProcAtualizaRecebidas
        Else
            ProcAtualizaTransf
    End If
End If

'contas à pagar, contas à receber e contas pagas
If TBContas.EOF = False And TBFIltro.EOF = False And TBAbrir.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento And TBContas!dt_Pagamento <= TBAbrir!DataBaixa Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBFIltro!Vencimento <= TBContas!dt_Pagamento And TBFIltro!Vencimento <= TBAbrir!DataBaixa Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBAbrir!DataBaixa <= TBFIltro!Vencimento And TBAbrir!DataBaixa <= TBContas!dt_Pagamento Then
        ProcAtualizaPagas
        GoTo Calcular
    End If
End If
'contas à pagar, contas à receber e contas recebidas
If TBContas.EOF = False And TBFIltro.EOF = False And TBItem.EOF = False And TBLISTA.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento And TBContas!dt_Pagamento <= TBItem!Data_pagamento Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBFIltro!Vencimento <= TBContas!dt_Pagamento And TBFIltro!Vencimento <= TBItem!Data_pagamento Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBFIltro!Vencimento And TBItem!Data_pagamento <= TBContas!dt_Pagamento Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
    Exit Sub
End If
'contas à pagar, contas pagas e contas recebidas
If TBContas.EOF = False And TBAbrir.EOF = False And TBItem.EOF = False Then
    If TBContas!dt_Pagamento <= TBAbrir!DataBaixa And TBContas!dt_Pagamento <= TBItem!Data_pagamento Then
        ProcAtualizaPagar
        GoTo Calcular
    End If
    If TBAbrir!DataBaixa <= TBContas!dt_Pagamento And TBAbrir!DataBaixa <= TBItem!Data_pagamento Then
        ProcAtualizaPagas
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBAbrir!DataBaixa And TBItem!Data_pagamento <= TBContas!dt_Pagamento Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
End If
'contas receber, contas pagas e contas recebidas
If TBFIltro.EOF = False And TBAbrir.EOF = False And TBItem.EOF = False Then
    If TBFIltro!Vencimento <= TBAbrir!DataBaixa And TBFIltro!Vencimento <= TBItem!Data_pagamento Then
        ProcAtualizaReceber
        GoTo Calcular
    End If
    If TBAbrir!DataBaixa <= TBFIltro!Vencimento And TBAbrir!DataBaixa <= TBItem!Data_pagamento Then
        ProcAtualizaPagas
        GoTo Calcular
    End If
    If TBItem!Data_pagamento <= TBFIltro!Vencimento And TBItem!Data_pagamento <= TBAbrir!DataBaixa Then
        ProcAtualizaRecebidas
        GoTo Calcular
    End If
End If
'contas à pagar e contas à receber / contas à receber e contas à pagar
If TBContas.EOF = False And TBFIltro.EOF = False Then
    If TBContas!dt_Pagamento <= TBFIltro!Vencimento Then
        ProcAtualizaPagar
    Else
        ProcAtualizaReceber
    End If
End If
'contas pagas e contas recebidas / contas recebidas e contas pagas
If TBAbrir.EOF = False And TBItem.EOF = False Then
    If TBAbrir!DataBaixa <= TBItem!Data_pagamento Then
        ProcAtualizaPagas
    Else
        ProcAtualizaRecebidas
    End If
End If

If TBContas.EOF = False Then ProcAtualizaPagar
If TBFIltro.EOF = False Then ProcAtualizaReceber
If TBAbrir.EOF = False Then ProcAtualizaPagas
If TBItem.EOF = False Then ProcAtualizaRecebidas
If TBLISTA.EOF = False Then ProcAtualizaTransf

Calcular:

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaPagar()
On Error GoTo tratar_erro

If TBContas.EOF = False Then
    Do While TBContas.EOF = False
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from tbl_Fluxo_de_caixa", Conexao, adOpenKeyset, adLockOptimistic
        TBproducao.AddNew
        TBproducao!IDintconta = TBContas!IDintconta
        If IsNull(TBContas!dt_Pagamento) = False Then TBproducao!Data = TBContas!dt_Pagamento
        Operacao = "À Debitar"
        TBproducao!Operacao = Operacao
        TBproducao!status = TBContas!Logsit
        If IsNull(TBContas!dbl_valorpagto) = False Then TBproducao!valor = TBContas!dbl_valorpagto
        If IsNull(TBContas!txt_ndocumento) = False Then TBproducao!int_NotaFiscal = TBContas!txt_ndocumento
        If IsNull(TBContas!Txt_pedido) = False Then TBproducao!Documento = TBContas!Txt_pedido
        If IsNull(TBContas!Txt_fornecedor) = False Then TBproducao!Descricao = Trim(TBContas!Txt_fornecedor)
        If IsNull(TBContas!tituloref) = False And TBContas!tituloref <> "" Then TBproducao!tituloref = TBContas!tituloref
        TBproducao!Bloqueado = False
        TBproducao!ID_empresa = IIf(IsNull(TBContas!ID_empresa), 0, TBContas!ID_empresa)
        TBContas!IDFluxo = TBproducao!IDFluxo
        TBContas.Update
        TBproducao.Update
        TBproducao.Close
        TBContas.MoveNext
        ProcVerifTodos
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaPagas()
On Error GoTo tratar_erro

If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from tbl_Fluxo_de_caixa", Conexao, adOpenKeyset, adLockOptimistic
        TBproducao.AddNew
        TBproducao!IDintconta = TBAbrir!IDintconta
        TBproducao!Data = TBAbrir!DataBaixa
        Operacao = "Débito"
        TBproducao!Operacao = Operacao
        TBproducao!status = TBAbrir!Logsit
        If IsNull(TBAbrir!ValorPago) = False Then TBproducao!valor = TBAbrir!ValorPago
        TBproducao!Bloqueado = False
        If TBAbrir!FormaBaixa = "SAQUE" Then TBproducao!Bloqueado = True
        If TBAbrir!FormaBaixa = "CHEQUE" Or TBAbrir!FormaBaixa = "CHEQUE PRÉ-DATADO" Or TBAbrir!FormaBaixa = "DOC" Or TBAbrir!FormaBaixa = "TED" Then
            If IsNull(TBAbrir!NDoctoBaixa) = False And TBAbrir!NDoctoBaixa <> "" Then
                TBproducao!Cheque = TBAbrir!NDoctoBaixa
                TBproducao!Bloqueado = True
                
                valor = 0
                Set TBMaquinas = CreateObject("adodb.recordset")
                If TBAbrir!FormaBaixa = "CHEQUE" Or TBAbrir!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                    TBMaquinas.Open "Select Sum(ValorPago) as Valor from tbl_ContasPagar where NDoctoBaixa = '" & TBAbrir!NDoctoBaixa & "' and Banco = '" & TBAbrir!Banco & "' and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO')", Conexao, adOpenKeyset, adLockOptimistic
                ElseIf TBAbrir!FormaBaixa = "DOC" Then
                        TBMaquinas.Open "Select Sum(ValorPago) as Valor from tbl_ContasPagar where NDoctoBaixa = '" & TBAbrir!NDoctoBaixa & "' and FormaBaixa = 'DOC' and Banco = '" & TBAbrir!Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Else
                        TBMaquinas.Open "Select Sum(ValorPago) as Valor from tbl_ContasPagar where NDoctoBaixa = '" & TBAbrir!NDoctoBaixa & "' and FormaBaixa = 'TED' and Banco = '" & TBAbrir!Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                End If
                If TBMaquinas.EOF = False Then
                    valor = IIf(IsNull(TBMaquinas!valor), 0, TBMaquinas!valor)
                End If
                TBMaquinas.Close
                If valor <> 0 Then
                    'Fluxo de Caixa
                    'Cria registro com o valor total da operação
                    If TBAbrir!FormaBaixa = "CHEQUE" Or TBAbrir!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                        Cheque = "Cheque n. " & TBAbrir!NDoctoBaixa
                    ElseIf TBAbrir!FormaBaixa = "DOC" Then
                            Cheque = "Doc n. " & TBAbrir!NDoctoBaixa
                        Else
                            Cheque = "Ted n. " & TBAbrir!NDoctoBaixa
                    End If
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Débito' and Descricao = '" & Cheque & "' and Instituicao = '" & TBAbrir!Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then TBGravar.AddNew
                    TBGravar!Operacao = "Débito"
                    TBGravar!Data = TBAbrir!DataBaixa
                    TBGravar!valor = valor
                    TBGravar!Bloqueado = False
                    Select Case TBAbrir!FormaBaixa
                        Case "CHEQUE": TBGravar!Descricao = Cheque
                        Case "CHEQUE PRÉ-DATADO": TBGravar!Descricao = Cheque
                        Case "DOC": TBGravar!Descricao = Cheque
                        Case "TED": TBGravar!Descricao = Cheque
                    End Select
                    TBGravar!Obs = TBGravar!Descricao
                    
                    If TBAbrir!FormaBaixa = "CHEQUE" Or TBAbrir!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                        If IsNull(TBAbrir!Bom_para) = True Or TBAbrir!Bom_para = "" Then
                            If TBAbrir!DataBaixa < Date Then TBGravar!Bloqueado = False Else TBGravar!Bloqueado = True
                        Else
                            If TBAbrir!Bom_para < Date Then TBGravar!Bloqueado = False Else TBGravar!Bloqueado = True
                        End If
                    End If
                    
                    TBGravar!status = "S"
                    TBGravar!Instituicao = TBAbrir!Banco
                    TBGravar!Hora = Now
                    TBGravar!Cheque = TBAbrir!NDoctoBaixa
                    TBGravar.Update
                    TBGravar.Close
                End If
            End If
        End If
        If IsNull(TBAbrir!txt_ndocumento) = False Then TBproducao!int_NotaFiscal = TBAbrir!txt_ndocumento
        If IsNull(TBAbrir!Txt_pedido) = False Then TBproducao!Documento = TBAbrir!Txt_pedido
        If IsNull(TBAbrir!Txt_fornecedor) = False Then TBproducao!Descricao = Trim(TBAbrir!Txt_fornecedor)
        If IsNull(TBAbrir!Banco) = False Then TBproducao!Instituicao = TBAbrir!Banco
        If IsNull(TBAbrir!tituloref) = False And TBAbrir!tituloref <> "" Then TBproducao!tituloref = TBAbrir!tituloref
        TBproducao!ID_empresa = IIf(IsNull(TBAbrir!ID_empresa), 0, TBAbrir!ID_empresa)
        TBproducao!Hora = Now
        Contador = 0
        Do While Contador <> 9999
            Contador = Contador + 1
        Loop
        TBproducao.Update
        TBAbrir!IDFluxo = TBproducao!IDFluxo
        TBAbrir.Update
        TBproducao.Close
        TBAbrir.MoveNext
        ProcVerifTodos
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaReceber()
On Error GoTo tratar_erro

If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from tbl_Fluxo_de_caixa", Conexao, adOpenKeyset, adLockOptimistic
        TBproducao.AddNew
        TBproducao!IDintconta = TBFIltro!IDintconta
        TBproducao!Data = TBFIltro!Vencimento
        Operacao = "À Creditar"
        TBproducao!Operacao = Operacao
        TBproducao!status = TBFIltro!Logsit
        If IsNull(TBFIltro!valor) = False Then TBproducao!valor = TBFIltro!valor
        If IsNull(TBFIltro!NFiscal) = False Then TBproducao!int_NotaFiscal = TBFIltro!NFiscal
        If IsNull(TBFIltro!txt_ndocumento) = False Then TBproducao!Documento = TBFIltro!txt_ndocumento
        If IsNull(TBFIltro!Nome_Razao) = False Then TBproducao!Descricao = Trim(TBFIltro!Nome_Razao)
        If IsNull(TBFIltro!tituloref) = False And TBFIltro!tituloref <> "" Then TBproducao!tituloref = TBFIltro!tituloref
        TBproducao!Bloqueado = False
        TBproducao!ID_empresa = IIf(IsNull(TBFIltro!ID_empresa), 0, TBFIltro!ID_empresa)
        If TBFIltro!IDtrocatitulo <> 0 Then
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select * from troca_titulo where ID = " & TBFIltro!IDtrocatitulo & " and Vlrtotalresgatado <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                TBproducao!Bloqueado = True
                'Fluxo de Caixa
                'Cria registro com o valor total da operação
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from tbl_Fluxo_de_caixa where descricao = 'Desconto de duplicata borderô n. " & TBFIltro!IDtrocatitulo & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = True Then TBGravar.AddNew
                TBGravar!Operacao = "Crédito"
                If IsNull(TBMaquinas!Data_operacao) = False And TBMaquinas!Data_operacao <> "" Then
                    TBGravar!Data = TBMaquinas!Data_operacao
                Else
                    TBGravar!Data = TBMaquinas!Data
                End If
                TBGravar!valor = TBMaquinas!Vlrtotalresgatado
                TBGravar!Descricao = "Desconto de duplicata borderô n. " & TBFIltro!IDtrocatitulo
                TBGravar!Instituicao = TBMaquinas!banco_recebedor
                TBGravar!status = "S"
                TBGravar!Hora = Now
                TBGravar!Cheque = TBFIltro!IDtrocatitulo
                TBGravar.Update
                TBGravar.Close
            End If
            TBMaquinas.Close
        End If
        TBFIltro!IDFluxo = TBproducao!IDFluxo
        TBFIltro.Update
        TBproducao.Update
        TBproducao.Close
        TBFIltro.MoveNext
        ProcVerifTodos
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaRecebidas()
On Error GoTo tratar_erro

If TBItem.EOF = False Then
    Do While TBItem.EOF = False
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from tbl_Fluxo_de_caixa", Conexao, adOpenKeyset, adLockOptimistic
        TBproducao.AddNew
        TBproducao!IDintconta = TBItem!IDintconta
        TBproducao!Data = TBItem!Data_pagamento
        Operacao = "Crédito"
        TBproducao!Operacao = Operacao
        TBproducao!status = TBItem!Logsit
        If IsNull(TBItem!valortitulorecebido) = False Then TBproducao!valor = TBItem!valortitulorecebido
        TBproducao!Bloqueado = False
        If TBItem!FormaBaixa = "CHEQUE" Or TBItem!FormaBaixa = "CHEQUE PRÉ-DATADO" Or TBItem!FormaBaixa = "DOC" Or TBItem!FormaBaixa = "TED" Then
            If IsNull(TBItem!NDoctoBaixa) = False And TBItem!NDoctoBaixa <> "" Then
                TBproducao!Bloqueado = False
                TBproducao!Cheque = TBItem!NDoctoBaixa
                valor = 0
                Set TBMaquinas = CreateObject("adodb.recordset")
                If TBItem!FormaBaixa = "CHEQUE" Or TBItem!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                    TBMaquinas.Open "Select Sum(valortitulorecebido) as valor from tbl_contas_receber where NDoctoBaixa = '" & TBItem!NDoctoBaixa & "' and Banco = '" & TBItem!Banco & "' and (FormaBaixa = 'CHEQUE' or FormaBaixa = 'CHEQUE PRÉ-DATADO')", Conexao, adOpenKeyset, adLockOptimistic
                ElseIf TBItem!FormaBaixa = "DOC" Then
                        TBMaquinas.Open "Select Sum(valortitulorecebido) as valor from tbl_contas_receber where NDoctoBaixa = '" & TBItem!NDoctoBaixa & "' and FormaBaixa = 'DOC' and Banco = '" & TBItem!Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Else
                        TBMaquinas.Open "Select Sum(valortitulorecebido) as valor from tbl_contas_receber where NDoctoBaixa = '" & TBItem!NDoctoBaixa & "' and FormaBaixa = 'TED' and Banco = '" & TBItem!Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                End If
                If TBMaquinas.EOF = False Then
                    valor = IIf(IsNull(TBMaquinas!valor), 0, TBMaquinas!valor)
                End If
                TBMaquinas.Close
                If valor <> 0 Then
                    'Fluxo de Caixa
                    'Cria registro com o valor total da operação
                    If TBItem!FormaBaixa = "CHEQUE" Or TBItem!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                        Cheque = "Cheque n. " & TBItem!NDoctoBaixa
                    ElseIf TBItem!FormaBaixa = "DOC" Then
                            Cheque = "Doc n. " & TBItem!NDoctoBaixa
                        Else
                            Cheque = "Ted n. " & TBItem!NDoctoBaixa
                    End If
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Crédito' and Descricao = '" & Cheque & "' and Instituicao = '" & TBItem!Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then TBGravar.AddNew
                    TBGravar!Operacao = "Crédito"
                    TBGravar!Data = TBItem!Data_pagamento
                    TBGravar!valor = valor
                    TBGravar!Bloqueado = False
                    Select Case TBItem!FormaBaixa
                        Case "CHEQUE": TBGravar!Descricao = Cheque
                        Case "CHEQUE PRÉ-DATADO": TBGravar!Descricao = Cheque
                        Case "DOC": TBGravar!Descricao = Cheque
                        Case "TED": TBGravar!Descricao = Cheque
                    End Select
                    TBGravar!Obs = TBGravar!Descricao
                    
                    If TBItem!FormaBaixa = "CHEQUE" Or TBItem!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                        If TBItem!Data_pagamento < Date Then TBGravar!Bloqueado = False Else TBGravar!Bloqueado = True
                    End If
                    
                    TBGravar!status = "S"
                    TBGravar!Instituicao = TBItem!Banco
                    TBGravar!Hora = Now
                    TBGravar!Cheque = TBItem!NDoctoBaixa
                    TBGravar.Update
                    TBGravar.Close
                End If
            End If
        End If
        If IsNull(TBItem!NFiscal) = False Then TBproducao!int_NotaFiscal = TBItem!NFiscal
        If IsNull(TBItem!txt_ndocumento) = False Then TBproducao!Documento = TBItem!txt_ndocumento
        If IsNull(TBItem!Nome_Razao) = False Then TBproducao!Descricao = Trim(TBItem!Nome_Razao)
        If IsNull(TBItem!Banco) = False Then TBproducao!Instituicao = TBItem!Banco
        If IsNull(TBItem!tituloref) = False And TBItem!tituloref <> "" Then TBproducao!tituloref = TBItem!tituloref
        TBproducao!ID_empresa = IIf(IsNull(TBItem!ID_empresa), 0, TBItem!ID_empresa)
        If TBItem!IDtrocatitulo <> 0 Then
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select * from troca_titulo where ID = " & TBItem!IDtrocatitulo & " and Vlrtotalresgatado <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                TBproducao!Bloqueado = True
                'Fluxo de Caixa
                'Cria registro com o valor total da operação
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from tbl_Fluxo_de_caixa where descricao = 'Desconto de duplicata borderô n. " & TBItem!IDtrocatitulo & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = True Then TBGravar.AddNew
                TBGravar!Operacao = "Crédito"
                If IsNull(TBMaquinas!Data_operacao) = False And TBMaquinas!Data_operacao <> "" Then
                    TBGravar!Data = TBMaquinas!Data_operacao
                Else
                    TBGravar!Data = TBMaquinas!Data
                End If
                TBGravar!valor = TBMaquinas!Vlrtotalresgatado
                TBGravar!Descricao = "Desconto de duplicata borderô n. " & TBItem!IDtrocatitulo
                TBGravar!Obs = TBGravar!Descricao
                TBGravar!Instituicao = TBMaquinas!banco_recebedor
                TBGravar!status = "S"
                TBGravar!Hora = Now
                TBGravar!Cheque = TBItem!IDtrocatitulo
                TBGravar!ID_empresa = IIf(IsNull(TBItem!ID_empresa), 0, TBItem!ID_empresa)
                TBGravar!Bloqueado = False
                TBGravar.Update
                TBGravar.Close
            End If
            TBMaquinas.Close
        End If
        TBproducao!Hora = Now
        Contador = 0
        Do While Contador <> 9999
            Contador = Contador + 1
        Loop
        TBproducao.Update
        TBItem!IDFluxo = TBproducao!IDFluxo
        TBItem.Update
        TBproducao.Close
        TBItem.MoveNext
        ProcVerifTodos
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaTransf()
On Error GoTo tratar_erro

If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select ID_empresa from tbl_Instituicoes where txt_Descricao = " & TBLISTA!banco_remetente, Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            IDempresa = TBMaquinas!ID_empresa
        End If
        TBMaquinas.Close
        If TBLISTA!Tipo = "S" Then
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBLISTA!IDFluxo), 0, TBLISTA!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = True Then TBproducao.AddNew
            TBproducao!IDintconta = TBLISTA!id_transf
            TBproducao!Operacao = "Débito"
            TBproducao!Data = TBLISTA!data_transf
            TBproducao!valor = TBLISTA!valor_transf
            TBproducao!Descricao = "Saque"
            TBproducao!Obs = TBproducao!Descricao
            TBproducao!Instituicao = TBLISTA!banco_remetente
            TBproducao!status = "S"
            TBproducao!Hora = Now
            TBproducao!Bloqueado = False
            TBproducao!ID_empresa = IDempresa
            TBproducao.Update
            TBLISTA!IDFluxo = TBproducao!IDFluxo
            TBLISTA.Update
        Else
            'Cria cheque na tabela de contas a pagar
            If TBLISTA!Tipo = "D" And TBLISTA!FormaBaixa = "CHEQUE" Then
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from tbl_ContasPagar where NDoctoBaixa = '" & TBLISTA!NDoctoBaixa & "' and Banco = '" & TBLISTA!banco_remetente & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = True Then TBMaquinas.AddNew
                TBMaquinas!Logsit = ""
                TBMaquinas!Txt_fornecedor = TBLISTA!banco_recebedor
                TBMaquinas!FormaBaixa = "CHEQUE"
                TBMaquinas!DataBaixa = TBLISTA!data_transf
                TBMaquinas!ValorPago = TBLISTA!valor_transf
                TBMaquinas!NDoctoBaixa = TBLISTA!NDoctoBaixa
                TBMaquinas!Banco = TBLISTA!banco_remetente
                TBMaquinas!Favorecido = ""
                TBMaquinas!status = "DEPÓSITO EM CHEQUE"
                TBMaquinas!resppag = TBLISTA!Responsavel
                TBMaquinas!ID_empresa = IDempresa
                TBAliquota.Close
                TBMaquinas.Update
                TBMaquinas.Close
            End If
                    
            'Fluxo de Caixa
            Set TBproducao = CreateObject("adodb.recordset")
            If TBLISTA!FormaBaixa <> "" Then
                If TBLISTA!Tipo = "T" Then
                    Select Case TBLISTA!FormaBaixa
                        Case "DOC": TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Crédito' and Descricao = '" & "Doc n. " & TBLISTA!NDoctoBaixa & "' and Instituicao = '" & TBLISTA!banco_recebedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Case "TED": TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Crédito' and Descricao = '" & "Ted n. " & TBLISTA!NDoctoBaixa & "' and Instituicao = '" & TBLISTA!banco_recebedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                    End Select
                Else
                    If TBLISTA!FormaBaixa = "Dinheiro" Then
                        TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Crédito' and Instituicao = '" & TBLISTA!banco_recebedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Else
                        TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Crédito' and Descricao = '" & "Cheque n. " & TBLISTA!NDoctoBaixa & "' and Instituicao = '" & TBLISTA!banco_recebedor & "'", Conexao, adOpenKeyset, adLockOptimistic
                    End If
                End If
            Else
                TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Crédito' and Descricao = 'Transferência' and Instituicao = '" & TBLISTA!banco_recebedor & "'", Conexao, adOpenKeyset, adLockOptimistic
            End If
            If TBproducao.EOF = True Then TBproducao.AddNew
            TBproducao!IDintconta = TBLISTA!id_transf
            TBproducao!Operacao = "Crédito"
            TBproducao!Data = TBLISTA!data_transf
            TBproducao!valor = TBLISTA!valor_transf
            If TBLISTA!Tipo = "T" Then
                Select Case TBLISTA!FormaBaixa
                    Case "DOC": TBproducao!Descricao = "Doc n. " & TBLISTA!NDoctoBaixa
                    Case "TED": TBproducao!Descricao = "Ted n. " & TBLISTA!NDoctoBaixa
                End Select
            Else
                If TBLISTA!FormaBaixa = "Dinheiro" Then
                    TBproducao!Descricao = "Depósito"
                Else
                    TBproducao!Descricao = "Cheque n. " & TBLISTA!NDoctoBaixa
                End If
            End If
            TBproducao!Obs = TBproducao!Descricao
            TBproducao!Instituicao = TBLISTA!banco_recebedor
            TBproducao!status = "S"
            TBproducao!Hora = Now
            If TBLISTA!NDoctoBaixa <> "" Then TBproducao!Cheque = TBLISTA!NDoctoBaixa
            If TBLISTA!FormaBaixa = "CHEQUE" Then TBproducao!Bloqueado = True Else TBproducao!Bloqueado = False
            TBproducao!ID_empresa = IDempresa
            TBproducao.Update
            TBproducao.Close
            
            Contador = 0
            Do While Contador <> 9999
                Contador = Contador + 1
            Loop
            
            Set TBproducao = CreateObject("adodb.recordset")
            If TBLISTA!FormaBaixa <> "" Then
                If TBLISTA!Tipo = "T" Then
                    Select Case TBLISTA!FormaBaixa
                        Case "DOC": TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Débito' and Descricao = '" & "Doc n. " & TBLISTA!NDoctoBaixa & "' and Instituicao = '" & TBLISTA!banco_remetente & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Case "TED": TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Débito' and Descricao = '" & "Ted n. " & TBLISTA!NDoctoBaixa & "' and Instituicao = '" & TBLISTA!banco_remetente & "'", Conexao, adOpenKeyset, adLockOptimistic
                    End Select
                Else
                    If TBLISTA!FormaBaixa = "Dinheiro" Then
                        TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Débito' and Instituicao = '" & TBLISTA!banco_remetente & "'", Conexao, adOpenKeyset, adLockOptimistic
                    Else
                        TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Débito' and Descricao = '" & "Cheque n. " & TBLISTA!NDoctoBaixa & "' and Instituicao = '" & TBLISTA!banco_remetente & "'", Conexao, adOpenKeyset, adLockOptimistic
                    End If
                End If
            Else
                TBproducao.Open "Select * from tbl_Fluxo_de_caixa where IDIntconta = " & TBLISTA!id_transf & " and Operacao = 'Débito' and Descricao = 'Transferência' and Instituicao = '" & TBLISTA!banco_remetente & "'", Conexao, adOpenKeyset, adLockOptimistic
            End If
            If TBproducao.EOF = True Then TBproducao.AddNew
            TBproducao!IDintconta = TBLISTA!id_transf
            TBproducao!Operacao = "Débito"
            TBproducao!Data = TBLISTA!data_transf
            TBproducao!valor = TBLISTA!valor_transf
            If TBLISTA!Tipo = "T" Then
                Select Case TBLISTA!FormaBaixa
                    Case "DOC": TBproducao!Descricao = "Doc n. " & TBLISTA!NDoctoBaixa
                    Case "TED": TBproducao!Descricao = "Ted n. " & TBLISTA!NDoctoBaixa
                End Select
            Else
                If TBLISTA!FormaBaixa = "Dinheiro" Then
                    TBproducao!Descricao = "Depósito"
                Else
                    TBproducao!Descricao = "Cheque n. " & TBLISTA!NDoctoBaixa
                End If
            End If
            TBproducao!Obs = TBproducao!Descricao
            TBproducao!Instituicao = TBLISTA!banco_remetente
            TBproducao!status = "S"
            TBproducao!Hora = Now
            If TBLISTA!NDoctoBaixa <> "" Then TBproducao!Cheque = TBLISTA!NDoctoBaixa
            TBproducao!ID_empresa = IDempresa
            If TBLISTA!FormaBaixa = "CHEQUE" Then TBproducao!Bloqueado = True Else TBproducao!Bloqueado = False
            TBproducao.Update
        End If
        TBproducao.Close
        TBLISTA.MoveNext
        ProcVerifTodos
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_fluxo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lst_fluxo, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_fluxo1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lst_fluxo1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_fluxo2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lst_fluxo2, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_Instituicoes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lst_Instituicoes
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
    ProcLimparTudo
Else
    ProcOrdenaListView lst_Instituicoes, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_Instituicoes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

lst_fluxo.ListItems.Clear
lst_fluxo1.ListItems.Clear
lst_fluxo2.ListItems.Clear
ProcLimpaCampos
ProcLimpaCampos1
ProcLimpaCampos2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

With USToolBar1
    .ButtonState(2) = 0
    Select Case SSTab1.Tab
        Case 0:
            lst_fluxo.Visible = True
            lst_fluxo1.Visible = False
            lst_fluxo2.Visible = False
            Frame2.Visible = True
            Frame4.Visible = False
            Frame5.Visible = False
        Case 1:
            lst_fluxo.Visible = False
            lst_fluxo1.Visible = True
            lst_fluxo2.Visible = False
            Frame2.Visible = False
            Frame4.Visible = True
            Frame5.Visible = False
        Case 2:
            .ButtonState(2) = 5
            lst_fluxo.Visible = False
            lst_fluxo1.Visible = False
            lst_fluxo2.Visible = True
            Frame2.Visible = False
            Frame4.Visible = False
            Frame5.Visible = True
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifProsseguir() As Boolean
On Error GoTo tratar_erro

NomeBanco = "txt_Descricao IS NULL and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
NomeBanco1 = "Instituicao IS NULL and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
NomeBancoRel = "ISNULL({tbl_Fluxo_de_caixa.Instituicao}) and {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
Permitido = False
ProcVerifProsseguir = True
With lst_Instituicoes
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Permitido = True
            NomeBanco = "(" & NomeBanco & " or txt_Descricao = '" & .ListItems.Item(InitFor).ListSubItems(4).Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ")"
            NomeBanco1 = "(" & NomeBanco1 & " or Instituicao = '" & .ListItems.Item(InitFor).ListSubItems(4).Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ")"
            NomeBancoRel = "(" & NomeBancoRel & " or {tbl_Fluxo_de_caixa.Instituicao} = '" & .ListItems.Item(InitFor).ListSubItems(4).Text & "' and {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & ")"
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) banco(s) na lista antes de filtrar."), vbExclamation, "CAPRIND v5.0"
    ProcVerifProsseguir = False
    Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 3: ProcAtualizar
    Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

