VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCargaMaq 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Carga de posto de trabalho"
   ClientHeight    =   9975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
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
      FormHeightDT    =   10440
      FormWidthDT     =   15480
      FormScaleHeightDT=   9975
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.CheckBox chk_HoraExtra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hora extra"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox chk_Feriado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Feriado"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3750
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chk_Domingo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Domingo"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2620
      TabIndex        =   6
      Top             =   1080
      Width           =   915
   End
   Begin VB.CheckBox chk_Sabado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sábado"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1550
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13800
      Top             =   270
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCargaMaq.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame2 
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
      Height          =   1035
      Left            =   55
      TabIndex        =   13
      Top             =   8910
      Width           =   15195
      Begin VB.TextBox txtTempoDisp 
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
         Left            =   4290
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Horas disponíveis no período."
         Top             =   180
         Width           =   2955
      End
      Begin VB.TextBox txtUltrapassadas 
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Horas ultrapassadas  no período"
         Top             =   570
         Width           =   2955
      End
      Begin VB.TextBox txtTempoRestante 
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Horas restantes no período."
         Top             =   180
         Width           =   2955
      End
      Begin VB.TextBox txtUtilizada 
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
         Left            =   4290
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Horas utilizadas no período."
         Top             =   570
         Width           =   2955
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas disponíveis no período :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2085
         TabIndex        =   21
         Top             =   180
         Width           =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horas ultrapassadas  no período : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   7500
         TabIndex        =   17
         Top             =   570
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horas restantes no período :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   7845
         TabIndex        =   16
         Top             =   180
         Width           =   2070
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Horas utilizadas no período :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1470
         TabIndex        =   15
         Top             =   570
         Width           =   2775
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6465
      Left            =   60
      TabIndex        =   8
      Top             =   2160
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11404
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "T"
         Text            =   "Cód. posto"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9534
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Disponivel"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Utilizado"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Restantes"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Ultrapassadas"
         Object.Width           =   3528
      EndProperty
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
      Height          =   825
      Left            =   55
      TabIndex        =   14
      Top             =   1320
      Width           =   15195
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
         ItemData        =   "frmCargaMaq.frx":348D
         Left            =   180
         List            =   "frmCargaMaq.frx":349A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   1845
      End
      Begin VB.ComboBox CmbTexto 
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
         Left            =   2040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   6255
      End
      Begin VB.TextBox txtdias 
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
         Left            =   13470
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Total de dias no período."
         Top             =   390
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   11040
         TabIndex        =   2
         ToolTipText     =   "Data final."
         Top             =   390
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
         Format          =   488570881
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "dias"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   14730
         TabIndex        =   27
         Top             =   510
         Width           =   285
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
         Left            =   682
         TabIndex        =   26
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "até:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   10725
         TabIndex        =   23
         Top             =   450
         Width           =   300
      End
      Begin VB.Label lbldata 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "01/0/2006"
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
         Left            =   9345
         TabIndex        =   22
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Período de:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   8490
         TabIndex        =   20
         Top             =   450
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   12840
         TabIndex        =   19
         Top             =   450
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   18
         Top             =   180
         Width           =   6255
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   24
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
      ButtonCaption2  =   "Ordens emitidas"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Lista de ordem emitidas por máquina (F7)"
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
      ButtonWidth2    =   85
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Relatório"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Relatório (F5)"
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
      ButtonLeft3     =   127
      ButtonTop3      =   2
      ButtonWidth3    =   51
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
      ButtonLeft4     =   180
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
      ButtonLeft5     =   184
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
      ButtonLeft6     =   222
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
      ButtonLeft7     =   250
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   25
      Top             =   8640
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
End
Attribute VB_Name = "frmCargaMaq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLISTA_Carga_Maq As ADODB.Recordset 'OK

Private Sub ProcListaOrdem()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
frmCargaMaq_ListaOrdem.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_Domingo_Click()
On Error GoTo tratar_erro

Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_Feriado_Click()
On Error GoTo tratar_erro

Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_HoraExtra_Click()
On Error GoTo tratar_erro

Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_Sabado_Click()
On Error GoTo tratar_erro

Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear
Select Case cmbfiltrarpor
    Case "Posto de trabalho": ProcCarregaComboPostoTrab cmbTexto, "Bloqueado = 'False'", True, True
    Case "Setor": ProcCarregaComboSetorPT cmbTexto, True
    Case "Grupo": ProcCarregaComboGrupoPT cmbTexto, True
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear

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
    Case vbKeyF7: ProcListaOrdem
    Case vbKeyReturn: Lista_DblClick
    Case vbKeyEscape: Unload Me
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
cmbfiltrarpor = "Setor"
lbldata.Caption = Format(Date, "dd/mm/yy")
msk_fltFim.Value = Date

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

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

ProcCarregaLista
Conexao.Execute "DELETE from CargaMaq_Total"
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from CargaMaq_total", Conexao, adOpenKeyset, adLockOptimistic
TBMaquinas.AddNew
TBMaquinas!Inicio = Date
TBMaquinas!Final = msk_fltFim.Value
TBMaquinas!TotalDisp = txtTempoDisp.Text
TBMaquinas!TotalUtil = IIf(Len(txtUtilizada) > 20, Left(txtUtilizada, 20), txtUtilizada)
TBMaquinas!Totalrest = IIf(Len(txtTempoRestante) > 20, Left(txtTempoRestante, 20), txtTempoRestante)
TBMaquinas!Totalultra = IIf(Len(txtUltrapassadas) > 20, Left(txtUltrapassadas, 20), txtUltrapassadas)
TBMaquinas.Update
TBMaquinas.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

TotalDispPeriodo = 0 'Tempo total disponivel por período
TotalUtilPeriodo = 0 'Tempo total utilizado por período
TotalRestMaquina = 0 'Tempo total restante ou ultrapassados por máquina
TotalRestPeriodo = 0 'Tempo total restante ou ultrapassados por período

Conexao.Execute "DELETE from CargaMaq_TotalMaq"
Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear

If cmbfiltrarpor = "Grupo" Then TextoFiltro = "Grupo" Else TextoFiltro = "Setor"
If cmbTexto = "" Then
    TextoFiltro1 = ""
Else
    If cmbfiltrarpor = "Posto de trabalho" Then
        TextoFiltro1 = "IDMaquina = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and "
    Else
        Select Case cmbfiltrarpor
            Case "Setor": TextoFiltro = "Setor"
            Case "Grupo": TextoFiltro = "Grupo"
        End Select
        TextoFiltro1 = TextoFiltro & " = '" & cmbTexto & "' and "
    End If
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from cadmaquinas where " & TextoFiltro1 & " Bloqueado = 'False' order by maquina", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from CargaMaq_TotalMaq", Conexao, adOpenKeyset, adLockOptimistic
            TBAbrir.AddNew
            
            .Add , , TBLISTA!maquina
            TBAbrir!maquina = TBLISTA!maquina
            .Item(.Count).SubItems(1) = TBLISTA!Descricao
            ProcVerificaTotalDisp
            ElapsedTime (TotalDisponivel)
            .Item(.Count).SubItems(2) = HoraTotal
            TBAbrir!TotalDisp = HoraTotal
            TotalDispPeriodo = TotalDispPeriodo + TotalDisponivel
            
            ProcVerificaTotalUtilizado
            TotalUtilizado = TEUSEG
            .Item(.Count).SubItems(3) = FormataTempo(TEUSEG)
            TBAbrir!TotalUtil = IIf(Len(.Item(.Count).SubItems(3)) > 20, Left(.Item(.Count).SubItems(3), 20), .Item(.Count).SubItems(3))
            TotalUtilPeriodo = TotalUtilPeriodo + TotalUtilizado
            
            ElapsedTime (TotalDisponivel)
            If TotalUtilizado > s Then
                TotalRestMaquina = TotalUtilizado - s
                .Item(.Count).SubItems(4) = "00:00:00"
                TBAbrir!Totalrest = "00:00:00"
                .Item(.Count).SubItems(5) = FormataTempo(TotalRestMaquina)
                TBAbrir!Totalultra = IIf(Len(.Item(.Count).SubItems(5)) > 20, Left(.Item(.Count).SubItems(5), 20), .Item(.Count).SubItems(5))
            Else
                TotalRestMaquina = s - TotalUtilizado
                .Item(.Count).SubItems(4) = FormataTempo(TotalRestMaquina)
                TBAbrir!Totalrest = IIf(Len(.Item(.Count).SubItems(4)) > 20, Left(.Item(.Count).SubItems(4), 20), .Item(.Count).SubItems(4))
                .Item(.Count).SubItems(5) = "00:00:00"
                TBAbrir!Totalultra = "00:00:00"
            End If
            TBAbrir.Update
            TBAbrir.Close
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close
ElapsedTime (TotalDispPeriodo)
txtTempoDisp.Text = HoraTotal
TEUSEG = TotalUtilPeriodo
txtUtilizada.Text = FormataTempo(TEUSEG)
ElapsedTime (TotalDispPeriodo)
If s > TotalUtilPeriodo Then
    TPUSEG = s - TotalUtilPeriodo
    txtTempoRestante.Text = FormataTempo(TPUSEG)
    txtUltrapassadas.Text = "00:00:00"
Else
    TPUSEG = TotalUtilPeriodo - s
    txtTempoRestante.Text = "00:00:00"
    txtUltrapassadas.Text = FormataTempo(TPUSEG)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaTotalDisp()
On Error GoTo tratar_erro

TotalDisponivel = 0 'Tempo total disponivel por maquina
Dias = 0
TempoFinal = msk_fltFim.Value
TempoInicio = Date
Do While TempoFinal >= TempoInicio
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select * from Feriados where Data_feriado = '" & Format(TempoInicio, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = True Or chk_Feriado.Value = 1 Then
        ProcVerificaDia
        ProcTotalHoraDia
    End If
    TBMaquinas.Close
    TempoInicio = TempoInicio + 1
    Dias = Dias + 1
Loop
txtdias = Dias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaTotalUtilizado()
On Error GoTo tratar_erro

TEUSEG = 0 'Tempo total utilizado por maquina
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select SUM(CASE WHEN OS.TTLPREVS - OS.TETTUTILSEG < 0 THEN 0 ELSE OS.TTLPREVS - OS.TETTUTILSEG END) AS TEUSEG from ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where P.Status <> 'Cancelada' and OS.maquina = '" & TBLISTA!maquina & "' AND OS.Prazofinal <= '" & Format(msk_fltFim.Value, "Short Date") & "' and OS.pronto = 'Não'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    TEUSEG = IIf(IsNull(TBproducao!TEUSEG), 0, TBproducao!TEUSEG)
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "Pcp_cargapostodetrabalho.rpt"
If cmbTexto = "" Then
    TextoFiltro1 = ""
Else
    If cmbfiltrarpor = "Posto de trabalho" Then
        TextoFiltro1 = "{CadMaquinas.IDMaquina} = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and "
    Else
        Select Case cmbfiltrarpor
            Case "Setor": TextoFiltro = "Setor"
            Case "Grupo": TextoFiltro = "Grupo"
        End Select
        TextoFiltro1 = "{CadMaquinas." & TextoFiltro & "} = '" & cmbTexto & "' and "
    End If
End If
ProcImprimirRel TextoFiltro1 & " {Producao.Status} <> 'Cancelada' and {ordemservico.pronto}= 'Não' and {ordemservico.Prazofinal}<= Date(" & Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcTotalHoraDia()
On Error GoTo tratar_erro

TextoFiltro = ""
Dataini = "00:00:00"
If chk_Sabado.Value = 0 And Diasemana = "Sabado" Then Exit Sub
If chk_Domingo.Value = 0 And Diasemana = "Domingo" Then Exit Sub
If chk_HoraExtra.Value = 0 Then TextoFiltro = " and (Percentual_HoraExtra IS NULL or Percentual_HoraExtra = 0)"

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select TotalTurno from cadmaqturnos where maquina ='" & TBLISTA!maquina & "' and diasemana = '" & Diasemana & "' " & TextoFiltro & " order by turno", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Do While TBproducao.EOF = False
        Dataini = Left(TBproducao!TotalTurno, 8)
        TotalDisponivel = TotalDisponivel + Dataini
        TBproducao.MoveNext
    Loop
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaDia()
On Error GoTo tratar_erro
                  
Diasemana = Weekday(TempoInicio)
Select Case Diasemana
    Case "1": Diasemana = "Domingo"
    Case "2": Diasemana = "Segunda"
    Case "3": Diasemana = "Terça"
    Case "4": Diasemana = "Quarta"
    Case "5": Diasemana = "Quinta"
    Case "6": Diasemana = "Sexta"
    Case "7": Diasemana = "Sabado"
End Select

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

If Lista.ListItems.Count = 0 Then Exit Sub
frmCargaMaq_ListaOrdem.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    If .ListItems.Count = 0 Then Exit Sub
    Caption = "PCP - Carga de posto de trabalho - (Cód. do posto : " & .SelectedItem & " - Descrição. : " & .SelectedItem.ListSubItems(1) & ")"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Caption = "PCP - Carga de posto de trabalho"
Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcListaOrdem
    Case 3: ProcImprimir
    'Case 5: ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
