VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProd_programacao_Ordem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Programação da produção - Localizar ordem"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmProd_programacao_Ordem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmProd_programacao_Ordem.frx":1042
   MousePointer    =   99  'Custom
   ScaleHeight     =   8640
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.CheckBox chkFabricacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Componente"
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
      Left            =   1965
      TabIndex        =   19
      Top             =   975
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkMontagem 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subconjunto"
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
      Left            =   225
      TabIndex        =   18
      Top             =   975
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Chk_saldo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Saldo diferente de zero"
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
      Left            =   6675
      TabIndex        =   17
      Top             =   975
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   11
      Top             =   885
      Width           =   13935
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4620
         TabIndex        =   15
         Top             =   210
         Width           =   3975
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
            MousePointer    =   99  'Custom
            TabIndex        =   5
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
            MousePointer    =   99  'Custom
            TabIndex        =   3
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
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   180
            Width           =   1275
         End
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
         ItemData        =   "frmProd_programacao_Ordem.frx":134C
         Left            =   180
         List            =   "frmProd_programacao_Ordem.frx":136B
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4365
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
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "``"
         Top             =   1050
         Width           =   13515
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
         Left            =   180
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Status."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   3645
         TabIndex        =   13
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   1935
         TabIndex        =   12
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   55
      TabIndex        =   10
      Top             =   2385
      Width           =   13935
      Begin VB.CheckBox ChkData 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dt. emissão"
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
         Left            =   420
         TabIndex        =   6
         Top             =   240
         Width           =   1245
      End
      Begin VB.CheckBox ChkData1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Prazo entrega"
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
         Left            =   1935
         TabIndex        =   7
         Top             =   240
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   7200
         TabIndex        =   9
         ToolTipText     =   "Data final para pesquisa."
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   104333313
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   8
         ToolTipText     =   "Data início para pesquisa."
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   104333315
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "à"
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
         Left            =   6945
         TabIndex        =   14
         Top             =   240
         Width           =   90
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5325
      Left            =   60
      TabIndex        =   16
      Top             =   2955
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   9393
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmProd_programacao_Ordem.frx":13C8
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Qtde. faturada"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "D"
         Text            =   "Prazo final"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Cod. de ref."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   2743
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Observação"
         Object.Width           =   2646
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   885
      Left            =   90
      TabIndex        =   20
      Top             =   0
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   1561
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
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
      ButtonKey1      =   "2"
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
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   48
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "14"
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
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "15"
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
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "16"
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
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5850
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProd_programacao_Ordem.frx":16E2
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   8340
      Width           =   13935
      _ExtentX        =   24580
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
Attribute VB_Name = "frmProd_programacao_Ordem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pesquisa_ordem As String 'ok
Dim StrSql_Ordem_programacao_LocalizarOrdem As String

Private Sub ChkData_Click()
On Error GoTo tratar_erro

If ChkData.Value = 1 Then
    ChkData1.Value = 0
    msk_fltInicio.Enabled = True
    msk_fltFim.Enabled = True
Else
    If ChkData1.Value = 0 Then
        msk_fltInicio.Value = Date
        msk_fltInicio.Enabled = False
        msk_fltFim.Value = Date
        msk_fltFim.Enabled = False
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChkData1_Click()
On Error GoTo tratar_erro

If ChkData1.Value = 1 Then
    ChkData.Value = 0
    msk_fltInicio.Enabled = True
    msk_fltFim.Enabled = True
Else
    If ChkData.Value = 0 Then
        msk_fltInicio.Value = Date
        msk_fltInicio.Enabled = False
        msk_fltFim.Value = Date
        msk_fltFim.Enabled = False
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Tipo" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'true'", False
        ElseIf cmbfiltrarpor = "Status" Then
                .AddItem "Aberta"
                .AddItem "Produzindo"
                .AddItem "Concluída"
                .AddItem "Cancelada"
                .AddItem "Aguardando"
                .AddItem "Entregue"
            Else
                .AddItem "Componente"
                .AddItem "Subconjunto"
                .AddItem "Produto final"
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: Ajuda
    'Case vbKeyF2: cmdFiltrar_Click
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcVerifColunas
cmbfiltrarpor = "Ordem"
txtTexto.Visible = True
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

ProcCarregaToolBar1 Me, 15195, 4, True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifColunas()
On Error GoTo tratar_erro

ProcCorrigeColunasForm Lista, "PCP/Programação da produção/Localizar ordem", 13, True, 800, 500, 800, 1100, 800, 950, 900, 1200, 1200, 1555, 1500, 800, 800, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

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
If chkFabricacao = 1 Or chkMontagem = 1 Or chkExpedicao = 1 Then
    If chkFabricacao = 1 Then Tipo = "and Tipo = 'F'"
    If chkMontagem = 1 Then Tipo = "and Tipo = 'M'"
    If chkExpedicao = 1 Then Tipo = "and Tipo = 'E'"
End If

If ChkData.Value = 0 And ChkData1.Value = 0 Then
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Código interno" Then
            If Optinicio.Value = True Then StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where desenho like '" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " " & Tipo & " order by Ordem desc"
            If Optmeio.Value = True Then StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where desenho like '%" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where desenho like '%" & txtTexto.Text & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
            End If
        End If
        If cmbfiltrarpor = "Código referencia" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where N_Referencia like '" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where N_Referencia like '%" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where N_Referencia like '%" & txtTexto.Text & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
            End If
        End If
        If cmbfiltrarpor = "Descrição" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where produto like '" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where produto like '%" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where produto like '%" & txtTexto.Text & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "Cliente" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where cliente like '" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where cliente like '%" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where cliente like '%" & txtTexto.Text & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "Ordem" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Ordem like '" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Ordem like '%" & txtTexto.Text & "%' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Ordem like '%" & txtTexto.Text & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "OS" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select producao.* FROM producao INNER JOIN Ordemservico ON producao.Ordem = Ordemservico.Ordem where Ordemservico.IDProducao = " & txtTexto.Text & " and producao.concluida = 'false' " & Tipo & " order by producao.Ordem desc"
            
        End If
        If cmbfiltrarpor = "Status" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Status = '" & cmbfamilia & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
        
        End If
        If cmbfiltrarpor = "Família" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select producao.* FROM producao INNER JOIN projproduto ON producao.Desenho = projproduto.Desenho where projproduto.classe = '" & cmbfamilia & "' and producao.concluida = 'false' " & Tipo & " order by producao.Ordem desc"
        
        End If
        If cmbfiltrarpor = "Tipo" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Tipo = '" & Tipo & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
        
        End If
    Else
        StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where concluida = 'false' " & Tipo & " order by Ordem desc"
    End If
Else
    If ChkData.Value = 1 Then Pesquisa_ordem = "data"
    If ChkData1.Value = 1 Then Pesquisa_ordem = "PrazoEntrega"
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Código interno" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where desenho like '" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where desenho like '%" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where desenho like '%" & txtTexto & "' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "Código referencia" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where N_Referencia like '" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where N_Referencia like '%" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where N_Referencia like '%" & txtTexto & "' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "Descrição" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where produto like '" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where produto like '%" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where produto like '%" & txtTexto & "' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "Cliente" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where cliente like '" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where cliente like '%" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where cliente like '%" & txtTexto & "' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "Ordem" Then
            If Optinicio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Ordem like '" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optmeio.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Ordem like '%" & txtTexto & "%' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
            If Optfim.Value = True Then
                StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Ordem like '%" & txtTexto & "' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
                
            End If
        End If
        If cmbfiltrarpor = "OS" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select producao.* FROM producao INNER JOIN Ordemservico ON producao.Ordem = Ordemservico.Ordem where Ordemservico.IDProducao = " & txtTexto.Text & " and (Producao." & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and producao.concluida = 'false' " & Tipo & " order by producao.Ordem desc"
            
        End If
        If cmbfiltrarpor = "Status" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Status = '" & cmbfamilia & "' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
            
        End If
        If cmbfiltrarpor = "Família" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select producao.* FROM producao INNER JOIN projproduto ON producao.Desenho = projproduto.Desenho where projproduto.classe = '" & cmbfamilia & "' and (producao." & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and producao.concluida = 'false' " & Tipo & " order by producao.Ordem desc"
            
        End If
        If cmbfiltrarpor = "Tipo" Then
            StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where Tipo = '" & Tipo & "' and (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'false' " & Tipo & " order by Ordem desc"
            
        End If
    Else
        StrSql_Ordem_programacao_LocalizarOrdem = "Select * from producao where (" & Pesquisa_ordem & ") Between '" & Format(msk_fltInicio.Value, "dd/mm/YYYY") & "' And '" & Format(msk_fltFim.Value, "dd/mm/yyyy") & "' and concluida = 'true' order by Ordem desc"
    End If
End If
ProcCarregaLista

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

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
With frmProd_programacao
    .txtOrdem = Lista.SelectedItem
    .ProcCarregaOrdem
End With
Unload Me
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro
Dim Condicao2 As String 'OK

Lista.ListItems.Clear
If StrSql_Ordem_programacao_LocalizarOrdem = "" Then Exit Sub
Condicao2 = "NÃO"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Ordem_programacao_LocalizarOrdem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        Permitido = True
'        If IsNull(TBLISTA!Qtde_NF) = False And TBLISTA!Qtde_NF <> "" And TBLISTA!Qtde_NF <> "0" Then
'            Qtde = IIf(IsNull(TBLISTA!Qtde_NF), 0, TBLISTA!Qtde_NF)
'        Else
'            Qtde = IIf(IsNull(TBLISTA!Quant), 0, TBLISTA!Quant)
'        End If
        'Cálcula saldo
'        Qtd = 0
'        Set TBAbrir = CreateObject("adodb.recordset")
'        TBAbrir.Open "select * from producao_nf where Ordem = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
'        Do While TBAbrir.EOF = False
'            Qtd = Qtd + IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde) + IIf(IsNull(TBAbrir!QtdeNC), 0, TBAbrir!QtdeNC)
'            TBAbrir.MoveNext
'        Loop
'        TBAbrir.Close
        If Chk_saldo.Value = 1 Then
            If Qtde - Qtd = 0 Then Permitido = False
        End If
        If Permitido = True Then
            With Lista.ListItems
                .Add , , TBLISTA!Ordem
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
'                If IsNull(TBLISTA!Qtde_NF) = False And TBLISTA!Qtde_NF <> "" And TBLISTA!Qtde_NF <> "0" Then
'                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Qtde_NF), "", Format(TBLISTA!Qtde_NF, "###,##0.0000"))
'                    Qtde = IIf(IsNull(TBLISTA!Qtde_NF), 0, TBLISTA!Qtde_NF)
'                Else
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Quant), "", Format(TBLISTA!Quant, "###,##0.0000"))
                    Qtde = IIf(IsNull(TBLISTA!Quant), 0, TBLISTA!Quant)
'                End If
                .Item(.Count).SubItems(3) = Format(Qtd, "###,##0.0000")
                .Item(.Count).SubItems(4) = Format(Qtde - Qtd, "###,##0.0000")
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!PrazoEntrega), "", Format(TBLISTA!PrazoEntrega, "dd/mm/yy"))
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Produto), "", TBLISTA!Produto)
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
                .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!status), "", TBLISTA!status)
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Obs), "", Trim(TBLISTA!Obs))
            End With
        End If
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

