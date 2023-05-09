VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmContas_receber_localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Financeiro - Contas a receber - Localizar"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8895
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk_bloqueada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bloqueada"
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
      Left            =   7440
      TabIndex        =   8
      Top             =   1568
      Width           =   1215
   End
   Begin VB.CheckBox Chk_devolucao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Incluir contas de devolução"
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
      Left            =   3210
      TabIndex        =   7
      Top             =   1568
      Value           =   1  'Checked
      Width           =   2805
   End
   Begin VB.CheckBox Chk_antecipacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Incluir contas de antecipação"
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
      Left            =   300
      TabIndex        =   6
      Top             =   1568
      Value           =   1  'Checked
      Width           =   2805
   End
   Begin VB.CheckBox Chk_transacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transação"
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
      Left            =   270
      TabIndex        =   15
      Top             =   4050
      Width           =   1185
   End
   Begin VB.CheckBox Optclassificacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo docto. :"
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
      Left            =   6390
      TabIndex        =   4
      Top             =   1125
      Width           =   1335
   End
   Begin VB.ComboBox cmbtipo_conta 
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
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmContas_Receber_Localizar.frx":0000
      Left            =   7770
      List            =   "frmContas_Receber_Localizar.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Tipo do documento."
      Top             =   1110
      Width           =   885
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
      ItemData        =   "frmContas_Receber_Localizar.frx":0004
      Left            =   1140
      List            =   "frmContas_Receber_Localizar.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   5115
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   60
      TabIndex        =   20
      Top             =   2250
      Width           =   8805
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   28
         Top             =   210
         Width           =   4785
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
            TabIndex        =   13
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
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
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
            Left            =   3930
            TabIndex        =   14
            Top             =   180
            Width           =   705
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
         ItemData        =   "frmContas_Receber_Localizar.frx":0008
         Left            =   180
         List            =   "frmContas_Receber_Localizar.frx":002A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3555
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
         Height          =   330
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmContas_Receber_Localizar.frx":00BD
         Left            =   180
         List            =   "frmContas_Receber_Localizar.frx":00BF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.Label Label3 
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
         Left            =   3645
         TabIndex        =   22
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label45 
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
         Left            =   1537
         TabIndex        =   21
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.CheckBox chkEmissao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissão"
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
      Left            =   1650
      TabIndex        =   16
      Top             =   4050
      Width           =   1005
   End
   Begin VB.CheckBox chkVencimento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vencimento"
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
      Left            =   2820
      TabIndex        =   17
      Top             =   4050
      Width           =   1485
   End
   Begin VB.CheckBox Optstatus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status :"
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
      Left            =   300
      TabIndex        =   9
      Top             =   1890
      Width           =   975
   End
   Begin VB.ComboBox cmbStatus 
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
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmContas_Receber_Localizar.frx":00C1
      Left            =   1320
      List            =   "frmContas_Receber_Localizar.frx":00C8
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Status."
      Top             =   1890
      Width           =   7365
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   60
      TabIndex        =   23
      Top             =   3780
      Width           =   8805
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   7320
         TabIndex        =   19
         ToolTipText     =   "Data final."
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
         Format          =   132448257
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   18
         ToolTipText     =   "Data inicio."
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
         Format          =   132448257
         CurrentDate     =   39057
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
         Left            =   6915
         TabIndex        =   25
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
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
         Left            =   5070
         TabIndex        =   24
         Top             =   240
         Width           =   300
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4590
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmContas_Receber_Localizar.frx":00EC
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   26
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
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
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ButtonHeight2   =   54
      ButtonUseMaskColor2=   0   'False
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
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
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
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      Left            =   240
      TabIndex        =   27
      Top             =   1110
      Width           =   825
   End
End
Attribute VB_Name = "frmContas_receber_localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_bloqueada_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_transacao_Click()
On Error GoTo tratar_erro

If Chk_transacao.Value = 1 Then
    chkEmissao.Value = 0
    chkVencimento.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_antecipacao_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkEmissao_Click()
On Error GoTo tratar_erro

If chkEmissao.Value = 1 Then
    Chk_transacao.Value = 0
    chkVencimento.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVencimento_Click()
On Error GoTo tratar_erro

If chkVencimento.Value = 1 Then
    Chk_transacao.Value = 0
    chkEmissao.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

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

If Chk_bloqueada.Value = 1 Then TextoFiltroBloq = "CR.Bloqueado = 'True'" Else TextoFiltroBloq = "CR.Bloqueado = 'False'"

If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Or cmbfiltrarpor = "Instituição" Or cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Local do desconto" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    
    With cmbTexto
        .Clear
        .AddItem ""
        Set TBLISTA = CreateObject("adodb.recordset")
        If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
            If Chk_antecipacao.Value = 1 Then TextoFiltro = "(CR.LogSit = 'N' or CR.Status = 'TÍTULO LIQUIDADO ANTECIPADO')" Else TextoFiltro = "CR.LogSit = 'N'"
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select F.int_codfamilia, F.Codigo, F.txt_descricao from (tbl_contas_receber CR INNER JOIN familia_financeiro FF ON CR.IdIntConta = FF.IDConta) INNER JOIN tbl_familia F ON F.int_codfamilia = FF.ID_PC where FF.tipoconta = 'R' and " & TextoFiltro & " and CR.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltroBloq & " Group by F.int_codfamilia, F.Codigo, F.txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    If cmbfiltrarpor = "Conta contábil" Then .AddItem TBLISTA!Txt_descricao & " - " & TBLISTA!CODIGO Else .AddItem TBLISTA!CODIGO & " - " & TBLISTA!Txt_descricao
                    .ItemData(cmbTexto.NewIndex) = TBLISTA!int_codfamilia
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        ElseIf cmbfiltrarpor = "Cliente" Then
                If Chk_antecipacao.Value = 1 Then TextoFiltro = "(LogSit = 'N' or Status = 'TÍTULO LIQUIDADO ANTECIPADO')" Else TextoFiltro = "LogSit = 'N'"
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select IDcliente, Nome_Razao from tbl_contas_receber CR where Nome_Razao is not null and " & TextoFiltro & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltroBloq & " Group by IDcliente, Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    Do While TBLISTA.EOF = False
                        .AddItem TBLISTA!Nome_Razao
                        .ItemData(cmbTexto.NewIndex) = TBLISTA!IDCliente
                        TBLISTA.MoveNext
                    Loop
                End If
                TBLISTA.Close
            ElseIf cmbfiltrarpor = "Instituição" Then
                    ProcCarregaComboBancoFinanceiro cmbTexto, "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Txt_descricao is not null", True
                ElseIf cmbfiltrarpor = "Local do desconto" Then
                        Set TBLISTA = CreateObject("adodb.recordset")
                        TBLISTA.Open "Select local_troca from troca_titulo where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and local_troca is not null Group by local_troca", Conexao, adOpenKeyset, adLockOptimistic
                        If TBLISTA.EOF = False Then
                            Do While TBLISTA.EOF = False
                                If TBLISTA!local_troca <> "" Then .AddItem TBLISTA!local_troca
                                TBLISTA.MoveNext
                            Loop
                        End If
                        TBLISTA.Close
                    ElseIf cmbfiltrarpor = "Status" Then
                            .AddItem "TÍTULO EM ABERTO"
                            If Chk_antecipacao.Value = 1 Then .AddItem "TÍTULO LIQUIDADO ANTECIPADO" Else .AddItem "TÍTULO RECEBIDO PARCIAL"
                            .AddItem "DUPLICATA DESCONTADA EM ABERTO"
        End If
    End With
Else
    txtTexto.Visible = True
    cmbTexto.Visible = False
    
    If txtTexto <> "" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
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

TextoFiltroClass = ""
TextoFiltroClassRel = ""
If Optclassificacao.Value = 1 Then
    TextoFiltroClass = " and (CR.Tipo_doc IS NULL or CR.Tipo_doc = '" & cmbtipo_conta & "')"
    TextoFiltroClassRel = " and {tbl_Contas_receber.Tipo_doc}= '" & cmbtipo_conta & "'"
    'and isnull({tbl_contas_receber.Tipo_doc})= true or {tbl_contas_receber.Tipo_doc} = ''
End If

If Chk_bloqueada.Value = 1 Then
    TextoFiltroBloq = " and CR.Bloqueado = 'True'"
    TextoFiltroBloqRel = " and {tbl_Contas_receber.Bloqueado} = True"
Else
    TextoFiltroBloq = " and CR.Bloqueado = 'False'"
    TextoFiltroBloqRel = " and {tbl_Contas_receber.Bloqueado} = False"
End If

If Chk_antecipacao.Value = 1 Then
    TextoFiltroAntecipacao = " and (CR.LogSit = 'N' or CR.Status = 'TÍTULO LIQUIDADO ANTECIPADO' and CR.Saldo_antecipacao > 0)"
    TextoFiltroAntecipacaoRel = " and ({tbl_Contas_receber.LogSit} = 'N' or {tbl_Contas_receber.Status} = 'TÍTULO LIQUIDADO ANTECIPADO' and {tbl_Contas_receber.Saldo_antecipacao} > 0)"
    TextoFiltroAntecipacao1 = " and CR.Antecipacao = 'True' and CR.Status = 'TÍTULO LIQUIDADO ANTECIPADO' and CR.Saldo_antecipacao > 0"
Else
    TextoFiltroAntecipacao = " and CR.Antecipacao = 'False' and CR.LogSit = 'N'"
    TextoFiltroAntecipacaoRel = " and {tbl_Contas_receber.Antecipacao} = False and {tbl_Contas_receber.LogSit} = 'N'"
    TextoFiltroAntecipacao1 = ""
End If

If Chk_devolucao.Value = 1 Then
    TextoFiltroDevolucao = " and (CR.LogSit = 'N' or CR.Devolucao = 'True')"
    TextoFiltroDevolucaoRel = " and ({tbl_Contas_receber.LogSit} = 'N' or {tbl_Contas_receber.Devolucao} = True)"
    TextoFiltroDevolucao1 = " and CR.Devolucao = 'True' and CR.LogSit = 'N'"
Else
    TextoFiltroDevolucao = " and CR.Devolucao = 'False' and CR.LogSit = 'N'"
    TextoFiltroDevolucaoRel = " and {tbl_Contas_receber.Devolucao} = False and {tbl_Contas_receber.LogSit} = 'N'"
    TextoFiltroDevolucao1 = ""
End If

TextoFiltroStatus = ""
TextoFiltroStatusRel = ""
If Optstatus.Value = 1 Then
    TextoFiltroStatus = " and CR.status = '" & cmbStatus & "'"
    TextoFiltroStatusRel = " and {tbl_Contas_receber.status} = '" & cmbStatus & "'"
End If

DataFiltro = ""
DataFiltroRel = ""
Data_receber = "CR.vencimento"
If Chk_transacao.Value = 1 Or chkVencimento.Value = 1 Or chkEmissao.Value = 1 Then
    If Chk_transacao.Value = 1 Then
        Data_receber = "CR.Data_transacao"
    ElseIf chkEmissao.Value = 1 Then
            Data_receber = "CR.Emissao"
        Else
            Data_receber = "CR.vencimento"
    End If
    DataFiltro = " and " & Data_receber & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    DataFiltroRel = " and {" & Replace(Data_receber, "CR.", "tbl_Contas_receber.") & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & Replace(Data_receber, "CR.", "tbl_Contas_receber.") & "} <= Date(" & _
                            Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
End If

NomeRel = "Contas_receber.rpt"

Imprimir = True
With frmContas_Receber
    .StrSql_Contas_Receber_AntecTotal = ""
    .StrSql_Contas_Receber_DevTotal = ""
    
    CamposFiltro = "CR.IDintconta, CR.emissao, CR.Vencimento, CR.Data_transacao, CR.Valor, CR.txt_ndocumento, CR.NFiscal, CR.Parcela, CR.Nome_Razao, CR.Responsavel, CR.ID_empresa, CR.IDduplicata, CR.Saldo_antecipacao, CR.Antecipacao"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from tbl_contas_receber CR"
    INNERJOINTEXTOSUM = "Select SUM(CR.Valor) AS TotContas from tbl_contas_receber CR"
    INNERJOINTEXTOSUMANTEC = "Select SUM(CR.Saldo_antecipacao) AS TotContas1 from tbl_contas_receber CR"
    OrdenarTexto = " group by " & CamposFiltro & " order by " & Data_receber & ", CR.IdIntConta"
    TextoFiltroPadrao1 = "CR.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & TextoFiltroClass & TextoFiltroBloq & DataFiltro & TextoFiltroStatus
    TextoFiltroPadrao = TextoFiltroPadrao1 & TextoFiltroAntecipacao & TextoFiltroDevolucao
    TextoFiltroPadraoSOMA = TextoFiltroPadrao1 & " and CR.Logsit = 'N' and CR.Antecipacao = 'False' and CR.Devolucao = 'False' and CR.status <> 'DUPLICATA DESCONTADA EM ABERTO'"
    TextoFiltroPadraoANTEC = TextoFiltroPadrao1 & TextoFiltroAntecipacao1 & " and CR.status <> 'DUPLICATA DESCONTADA EM ABERTO'"
    TextoFiltroPadraoDEV = TextoFiltroPadrao1 & TextoFiltroDevolucao1 & " and CR.status <> 'DUPLICATA DESCONTADA EM ABERTO'"
    TextoFiltroPadraoDESC = TextoFiltroPadrao & " and CR.status = 'DUPLICATA DESCONTADA EM ABERTO'"
    TextoFiltroPadraoRel = "{tbl_contas_receber.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & TextoFiltroClassRel & TextoFiltroBloqRel & DataFiltroRel & TextoFiltroStatusRel & TextoFiltroAntecipacaoRel & TextoFiltroDevolucaoRel
    
    If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
        If cmbTexto.Visible = True Then
            If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
                NomeRel = "Contas_receber_conta contabil.rpt"
                
                INNERJOINPADRAO = " from tbl_contas_receber CR INNER JOIN familia_financeiro FF ON CR.IdIntConta = FF.idconta"
                INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
                INNERJOINTEXTOSUM = "Select SUM(CR.valor) AS TotContas " & INNERJOINPADRAO
                INNERJOINTEXTOSUMANTEC = "Select SUM(CR.Saldo_antecipacao) AS TotContas1 " & INNERJOINPADRAO
                TextoFiltro = "FF.ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and FF.tipoconta = 'R'"
                .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
                .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoSOMA
                If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " and " & TextoFiltroPadraoANTEC
                If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDEV
                .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDESC
                .FormulaRel_Contas_Receber = "{familia_financeiro.ID_PC} = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and {familia_financeiro.tipoconta} = 'R' and " & TextoFiltroPadraoRel
            ElseIf cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Local do desconto" Then
                    If cmbTexto = "DUPLICATA DESCONTADA EM ABERTO" Or cmbfiltrarpor = "Local do desconto" Then
                        If cmbTexto = "DUPLICATA DESCONTADA EM ABERTO" Then
                            NomeRel = "Contas_receber_descontada_local_desconto.rpt"
                            TextoFiltro = "CR.status = '" & cmbTexto & "'"
                            TextoFiltroRel = "tbl_contas_receber.status"
                        Else
                            TextoFiltro = "TT.local_troca = '" & cmbTexto & "'"
                            TextoFiltroRel = "troca_titulo.local_troca"
                        End If
                        
                        INNERJOINPADRAO = " from tbl_contas_receber CR INNER JOIN troca_titulo TT ON TT.id = CR.Idtrocatitulo"
                        INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
                        INNERJOINTEXTOSUM = "Select SUM(CR.valor) AS TotContas " & INNERJOINPADRAO
                        INNERJOINTEXTOSUMANTEC = "Select SUM(CR.Saldo_antecipacao) AS TotContas1 " & INNERJOINPADRAO
                        
                        .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
                        .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoSOMA
                        If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " and " & TextoFiltroPadraoANTEC
                        If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDEV
                        .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDESC
                        .FormulaRel_Contas_Receber = "{" & TextoFiltroRel & "} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
                    Else
                        TextoFiltro = "CR.status = '" & cmbTexto & "'"
                        .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
                        .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoSOMA
                        If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & "and " & TextoFiltroPadraoANTEC
                        If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDEV
                        .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDESC
                        .FormulaRel_Contas_Receber = "{tbl_Contas_receber.status} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
                    End If
                Else
                    If cmbfiltrarpor = "Instituição" Then TextoFiltro = "CR.Banco" Else TextoFiltro = "CR.Nome_Razao"
                    .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadrao & OrdenarTexto
                    .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoSOMA
                    If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoANTEC
                    If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoDEV
                    .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoDESC
                    .FormulaRel_Contas_Receber = "{" & Replace(TextoFiltro, "CR.", "tbl_contas_receber.") & "} = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoRel
            End If
        Else
            If cmbfiltrarpor = "Pedido interno" Then
                INNERJOINPADRAO = " from tbl_contas_receber CR LEFT JOIN tbl_proposta_nota PN ON PN.Id_nota = CR.Id_nota"
                INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
                INNERJOINTEXTOSUM = "Select SUM(CR.valor) AS TotContas " & INNERJOINPADRAO
                INNERJOINTEXTOSUMANTEC = "Select SUM(CR.Saldo_antecipacao) AS TotContas1 " & INNERJOINPADRAO
                TextoFiltro = "PN.proposta " & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
                .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
                .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoSOMA
                If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " and " & TextoFiltroPadraoANTEC
                If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDEV
                .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDESC
                .FormulaRel_Contas_Receber = "{tbl_proposta_nota.proposta}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " And " & TextoFiltroPadraoRel
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open .StrSql_Contas_Receber, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = True Then
                    TextoFiltro = "CR.proposta"
                    .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao & OrdenarTexto
                    .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoSOMA
                    If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoANTEC
                    If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoDEV
                    .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoDESC
                    .FormulaRel_Contas_Receber = "{" & Replace(TextoFiltro, "CR.", "tbl_contas_receber.") & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
                End If
            ElseIf cmbfiltrarpor = "Nosso número" Then
                    INNERJOINPADRAO = " from tbl_contas_receber CR INNER JOIN tbl_Detalhes_Recebimento_Nboletos DRB ON DRB.IDContaReceber = CR.IdIntConta"
                    INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
                    INNERJOINTEXTOSUM = "Select SUM(CR.valor) AS TotContas " & INNERJOINPADRAO
                    INNERJOINTEXTOSUMANTEC = "Select SUM(CR.Saldo_antecipacao) AS TotContas1 " & INNERJOINPADRAO
                    TextoFiltro = "DRB.Nosso_numero " & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
                    .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
                    .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoSOMA
                    If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " and " & TextoFiltroPadraoANTEC
                    If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDEV
                    .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDESC
                    .FormulaRel_Contas_Receber = "{tbl_Detalhes_Recebimento_Nboletos.Nosso_numero}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
                    Imprimir = False
                ElseIf cmbfiltrarpor = "Valor" Then
                        valor = txtTexto
                        NovoValor = Replace(valor, ",", ".")
                        TextoFiltro = "CR.Valor"
                        .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao & OrdenarTexto
                        .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadraoSOMA
                        If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadraoANTEC
                        If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadraoDEV
                        .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadraoDESC
                        .FormulaRel_Contas_Receber = "{" & Replace(TextoFiltro, "CR.", "tbl_contas_receber.") & "} = " & NovoValor & " and " & TextoFiltroPadraoRel
                    Else
                        If cmbfiltrarpor = "Nota fiscal" Then
                            If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
                            TextoFiltro = "CR.Nfiscal"
                        Else
                            TextoFiltro = "CR.IDduplicata"
                        End If
                        .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao & OrdenarTexto
                        .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoSOMA
                        If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoANTEC
                        If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoDEV
                        .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoDESC
                        .FormulaRel_Contas_Receber = "{" & Replace(TextoFiltro, "CR.", "tbl_contas_receber.") & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
            End If
        End If
    Else
        .StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltroPadrao & OrdenarTexto
        .StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltroPadraoSOMA
        If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Receber_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltroPadraoANTEC
        If Chk_devolucao.Value = 1 Then .StrSql_Contas_Receber_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltroPadraoDEV
        .StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltroPadraoDESC
        .FormulaRel_Contas_Receber = TextoFiltroPadraoRel
    End If
    .ProcSalvarDadosRel Chk_transacao, chkEmissao, chkVencimento, msk_fltInicio, msk_fltFim
    
    .ProcCarregaLista (1)
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8805, 5, True

ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Cliente"
msk_fltFim.Value = Date
msk_fltInicio.Value = Date

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

Private Sub Optclassificacao_Click()
On Error GoTo tratar_erro

With cmbtipo_conta
    If Optclassificacao.Value = 1 Then
        ProcCarregaComboTipoDocto cmbtipo_conta, "Tipo = 'R'"
        .Enabled = True
        .SetFocus
    Else
        .Enabled = False
        .ListIndex = -1
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optstatus_Click()
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    If Optstatus.Value = 1 Then
        cmbStatus.Enabled = True
        cmbStatus.SetFocus
        .AddItem "Local do desconto"
        .Text = "Local do desconto"
    Else
        cmbStatus.Enabled = False
        cmbStatus.ListIndex = -1
        .AddItem "Nota fiscal"
        .AddItem "Pedido interno"
        .AddItem "Cliente"
        .AddItem "Conta contábil"
        .AddItem "Código da conta contábil"
        .AddItem "Instituição"
        .AddItem "Status"
        .AddItem "Número da duplicata"
        .Text = "Cliente"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" And (cmbfiltrarpor = "Nota fiscal" Or cmbfiltrarpor = "Valor") Then
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

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    Select Case cmbfiltrarpor
        Case "Nota fiscal": txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
        Case "Valor": txtTexto = Format(txtTexto, "###,##0.00")
    End Select
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
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
