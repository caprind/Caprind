VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmcompras_fornecedores_localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Compras - Fornecedores - Localizar"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10575
   ClipControls    =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAprovada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fonte aprovada pelo cliente"
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
      Left            =   2850
      TabIndex        =   14
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CheckBox Chk_prospecto 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prospecto"
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
      Left            =   210
      TabIndex        =   17
      Top             =   1080
      Width           =   1155
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   6450
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmcompras_fornecedores_localizar.frx":0000
      Count           =   1
   End
   Begin VB.CheckBox chkAvaliado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Avaliado"
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
      Left            =   1620
      TabIndex        =   13
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CheckBox chkCertificado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado"
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
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
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
      Height          =   1575
      Left            =   1470
      TabIndex        =   18
      Top             =   1290
      Width           =   9045
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4080
         TabIndex        =   26
         Top             =   210
         Width           =   4785
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
            TabIndex        =   6
            Top             =   180
            Width           =   705
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
            TabIndex        =   4
            Top             =   180
            Width           =   1275
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
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
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
            TabIndex        =   5
            Top             =   180
            Width           =   1155
         End
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
         TabIndex        =   7
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1110
         Width           =   8685
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   10
         ToolTipText     =   "Número do CPF."
         Top             =   1110
         Visible         =   0   'False
         Width           =   8685
         _ExtentX        =   15319
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
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
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
         ItemData        =   "frmcompras_fornecedores_localizar.frx":21ED
         Left            =   180
         List            =   "frmcompras_fornecedores_localizar.frx":220F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3825
      End
      Begin MSMask.MaskEdBox txtcnpj 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         ToolTipText     =   "Número do CNPJ."
         Top             =   1110
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         MouseIcon       =   "frmcompras_fornecedores_localizar.frx":2297
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
      Begin VB.ComboBox cmbstatus 
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
         ItemData        =   "frmcompras_fornecedores_localizar.frx":25B1
         Left            =   180
         List            =   "frmcompras_fornecedores_localizar.frx":25BE
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Status."
         Top             =   1110
         Width           =   8685
      End
      Begin VB.ComboBox cmbFamilia 
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Família."
         Top             =   1110
         Width           =   8685
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3787
         TabIndex        =   21
         Top             =   900
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
         Left            =   1672
         TabIndex        =   20
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1575
      Left            =   55
      TabIndex        =   19
      Top             =   1290
      Width           =   1305
      Begin VB.OptionButton optFisica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Física"
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
         TabIndex        =   1
         Top             =   990
         Width           =   855
      End
      Begin VB.OptionButton optJuridica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jurídica"
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
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   60
      TabIndex        =   22
      Top             =   2850
      Width           =   10455
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   8970
         TabIndex        =   16
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
         Format          =   489816065
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   7050
         TabIndex        =   15
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
         Format          =   489816065
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8565
         TabIndex        =   24
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por período :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6060
         TabIndex        =   23
         Top             =   210
         Width           =   930
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   25
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
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
      ButtonLeft2     =   46
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   50
      ButtonTop3      =   2
      ButtonWidth3    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   93
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
End
Attribute VB_Name = "frmcompras_fornecedores_localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAprovada_Click()
On Error GoTo tratar_erro

Frame5.Enabled = False
If chkAprovada.Value = 1 Then
    chkAvaliado.Value = 0
    chkCertificado.Value = 0
    Frame5.Enabled = True
    msk_fltInicio.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkAvaliado_Click()
On Error GoTo tratar_erro

Frame5.Enabled = False
If chkAvaliado.Value = 1 Then
    chkCertificado.Value = 0
    chkAprovada.Value = 0
    Frame5.Enabled = True
    msk_fltInicio.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkcertificado_Click()
On Error GoTo tratar_erro

Frame5.Enabled = False
If chkCertificado.Value = 1 Then
    chkAvaliado.Value = 0
    chkAprovada.Value = 0
    Frame5.Enabled = True
    msk_fltInicio.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia.Text <> "" Then
    txtTexto.Text = ""
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Razão social" Or cmbfiltrarpor = "Nome fantasia" Or cmbfiltrarpor = "Cidade" Or cmbfiltrarpor = "Código do fornecedor" Or cmbfiltrarpor = "Centro de custo" Then
    txtTexto.Visible = True
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Segmento" Or cmbfiltrarpor = "Centro de custo" Or cmbfiltrarpor = "Regime tributário" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
    ElseIf cmbfiltrarpor = "Segmento" Then
            ProcCarregaComboSegmento
        ElseIf cmbfiltrarpor = "Regime tributário" Then
                ProcCarregaComboRegime
            Else
                ProcCarregaComboSetor cmbfamilia, "US.Setor IS NOT NULL", "", False, True, False, "", False, False
    End If
End If
If cmbfiltrarpor = "Status" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = True
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optJuridica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = True
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optFisica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbstatus_Click()
On Error GoTo tratar_erro

If cmbStatus.Text <> "" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With frmCompras_fornecedores
    If optFisica.Value = True Then
        TipoPessoa = "CF.pessoa = 'FÍSICA'"
        TipoPessoaRel = "{Compras_fornecedores.pessoa} = 'FÍSICA'"
        CPFCNPJ = "cpf_cnpj = '" & txtCpf.Text & "'"
        CPFCNPJRel = "{Compras_fornecedores.cpf_cnpj} = '" & txtCpf.Text & "'"
    Else
        TipoPessoa = "CF.pessoa = 'JURÍDICA'"
        TipoPessoaRel = "{Compras_fornecedores.pessoa} = 'JURÍDICA'"
        CPFCNPJ = "cpf_cnpj = '" & txtcnpj.Text & "'"
        CPFCNPJRel = "{Compras_fornecedores.cpf_cnpj} = '" & txtcnpj.Text & "'"
    End If
    If chkCertificado.Value = 1 Then
        TipoFornecedor = "CF.Fornecedor = 'C'"
        TipoFornecedorRel = "{Compras_fornecedores.Fornecedor} = 'C'"
    ElseIf chkAvaliado.Value = 1 Then
            TipoFornecedor = "CF.Fornecedor = 'A'"
            TipoFornecedorRel = "{Compras_fornecedores.Fornecedor} = 'A'"
        ElseIf chkAprovada.Value = 1 Then
                TipoFornecedor = "CF.Fornecedor = 'F'"
                TipoFornecedorRel = "{Compras_fornecedores.Fornecedor} = 'F'"
            Else
                TipoFornecedor = "CF.nome_razao <> 'Null'"
                TipoFornecedorRel = "{Compras_fornecedores.nome_razao} <> 'Null'"
    End If
    If Chk_prospecto.Value = 1 Then
        Prospecto = "CF.Prospecto = 'True'"
        ProspectoRel = "{Compras_fornecedores.Prospecto} = True"
    Else
        Prospecto = "CF.Prospecto = 'False'"
        ProspectoRel = "{Compras_fornecedores.Prospecto} = False"
    End If
    
    If cmbfiltrarpor = "Regime tributário" Then
        If cmbfamilia = "Lucro presumido" Then
            TextoRegime = "CF.Presumido = 'True'"
            TextoRegimeRel = "{Compras_fornecedores.Presumido} = True"
        ElseIf cmbfamilia = "CF.Simples nacional" Then
                TextoRegime = "Simples = 'True'"
                TextoRegimeRel = "{Compras_fornecedores.Simples} = True"
            ElseIf cmbfamilia = "Lucro real" Then
                    TextoRegime = "CF.Real = 'True'"
                    TextoRegimeRel = "{Compras_fornecedores.Real} = True"
                Else
                    TextoRegime = "MEI = 'True'"
                    TextoRegimeRel = "{Compras_fornecedores.MEI} = True"
        End If
    End If

    CamposFiltro = "CF.IDCliente, CF.DTCadastro, CF.Responsavel, CF.Nome_Razao, CF.DtValidacao, CF.ID"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from ((Compras_fornecedores CF LEFT JOIN compras_fornecedores_familia CFF ON CF.IDCliente = CFF.IDCliente) LEFT JOIN compras_fornecedores_segmentos CFS ON CF.IDCliente = CFS.IDFornecedor) LEFT JOIN Usuarios_Setor US ON US.ID = CF.ID_CC"
    TextoFiltroPadrao = TipoPessoa & " and " & TipoFornecedor & " and " & Prospecto & " group by " & CamposFiltro & " order by CF.nome_razao"
    TextoFiltroPadraoRel = TipoPessoaRel & " and " & TipoFornecedorRel & " and " & ProspectoRel & IIf(cmbfiltrarpor <> "Status", " and {Compras_fornecedores.status} = 'Liberado'", "")
    
    If txtTexto <> "" Or cmbfamilia <> "" Or cmbStatus <> "" Or txtcnpj <> "__.___.___/____-__" Or txtCpf <> "___.___.___-__" Then
        If cmbfiltrarpor = "Status" Then
            .StrSql_Fornecedor = INNERJOINTEXTO & " where CF.status = '" & cmbStatus.Text & "' and " & TextoFiltroPadrao
            .FormulaRel_Fornecedor = "{Compras_fornecedores.status} = '" & cmbStatus.Text & "' and " & TextoFiltroPadraoRel
        ElseIf cmbfiltrarpor = "Família" Then
                .StrSql_Fornecedor = INNERJOINTEXTO & " where CFF.Familia = '" & cmbfamilia & "' and CFF.tipo = 'F' and " & TextoFiltroPadrao
                .FormulaRel_Fornecedor = "{compras_fornecedores_familia.Familia} = '" & cmbfamilia & "' and {compras_fornecedores_familia.tipo} = 'F' and " & TextoFiltroPadraoRel
            ElseIf cmbfiltrarpor = "CNPJ/CPF" Then
                        .StrSql_Fornecedor = INNERJOINTEXTO & " where " & CPFCNPJ & " and " & TextoFiltroPadrao
                        .FormulaRel_Fornecedor = CPFCNPJRel & " and " & TextoFiltroPadraoRel
                    ElseIf cmbfiltrarpor = "Segmento" Then
                            .StrSql_Fornecedor = INNERJOINTEXTO & " where CFS.segmento = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                            .FormulaRel_Fornecedor = "{compras_fornecedores_segmentos.Segmento} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                        ElseIf cmbfiltrarpor = "Centro de custo" Then
                                .StrSql_Fornecedor = INNERJOINTEXTO & " where US.Setor = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                                .FormulaRel_Fornecedor = "{Usuarios_Setor.Setor} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
                            ElseIf cmbfiltrarpor = "Código do fornecedor" Then
                                    .StrSql_Fornecedor = INNERJOINTEXTO & " where CF.IDCliente = " & txtTexto & " and " & TextoFiltroPadrao
                                    .FormulaRel_Fornecedor = "{Compras_fornecedores.IDCliente} = " & txtTexto & " and " & TextoFiltroPadraoRel
                                    ElseIf cmbfiltrarpor = "Regime tributário" Then
                                        .StrSql_Fornecedor = INNERJOINTEXTO & " where " & TextoRegime & " and " & TextoFiltroPadrao
                                        .FormulaRel_Fornecedor = TextoRegimeRel & " And " & TextoFiltroPadraoRel
                                    Else
                                        Select Case cmbfiltrarpor
                                            Case "Razão social": TextoFiltro = "CF.nome_razao"
                                            Case "Nome fantasia": TextoFiltro = "CF.nomefantasia"
                                            Case "Cidade": TextoFiltro = "CF.cidade"
                                            Case "Centro de custo": TextoFiltro = "US.Setor"
                                        End Select
                                        If Left(TextoFiltro, 2) = "CF" Then TextoFiltroRel = Replace(TextoFiltro, "CF.", "Compras_fornecedores.") Else TextoFiltroRel = Replace(TextoFiltro, "US.", "Usuarios_Setor.")
                                        .StrSql_Fornecedor = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                                        .FormulaRel_Fornecedor = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
        End If
    Else
        .StrSql_Fornecedor = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        .FormulaRel_Fornecedor = TextoFiltroPadraoRel
    End If
    .ProcAtualizalista (1)
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10455, 5, True
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
cmbfiltrarpor = "Razão social"
msk_fltFim = Date
msk_fltInicio = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optFisica_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Regime tributário" Then ProcCarregaComboRegime

If optFisica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = False
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optJuridica_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Regime tributário" Then ProcCarregaComboRegime

If optJuridica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = True
    txtCpf.Visible = False
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcnpj_Change()
On Error GoTo tratar_erro
  
If txtcnpj.Text <> "__.___.___/____-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCpf_Change()
On Error GoTo tratar_erro
  
If txtCpf.Text <> "___.___.___-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboSegmento()
On Error GoTo tratar_erro

With cmbfamilia
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select ID, segmento from Segmentos where Tipo = 'C' order by segmento", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        .AddItem ""
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Segmento) = False And TBCarregarCombo!Segmento <> "" Then
                .AddItem TBCarregarCombo!Segmento
                .ItemData(.NewIndex) = TBCarregarCombo!ID
            End If
            TBCarregarCombo.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboRegime()
On Error GoTo tratar_erro

With cmbfamilia
    .Clear
    If optJuridica.Value = True Then
        .AddItem "Lucro presumido"
        .AddItem "Lucro real"
        .AddItem "Simples nacional"
    Else
        .AddItem "MEI"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
