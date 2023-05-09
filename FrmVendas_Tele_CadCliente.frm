VERSION 5.00
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frmVendas_Tele_CadCliente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Vendas  Telemarketing - Cadastrar prospecto"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10635
   ClipControls    =   0   'False
   Icon            =   "FrmVendas_Tele_CadCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Digite os dados do prospecto abaixo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   120
      TabIndex        =   4
      Top             =   1275
      Width           =   10380
      Begin VB.ComboBox cmbCidade 
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
         ItemData        =   "FrmVendas_Tele_CadCliente.frx":000C
         Left            =   1035
         List            =   "FrmVendas_Tele_CadCliente.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Cidade."
         Top             =   1170
         Width           =   4860
      End
      Begin VB.TextBox txtnomerazao 
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
         Left            =   195
         TabIndex        =   0
         ToolTipText     =   "Razão social."
         Top             =   570
         Width           =   9990
      End
      Begin VB.TextBox txtidcliente 
         Alignment       =   2  'Centralizar
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
         ToolTipText     =   "Código do cliente."
         Top             =   570
         Width           =   915
      End
      Begin VB.TextBox txttel01 
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
         Left            =   5910
         TabIndex        =   2
         ToolTipText     =   "Telefone."
         Top             =   1170
         Width           =   1420
      End
      Begin VB.CommandButton imgemail 
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
         Height          =   315
         Left            =   14685
         Picture         =   "FrmVendas_Tele_CadCliente.frx":0010
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Enviar e-mail para o cliente (F8)"
         Top             =   990
         Width           =   315
      End
      Begin VB.ComboBox cmbUF 
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
         ItemData        =   "FrmVendas_Tele_CadCliente.frx":041A
         Left            =   165
         List            =   "FrmVendas_Tele_CadCliente.frx":041C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "UF."
         Top             =   1170
         Width           =   855
      End
      Begin ControlesUteis.txt txtEmail 
         Height          =   360
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "E-mail."
         Top             =   1860
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   635
         Tamanho         =   10005
         Tipo            =   2
         Text            =   ""
         FocusColor      =   16777215
         ShowCaption     =   0   'False
         Caption         =   ""
         MaxLength       =   255
         BackColor       =   14737632
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
      Begin ControlesUteis.txt txtSite 
         Height          =   360
         Left            =   7380
         TabIndex        =   3
         ToolTipText     =   "E-mail."
         Top             =   1170
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   635
         Tamanho         =   2805
         Tipo            =   2
         Text            =   ""
         FocusColor      =   16777215
         ShowCaption     =   0   'False
         Caption         =   ""
         MaxLength       =   255
         BackColor       =   14737632
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparente
         Caption         =   "Cidade"
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
         Index           =   3
         Left            =   3225
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Site"
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
         Index           =   3
         Left            =   8640
         TabIndex        =   12
         Top             =   945
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparente
         Caption         =   "UF"
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
         Index           =   5
         Left            =   450
         TabIndex        =   11
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Nome Razão social"
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
         Left            =   4530
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Telefone"
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
         Index           =   1
         Left            =   6300
         TabIndex        =   9
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "E-mail"
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
         Index           =   2
         Left            =   4920
         TabIndex        =   8
         Top             =   1635
         Width           =   420
      End
   End
   Begin DrawSuite2014.USImageList USImageList1 
      Left            =   4505
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "FrmVendas_Tele_CadCliente.frx":041E
      Count           =   1
   End
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10620
      _ExtentX        =   18733
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   38
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
      ButtonLeft2     =   42
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   46
      ButtonTop3      =   2
      ButtonWidth3    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
End
Attribute VB_Name = "frmVendas_Tele_CadCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function FormataTelefone(ByVal text As String) As String
Dim i As Long

' ignora vazio
If Len(text) = 0 Then Exit Function

 'verifica valores invalidos
  For i = Len(text) To 1 Step -1
    If InStr("0123456789", Mid$(text, i, 1)) = 0 Then
       text = Left$(text, i - 1) & Mid$(text, i + 1)
    End If
  Next
  ' ajusta a posicao correta
  If Len(text) <= 7 Then
     FormataTelefone = Format$(text, "!@@@-@@@@")
  ElseIf Len(text) > 7 And Len(text) <= 9 Then
     FormataTelefone = Format$(text, "!(@@) @@@-@@@@")
  ElseIf Len(text) > 9 Then
     FormataTelefone = Format$(text, "!(@@) @@@@-@@@@")
  End If
End Function

Private Sub ProcSalvar()
On Error GoTo tratar_erro
Novo = True
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from clientes ", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IDCliente from clientes order by idcliente desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtidcliente.text = TBAbrir!IDCliente + 1
    Else
        txtidcliente.text = 1
    End If
    TBAbrir.Close
    TBGravar!IDCliente = txtidcliente.text
    TBGravar!Status = "Liberado"
    TBGravar!nomefantasia = IIf(txtnomefantasia = "", Null, txtnomefantasia)
    TBGravar!IDCliente = txtidcliente.text
    TBGravar!Data = IIf(txtData = "", Date, txtData)
    TBGravar!Responsavel = pubUsuario
    TBGravar!NomeRazao = Replace(txtnomerazao, "'", " ")
    TBGravar!UF = cmbUF.text
    TBGravar!Cidade = cmbCidade.text
    TBGravar!Tel01 = txttel01.text
    TBGravar!Email = IIf(txtEmail.text = "", Null, LCase(txtEmail.text))
    TBGravar!Prospecto = True
    TBGravar!DtValidacao = Date
    TBGravar!RespValidacao = pubUsuario
    TBGravar!Site = txtSite.text
    TBGravar!Tipo = "JP"
    TBGravar.Update
    TBGravar.Close
Novo = False

MsgBox ("Cliente cadastrado com sucesso!"), vbInformation
frmVendas_Tele_Clientes.ProcCarregaLista
Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub cmbUF_LostFocus()
On Error GoTo tratar_erro

If cmbUF.text <> "EX" Then
        ProcCarregaComboCidade cmbCidade, "Sigla_UF = '" & cmbUF & "'", False
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10715, 5, True
ProcCarregaComboUF cmbUF, "UF is not null", "Nacional"
cmbUF.text = "SP"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txttel01_Validate(Cancel As Boolean)
On Error GoTo tratar_erro

If Not IsNumeric(txttel01.text) Or Len(txttel01.text) < 4 Then
   keepfocus = True
   MsgBox "Informe um valor valido !", vbInformation, "Formatando telefone"
   Exit Sub
End If
txttel01.text = FormataTelefone(txttel01.text)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

