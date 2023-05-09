VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmOpcoesGeral_ConfRelatorio 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações do sistema - Opções gerais - Configurar dados da empresa nos relatórios"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10470
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Framedetalhes 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   2205
      Left            =   60
      TabIndex        =   8
      Top             =   990
      Width           =   10335
      Begin VB.PictureBox Cor_fonte 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   6570
         ScaleHeight     =   255
         ScaleWidth      =   795
         TabIndex        =   9
         Top             =   1635
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   6300
         TabIndex        =   10
         Top             =   270
         Width           =   3825
         Begin VB.CheckBox Chk_negrito 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Negrito"
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
            Left            =   90
            TabIndex        =   1
            Top             =   210
            Width           =   915
         End
         Begin VB.CheckBox Chk_italico 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Itálico"
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
            Left            =   1290
            TabIndex        =   2
            Top             =   210
            Width           =   915
         End
         Begin VB.CheckBox Chk_sublinhado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sublinhado"
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
            Left            =   2490
            TabIndex        =   3
            Top             =   210
            Width           =   1245
         End
      End
      Begin VB.ComboBox Cmb_fonte 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6300
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Fonte."
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox Cmb_tamanho_fonte 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmOpcoesGeral_ConfRelatorio.frx":0000
         Left            =   9210
         List            =   "frmOpcoesGeral_ConfRelatorio.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Tamanho."
         Top             =   1080
         Width           =   915
      End
      Begin VB.CommandButton Cmd_cor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Mudar cor das instruções de trabalho."
         Top             =   1530
         Width           =   3855
      End
      Begin RichTextLib.RichTextBox txtDescricao 
         Height          =   1665
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Dados da empresa nos relatórios."
         Top             =   390
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2937
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"frmOpcoesGeral_ConfRelatorio.frx":0004
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dados da empresa nos relatórios"
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
         Left            =   1935
         TabIndex        =   13
         Top             =   180
         Width           =   2355
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fonte"
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
         Index           =   5
         Left            =   7507
         TabIndex        =   12
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho"
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
         Index           =   6
         Left            =   9270
         TabIndex        =   11
         Top             =   870
         Width           =   795
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7350
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmOpcoesGeral_ConfRelatorio.frx":0082
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   33
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   38
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
      ButtonLeft3     =   77
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonUseMaskColor3=   0   'False
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
      ButtonLeft4     =   81
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
      ButtonLeft5     =   119
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
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
      ButtonLeft6     =   147
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
End
Attribute VB_Name = "frmOpcoesGeral_ConfRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_empresa_rel As Boolean

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Framedetalhes.Enabled = False Then
   ProcVerificaSalvar
   Exit Sub
End If
Acao = "salvar"
If txtdescricao = "" Then
    NomeCampo = "os dados da empresa nos relatórios"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If

With frmOpcoesGeral
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Dados_relatorio from empresa where codigo = " & .txtIDEmpresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir!Dados_relatorio = txtdescricao.TextRTF
        TBAbrir.Update
    End If
    TBAbrir.Close

    If Novo_empresa_rel = True Then
        USMsgBox ("Novos dados da empresa nos relatórios."), vbInformation, "CAPRIND v5.0"
        Evento = "Novos dados da empresa nos relatórios"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar dados da empresa nos relatórios"
    End If
    '==================================
    Modulo = "Configuração do sistema/Opções gerais"
    ID_documento = .txtIDEmpresa
    Documento = "Empresa: " & .txtEmpresa
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_empresa_rel = False
End With

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

With frmOpcoesGeral
    TextoFiltro = ""
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Empresa, Endereco, Bairro, CEP, Cidade, UF, Telefone, Fax, Email, Site, CNPJ, Ie from empresa where codigo = " & .txtIDEmpresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TextoFiltro = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
        TextoFiltro = TextoFiltro & vbCrLf & IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco) & " - " & IIf(IsNull(TBAbrir!Bairro), "", TBAbrir!Bairro) & " - " & IIf(IsNull(TBAbrir!CEP), "", TBAbrir!CEP) & " - " & IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade) & " - " & IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF)
        TextoFiltro = TextoFiltro & vbCrLf & "Tel: " & IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone) & " - Fax: " & IIf(IsNull(TBAbrir!Fax), "", TBAbrir!Fax) & " - E-mail: " & IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
        TextoFiltro = TextoFiltro & vbCrLf & "Site: " & IIf(IsNull(TBAbrir!Site), "", TBAbrir!Site) & " - CNPJ: " & IIf(IsNull(TBAbrir!CNPJ), "", TBAbrir!CNPJ) & " - IE: " & IIf(IsNull(TBAbrir!IE), "", TBAbrir!IE)
    End If
    TBAbrir.Close
    txtdescricao = TextoFiltro
End With
Novo_empresa_rel = True
Framedetalhes.Enabled = True
txtdescricao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_italico_Click()
On Error GoTo tratar_erro

txtdescricao.SelItalic = Chk_italico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_negrito_Click()
On Error GoTo tratar_erro

txtdescricao.SelBold = Chk_negrito

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_sublinhado_Click()
On Error GoTo tratar_erro

txtdescricao.SelUnderline = Chk_sublinhado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_fonte_Click()
On Error GoTo tratar_erro

txtdescricao.SelFontName = Cmb_fonte

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tamanho_fonte_Click()
On Error GoTo tratar_erro

txtdescricao.SelFontSize = Cmb_tamanho_fonte

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_cor_Click()
On Error GoTo tratar_erro

With CommonDialog1
    .Color = Cor_fonte.BackColor
    .ShowColor
End With
Cor_fonte.BackColor = CommonDialog1.Color
txtdescricao.SelColor = CommonDialog1.Color

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10335, 6, True
ProcCarregaComboFontes Cmb_fonte
ProcCarregaComboTamanhoFonte Cmb_tamanho_fonte, 7, 16

Novo_empresa_rel = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Dados_relatorio from empresa where codigo = " & frmOpcoesGeral.txtIDEmpresa & " and Dados_relatorio IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtdescricao.TextRTF = TBAbrir!Dados_relatorio
    Framedetalhes.Enabled = True
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDescricao_Change()
On Error GoTo tratar_erro

With txtdescricao
    Cor_fonte.BackColor = IIf(IsNull(.SelColor), Cor_fonte.BackColor, .SelColor)
    Chk_negrito.Value = IIf(IsNull(.SelBold), 2, Abs(.SelBold))
    Chk_italico.Value = IIf(IsNull(.SelItalic), 2, Abs(.SelItalic))
    Chk_sublinhado.Value = IIf(IsNull(.SelUnderline), 2, Abs(.SelUnderline))
    Cmb_fonte = IIf(IsNull(.SelFontName), "", .SelFontName)
    Cmb_tamanho_fonte = IIf(IsNull(.SelFontSize), "", .SelFontSize)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
