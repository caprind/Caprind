VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCertificado_qualidade_analisequimica 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Ensaios - Certificado da qualidade - Cadastro de elementos químicos por produto"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   6120
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Text            =   "0"
      Top             =   3870
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3435
      Left            =   60
      TabIndex        =   10
      Top             =   2415
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Produto"
         Object.Width           =   17383
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1545
      Left            =   55
      TabIndex        =   11
      Top             =   840
      Width           =   10545
      Begin VB.Frame Frame7_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   7035
         TabIndex        =   22
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto7 
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
            MaxLength       =   50
            TabIndex        =   7
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame6_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   5895
         TabIndex        =   21
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto6 
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
            MaxLength       =   50
            TabIndex        =   6
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame5_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   4755
         TabIndex        =   20
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto5 
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
            MaxLength       =   50
            TabIndex        =   5
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame8_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   8175
         TabIndex        =   19
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto8 
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
            MaxLength       =   50
            TabIndex        =   8
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame4_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   3615
         TabIndex        =   18
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto4 
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
            MaxLength       =   50
            TabIndex        =   4
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame3_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   2460
         TabIndex        =   17
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto3 
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
            MaxLength       =   50
            TabIndex        =   3
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame9_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   9330
         TabIndex        =   16
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto9 
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
            MaxLength       =   50
            TabIndex        =   9
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame2_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   1320
         TabIndex        =   15
         Top             =   810
         Width           =   1035
         Begin VB.TextBox txtTexto2 
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
            MaxLength       =   50
            TabIndex        =   2
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame1_carcaca 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   180
         TabIndex        =   14
         Top             =   810
         Width           =   1035
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
            Left            =   195
            MaxLength       =   50
            TabIndex        =   1
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.TextBox txtProduto 
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
         Height          =   315
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Produto."
         Top             =   450
         Width           =   10155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto"
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
         Left            =   4972
         TabIndex        =   12
         Top             =   240
         Width           =   570
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   23
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
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
      ButtonCaption3  =   "Excluir"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Excluir (F4)"
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
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   39
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   118
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonUseMaskColor4=   0   'False
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
      ButtonLeft5     =   122
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
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
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
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5130
         Top             =   90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCertificado_qualidade_analisequimica.frx":0000
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   5850
      Width           =   10545
      _ExtentX        =   18600
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
Attribute VB_Name = "frmCertificado_qualidade_analisequimica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Analise As Boolean 'OK

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Certificado_Analise order by Produto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Produto), "", TBLISTA!Produto)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If txtId.Text = 0 Then
    USMsgBox ("Informe o produto na lista antes de excluir."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If Sair = True Then GoTo Pula
    If USMsgBox("Deseja realmente excluir os elementos químicos desse produto?", vbYesNo) = vbYes Then
Pula:
    'Verifica elemento no analise quimico

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Certificado_Quimica where id_elemento = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Não é permitido excluir os elementos químicos, pois eles já estão sendo utilizados no módulo da qualidade."), vbInformation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
    Conexao.Execute "DELETE from Certificado_Analise where id = " & txtId
    USMsgBox ("Elementos químicos excluídos com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/Ensaios/Controle de certificados/Elementos químicos"
    Evento = "Excluir"
    Documento = txtId.Text
    ProcGravaEvento
    '==================================
    ProcLimpaCampos
    ProcCarregaLista
    Frame1.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Novo_Analise = True Then Exit Sub
ProcLimpaCampos
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Certificado_Analise", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
Novo_Analise = True
Frame1.Enabled = True
txtProduto.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtProduto = ""
txtTexto = ""
txtTexto2 = ""
txtTexto3 = ""
txtTexto4 = ""
txtTexto5 = ""
txtTexto6 = ""
txtTexto7 = ""
txtTexto8 = ""
txtTexto9 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtProduto.Text = "" Then
    USMsgBox ("Informe o produto antes de salvar."), vbInformation, "CAPRIND v5.0"
    txtProduto.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Certificado_Analise where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
If Novo_Analise = True Then
    USMsgBox ("Novos elementos químicos cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
TBGravar!Produto = txtProduto
TBGravar!Texto = txtTexto
TBGravar!Texto2 = txtTexto2
TBGravar!Texto3 = txtTexto3
TBGravar!texto4 = txtTexto4
TBGravar!Texto5 = txtTexto5
TBGravar!Texto6 = txtTexto6
TBGravar!Texto7 = txtTexto7
TBGravar!Texto8 = txtTexto8
TBGravar!Texto9 = txtTexto9
TBGravar.Update
TBGravar.Close
'==================================
Modulo = "Qualidade/Ensaios/Controle de certificados/Elementos químicos"
Documento = txtId
ProcGravaEvento
'==================================
ProcCarregaLista
Novo_Analise = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: Ajuda
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
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

ProcCarregaToolBar1 Me, 10530, 7, True
ProcLimpaVariaveisPrincipais
ProcCarregaLista

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

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Analise = True Then
    If USMsgBox("Os elementos químicos ainda não foram salvos, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Analise = True Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        If txtId <> "" Then
            Sair = True
            ProcExcluir
        End If
    End If
End If
Conexao.Execute "DELETE from Certificado_Analise WHERE Produto is null"
Novo_Analise = False
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
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Certificado_Analise where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    With frmCertificado_qualidade
        .txtID_Elemento = TBLISTA!ID
        .txtProduto = TBLISTA!Produto
        .txtTexto_carcaca1 = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
        .txtTexto_carcaca2 = IIf(IsNull(TBLISTA!Texto2), "", TBLISTA!Texto2)
        .txtTexto_carcaca3 = IIf(IsNull(TBLISTA!Texto3), "", TBLISTA!Texto3)
        .txtTexto_carcaca4 = IIf(IsNull(TBLISTA!texto4), "", TBLISTA!texto4)
        .txtTexto_carcaca5 = IIf(IsNull(TBLISTA!Texto5), "", TBLISTA!Texto5)
        .txtTexto_carcaca6 = IIf(IsNull(TBLISTA!Texto6), "", TBLISTA!Texto6)
        .txtTexto_carcaca7 = IIf(IsNull(TBLISTA!Texto7), "", TBLISTA!Texto7)
        .txtTexto_carcaca8 = IIf(IsNull(TBLISTA!Texto8), "", TBLISTA!Texto8)
        .txtTexto_carcaca9 = IIf(IsNull(TBLISTA!Texto9), "", TBLISTA!Texto9)
    End With
End If
TBLISTA.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
txtId.Text = Lista.SelectedItem
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Certificado_Analise where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtProduto = TBLISTA!Produto
    txtTexto = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
    txtTexto2 = IIf(IsNull(TBLISTA!Texto2), "", TBLISTA!Texto2)
    txtTexto3 = IIf(IsNull(TBLISTA!Texto3), "", TBLISTA!Texto3)
    txtTexto4 = IIf(IsNull(TBLISTA!texto4), "", TBLISTA!texto4)
    txtTexto5 = IIf(IsNull(TBLISTA!Texto5), "", TBLISTA!Texto5)
    txtTexto6 = IIf(IsNull(TBLISTA!Texto6), "", TBLISTA!Texto6)
    txtTexto7 = IIf(IsNull(TBLISTA!Texto7), "", TBLISTA!Texto7)
    txtTexto8 = IIf(IsNull(TBLISTA!Texto8), "", TBLISTA!Texto8)
    txtTexto9 = IIf(IsNull(TBLISTA!Texto9), "", TBLISTA!Texto9)
End If
TBLISTA.Close
Frame1.Enabled = True
Novo_Analise = False

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
    Case 3: ProcExcluir
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

