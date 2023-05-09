VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmproj_produto_similar 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Produtos e serviços - Cadastro de produtos similares"
   ClientHeight    =   10035
   ClientLeft      =   1695
   ClientTop       =   1500
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   10
      Top             =   9750
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
      SearchText      =   ""
      Value           =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   1455
      Left            =   55
      TabIndex        =   6
      Top             =   990
      Width           =   15195
      Begin VB.TextBox Txt_ID_similar 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Height          =   315
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   14
         Text            =   "0"
         Top             =   390
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Cmd_localizar_prod 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14700
         Picture         =   "frmproj_produto_similar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Localizar produtos similares."
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox Txt_descricao_similar 
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Descrição similar."
         Top             =   960
         Width           =   12525
      End
      Begin VB.TextBox Txt_cod_interno_similar 
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
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código interno similar."
         Top             =   960
         Width           =   1965
      End
      Begin VB.TextBox Txt_cod_interno 
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
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   1965
      End
      Begin VB.TextBox Txt_descricao 
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   12855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição similar"
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
         Left            =   7837
         TabIndex        =   12
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. int. similar"
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
         Left            =   517
         TabIndex        =   11
         Top             =   750
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   547
         TabIndex        =   9
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   8242
         TabIndex        =   8
         Top             =   180
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   7290
      Left            =   60
      TabIndex        =   5
      Top             =   2460
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12859
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   23116
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   13
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   118
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
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7230
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmproj_produto_similar.frx":0102
         Count           =   1
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4680
      TabIndex        =   7
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "frmproj_produto_similar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Similar As Boolean 'OK

Private Sub Cmd_localizar_prod_Click()
On Error GoTo tratar_erro

Engenharia = True
frmQualidadePPAP_LocalizarProduto.Show 1

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
    Case vbKeyF4: ProcExcluir
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 7, True
ProcLimpaVariaveisPrincipais
ProcRemoveObjetosResize Me

If Engenharia_Produtos = True Then
    Caption = "Engenharia - Produtos e serviços - Cadastro de produtos similares"
    Formulario = "Engenharia/Produtos e serviços/Cadastro de produtos similares"
End If
If Compras_Produtos = True Then
    Caption = "Compras - Produtos e serviços - Cadastro de produtos similares"
    Formulario = "Compras/Produtos e serviços/Cadastro de produtos similares"
End If
If Vendas_Produtos = True Then
    Caption = "Vendas - Produtos e serviços - Cadastro de produtos similares"
    Formulario = "Vendas/Produtos e serviços/Cadastro de produtos similares"
End If

With frmproj_produto
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID_similar from Projproduto where Codproduto = " & .txtcodproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_ID_similar = IIf(IsNull(TBAbrir!ID_similar), 0, TBAbrir!ID_similar)
        Txt_cod_interno = .txtdesenhoproduto
        Txt_descricao = .txtDescricaoProduto
    End If
    TBAbrir.Close
End With

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Codproduto, Desenho, Descricao from Projproduto where ID_similar = " & Txt_ID_similar & " and Desenho <> '" & Txt_cod_interno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
       With Lista.ListItems
            .Add , , TBLISTA!Codproduto
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
Else
    Conexao.Execute "UPDATE Projproduto Set ID_similar = NULL where Desenho = '" & Txt_cod_interno & "'"
    Txt_ID_similar = FunVerifApagaIDSimilar(Txt_ID_similar)
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Engenharia_Produtos = True Then Formulario = "Engenharia/Produtos e serviços/Cadastro de produtos similares"
If Compras_Produtos = True Then Formulario = "Compras/Produtos e serviços/Cadastro de produtos similares"
If Vendas_Produtos = True Then Formulario = "Vendas/Produtos e serviços/Cadastro de produtos similares"
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s) similar(es)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "UPDATE Projproduto Set ID_similar = NULL where Codproduto = " & .ListItems(InitFor)
            '==================================
            Modulo = Formulario
            Evento = "Excluir"
            ID_documento = frmproj_produto.txtcodproduto
            Documento = "Cód. interno: " & Txt_cod_interno
            Documento1 = "Cód. interno similar: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produtos(s) similar(es) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) similar(es) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
    Novo_Similar = False
    Frame2.Enabled = False
    Txt_ID_similar = FunVerifApagaIDSimilar(Txt_ID_similar)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_cod_interno_similar = ""
Txt_descricao_similar = ""
CodigoLista = 0

If Engenharia_Produtos = True Then TextoForm = "Engenharia"
If Compras_Produtos = True Then Formulario = "Compras"
If Vendas_Produtos = True Then Formulario = "Vendas"
Caption = TextoForm & " - Produtos e serviços - Cadastro de produtos similares"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_cod_interno_similar = "" Then
    NomeCampo = "o código interno do produto similar"
    ProcVerificaAcao
    Cmd_localizar_prod_Click
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select Codproduto from Projproduto where Desenho = '" & Txt_cod_interno_similar & "' and ID_similar = " & Txt_ID_similar, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido salvar este produto " & Txt_cod_interno_similar & ", pois o mesmo já foi cadastrado como similar ao produto " & Txt_cod_interno & "."), vbExclamation, "CAPRIND v5.0"
    TBGravar.Close
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID_similar from Projproduto where Desenho = '" & Txt_cod_interno_similar & "' and ID_similar IS NOT NULL and ID_similar <> 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Projproduto_similar where ID = " & Txt_ID_similar, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        TBGravar.Update
        Txt_ID_similar = TBGravar!ID
    End If
    TBGravar.Close
Else
    Txt_ID_similar = TBAbrir!ID_similar
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select ID_similar from Projproduto where Desenho = '" & Txt_cod_interno & "' and ID_similar IS NOT NULL and ID_similar <> 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Conexao.Execute "UPDATE Projproduto Set ID_similar = " & Txt_ID_similar & " where ID_similar = " & TBFI!ID_similar
        FunVerifApagaIDSimilar TBFI!ID_similar
    End If
    TBFI.Close
End If
TBAbrir.Close
Conexao.Execute "UPDATE Projproduto Set ID_similar = " & Txt_ID_similar & " where Desenho = '" & Txt_cod_interno & "'"
Conexao.Execute "UPDATE Projproduto Set ID_similar = " & Txt_ID_similar & " where Desenho = '" & Txt_cod_interno_similar & "'"

If Novo_Similar = True Then
    USMsgBox ("Novo produto similar cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
ProcCarregaLista
If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If
'==================================
Modulo = Formulario
ID_documento = frmproj_produto.txtcodproduto
Documento = "Cód. interno: " & Txt_cod_interno
Documento1 = "Cód. interno similar: " & Txt_cod_interno_similar
ProcGravaEvento
'==================================
Novo_Similar = False
Txt_ID_similar = FunVerifApagaIDSimilar(Txt_ID_similar)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_Similar = True
Frame2.Enabled = True
Cmd_localizar_prod_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Similar = True Then
    If USMsgBox("O produto similar ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Similar = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Similar = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Desenho, Descricao from Projproduto where Codproduto = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

Txt_cod_interno_similar = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
Txt_descricao_similar = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
Novo_Similar = False
Frame2.Enabled = True

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
