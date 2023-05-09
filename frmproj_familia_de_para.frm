VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmproj_familia_de_para 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Famílias - De, para"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6315
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   2760
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1425
      Left            =   55
      TabIndex        =   3
      Top             =   1020
      Width           =   6225
      Begin VB.ComboBox Cmb_para 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Para."
         Top             =   960
         Width           =   5865
      End
      Begin VB.ComboBox Cmb_de 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "De."
         Top             =   390
         Width           =   5865
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para"
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
         Left            =   2947
         TabIndex        =   5
         Top             =   750
         Width           =   330
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
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
         Left            =   3015
         TabIndex        =   4
         Top             =   180
         Width           =   195
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3150
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmproj_familia_de_para.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
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
      ButtonKey3      =   "6"
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
      ButtonKey4      =   "7"
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
      ButtonKey5      =   "8"
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   6
      Top             =   2460
      Width           =   6225
      _ExtentX        =   10980
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
End
Attribute VB_Name = "frmproj_familia_de_para"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcGravar()
On Error GoTo tratar_erro

Acao = "salvar"
If Cmb_de = "" Then
    NomeCampo = "a família de"
    ProcVerificaAcao
    Cmb_de.SetFocus
    Exit Sub
End If
If Cmb_para = "" Then
    NomeCampo = "a família para"
    ProcVerificaAcao
    Cmb_de.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where Classe = '" & Cmb_de & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Conexao.Execute "Update CFI Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update Compras_fornecedores_familia Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update Compras_pedido_lista Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update Estoque_Controle Set Classe = '" & Cmb_para & "' where Classe = '" & Cmb_de & "'"
    Conexao.Execute "Update Estoque_movimentacao Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update Instrumentos Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update Requisicao_materiais_lista Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update tbl_Detalhes_Nota Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update Vendas_analise Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    Conexao.Execute "Update Vendas_analise_setores Set Familia = '" & Cmb_para & "' where Familia = '" & Cmb_de & "'"
    
    TBProduto.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBProduto.RecordCount
    PBLista.Value = 1
    contador = 0
    TBProduto.MoveFirst
    Do While TBProduto.EOF = False
        qt = Len(TBProduto!Desenho)
        If qt > 6 Then
            Set TBFamilia = CreateObject("adodb.recordset")
            TBFamilia.Open "Select * from projfamilia where familia = '" & Cmb_para & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFamilia.EOF = False Then Letra = TBFamilia!Letra
            TBFamilia.Close
            
            CompLetra = Len(Letra)
            Letra1 = ""
            If TBProduto!SubTipoItem = 0 Or TBProduto!SubTipoItem = 1 Or TBProduto!SubTipoItem = 4 Or TBProduto!SubTipoItem = 5 Then
                If Left(TBProduto!Desenho, CompLetra) = Letra Then Letra1 = Left(TBProduto!Desenho, qt - 6) Else Letra1 = Right(TBProduto!Desenho, qt - 6)
            Else
                Letra1 = Right(TBProduto!Desenho, qt - 6)
            End If
            'Verifica se o código do produto está vinculado a família selecionada
            Set TBFamilia = CreateObject("adodb.recordset")
            TBFamilia.Open "Select * from projfamilia where Familia = '" & Cmb_para & "' and Letra = '" & Letra1 & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFamilia.EOF = False Then
                TBProduto!CodManual = False
            Else
                TBProduto!CodManual = True
            End If
            TBFamilia.Close
        Else
            TBProduto!CodManual = True
        End If
        TBProduto!Classe = Cmb_para
        TBProduto.Update
        
        TBProduto.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBProduto.Close

USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = Formulario
Evento = "Alterar De, para"
ID_documento = 0
Documento = "De: " & Cmb_de & " - Para: " & Cmb_para
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_De_Click()
On Error GoTo tratar_erro

If Cmb_de <> "" Then ProcCarregaComboPara
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcGravar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6255, 5, True
If Compras_Familia = True Then Caption = "Compras - Famílias - De, para"
If Vendas_Familia = True Then Caption = "Vendas - Famílias - De, para"
If Qualidade_Familia = True Then Caption = "Qualidade - Famílias - De, para"
ProcCarregaComboDe
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboDe()
On Error GoTo tratar_erro

With Cmb_de
    .Clear
    TextoFiltro = ""
    If Compras_Familia = True Then
        Caption = "Compras - Famílias - De, para"
        TextoFiltro = "where compras = 'True' or qualidade = 'True'"
    ElseIf Vendas_Familia = True Then
            Caption = "Vendas - Famílias - De, para"
            TextoFiltro = "where Vendas = 'True'"
        ElseIf Qualidade_Familia = True Then
                Caption = "Qualidade - Famílias - De, para"
                TextoFiltro = "where compras = 'True' or qualidade = 'True'"
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Familia from Projfamilia " & TextoFiltro & " group by Familia", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Familia
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboPara()
On Error GoTo tratar_erro

With Cmb_para
    .Clear
    TextoFiltro = ""
    If Compras_Familia = True Then
        Caption = "Compras - Famílias - De, para"
        TextoFiltro = "and (compras = 'True' or qualidade = 'True')"
    ElseIf Vendas_Familia = True Then
            Caption = "Vendas - Famílias - De, para"
            TextoFiltro = "and Vendas = 'True'"
        ElseIf Qualidade_Familia = True Then
                Caption = "Qualidade - Famílias - De, para"
                TextoFiltro = "and (compras = 'True' or qualidade = 'True')"
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Familia from Projfamilia where Familia <> '" & cmd_de & "' " & TextoFiltro & " group by Familia", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Familia
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With

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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcGravar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
