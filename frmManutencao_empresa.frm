VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmManutencao_empresa 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manutenção - Equipamentos - Empresa"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3300
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmManutencao_empresa.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   3
      Top             =   0
      Width           =   4185
      _ExtentX        =   7382
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
      Height          =   855
      Left            =   55
      TabIndex        =   0
      Top             =   990
      Width           =   4185
      Begin VB.ComboBox Cmb_empresa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   3825
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   1770
         TabIndex        =   2
         Top             =   180
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmManutencao_empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcGravar()
On Error GoTo tratar_erro

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from compras_requisicao", Conexao, adOpenKeyset, adLockOptimistic
TBCompras.AddNew
ProcCriarNovoNumero
TBCompras!Requisicaotexto = a
TBCompras!solicitado = pubUsuario
TBCompras!setorsolic = pubSetor
TBCompras!Data_Solicitacao = Date
TBCompras!status = "ABERTA"
TBCompras!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBCompras.Update
Cont_solicitacao = TBCompras!ID_Requisicao
TBCompras.Close

Conexao.Execute "Update Manutencao_data Set Solicitacao = '" & a & "' where ID = " & frmManutencao.txtIDData
frmManutencao.txtSolicitacao = a

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from Manutencao_defeito where idManutencao = " & frmManutencao.txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBCompras_Pedido = CreateObject("adodb.recordset")
        TBCompras_Pedido.Open "Select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
        TBCompras_Pedido.AddNew
        TBCompras_Pedido!ID_Requisicao = Cont_solicitacao
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from Compras_pedido_lista WHERE ID_Requisicao = " & Cont_solicitacao & " order by codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            TBFIltro.MoveLast
            TBCompras_Pedido!CODIGO = TBFIltro!CODIGO + 1
        Else
            TBCompras_Pedido!CODIGO = 1
        End If
        TBCompras_Pedido!Status_Item = "REQUISIT."
        TBCompras_Pedido!Un = TBAbrir!Unidade
        TBCompras_Pedido!Familia = TBAbrir!Familia
        TBCompras_Pedido!solicitado = pubUsuario
        TBCompras_Pedido!setorsolic = pubSetor
        TBCompras_Pedido!Descricao = TBAbrir!Descricao
        TBCompras_Pedido!quant_req = TBAbrir!Qtde
        TBCompras_Pedido!Desenho = TBAbrir!Desenho
        TBCompras_Pedido.Update
        TBCompras_Pedido.Close
        TBAbrir.MoveNext
    Loop
End If

Unload Me

With frmCompras_Requisicao
    .ProcLimpaCampos
    .txtNumero.Text = a
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from compras_requisicao where ID_Requisicao = " & Cont_solicitacao, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = False Then
        ProcCarregaComboEmpresa .Cmb_empresa, False
        If .PBLista.Value = 0 Then .PBLista = 100
        .ProcAbrir
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriarNovoNumero()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_requisicao where Year (Data_Solicitacao) = '" & Year(Date) & "' order by ID_Requisicao", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    Numero = Left(TBAbrir!Requisicaotexto, Len(TBAbrir!Requisicaotexto) - 3)
    Cont_solicitacao = Right(Numero, 5) + 1
Else
    Cont_solicitacao = 1
End If
TBAbrir.Close

a = Cont_solicitacao
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: a = "SOL-0000" & Cont_solicitacao & "/" & Ano
    Case 2: a = "SOL-000" & Cont_solicitacao & "/" & Ano
    Case 3: a = "SOL-00" & Cont_solicitacao & "/" & Ano
    Case 4: a = "SOL-0" & Cont_solicitacao & "/" & Ano
    Case 5: a = "SOL-" & Cont_solicitacao & "/" & Ano
End Select

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
 
ProcCarregaComboEmpresa Cmb_empresa, False
 
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
